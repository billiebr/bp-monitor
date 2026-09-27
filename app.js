(function(){
  const $ = id => document.getElementById(id);
  const sys = $('sys'), dia = $('dia'), pulse = $('pulse'), notes = $('notes');
  const sysDisp = $('sysDisp'), diaDisp = $('diaDisp');
  const saveBtn = $('saveBtn'), status = $('status'), pendingEl = $('pending');
  const photoCamera = $('photoCamera'), photoGallery = $('photoGallery'), photoPreview = $('photoPreview');
  const reading = $('reading');
  const scriptUrlInput = $('scriptUrl'), scriptTokenInput = $('scriptToken');
  const saveUrlBtn = $('saveUrlBtn'), configBlock = $('configBlock');
  const historyList = $('historyList');

  const STORAGE_URL = 'bp_tracker_script_url';
  const STORAGE_TOKEN = 'bp_tracker_token';
  const STORAGE_HISTORY = 'bp_tracker_history';
  const STORAGE_QUEUE = 'bp_tracker_queue';

  function load(key, fallback){ try { const v = localStorage.getItem(key); return v == null ? fallback : JSON.parse(v); } catch(e){ return fallback; } }
  function store(key, v){ try { localStorage.setItem(key, JSON.stringify(v)); } catch(e){} }
  function getUrl(){ try { return localStorage.getItem(STORAGE_URL) || ''; } catch(e){ return ''; } }
  function getToken(){ try { return localStorage.getItem(STORAGE_TOKEN) || ''; } catch(e){ return ''; } }
  function getHistory(){ return load(STORAGE_HISTORY, []); }
  function setHistory(list){ store(STORAGE_HISTORY, list.slice(0,20)); }
  let queue = load(STORAGE_QUEUE, []);   // readings not yet confirmed by the sheet

  function setStatus(text, cls){
    status.textContent = text;
    status.className = 'status' + (cls ? ' ' + cls : '');
  }

  function renderHistory(){
    const list = getHistory();
    if(!list.length){ historyList.innerHTML = '<div class="empty">אין עדיין מדידות</div>'; return; }
    historyList.innerHTML = list.map((e, i) => `
      <div class="entry">
        <div>
          <div class="val">${escapeHtml(e.sys)}/${escapeHtml(e.dia)}${e.pulse ? ` <span class="when">· דופק ${escapeHtml(e.pulse)}</span>` : ''}</div>
          ${e.notes ? `<div class="note">${escapeHtml(e.notes)}</div>` : ''}
        </div>
        <div class="side">
          <div class="when">${escapeHtml(e.when)}${queue.some(q => q.id === e.id) ? ' ⏳' : ''}</div>
          <button class="del" type="button" data-index="${i}" aria-label="מחיקת המדידה" title="מחיקה">🗑</button>
        </div>
      </div>
    `).join('');
  }

  // Delete a reading: from the sheet (by id), from the send queue, and from this list.
  historyList.addEventListener('click', async (ev) => {
    const btn = ev.target.closest('.del');
    if(!btn) return;
    const list = getHistory();
    const idx = Number(btn.dataset.index);
    const e = list[idx];
    if(!e) return;
    if(!confirm(`למחוק את המדידה ${e.sys}/${e.dia} (${e.when})?\nהיא תימחק גם מהגיליון.`)) return;

    btn.disabled = true;
    const queued = e.id && queue.some(q => q.id === e.id);
    if(e.id && !queued){
      setStatus('מוחק...');
      try{
        await api({ action: 'delete', id: e.id });
      }catch(err){
        btn.disabled = false;
        setStatus('המחיקה נכשלה: ' + err.message, 'err');
        return;
      }
    }
    if(queued){ queue = queue.filter(q => q.id !== e.id); store(STORAGE_QUEUE, queue); renderPending(); }
    setHistory(list.filter((x, j) => j !== idx));
    renderHistory();
    setStatus(e.id ? 'המדידה נמחקה ✓' : 'נמחקה מהרשימה — מדידה ישנה, יש למחוק אותה ידנית גם בגיליון', e.id ? 'ok' : 'err');
  });

  function renderPending(){
    pendingEl.textContent = queue.length ? `${queue.length} מדידות ממתינות לשליחה — יישלחו כשיהיה חיבור` : '';
  }

  function escapeHtml(s){
    return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  }

  function updateDisplay(){
    sysDisp.textContent = sys.value || '--';
    diaDisp.textContent = dia.value || '--';
  }
  sys.addEventListener('input', updateDisplay);
  dia.addEventListener('input', updateDisplay);

  // Talks to the Apps Script web app. text/plain avoids a CORS preflight,
  // and the JSON reply tells us whether the sheet really accepted it.
  async function api(payload){
    const url = getUrl(), token = getToken();
    if(!url || !token) throw new Error('קודם צריך להגדיר חיבור לגיליון (למטה)');
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(Object.assign({ token }, payload))
    });
    const text = await res.text();
    let data;
    try { data = JSON.parse(text); }
    catch(e){ throw new Error('הקוד שרץ בגיליון הוא לא הגרסה של האפליקציה הזו — צריך להדביק את Code.gs מחדש ולפרסם גרסה חדשה (New version)'); }
    if(!data.ok) throw new Error(data.error || 'שגיאה');
    return data;
  }

  // Shrink the photo before upload — faster and still sharp enough to read.
  function resizeImage(file, max){
    return new Promise((resolve, reject) => {
      const url = URL.createObjectURL(file);
      const img = new Image();
      img.onload = () => {
        const scale = Math.min(1, max / Math.max(img.width, img.height));
        const c = document.createElement('canvas');
        c.width = Math.round(img.width * scale);
        c.height = Math.round(img.height * scale);
        c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
        URL.revokeObjectURL(url);
        resolve(c.toDataURL('image/jpeg', 0.9));
      };
      img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('לא הצלחתי לפתוח את התמונה')); };
      img.src = url;
    });
  }


  // ---- full-screen photo viewer with the three fields underneath ----
  const viewer = $('viewer'), viewerImg = $('viewerImg'), viewerImgWrap = $('viewerImgWrap');
  const vSys = $('vSys'), vDia = $('vDia'), vPulse = $('vPulse');
  const pairs = [[vSys, sys], [vDia, dia], [vPulse, pulse]];

  function openViewer(focusEmpty){
    if(!photoPreview.src || photoPreview.style.display === 'none') return;
    viewerImg.src = photoPreview.src;
    viewerImgWrap.classList.remove('zoomed');
    pairs.forEach(([v, main]) => { v.value = main.value; });
    viewer.classList.add('open');
    document.body.style.overflow = 'hidden';
    if(focusEmpty){
      const empty = pairs.find(([v]) => !v.value);
      if(empty) empty[0].focus();
    }
  }
  function closeViewer(){
    viewer.classList.remove('open');
    document.body.style.overflow = '';
  }
  pairs.forEach(([v, main]) => v.addEventListener('input', () => { main.value = v.value; updateDisplay(); }));
  photoPreview.addEventListener('click', () => openViewer(false));
  viewerImg.addEventListener('click', (e) => {
    const zoomed = viewerImgWrap.classList.toggle('zoomed');
    if(zoomed){
      // keep the tapped spot under the finger
      const r = viewerImg.getBoundingClientRect();
      const fx = (e.clientX - r.left) / r.width, fy = (e.clientY - r.top) / r.height;
      requestAnimationFrame(() => {
        viewerImgWrap.scrollLeft = fx * viewerImg.clientWidth - viewerImgWrap.clientWidth / 2;
        viewerImgWrap.scrollTop = fy * viewerImg.clientHeight - viewerImgWrap.clientHeight / 2;
      });
    }
  });
  $('viewerClose').addEventListener('click', closeViewer);
  $('viewerDone').addEventListener('click', closeViewer);
  document.addEventListener('keydown', (e) => { if(e.key === 'Escape') closeViewer(); });


  // ---- 5-minute rest timer before measuring ----
  const REST_MS = 5 * 60 * 1000;
  const STORAGE_REST = 'bp_tracker_rest_end';
  const rest = $('rest'), restTime = $('restTime'), restProg = $('restProg');
  const restCancel = $('restCancel'), restClose = $('restClose'), restTitle = $('restTitle');
  const RING = 339.3;
  let restTick = null, wakeLock = null;

  async function keepAwake(on){
    try{
      if(on && 'wakeLock' in navigator && !wakeLock){ wakeLock = await navigator.wakeLock.request('screen'); }
      if(!on && wakeLock){ await wakeLock.release(); wakeLock = null; }
    }catch(e){ wakeLock = null; }
  }

  function beep(){
    try{
      const ctx = new (window.AudioContext || window.webkitAudioContext)();
      [0, 0.35, 0.7].forEach(t => {
        const o = ctx.createOscillator(), g = ctx.createGain();
        o.frequency.value = 880; o.connect(g); g.connect(ctx.destination);
        g.gain.setValueAtTime(0.0001, ctx.currentTime + t);
        g.gain.exponentialRampToValueAtTime(0.3, ctx.currentTime + t + 0.02);
        g.gain.exponentialRampToValueAtTime(0.0001, ctx.currentTime + t + 0.25);
        o.start(ctx.currentTime + t); o.stop(ctx.currentTime + t + 0.3);
      });
    }catch(e){}
    try{ navigator.vibrate && navigator.vibrate([300, 150, 300, 150, 300]); }catch(e){}
  }

  function renderRest(){
    const end = Number(load(STORAGE_REST, 0));
    const left = Math.max(0, end - Date.now());
    const secs = Math.ceil(left / 1000);
    restTime.textContent = Math.floor(secs / 60) + ':' + String(secs % 60).padStart(2, '0');
    restProg.style.strokeDashoffset = String(RING * (1 - left / REST_MS));
    if(left <= 0) finishRest();
  }

  function startRest(){
    store(STORAGE_REST, Date.now() + REST_MS);
    openRest();
  }

  function openRest(){
    rest.classList.remove('done');
    restTitle.textContent = 'מנוחה לפני מדידה';
    restCancel.hidden = false; restClose.hidden = true;
    restProg.style.transition = '';
    rest.classList.add('open');
    document.body.style.overflow = 'hidden';
    keepAwake(true);
    clearInterval(restTick);
    renderRest();
    restTick = setInterval(renderRest, 250);
  }

  function finishRest(){
    clearInterval(restTick); restTick = null;
    try{ localStorage.removeItem(STORAGE_REST); }catch(e){}
    restTime.textContent = '✓';
    restProg.style.transition = 'none';
    restProg.style.strokeDashoffset = '0';   // full green ring = done
    rest.classList.add('done');
    restTitle.textContent = 'המנוחה הסתיימה';
    restCancel.hidden = true; restClose.hidden = false;
    keepAwake(false);
    beep();
  }

  function closeRest(){
    clearInterval(restTick); restTick = null;
    try{ localStorage.removeItem(STORAGE_REST); }catch(e){}
    rest.classList.remove('open', 'done');
    document.body.style.overflow = '';
    keepAwake(false);
  }

  $('restBtn').addEventListener('click', startRest);
  restCancel.addEventListener('click', closeRest);
  restClose.addEventListener('click', closeRest);
  // Coming back to the app (screen was locked / switched apps): catch up, and re-take the wake lock.
  document.addEventListener('visibilitychange', () => {
    if(!document.hidden && rest.classList.contains('open') && !rest.classList.contains('done')){ keepAwake(true); renderRest(); }
  });
  // A rest that was running when the app was closed continues where it left off.
  if(Number(load(STORAGE_REST, 0)) > Date.now()) openRest();


  // ---- link to the Google Sheet (the sheet script reports its own address) ----
  const STORAGE_SHEET = 'bp_tracker_sheet_url';
  const sheetLink = $('sheetLink');
  function setSheetUrl(url){
    if(!url || !/^https:\/\/docs\.google\.com\//.test(url)) return;
    store(STORAGE_SHEET, url);
    sheetLink.href = url;
    sheetLink.hidden = false;
  }
  setSheetUrl(load(STORAGE_SHEET, ''));
  // Already connected from before this feature? Ask the sheet for its address once.
  if(!load(STORAGE_SHEET, '') && getUrl() && getToken()){
    api({ action: 'ping' }).then(r => setSheetUrl(r.sheetUrl)).catch(() => {});
  }


  // ---- measurement date & time (always 24h, always local time) ----
  const mDate = $('mDate'), mTime = $('mTime'), timeSrc = $('timeSrc');
  const pad2 = n => String(n).padStart(2, '0');
  function fmt24(d){ return pad2(d.getDate()) + '/' + pad2(d.getMonth() + 1) + '/' + d.getFullYear() + ' ' + pad2(d.getHours()) + ':' + pad2(d.getMinutes()); }
  function setWhen(d, label){
    mDate.value = d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
    mTime.value = pad2(d.getHours()) + ':' + pad2(d.getMinutes());
    mTime.classList.remove('bad');
    timeSrc.textContent = label || '';
  }
  function readWhen(){
    const m = /^(\d{1,2})[:.]?(\d{2})$/.exec(mTime.value.trim());
    const dm = /^(\d{4})-(\d{2})-(\d{2})$/.exec(mDate.value);
    if(!m || !dm || +m[1] > 23 || +m[2] > 59) return null;
    return new Date(+dm[1], +dm[2] - 1, +dm[3], +m[1], +m[2]);
  }
  // Type "1430" and get "14:30".
  mTime.addEventListener('input', () => {
    const digits = mTime.value.replace(/\D/g, '').slice(0, 4);
    mTime.value = digits.length > 2 ? digits.slice(0, 2) + ':' + digits.slice(2) : digits;
    mTime.classList.remove('bad');
    timeSrc.textContent = '';
  });
  mDate.addEventListener('input', () => { timeSrc.textContent = ''; });
  let whenTouched = false;
  [mDate, mTime].forEach(el => el.addEventListener('input', () => { whenTouched = true; }));
  setWhen(new Date());
  // Keep "now" fresh while the app sits open, unless the user set a time.
  setInterval(() => { if(!whenTouched && !photoTime) setWhen(new Date()); }, 30000);
  let photoTime = null;

  let source = 'ידני';

  async function handlePhoto(input){
    const file = input.files[0];
    if(!file) return;
    photoPreview.src = URL.createObjectURL(file);
    photoPreview.style.display = 'block';
    // A photo from the gallery keeps the time it was taken — use that, not "now".
    const t = file.lastModified ? new Date(file.lastModified) : null;
    const ageMs = t ? Date.now() - t.getTime() : Infinity;
    if(!whenTouched && ageMs > 2 * 60 * 1000 && ageMs < 14 * 24 * 3600 * 1000){
      photoTime = t; setWhen(t, '(לפי שעת הצילום)');
    }else if(!whenTouched){
      photoTime = null; setWhen(new Date());
    }
    setStatus('');
    if(!getUrl() || !getToken()){
      setStatus('כדי לקרוא את התמונה צריך להגדיר חיבור לגיליון (למטה). אפשר להקליד ידנית.', 'err');
      configBlock.open = true;
      return;
    }
    reading.classList.add('show');
    try{
      const dataUrl = await resizeImage(file, 2400);
      const res = await api({ action: 'analyze', image: dataUrl.split(',')[1], mediaType: 'image/jpeg' });
      const r = res.reading;
      if(r.readable){
        sys.value = r.systolic;
        dia.value = r.diastolic;
        pulse.value = r.pulse || '';
        if(r.note && !notes.value.trim()) notes.value = r.note;
        source = 'צילום';
        updateDisplay();
        setStatus('בדקו שהמספרים נכונים ולחצו "שמור מדידה" (הקישו על התמונה להגדלה)', 'ok');
      }else if(r.systolic || r.diastolic || r.pulse){
        // Partial read (e.g. glare over one number): fill what we have, leave the rest to type.
        if(r.systolic) sys.value = r.systolic;
        if(r.diastolic) dia.value = r.diastolic;
        if(r.pulse) pulse.value = r.pulse;
        source = 'צילום';
        updateDisplay();
        const missing = [!r.systolic && 'סיסטולי', !r.diastolic && 'דיאסטולי', !r.pulse && 'דופק'].filter(Boolean).join(', ');
        setStatus('חלק מהמסך לא היה ברור (בוהק?) — השלימו ידנית: ' + missing, 'err');
        openViewer(true);
      }else{
        setStatus('לא הצלחתי לקרוא את המסך בבירור — נסו לצלם שוב או הקלידו ידנית' + (r.reason ? ' (' + r.reason + ')' : ''), 'err');
        openViewer(true);
      }
    }catch(err){
      setStatus((err && err.message ? err.message : 'הקריאה נכשלה') + ' — אפשר להקליד ידנית', 'err');
    }finally{
      reading.classList.remove('show');
    }
  }
  photoCamera.addEventListener('change', () => handlePhoto(photoCamera));
  photoGallery.addEventListener('change', () => handlePhoto(photoGallery));

  // load stored config
  scriptUrlInput.value = getUrl();
  scriptTokenInput.value = getToken();
  if(!getUrl() || !getToken()){ configBlock.open = true; }

  saveUrlBtn.addEventListener('click', async () => {
    const v = scriptUrlInput.value.trim(), t = scriptTokenInput.value.trim();
    if(!v || !t){ setStatus('צריך למלא כתובת וקוד גישה', 'err'); return; }
    try { localStorage.setItem(STORAGE_URL, v); localStorage.setItem(STORAGE_TOKEN, t); } catch(e){}
    setStatus('בודק חיבור...');
    try{
      const pong = await api({ action: 'ping' });
      setSheetUrl(pong.sheetUrl);
      setStatus('מחובר לגיליון ✓', 'ok');
      configBlock.open = false;
      flushQueue();
    }catch(err){
      setStatus('החיבור נכשל: ' + err.message, 'err');
    }
  });

  let flushing = null;
  function flushQueue(){
    if(flushing) return flushing;   // a save that arrives mid-flush waits for the same run
    if(!queue.length || !getUrl()) return Promise.resolve();
    flushing = (async () => {
      try{
        for(const item of queue.slice()){
          await api({ action: 'save', reading: item });
          queue = queue.filter(q => q.id !== item.id);
          store(STORAGE_QUEUE, queue);
        }
      }catch(e){
        // keep the rest for the next attempt
      }finally{
        flushing = null;
        renderPending();
        renderHistory();
      }
    })();
    return flushing;
  }
  window.addEventListener('online', flushQueue);
  document.addEventListener('visibilitychange', () => { if(!document.hidden) flushQueue(); });

  saveBtn.addEventListener('click', async () => {
    const s = sys.value.trim(), d = dia.value.trim(), p = pulse.value.trim();
    if(!s || !d){
      setStatus('נא למלא סיסטולי ודיאסטולי', 'err');
      return;
    }
    if(Number(d) >= Number(s)){
      setStatus('הדיאסטולי (תחתון) צריך להיות נמוך מהסיסטולי (עליון)', 'err');
      return;
    }
    if(!getUrl() || !getToken()){
      setStatus('קודם צריך להגדיר חיבור לגיליון (למטה)', 'err');
      configBlock.open = true;
      return;
    }

    const now = readWhen();
    if(!now){
      mTime.classList.add('bad'); mTime.focus();
      setStatus('נא לכתוב שעה בפורמט 24 שעות, למשל 08:15 או 21:40', 'err');
      return;
    }
    if(now.getTime() > Date.now() + 5 * 60 * 1000){
      mTime.classList.add('bad'); mTime.focus();
      setStatus('מועד המדידה בעתיד — כדאי לבדוק את התאריך והשעה', 'err');
      return;
    }

    saveBtn.disabled = true;
    setStatus('שולח...');

    const item = {
      id: (crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random().toString(36).slice(2)),
      systolic: s,
      diastolic: d,
      pulse: p,
      note: notes.value.trim(),
      source: source,
      takenAt: now.toISOString()
    };
    queue.push(item);
    store(STORAGE_QUEUE, queue);

    const hist = getHistory();
    hist.unshift({
      id: item.id, sys: s, dia: d, pulse: p, notes: item.note,
      when: fmt24(now)
    });
    setHistory(hist);

    sys.value = ''; dia.value = ''; pulse.value = ''; notes.value = '';
    updateDisplay();
    photoPreview.style.display = 'none'; photoCamera.value = ''; photoGallery.value = '';
    source = 'ידני';
    photoTime = null; whenTouched = false; setWhen(new Date());

    await flushQueue();
    if(queue.some(q => q.id === item.id)) await flushQueue();
    if(queue.some(q => q.id === item.id)){
      setStatus(navigator.onLine ? 'השליחה נכשלה — המדידה שמורה ותישלח שוב אוטומטית' : 'אין אינטרנט — המדידה שמורה ותישלח כשיחזור החיבור', 'err');
    }else{
      setStatus('נשמר בגיליון ✓', 'ok');
    }
    saveBtn.disabled = false;
  });

  if('serviceWorker' in navigator){
    window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch(() => {}));
  }

  renderHistory();
  renderPending();
  flushQueue();
})();
