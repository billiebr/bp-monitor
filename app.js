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
    historyList.innerHTML = list.map(e => `
      <div class="entry">
        <div>
          <div class="val">${escapeHtml(e.sys)}/${escapeHtml(e.dia)}${e.pulse ? ` <span class="when">· דופק ${escapeHtml(e.pulse)}</span>` : ''}</div>
          ${e.notes ? `<div class="note">${escapeHtml(e.notes)}</div>` : ''}
        </div>
        <div class="when">${escapeHtml(e.when)}${queue.some(q => q.id === e.id) ? ' ⏳' : ''}</div>
      </div>
    `).join('');
  }

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
        resolve(c.toDataURL('image/jpeg', 0.85));
      };
      img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('לא הצלחתי לפתוח את התמונה')); };
      img.src = url;
    });
  }

  let source = 'ידני';

  async function handlePhoto(input){
    const file = input.files[0];
    if(!file) return;
    photoPreview.src = URL.createObjectURL(file);
    photoPreview.style.display = 'block';
    setStatus('');
    if(!getUrl() || !getToken()){
      setStatus('כדי לקרוא את התמונה צריך להגדיר חיבור לגיליון (למטה). אפשר להקליד ידנית.', 'err');
      configBlock.open = true;
      return;
    }
    reading.classList.add('show');
    try{
      const dataUrl = await resizeImage(file, 1600);
      const res = await api({ action: 'analyze', image: dataUrl.split(',')[1], mediaType: 'image/jpeg' });
      const r = res.reading;
      if(r.readable){
        sys.value = r.systolic;
        dia.value = r.diastolic;
        pulse.value = r.pulse || '';
        if(r.note && !notes.value.trim()) notes.value = r.note;
        source = 'צילום';
        updateDisplay();
        setStatus('בדקו שהמספרים נכונים ולחצו "שמור מדידה"', 'ok');
      }else{
        setStatus('לא הצלחתי לקרוא את המסך בבירור — נסו לצלם שוב או הקלידו ידנית' + (r.reason ? ' (' + r.reason + ')' : ''), 'err');
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
      await api({ action: 'ping' });
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

    saveBtn.disabled = true;
    setStatus('שולח...');

    const now = new Date();
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
      when: now.toLocaleDateString('he-IL') + ' ' + now.toLocaleTimeString('he-IL', {hour:'2-digit', minute:'2-digit'})
    });
    setHistory(hist);

    sys.value = ''; dia.value = ''; pulse.value = ''; notes.value = '';
    updateDisplay();
    photoPreview.style.display = 'none'; photoCamera.value = ''; photoGallery.value = '';
    source = 'ידני';

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
