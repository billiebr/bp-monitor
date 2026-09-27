/* מד לחץ דם — צד הלקוח (PWA). הנתונים נשמרים ב-Google Sheets דרך Apps Script. */
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };
  var LS = {
    get: function (k, d) { try { var v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
    set: function (k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* storage unavailable */ } },
  };

  var config = LS.get('bp.config', { url: '', token: '' });
  var queue = LS.get('bp.queue', []);         // readings waiting to reach the sheet
  var cache = LS.get('bp.readings', []);      // last list pulled from the sheet
  var source = 'ידני';
  var syncing = false;

  // ------------------------------------------------------------ helpers

  function toast(msg, ms) {
    var t = $('toast');
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toast._t);
    toast._t = setTimeout(function () { t.classList.remove('show'); }, ms || 2800);
  }

  function uid() {
    return (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2));
  }

  function localInputValue(d) {
    var p = function (n) { return String(n).padStart(2, '0'); };
    return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + 'T' + p(d.getHours()) + ':' + p(d.getMinutes());
  }

  function fmtDate(iso) {
    var d = new Date(iso);
    if (isNaN(d)) return iso;
    return d.toLocaleDateString('he-IL', { day: 'numeric', month: 'numeric', year: '2-digit' }) + ' ' +
      d.toLocaleTimeString('he-IL', { hour: '2-digit', minute: '2-digit' });
  }

  // Same thresholds as classify_ in Code.gs (ESC categories).
  function classify(sys, dia) {
    if (!sys || !dia) return null;
    if (sys >= 180 || dia >= 110) return { t: 'יתר לחץ דם דרגה 3', c: 'cat-3' };
    if (sys >= 160 || dia >= 100) return { t: 'יתר לחץ דם דרגה 2', c: 'cat-2' };
    if (sys >= 140 || dia >= 90) return { t: 'יתר לחץ דם דרגה 1', c: 'cat-1' };
    if (sys >= 130 || dia >= 85) return { t: 'תקין-גבוה', c: 'cat-high-normal' };
    if (sys < 90 || dia < 60) return { t: 'נמוך', c: 'cat-low' };
    return { t: 'תקין', c: 'cat-ok' };
  }

  function api(payload) {
    if (!config.url || !config.token) return Promise.reject(new Error('צריך להגדיר קודם את החיבור לגיליון'));
    payload.token = config.token;
    // text/plain keeps this a "simple" request, so Apps Script needs no CORS preflight.
    return fetch(config.url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload),
      redirect: 'follow',
    })
      .then(function (r) { return r.json(); })
      .then(function (data) {
        if (!data.ok) throw new Error(data.error || 'שגיאה');
        return data;
      });
  }

  // Shrink the camera photo before upload: faster, cheaper, same legibility.
  function resizeImage(file, max) {
    return new Promise(function (resolve, reject) {
      var url = URL.createObjectURL(file);
      var img = new Image();
      img.onload = function () {
        var scale = Math.min(1, max / Math.max(img.width, img.height));
        var c = document.createElement('canvas');
        c.width = Math.round(img.width * scale);
        c.height = Math.round(img.height * scale);
        c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
        URL.revokeObjectURL(url);
        resolve(c.toDataURL('image/jpeg', 0.85));
      };
      img.onerror = function () { URL.revokeObjectURL(url); reject(new Error('לא הצלחתי לפתוח את התמונה')); };
      img.src = url;
    });
  }

  // ------------------------------------------------------------ tabs

  function showTab(name) {
    document.querySelectorAll('.view').forEach(function (v) { v.classList.toggle('active', v.id === 'view-' + name); });
    document.querySelectorAll('.tabs button').forEach(function (b) { b.classList.toggle('active', b.dataset.tab === name); });
    if (name === 'history') { renderHistory(); loadHistory(); }
    window.scrollTo(0, 0);
  }
  document.querySelectorAll('.tabs button').forEach(function (b) {
    b.addEventListener('click', function () { showTab(b.dataset.tab); });
  });
  document.querySelectorAll('[data-goto]').forEach(function (a) {
    a.addEventListener('click', function (e) { e.preventDefault(); showTab(a.dataset.goto); });
  });

  // ------------------------------------------------------------ measure

  var form = $('form');

  function openForm(values) {
    values = values || {};
    $('sys').value = values.systolic || '';
    $('dia').value = values.diastolic || '';
    $('pulse').value = values.pulse || '';
    $('note').value = values.note || '';
    $('taken').value = localInputValue(new Date());
    form.hidden = false;
    $('capture').hidden = true;
    $('last').hidden = true;
    updateCategory();
  }

  function resetMeasure() {
    form.hidden = true;
    form.reset();
    $('preview-wrap').hidden = true;
    $('preview').removeAttribute('src');
    $('capture').hidden = false;
    $('photo').value = '';
    source = 'ידני';
  }

  function updateCategory() {
    var cat = classify(Number($('sys').value), Number($('dia').value));
    var el = $('category');
    el.className = 'category' + (cat ? ' ' + cat.c : '');
    el.textContent = cat ? cat.t : '';
  }
  ['sys', 'dia'].forEach(function (id) { $(id).addEventListener('input', updateCategory); });

  $('manual').addEventListener('click', function () { source = 'ידני'; openForm(); $('sys').focus(); });
  $('cancel').addEventListener('click', resetMeasure);

  $('photo').addEventListener('change', function () {
    var file = this.files && this.files[0];
    if (!file) return;
    source = 'צילום';
    $('preview-wrap').hidden = false;
    $('analyzing').hidden = false;
    $('capture').hidden = true;
    $('last').hidden = true;
    form.hidden = true;

    resizeImage(file, 1600)
      .then(function (dataUrl) {
        $('preview').src = dataUrl;
        return api({ action: 'analyze', image: dataUrl.split(',')[1], mediaType: 'image/jpeg' });
      })
      .then(function (res) {
        var r = res.reading;
        $('analyzing').hidden = true;
        if (!r.readable) {
          openForm({ note: r.note || '' });
          toast('לא הצלחתי לקרוא את המסך בבירור — אפשר להקליד את הערכים', 4000);
          $('sys').focus();
          return;
        }
        openForm({ systolic: r.systolic, diastolic: r.diastolic, pulse: r.pulse || '', note: r.note || '' });
        toast('בדקי שהמספרים נכונים ושמרי ✔');
      })
      .catch(function (err) {
        $('analyzing').hidden = true;
        openForm();
        toast((err && err.message) || 'הקריאה נכשלה — אפשר להקליד ידנית', 4500);
      });
  });

  form.addEventListener('submit', function (e) {
    e.preventDefault();
    var sys = Number($('sys').value), dia = Number($('dia').value);
    if (!(dia < sys)) { toast('הערך הדיאסטולי צריך להיות נמוך מהסיסטולי'); return; }
    var reading = {
      id: uid(),
      systolic: sys,
      diastolic: dia,
      pulse: $('pulse').value ? Number($('pulse').value) : '',
      takenAt: new Date($('taken').value).toISOString(),
      note: $('note').value.trim(),
      source: source,
    };
    queue.push(reading);
    LS.set('bp.queue', queue);
    // Show it in history right away; the sheet copy replaces it after sync.
    cache.unshift(Object.assign({ category: (classify(sys, dia) || {}).t, pending: true }, reading));
    LS.set('bp.readings', cache);
    showLast(reading);
    resetMeasure();
    sync(true);
  });

  function showLast(r) {
    var cat = classify(r.systolic, r.diastolic);
    $('last').innerHTML =
      '<div class="sub">מדידה אחרונה</div>' +
      '<div class="big">' + r.systolic + '/' + r.diastolic + '</div>' +
      (r.pulse ? '<div class="sub">דופק ' + r.pulse + '</div>' : '') +
      (cat ? '<div class="category ' + cat.c + '">' + cat.t + '</div>' : '');
    $('last').hidden = false;
  }

  // ------------------------------------------------------------ sync (offline queue)

  function updatePending() {
    var el = $('pending');
    el.hidden = queue.length === 0;
    el.textContent = queue.length + ' ממתינות לשליחה';
  }

  function sync(announce) {
    updatePending();
    if (syncing || !queue.length || !config.url) return Promise.resolve();
    syncing = true;
    var chain = Promise.resolve();
    queue.slice().forEach(function (r) {
      chain = chain.then(function () {
        return api({ action: 'save', reading: r }).then(function () {
          queue = queue.filter(function (q) { return q.id !== r.id; });
          LS.set('bp.queue', queue);
          updatePending();
        });
      });
    });
    return chain
      .then(function () { if (announce) toast('נשמר בגיליון ✔'); loadHistory(true); })
      .catch(function (err) {
        if (announce) toast(navigator.onLine ? 'השמירה נכשלה: ' + err.message + ' — אנסה שוב' : 'אין אינטרנט — אשלח כשיחזור החיבור', 4000);
      })
      .then(function () { syncing = false; updatePending(); });
  }

  window.addEventListener('online', function () { sync(true); });
  document.addEventListener('visibilitychange', function () { if (!document.hidden) sync(false); });
  setInterval(function () { sync(false); }, 60000);

  // ------------------------------------------------------------ history

  function loadHistory(silent) {
    if (!config.url) return;
    api({ action: 'list', limit: 300 })
      .then(function (res) {
        var pending = queue.map(function (r) {
          return Object.assign({ category: (classify(r.systolic, r.diastolic) || {}).t, pending: true }, r);
        });
        cache = pending.concat(res.readings);
        LS.set('bp.readings', cache);
        if (res.sheetUrl) { LS.set('bp.sheetUrl', res.sheetUrl); }
        renderHistory();
      })
      .catch(function (err) { if (!silent) toast('לא הצלחתי לטעון מהגיליון: ' + err.message); });
  }
  $('refresh').addEventListener('click', function () { loadHistory(false); });

  var CAT_CLASS = {
    'נמוך': 'cat-low', 'תקין': 'cat-ok', 'תקין-גבוה': 'cat-high-normal',
    'יתר לחץ דם דרגה 1': 'cat-1', 'יתר לחץ דם דרגה 2': 'cat-2', 'יתר לחץ דם דרגה 3': 'cat-3',
  };

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function renderHistory() {
    var sheetUrl = LS.get('bp.sheetUrl', '');
    $('sheet-link').hidden = !sheetUrl;
    if (sheetUrl) $('sheet-link').href = sheetUrl;

    var list = $('list');
    if (!cache.length) {
      list.innerHTML = '<li class="empty">עוד אין מדידות</li>';
    } else {
      list.innerHTML = cache.map(function (r) {
        return '<li><span class="val">' + esc(r.systolic) + '/' + esc(r.diastolic) + '</span>' +
          '<span class="meta">' + esc(fmtDate(r.takenAt)) + (r.pulse ? ' · דופק ' + esc(r.pulse) : '') +
          (r.note ? '<br>' + esc(r.note) : '') + (r.pending ? '<br>⏳ ממתין לשליחה' : '') + '</span>' +
          (r.category ? '<span class="badge ' + (CAT_CLASS[r.category] || '') + '">' + esc(r.category) + '</span>' : '') +
          '</li>';
      }).join('');
    }
    renderStats();
    renderChart();
  }

  function avg(arr, key) {
    var v = arr.map(function (r) { return Number(r[key]); }).filter(function (n) { return n > 0; });
    return v.length ? Math.round(v.reduce(function (a, b) { return a + b; }, 0) / v.length) : null;
  }

  function renderStats() {
    var weekAgo = Date.now() - 7 * 864e5;
    var week = cache.filter(function (r) { return new Date(r.takenAt).getTime() >= weekAgo; });
    var s = avg(week, 'systolic'), d = avg(week, 'diastolic'), p = avg(week, 'pulse');
    $('stats').innerHTML =
      '<div class="stat"><b>' + (s ? s + '/' + d : '—') + '</b><small>ממוצע 7 ימים</small></div>' +
      '<div class="stat"><b>' + (p || '—') + '</b><small>דופק ממוצע</small></div>' +
      '<div class="stat"><b>' + week.length + '</b><small>מדידות השבוע</small></div>' +
      '<div class="stat"><b>' + cache.length + '</b><small>סה״כ מדידות</small></div>';
  }

  function renderChart() {
    var pts = cache.slice(0, 30).reverse();
    var el = $('chart');
    if (pts.length < 2) { el.innerHTML = '<p class="empty">הגרף יופיע אחרי שתי מדידות</p>'; return; }
    var W = 340, H = 180, L = 30, R = 8, T = 10, B = 20;
    var vals = [];
    pts.forEach(function (r) { vals.push(+r.systolic, +r.diastolic); if (+r.pulse) vals.push(+r.pulse); });
    var min = Math.floor((Math.min.apply(null, vals) - 5) / 10) * 10;
    var max = Math.ceil((Math.max.apply(null, vals) + 5) / 10) * 10;
    var x = function (i) { return L + (i * (W - L - R)) / (pts.length - 1); };
    var y = function (v) { return T + ((max - v) * (H - T - B)) / (max - min); };
    var line = function (key, color) {
      var d = '', dots = '';
      pts.forEach(function (r, i) {
        var v = +r[key];
        if (!v) return;
        d += (d ? 'L' : 'M') + x(i).toFixed(1) + ' ' + y(v).toFixed(1);
        dots += '<circle cx="' + x(i).toFixed(1) + '" cy="' + y(v).toFixed(1) + '" r="2.6" fill="' + color + '"/>';
      });
      return '<path d="' + d + '" fill="none" stroke="' + color + '" stroke-width="2" stroke-linejoin="round"/>' + dots;
    };
    var grid = '';
    for (var g = min; g <= max; g += (max - min > 80 ? 20 : 10)) {
      grid += '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y(g) + '" y2="' + y(g) + '" stroke="currentColor" stroke-opacity=".12"/>' +
        '<text x="' + (L - 4) + '" y="' + (y(g) + 3) + '" font-size="9" text-anchor="end" fill="currentColor" fill-opacity=".55">' + g + '</text>';
    }
    // Reference lines at 140/90.
    var ref = [140, 90].filter(function (v) { return v > min && v < max; }).map(function (v) {
      return '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y(v) + '" y2="' + y(v) + '" stroke="#f59e0b" stroke-dasharray="4 3" stroke-opacity=".7"/>';
    }).join('');
    var first = new Date(pts[0].takenAt), last = new Date(pts[pts.length - 1].takenAt);
    var lbl = function (d) { return d.toLocaleDateString('he-IL', { day: 'numeric', month: 'numeric' }); };
    el.innerHTML = '<svg viewBox="0 0 ' + W + ' ' + H + '" direction="ltr" role="img" aria-label="גרף לחץ דם">' + grid + ref +
      line('systolic', '#dc2626') + line('diastolic', '#2563eb') + line('pulse', '#16a34a') +
      '<text x="' + L + '" y="' + (H - 4) + '" font-size="9" fill="currentColor" fill-opacity=".55">' + lbl(first) + '</text>' +
      '<text x="' + (W - R) + '" y="' + (H - 4) + '" font-size="9" text-anchor="end" fill="currentColor" fill-opacity=".55">' + lbl(last) + '</text>' +
      '</svg>';
  }

  // ------------------------------------------------------------ settings

  function fillSettings() {
    $('api-url').value = config.url || '';
    $('api-token').value = config.token || '';
    $('setup-hint').hidden = !!(config.url && config.token);
  }

  $('settings').addEventListener('submit', function (e) {
    e.preventDefault();
    config = { url: $('api-url').value.trim(), token: $('api-token').value.trim() };
    LS.set('bp.config', config);
    fillSettings();
    var st = $('settings-status');
    st.className = 'status';
    st.textContent = 'בודקת חיבור…';
    api({ action: 'ping' })
      .then(function (res) {
        if (res.sheetUrl) LS.set('bp.sheetUrl', res.sheetUrl);
        st.className = 'status ok';
        st.textContent = 'מחובר לגיליון ✔';
        sync(false);
        loadHistory(true);
      })
      .catch(function (err) {
        st.className = 'status err';
        st.textContent = 'החיבור נכשל: ' + err.message;
      });
  });

  // Allow one-tap setup: …/index.html#url=<exec url>&token=<code>
  (function importFromHash() {
    if (!location.hash || location.hash.indexOf('url=') < 0) return;
    var p = new URLSearchParams(location.hash.slice(1));
    if (p.get('url') && p.get('token')) {
      config = { url: p.get('url'), token: p.get('token') };
      LS.set('bp.config', config);
      history.replaceState(null, '', location.pathname);
      toast('החיבור לגיליון הוגדר ✔');
    }
  })();

  // ------------------------------------------------------------ install / service worker

  var deferredPrompt = null;
  window.addEventListener('beforeinstallprompt', function (e) {
    e.preventDefault();
    deferredPrompt = e;
    $('install').hidden = false;
  });
  $('install').addEventListener('click', function () {
    if (!deferredPrompt) return;
    deferredPrompt.prompt();
    deferredPrompt = null;
    $('install').hidden = true;
  });

  if ('serviceWorker' in navigator) {
    window.addEventListener('load', function () { navigator.serviceWorker.register('sw.js').catch(function () {}); });
  }

  // ------------------------------------------------------------ start

  fillSettings();
  updatePending();
  if (cache.length) showLast(cache[0]);
  if (!config.url) showTab('measure');
  sync(false);
})();
