/* dash.js — pengganti runtime dashboard Claude untuk versi web.
   Memuat data.json (hasil update_data.py) dan CSV laporan harian (Google Sheets "Publish to web"). */
(function () {
  const CFG = window.OH_CONFIG || {};
  const DEFAULT_PARAMS = { tahun: new Date().getFullYear(), up: 'Semua UP', tahun_b: new Date().getFullYear() - 1, up_b: 'Semua UP', up_r: 'Semua UP', rwin: '90', up_g: 'Semua UP', div_g: 'Semua Divisi' };
  const DS = {};            // dataset id -> {status, data, message, meta}
  const CALCS = {}, LOADERS = {}, CBS = [];
  let memo = {}, params = { ...DEFAULT_PARAMS }, meta = {};

  const css = getComputedStyle(document.documentElement);
  const tok = (n, f) => (css.getPropertyValue(n).trim() ? `var(${n})` : f);

  const dash = {
    colors: ['--s1', '--s2', '--s3', '--s4', '--s5', '--s6', '--s7', '--s8'].map((n) => tok(n, '#888')),
    calc(id, o) { CALCS[id] = o; },
    loader(id, o) { LOADERS[id] = o; },
    data(id) {
      if (CALCS[id]) return evalCalc(id);
      const d = DS[id] || { status: 'loading', data: [] };
      return { ...d, columns: d.data && d.data[0] ? Object.keys(d.data[0]) : [], meta: d.meta || [], refreshing: false };
    },
    onData(fn) { CBS.push(fn); safe(fn); },
    params() { return { ...params }; },
    setParams(p) { params = { ...params, ...p }; notify(); },
    resetParams() { params = { ...DEFAULT_PARAMS }; notify(); },
    setLink(name) { try { history.replaceState(null, '', '#' + name); } catch (e) { /* abaikan */ } },
    refresh() { load(); },
    meta: () => meta,
  };
  window.dash = dash;

  function safe(fn) { try { fn(); } catch (e) { console.error(e); } }
  let pending = false;
  function notify() {
    memo = {};
    if (pending) return;
    pending = true;
    requestAnimationFrame(() => { pending = false; CBS.forEach(safe); });
  }

  function evalCalc(id) {
    if (memo[id]) return memo[id];
    const c = CALCS[id];
    const ins = c.inputs.map((i) => dash.data(i));
    let res;
    const bad = ins.find((d) => d.status !== 'ok');
    if (bad) res = { status: bad.status === 'loading' ? 'loading' : bad.status, data: [], message: bad.message };
    else {
      try { res = { status: 'ok', data: c.fn(...ins.map((d) => d.data)) }; }
      catch (e) { console.error('calc', id, e); res = { status: 'error', data: [], message: String(e) }; }
    }
    res.columns = res.data && res.data[0] ? Object.keys(res.data[0]) : [];
    res.meta = [];
    memo[id] = res;
    return res;
  }

  function parseCSV(text) {
    const rows = []; let row = [], cell = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) {
        if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; }
        else cell += ch;
      } else if (ch === '"') q = true;
      else if (ch === ',') { row.push(cell); cell = ''; }
      else if (ch === '\n') { row.push(cell); rows.push(row); row = []; cell = ''; }
      else if (ch !== '\r') cell += ch;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows;
  }

  async function load() {
    // data Excel
    try {
      const r = await fetch('data.json?v=' + Date.now(), { cache: 'no-store' });
      if (!r.ok) throw new Error('data.json tidak ditemukan (' + r.status + ')');
      const j = await r.json();
      meta = { dibuat: j.dibuat, sumber: j.sumber || {} };
      ['jadwal', 'ditgas', 'om4'].forEach((k) => { DS[k] = { status: 'ok', data: j[k] || [] }; });
    } catch (e) {
      ['jadwal', 'ditgas', 'om4'].forEach((k) => { DS[k] = { status: 'error', data: [], message: String(e) }; });
    }
    notify();
    // laporan harian
    if (!CFG.harianCsvUrl) {
      ['harian_po', 'harian_fo'].forEach((k) => { DS[k] = { status: 'connect', data: [], message: 'harianCsvUrl belum diisi di config.js' }; });
    } else {
      try {
        const sep = CFG.harianCsvUrl.includes('?') ? '&' : '?';
        const r = await fetch(CFG.harianCsvUrl + sep + 't=' + Date.now(), { cache: 'no-store' });
        if (!r.ok) throw new Error('Gagal membaca CSV (' + r.status + ')');
        const answer = { values: parseCSV(await r.text()) };
        ['harian_po', 'harian_fo'].forEach((k) => {
          try { const out = LOADERS[k].fn({ answer }); DS[k] = { status: 'ok', data: out.data || out, meta: out.meta || [] }; }
          catch (e) { DS[k] = { status: 'error', data: [], message: String(e) }; }
        });
      } catch (e) {
        ['harian_po', 'harian_fo'].forEach((k) => { DS[k] = { status: 'error', data: [], message: String(e) }; });
      }
    }
    notify();
  }

  // tema terang/gelap: ingat pilihan pembaca
  const html = document.documentElement;
  try { const m = localStorage.getItem('oh-mode'); if (m === 'dark' || m === 'light') { html.setAttribute('data-mode', m); html.setAttribute('data-theme', m); } } catch (e) { /* abaikan */ }
  new MutationObserver(() => { if (html.classList.contains('exporting')) return; try { const m = html.getAttribute('data-mode'); if (m) localStorage.setItem('oh-mode', m); } catch (e) { /* abaikan */ } })
    .observe(html, { attributes: true, attributeFilter: ['data-mode'] });

  let rt;
  window.addEventListener('resize', () => { clearTimeout(rt); rt = setTimeout(notify, 150); });
  dash.__start = () => {
    load();
    const ms = (Number(CFG.refreshMenit) || 0) * 60000;
    if (ms > 0) setInterval(load, ms);
  };
})();
