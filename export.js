/* export.js — tombol "Unduh PDF" di setiap tab. Output: A4 potret, tampilan terang. */
(function () {
  const LIBS = [
    'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js',
  ];
  const TAB_NAMA = {
    jadwal: 'Jadwal OH NRE', monitoring: 'Monitoring Progress Harian', biaya: 'Rencana & Realisasi Biaya',
    rapat: 'Jadwal Rapat & Persiapan', ditgas: 'Monitoring OH Ditgas 2026',
  };
  const W_EXPORT = 1040; // lebar halaman saat dirender ke PDF (px)

  const loadScript = (src) => new Promise((ok, no) => {
    if ([...document.scripts].some((s) => s.src === src)) return ok();
    const s = document.createElement('script'); s.src = src; s.onload = ok; s.onerror = () => no(new Error('Gagal memuat ' + src));
    document.head.appendChild(s);
  });
  const wait = (ms) => new Promise((r) => setTimeout(r, ms));
  const frame = () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));

  function tambahTombol() {
    document.querySelectorAll('.oh-page').forEach((sec) => {
      if (sec.querySelector('.oh-pdf')) return;
      const b = document.createElement('button');
      b.type = 'button'; b.className = 'oh-pdf'; b.dataset.tab = sec.id;
      b.innerHTML = '<svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true"><path d="M8 2v8m0 0L5 7m3 3l3-3M3 12.5h10" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg><span>Unduh PDF</span>';
      b.addEventListener('click', () => ekspor(sec.id, b));
      const head = sec.querySelector('header.oh-head');
      const ctrl = head && head.querySelector('.oh-ctrl');
      if (ctrl) ctrl.appendChild(b);
      else if (head) { const c = document.createElement('div'); c.className = 'oh-ctrl'; c.appendChild(b); head.appendChild(c); }
      else sec.prepend(b);
    });
  }

  let jalan = false;
  async function ekspor(tab, btn) {
    if (jalan) return;
    jalan = true;
    const label = btn.querySelector('span'); const teks = label.textContent;
    label.textContent = 'Menyiapkan…'; btn.disabled = true;
    const html = document.documentElement, root = document.getElementById('dash-root');
    const modeLama = html.getAttribute('data-mode'), themeLama = html.getAttribute('data-theme');
    try {
      await Promise.all(LIBS.map(loadScript));
      html.classList.add('exporting');
      html.setAttribute('data-mode', 'light'); html.setAttribute('data-theme', 'light');
      root.style.width = W_EXPORT + 'px';
      window.dispatchEvent(new Event('resize'));
      await wait(450); await frame();
      if (document.fonts && document.fonts.ready) await document.fonts.ready;

      const rootTop = root.getBoundingClientRect().top;
      const breaks = new Set();
      root.querySelectorAll('.oh-top, header.oh-head, .oh-kpis, .oh-kpi, .oh-card, .oh-pocard, tr, .oh-legend, .web-foot, .oh-page > .note, svg > g')
        .forEach((el) => { if (el.offsetParent !== null) breaks.add(Math.round(el.getBoundingClientRect().top - rootTop)); });

      const scale = 2;
      const canvas = await window.html2canvas(root, { scale, backgroundColor: getComputedStyle(document.body).backgroundColor || '#ffffff', useCORS: true, logging: false, windowWidth: W_EXPORT + 80 });

      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4', compress: true });
      const M = 10, PW = 210, PH = 297, FOOT = 8;
      const cw = PW - 2 * M, ch = PH - 2 * M - FOOT;
      const pxPerMm = canvas.width / scale / cw;          // px CSS per mm
      const pageH = Math.floor(ch * pxPerMm);              // tinggi konten per halaman (px CSS)
      const totalH = canvas.height / scale;
      const pts = [...breaks].filter((y) => y > 0 && y < totalH).sort((a, b) => a - b);
      const slices = []; let y0 = 0;
      while (y0 < totalH - 2) {
        let y1 = Math.min(totalH, y0 + pageH);
        if (y1 < totalH) {
          const cand = pts.filter((p) => p > y0 + pageH * 0.45 && p <= y0 + pageH);
          if (cand.length) y1 = cand[cand.length - 1];
        }
        slices.push([y0, y1]); y0 = y1;
      }
      const tgl = new Date().toLocaleString('id-ID', { day: 'numeric', month: 'long', year: 'numeric', hour: '2-digit', minute: '2-digit' });
      slices.forEach(([a, b], i) => {
        if (i) pdf.addPage();
        const part = document.createElement('canvas');
        part.width = canvas.width; part.height = Math.ceil((b - a) * scale);
        const ctx = part.getContext('2d');
        ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, part.width, part.height);
        ctx.drawImage(canvas, 0, Math.floor(a * scale), canvas.width, part.height, 0, 0, canvas.width, part.height);
        pdf.addImage(part.toDataURL('image/jpeg', 0.92), 'JPEG', M, M, cw, (b - a) / pxPerMm, undefined, 'FAST');
        pdf.setFontSize(8); pdf.setTextColor(110);
        pdf.text(`Dashboard OH · OM NRE  —  ${TAB_NAMA[tab] || tab}`, M, PH - M + 2);
        pdf.text(`Dicetak ${tgl}  ·  Halaman ${i + 1} dari ${slices.length}`, PW - M, PH - M + 2, { align: 'right' });
      });
      const d = new Date(); const ymd = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
      pdf.save(`Dashboard-OH_${(TAB_NAMA[tab] || tab).replace(/[^A-Za-z0-9]+/g, '-')}_${ymd}.pdf`);
    } catch (e) {
      console.error(e);
      alert('PDF gagal dibuat: ' + e.message);
    } finally {
      root.style.width = '';
      if (modeLama) html.setAttribute('data-mode', modeLama); else html.removeAttribute('data-mode');
      if (themeLama) html.setAttribute('data-theme', themeLama); else html.removeAttribute('data-theme');
      html.classList.remove('exporting');
      window.dispatchEvent(new Event('resize'));
      label.textContent = teks; btn.disabled = false; jalan = false;
    }
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', tambahTombol); else tambahTombol();
})();
