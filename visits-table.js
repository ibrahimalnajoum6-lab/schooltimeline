// ===== جدول الزيارات بالتصميم المعتمد (التاريخ | اليوم | فريق الذكور | المدرسة | فريق الإناث | المدرسة 2) + تحميل Excel / PDF =====
(function () {
  const DAYS_AR = ['الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];
  const HEAD = ['التاريخ', 'اليوم', 'فريق الذكور', 'المدرسة', 'فريق الإناث', 'المدرسة 2'];
  // ألوان الأعمدة كما في الصورة
  const HEAD_BG = ['#5B9BD5', '#5B9BD5', '#5B9BD5', '#5B9BD5', '#FFD966', '#5B9BD5'];
  const BODY_BG = ['#BDD7EE', '#FFFFFF', '#DDEBF7', '#FFFFFF', '#FFD966', '#E2EFDA'];
  const LINE = '#9DC3E6';

  const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const COLL = new Intl.Collator('ar');

  function parseVisit(v) {
    const t = String(v || '').trim();
    const m = t.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) {
      const d = new Date(+m[1], +m[2] - 1, +m[3]);
      return { key: t, date: m[3] + '/' + m[2] + '/' + m[1], day: DAYS_AR[d.getDay()], sort: t };
    }
    if (DAYS_AR.includes(t)) return { key: 'z-' + t, date: '', day: t, sort: '9' + DAYS_AR.indexOf(t) };
    return { key: 'zz', date: '', day: '', sort: 'zzz' };
  }

  const teamText = r => (r && r.supervisors && r.supervisors.length) ? r.supervisors.map(s => s.name).join('\n') : '';

  // يبني صفوف الجدول: لكل يوم زيارة يقابَل فريق الذكور/مدرسته مع فريق الإناث/مدرستها
  window.buildVisitsRows = function (filter) {
    const f = filter || {};
    const q = String(f.q || '').trim();
    const list = (window.distributionResults || distributionResults || []).filter(r => {
      if (f.gender && f.gender !== 'ALL' && r.gender !== f.gender) return false;
      if (q && !(String(r.schoolName) + ' ' + String(r.zone) + ' ' + teamText(r)).includes(q)) return false;
      return true;
    });
    const groups = {};
    list.forEach(r => {
      const p = parseVisit(r.visitDate);
      const g = groups[p.key] = groups[p.key] || { p, m: [], f: [] };
      (r.gender === 'ذكور' ? g.m : g.f).push(r);
    });
    const rows = [];
    Object.values(groups).sort((a, b) => a.p.sort < b.p.sort ? -1 : a.p.sort > b.p.sort ? 1 : 0).forEach(g => {
      const byZone = (a, b) => COLL.compare(a.zone || '', b.zone || '') || COLL.compare(a.schoolName, b.schoolName);
      g.m.sort(byZone); g.f.sort(byZone);
      const n = Math.max(g.m.length, g.f.length);
      for (let i = 0; i < n; i++) {
        const m = g.m[i], fe = g.f[i];
        rows.push({
          cells: [g.p.date, g.p.day, teamText(m), m ? m.schoolName : '', teamText(fe), fe ? fe.schoolName : ''],
          mId: m && m.schoolId, fId: fe && fe.schoolId
        });
      }
    });
    return rows;
  };

  // ---------- العرض على الشاشة ----------
  const CSS = `
  #vt-wrap{background:#fff;border:1px solid #e2e8f0;border-radius:16px;overflow:hidden}
  #vt-scroll{overflow-x:auto}
  table.vt{width:100%;min-width:760px;border-collapse:collapse;direction:rtl;font-family:'Cairo',sans-serif;font-size:13px}
  table.vt th{color:#fff;font-weight:700;padding:9px 8px;text-align:center;border:1px solid ${LINE};white-space:nowrap}
  table.vt td{padding:8px;text-align:center;border:1px solid ${LINE};vertical-align:middle;line-height:1.7;white-space:pre-line}
  table.vt td.vt-link{cursor:pointer}
  table.vt td.vt-link:hover{text-decoration:underline}
  table.vt th::after{content:'\\25BE';display:inline-block;margin-inline-start:6px;background:#fff;color:#555;font-size:9px;padding:0 3px;border:1px solid #999;line-height:14px}
  .vt-bar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;justify-content:space-between;padding:12px 14px;border-bottom:1px solid #e2e8f0;background:#f8fafc}
  .vt-btn{display:inline-flex;align-items:center;gap:6px;font-size:12px;font-weight:700;padding:8px 14px;border-radius:12px;border:1px solid #cbd5e1;background:#fff;color:#0f172a;cursor:pointer}
  .vt-btn:hover{background:#f1f5f9}
  .vt-btn.ex{color:#15803d;border-color:#86efac;background:#f0fdf4}
  .vt-btn.pd{color:#b91c1c;border-color:#fca5a5;background:#fef2f2}
  @media print{#vt-bar{display:none!important}}
  `;

  function tableHtml(rows, forExport) {
    const th = HEAD.map((h, i) => `<th style="background:${HEAD_BG[i]}">${esc(h)}</th>`).join('');
    const body = rows.map(r => '<tr>' + r.cells.map((c, i) => {
      const id = i === 3 ? r.mId : i === 5 ? r.fId : null;
      const link = (!forExport && id && i !== 0) ? ` class="vt-link" onclick="openTeamBuilder('${id}')" title="تحديد/تعديل فريق الزيارة"` : '';
      return `<td${link} style="background:${BODY_BG[i]}">${esc(c)}</td>`;
    }).join('') + '</tr>').join('');
    return `<table class="vt"${forExport ? ' style="width:1100px;min-width:0"' : ''}><thead><tr>${th}</tr></thead><tbody>${body}</tbody></table>`;
  }

  function ensureUi() {
    if (!document.getElementById('vt-style')) {
      const st = document.createElement('style'); st.id = 'vt-style'; st.textContent = CSS; document.head.appendChild(st);
    }
    if (document.getElementById('vt-wrap')) return;
    const view = document.getElementById('view-distribution'); if (!view) return;
    // إخفاء الجدول القديم وبطاقاته
    const cards = document.getElementById('distribution-cards'); if (cards) cards.style.display = 'none';
    const old = view.querySelector('.hidden.md\\:block'); if (old) old.style.display = 'none';
    const wrap = document.createElement('div'); wrap.id = 'vt-wrap';
    wrap.innerHTML = `
      <div class="vt-bar no-print" id="vt-bar">
        <div style="font-size:12px;color:#475569;font-weight:600">جدول الزيارات الميدانية</div>
        <div style="display:flex;gap:8px;flex-wrap:wrap">
          <button class="vt-btn ex" onclick="downloadVisitsExcel()"><i class="fa-solid fa-file-excel"></i> تحميل Excel</button>
          <button class="vt-btn pd" onclick="downloadVisitsPdf()"><i class="fa-solid fa-file-pdf"></i> تحميل PDF</button>
        </div>
      </div>
      <div id="vt-scroll"></div>`;
    view.appendChild(wrap);
  }

  window.renderDistributionTable = function () {
    ensureUi();
    const box = document.getElementById('vt-scroll'); if (!box) return;
    const g = (document.getElementById('filter-dist-gender') || {}).value;
    const q = (document.getElementById('search-dist-school') || {}).value;
    const rows = window.buildVisitsRows({ gender: g, q });
    box.innerHTML = rows.length ? tableHtml(rows, false)
      : '<div style="padding:32px;text-align:center;color:#94a3b8;font-size:13px">لا يوجد جدول زيارات حالياً. اضغط «جدولة الزيارات» لإنشائه</div>';
  };

  // ---------- تحميل Excel (بالتنسيق واللون واتجاه RTL) ----------
  function loadScript(src) {
    return new Promise((res, rej) => {
      const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = () => rej(new Error('تعذّر تحميل المكتبة: ' + src));
      document.head.appendChild(s);
    });
  }
  const LIBS = {
    exceljs: 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js',
    h2c: 'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js',
    jspdf: 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js'
  };
  const argb = h => 'FF' + h.replace('#', '').toUpperCase();
  const fileStamp = () => new Date().toISOString().slice(0, 10);
  function saveBlob(blob, name) {
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }
  function currentRows() {
    const rows = window.buildVisitsRows({});
    if (!rows.length) { showToast('لا يوجد جدول زيارات للتصدير. اضغط «جدولة الزيارات» أولاً', 'error'); return null; }
    return rows;
  }

  window.downloadVisitsExcel = async function () {
    const rows = currentRows(); if (!rows) return;
    try {
      if (!window.ExcelJS) await loadScript(LIBS.exceljs);
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('جدول الزيارات', { views: [{ rightToLeft: true, state: 'frozen', ySplit: 1 }] });
      const widths = [14, 12, 34, 40, 34, 40];
      ws.columns = HEAD.map((h, i) => ({ header: h, key: 'c' + i, width: widths[i] }));
      const border = { style: 'thin', color: { argb: argb(LINE) } };
      const bd = { top: border, left: border, bottom: border, right: border };
      const hr = ws.getRow(1); hr.height = 26;
      hr.eachCell((c, i) => {
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(HEAD_BG[i - 1]) } };
        c.font = { name: 'Cairo', bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' }; c.border = bd;
      });
      rows.forEach(r => {
        const row = ws.addRow(r.cells);
        const lines = Math.max(1, ...r.cells.map(t => String(t).split('\n').length));
        row.height = Math.max(22, lines * 18);
        row.eachCell({ includeEmpty: true }, (c, i) => {
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(BODY_BG[i - 1]) } };
          c.font = { name: 'Cairo', size: 11 };
          c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; c.border = bd;
        });
      });
      ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: 6 } };
      ws.pageSetup = { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
      const buf = await wb.xlsx.writeBuffer();
      saveBlob(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'جدول_الزيارات_' + fileStamp() + '.xlsx');
      showToast('تم تحميل ملف Excel');
    } catch (e) { console.error(e); showToast(e.message || 'تعذّر إنشاء ملف Excel', 'error'); }
  };

  // ---------- تحميل PDF (صفحات A4 أفقية، الخط العربي يُرسم من الصفحة نفسها) ----------
  window.downloadVisitsPdf = async function () {
    const rows = currentRows(); if (!rows) return;
    try {
      if (!window.html2canvas) await loadScript(LIBS.h2c);
      if (!(window.jspdf && window.jspdf.jsPDF)) await loadScript(LIBS.jspdf);
      if (document.fonts && document.fonts.ready) await document.fonts.ready;
      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
      const PW = 297, PH = 210, M = 8;
      const PER_PAGE = 11;
      const host = document.createElement('div');
      host.style.cssText = 'position:fixed;left:-99999px;top:0;background:#fff;padding:14px;direction:rtl;font-family:Cairo,sans-serif';
      document.body.appendChild(host);
      const pages = Math.ceil(rows.length / PER_PAGE);
      for (let p = 0; p < pages; p++) {
        host.innerHTML = `<div style="text-align:center;font-weight:800;font-size:20px;margin-bottom:10px;color:#0f2f5c">جدول الزيارات الميدانية</div>` +
          tableHtml(rows.slice(p * PER_PAGE, (p + 1) * PER_PAGE), true) +
          `<div style="text-align:center;font-size:11px;color:#64748b;margin-top:8px">صفحة ${p + 1} من ${pages}</div>`;
        const canvas = await html2canvas(host, { scale: 2, backgroundColor: '#ffffff' });
        const iw = PW - 2 * M; let ih = canvas.height * iw / canvas.width;
        let w = iw; const maxH = PH - 2 * M;
        if (ih > maxH) { w = iw * maxH / ih; ih = maxH; }
        if (p > 0) pdf.addPage();
        pdf.addImage(canvas.toDataURL('image/png'), 'PNG', (PW - w) / 2, M, w, ih);
      }
      host.remove();
      pdf.save('جدول_الزيارات_' + fileStamp() + '.pdf');
      showToast('تم تحميل ملف PDF');
    } catch (e) { console.error(e); showToast(e.message || 'تعذّر إنشاء ملف PDF', 'error'); }
  };

  window.addEventListener('DOMContentLoaded', () => { try { ensureUi(); window.renderDistributionTable(); } catch (e) { console.error(e); } });
})();
