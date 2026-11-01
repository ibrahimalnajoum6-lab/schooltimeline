// ===== جدولة الزيارات حسب المنطقة والمسافة + جدول الزيارات الأسبوعي (A4) + تحميل Excel / PDF =====
// الأحد: ريف بعيد (50 كم فأكثر) | الإثنين: مدينة حلب | الثلاثاء-الخميس: ريف قريب (أقل من 50 كم)
// فريق الذكور وفريق الإناث يزوران في اليوم نفسه مدرستين من المنطقة نفسها ما أمكن.
(function () {
  const DAYS_AR = ['الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];
  const HEAD = ['التاريخ', 'اليوم', 'فريق الذكور', 'المدرسة', 'فريق الإناث', 'المدرسة '];
  const HEAD_BG = ['#5B9BD5', '#5B9BD5', '#5B9BD5', '#5B9BD5', '#FFD966', '#5B9BD5'];
  const BODY_BG = ['#BDD7EE', '#FFFFFF', '#DDEBF7', '#FFFFFF', '#FFD966', '#E2EFDA'];
  const LINE = '#9DC3E6';
  const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const COLL = new Intl.Collator('ar');

  // ---------- تواريخ (محلية، بلا انزياح المنطقة الزمنية) ----------
  const pad = n => String(n).padStart(2, '0');
  const iso = d => d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
  const fromIso = t => { const m = String(t || '').match(/^(\d{4})-(\d{2})-(\d{2})$/); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null; };
  const addDays = (d, k) => { const x = new Date(d); x.setDate(x.getDate() + k); return x; };
  const sundayOf = d => addDays(d, -d.getDay());
  const dm = d => pad(d.getDate()) + '/' + pad(d.getMonth() + 1);
  const dmy = d => dm(d) + '/' + d.getFullYear();

  function parseVisit(v) {
    const t = String(v || '').trim(), d = fromIso(t);
    if (d) return { key: t, date: dmy(d), day: DAYS_AR[d.getDay()], sort: t, week: iso(sundayOf(d)) };
    if (DAYS_AR.includes(t)) return { key: 'z-' + t, date: '', day: t, sort: '9' + DAYS_AR.indexOf(t), week: '' };
    return { key: 'zz', date: '', day: '', sort: 'zzz', week: '' };
  }

  // ---------- المنطقة والمسافة من config.js ----------
  const nk = s => String(s ?? '').trim().replace(/[أإآ]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي').replace(/\s+/g, ' ');
  let INFO_MAP = null;
  window.schoolInfo = function (name, fallbackZone) {
    if (!INFO_MAP) { INFO_MAP = {}; const o = (window.APP_CONFIG || {}).SCHOOL_INFO || {}; Object.keys(o).forEach(k => INFO_MAP[nk(k)] = o[k]); }
    const i = INFO_MAP[nk(name)];
    if (!i) return { zone: fallbackZone || 'غير محدد', km: 0, kind: 'near', known: false };
    const km = Number(i.km) || 0;
    return { zone: i.zone, km, kind: i.zone === 'المدينة' ? 'city' : km >= 50 ? 'far' : 'near', known: true };
  };

  const teamText = r => (r && r.supervisors && r.supervisors.length) ? r.supervisors.map(s => s.name).join('\n') : '';
  const zoneOf = r => schoolInfo(r.schoolName, r.zone).zone;
  const byZoneName = (a, b) => COLL.compare(zoneOf(a), zoneOf(b)) || COLL.compare(a.schoolName, b.schoolName);

  // ---------- بناء صفوف الجدول (مقابلة الذكور والإناث + رقم الأسبوع) ----------
  window.buildVisitsRows = function (filter) {
    const f = filter || {}, q = String(f.q || '').trim();
    const all = distributionResults || [];
    const starts = [...new Set(all.map(r => parseVisit(r.visitDate).week).filter(Boolean))].sort();
    const wno = {}; starts.forEach((k, i) => wno[k] = i + 1);
    const list = all.filter(r => {
      if (f.gender && f.gender !== 'ALL' && r.gender !== f.gender) return false;
      if (q && !(String(r.schoolName) + ' ' + zoneOf(r) + ' ' + teamText(r)).includes(q)) return false;
      return true;
    });
    const groups = {};
    list.forEach(r => {
      const p = parseVisit(r.visitDate);
      const g = groups[p.key] = groups[p.key] || { p, slots: new Map(), m: [], f: [] };
      const side = r.gender === 'ذكور' ? 'm' : 'f';
      if (r.slot != null) {
        const s = g.slots.get(r.slot) || {}; g.slots.set(r.slot, s);
        if (!s[side]) { s[side] = r; return; }
      }
      g[side].push(r);
    });
    const rows = [];
    Object.values(groups).sort((a, b) => a.p.sort < b.p.sort ? -1 : a.p.sort > b.p.sort ? 1 : 0).forEach(g => {
      g.m.sort(byZoneName); g.f.sort(byZoneName);
      const pairs = [...g.slots.entries()].sort((a, b) => a[0] - b[0]).map(e => e[1]);
      pairs.forEach(pr => { if (!pr.m && g.m.length) pr.m = g.m.shift(); if (!pr.f && g.f.length) pr.f = g.f.shift(); });
      while (g.m.length || g.f.length) pairs.push({ m: g.m.shift(), f: g.f.shift() });
      pairs.forEach(pr => rows.push({
        cells: [g.p.date, g.p.day, teamText(pr.m), pr.m ? pr.m.schoolName : '', teamText(pr.f), pr.f ? pr.f.schoolName : ''],
        mId: pr.m && pr.m.schoolId, fId: pr.f && pr.f.schoolId, week: wno[g.p.week] || 0, weekStart: g.p.week
      }));
    });
    return rows;
  };

  // تجميع الصفوف حسب الأسبوع (لكل أسبوع صفحة A4)
  window.groupVisitsByWeek = function (rows) {
    const map = new Map();
    rows.forEach(r => { const k = r.week; if (!map.has(k)) map.set(k, { no: k, start: r.weekStart, rows: [] }); map.get(k).rows.push(r); });
    return [...map.values()].sort((a, b) => (a.no || 999) - (b.no || 999)).map(g => {
      const s = fromIso(g.start);
      g.label = g.no && s ? 'الأسبوع ' + g.no + ' — من ' + dm(s) + ' إلى ' + dmy(addDays(s, 4)) : 'زيارات غير مؤرخة';
      g.sheet = g.no ? 'الأسبوع ' + g.no : 'غير مؤرخة';
      return g;
    });
  };

  // ---------- العرض على الشاشة ----------
  const CSS = `
  #vt-wrap{background:#fff;border:1px solid #e2e8f0;border-radius:16px;overflow:hidden}
  .vt-scroll{overflow-x:auto}
  table.vt{width:100%;min-width:720px;border-collapse:collapse;direction:rtl;font-family:'Cairo',sans-serif;font-size:13px}
  table.vt th{color:#fff;font-weight:700;padding:8px;text-align:center;border:1px solid ${LINE};white-space:nowrap}
  table.vt td{padding:7px 8px;text-align:center;border:1px solid ${LINE};vertical-align:middle;line-height:1.7;white-space:pre-line}
  table.vt td.vt-link{cursor:pointer}
  table.vt td.vt-link:hover{text-decoration:underline}
  table.vt th::after{content:'\\25BE';display:inline-block;margin-inline-start:6px;background:#fff;color:#555;font-size:9px;padding:0 3px;border:1px solid #999;line-height:14px}
  .vt-bar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;justify-content:space-between;padding:12px 14px;border-bottom:1px solid #e2e8f0;background:#f8fafc}
  .vt-wh{padding:9px 14px;font-weight:800;font-size:13px;color:#0f2f5c;background:#eaf2fb;border-bottom:1px solid ${LINE};border-top:1px solid ${LINE}}
  .vt-btn{display:inline-flex;align-items:center;gap:6px;font-size:12px;font-weight:700;padding:8px 14px;border-radius:12px;border:1px solid #cbd5e1;background:#fff;color:#0f172a;cursor:pointer}
  .vt-btn:hover{background:#f1f5f9}
  .vt-btn.ex{color:#15803d;border-color:#86efac;background:#f0fdf4}
  .vt-btn.pd{color:#b91c1c;border-color:#fca5a5;background:#fef2f2}
  `;

  function tableHtml(rows) {
    const th = HEAD.map((h, i) => `<th style="background:${HEAD_BG[i]}">${esc(h)}</th>`).join('');
    const body = rows.map(r => '<tr>' + r.cells.map((c, i) => {
      const id = i === 3 ? r.mId : i === 5 ? r.fId : null;
      const link = id ? ` class="vt-link" onclick="openTeamBuilder('${id}')" title="تحديد/تعديل فريق الزيارة"` : '';
      return `<td${link} style="background:${BODY_BG[i]}">${esc(c)}</td>`;
    }).join('') + '</tr>').join('');
    return `<table class="vt"><thead><tr>${th}</tr></thead><tbody>${body}</tbody></table>`;
  }

  function ensureUi() {
    if (!document.getElementById('vt-style')) { const st = document.createElement('style'); st.id = 'vt-style'; st.textContent = CSS; document.head.appendChild(st); }
    if (document.getElementById('vt-wrap')) return;
    const view = document.getElementById('view-distribution'); if (!view) return;
    const cards = document.getElementById('distribution-cards'); if (cards) cards.style.display = 'none';
    const old = view.querySelector('.hidden.md\\:block'); if (old) old.style.display = 'none';
    const wrap = document.createElement('div'); wrap.id = 'vt-wrap';
    wrap.innerHTML = `
      <div class="vt-bar no-print" id="vt-bar">
        <div style="font-size:12px;color:#475569;font-weight:600">جدول الزيارات الميدانية — صفحة A4 لكل أسبوع</div>
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
    const weeks = window.groupVisitsByWeek(window.buildVisitsRows({ gender: g, q }));
    if (window.__vtKey !== g + '|' + q) { window.__vtKey = g + '|' + q; window.__vtLim = 3; }
    if (!weeks.length) { box.innerHTML = '<div style="padding:32px;text-align:center;color:#94a3b8;font-size:13px">لا يوجد جدول زيارات حالياً. اضغط «جدولة الزيارات» لإنشائه</div>'; return; }
    const shown = weeks.slice(0, window.__vtLim);
    const more = weeks.length > shown.length ? `<div style="padding:12px;text-align:center"><button class="vt-btn" onclick="window.__vtLim+=3;renderDistributionTable()">عرض المزيد من الأسابيع (${weeks.length - shown.length})</button></div>` : '';
    box.innerHTML = shown.map(w => `<div class="vt-wh">${esc(w.label)} <span style="font-weight:600;color:#64748b">· ${w.rows.length} صفاً</span></div><div class="vt-scroll">${tableHtml(w.rows)}</div>`).join('') + more;
  };

  // ---------- أدوات مشتركة ----------
  const fileStamp = () => iso(new Date());
  function saveBlob(blob, name) {
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 8000);
  }
  function currentWeeks() {
    const rows = window.buildVisitsRows({});
    if (!rows.length) { showToast('لا يوجد جدول زيارات للتصدير. اضغط «جدولة الزيارات» أولاً', 'error'); return null; }
    return window.groupVisitsByWeek(rows);
  }

  // ---------- Excel: ورقة A4 لكل أسبوع، تُبنى مباشرة بلا أي مكتبة ----------
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = b => { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  function zipStore(files) {
    const enc = new TextEncoder(), parts = [], central = []; let off = 0;
    const u16 = v => [v & 255, (v >> 8) & 255], u32 = v => [v & 255, (v >> 8) & 255, (v >> 16) & 255, (v >>> 24) & 255];
    files.forEach(([name, text]) => {
      const nb = enc.encode(name), db = enc.encode(text), crc = crc32(db);
      const head = new Uint8Array([0x50, 0x4b, 3, 4, ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21), ...u32(crc), ...u32(db.length), ...u32(db.length), ...u16(nb.length), ...u16(0)]);
      parts.push(head, nb, db);
      central.push(new Uint8Array([0x50, 0x4b, 1, 2, ...u16(20), ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21), ...u32(crc), ...u32(db.length), ...u32(db.length), ...u16(nb.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(off)]), nb);
      off += head.length + nb.length + db.length;
    });
    let csz = 0; central.forEach(c => csz += c.length);
    const end = new Uint8Array([0x50, 0x4b, 5, 6, 0, 0, 0, 0, ...u16(files.length), ...u16(files.length), ...u32(csz), ...u32(off), 0, 0]);
    return new Blob([...parts, ...central, end], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }
  const xe = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  const XH = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

  function buildXlsx(weeks) {
    const hex = h => 'FF' + h.replace('#', '').toUpperCase();
    const fills = [...new Set([...HEAD_BG, ...BODY_BG])];
    const fillXml = '<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
      fills.map(c => `<fill><patternFill patternType="solid"><fgColor rgb="${hex(c)}"/><bgColor indexed="64"/></patternFill></fill>`).join('');
    const fid = c => 2 + fills.indexOf(c);
    const al = '<alignment horizontal="center" vertical="center" wrapText="1"/>';
    // 0 افتراضي | 1..6 ترويسة | 7..12 جسم | 13 عنوان الأسبوع
    const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
      .concat(HEAD_BG.map(c => `<xf numFmtId="0" fontId="1" fillId="${fid(c)}" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">${al}</xf>`))
      .concat(BODY_BG.map(c => `<xf numFmtId="49" fontId="0" fillId="${fid(c)}" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1">${al}</xf>`))
      .concat([`<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1">${al}</xf>`]);
    const styles = XH + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<fonts count="3"><font><sz val="11"/><name val="Arial"/></font><font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Arial"/></font><font><b/><sz val="14"/><color rgb="FF0F2F5C"/><name val="Arial"/></font></fonts>' +
      `<fills count="${2 + fills.length}">${fillXml}</fills>` +
      `<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="${hex(LINE)}"/></left><right style="thin"><color rgb="${hex(LINE)}"/></right><top style="thin"><color rgb="${hex(LINE)}"/></top><bottom style="thin"><color rgb="${hex(LINE)}"/></bottom><diagonal/></border></borders>` +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
    const col = i => 'ABCDEF'[i];
    const cell = (r, i, v, st) => `<c r="${col(i)}${r}" s="${st}" t="inlineStr"><is><t xml:space="preserve">${xe(v)}</t></is></c>`;
    const widths = [12, 10, 26, 32, 26, 32];

    const sheetXml = w => {
      let sd = `<row r="1" ht="30" customHeight="1">${cell(1, 0, 'جدول الزيارات الميدانية — ' + w.label, 13)}</row>`;
      sd += `<row r="2" ht="24" customHeight="1">${HEAD.map((h, i) => cell(2, i, h, 1 + i)).join('')}</row>`;
      w.rows.forEach((r, k) => {
        const ln = Math.max(1, ...r.cells.map(t => String(t).split('\n').length));
        sd += `<row r="${k + 3}" ht="${Math.max(22, ln * 16)}" customHeight="1">${r.cells.map((v, i) => cell(k + 3, i, v, 7 + i)).join('')}</row>`;
      });
      return XH + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' +
        '<sheetViews><sheetView rightToLeft="1" workbookViewId="0"><pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
        `<cols>${widths.map((x, i) => `<col min="${i + 1}" max="${i + 1}" width="${x}" customWidth="1"/>`).join('')}</cols>` +
        `<sheetData>${sd}</sheetData><autoFilter ref="A2:F${w.rows.length + 2}"/><mergeCells count="1"><mergeCell ref="A1:F1"/></mergeCells>` +
        '<pageMargins left="0.3" right="0.3" top="0.4" bottom="0.4" header="0.2" footer="0.2"/><pageSetup paperSize="9" orientation="portrait" fitToWidth="1" fitToHeight="1"/></worksheet>';
    };

    const ct = XH + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      weeks.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') + '</Types>';
    const rels = XH + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
    const wb = XH + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' +
      weeks.map((w, i) => `<sheet name="${xe(w.sheet)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') + '</sheets></workbook>';
    const wbRels = XH + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      weeks.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') +
      `<Relationship Id="rId${weeks.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
    return zipStore([['[Content_Types].xml', ct], ['_rels/.rels', rels], ['xl/workbook.xml', wb], ['xl/_rels/workbook.xml.rels', wbRels], ['xl/styles.xml', styles]]
      .concat(weeks.map((w, i) => ['xl/worksheets/sheet' + (i + 1) + '.xml', sheetXml(w)])));
  }
  window.buildVisitsXlsx = buildXlsx;
  window.downloadVisitsExcel = function () {
    const weeks = currentWeeks(); if (!weeks) return;
    try { saveBlob(buildXlsx(weeks), 'جدول_الزيارات_' + fileStamp() + '.xlsx'); showToast('تم تحميل ملف Excel (ورقة لكل أسبوع)'); }
    catch (e) { console.error(e); showToast('تعذّر إنشاء ملف Excel', 'error'); }
  };

  // ---------- PDF: صفحة A4 عمودية لكل أسبوع (تُصغَّر تلقائياً لتكفي صفحة واحدة) ----------
  window.visitsPrintHtml = function (weeks) {
    const th = HEAD.map((h, i) => `<th style="background:${HEAD_BG[i]}">${esc(h)}</th>`).join('');
    const page = w => `<section class="pg"><div class="in"><h1>جدول الزيارات الميدانية</h1><h2>${esc(w.label)}</h2>
<table><colgroup><col style="width:11%"><col style="width:8%"><col style="width:22%"><col style="width:19%"><col style="width:22%"><col style="width:18%"></colgroup><thead><tr>${th}</tr></thead><tbody>${w.rows.map(r => '<tr>' + r.cells.map((c, i) => `<td style="background:${BODY_BG[i]}">${esc(c)}</td>`).join('') + '</tr>').join('')}</tbody></table></div></section>`;
    return `<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><title>جدول الزيارات</title>
<link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;700&display=swap" rel="stylesheet">
<style>@page{size:A4 portrait;margin:8mm}*{-webkit-print-color-adjust:exact;print-color-adjust:exact;box-sizing:border-box}
html,body{margin:0;padding:0}body{font-family:'Cairo','Segoe UI',Tahoma,sans-serif;direction:rtl}
.pg{height:279mm;overflow:hidden;page-break-after:always;break-after:page}.pg:last-child{page-break-after:auto;break-after:auto}
h1{text-align:center;font-size:17px;margin:0 0 2px;color:#0f2f5c}h2{text-align:center;font-size:13px;margin:0 0 8px;color:#334155}
table{width:100%;table-layout:fixed;border-collapse:collapse;font-size:11px}tr{page-break-inside:avoid}
th{color:#fff;padding:6px 4px;border:1px solid ${LINE};font-size:12px}
td{padding:5px 4px;text-align:center;border:1px solid ${LINE};white-space:pre-line;line-height:1.55}</style></head><body>
${weeks.map(page).join('')}
<script>function fit(){document.querySelectorAll('.pg').forEach(function(p){var i=p.firstElementChild,z=1;i.style.zoom=1;
while(i.getBoundingClientRect().height>p.clientHeight-2&&z>0.3){z-=0.03;i.style.zoom=z;}});}
window.addEventListener('load',function(){fit();if(document.fonts&&document.fonts.ready)document.fonts.ready.then(fit);});</script></body></html>`;
  };
  window.downloadVisitsPdf = function () {
    const weeks = currentWeeks(); if (!weeks) return;
    const fr = document.createElement('iframe');
    fr.style.cssText = 'position:fixed;right:0;bottom:0;width:210mm;height:297mm;border:0;visibility:hidden';
    document.body.appendChild(fr);
    fr.srcdoc = window.visitsPrintHtml(weeks);
    fr.onload = () => {
      setTimeout(() => {
        try { fr.contentWindow.fit && fr.contentWindow.fit(); fr.contentWindow.focus(); fr.contentWindow.print(); }
        catch (e) { console.error(e); showToast('تعذّر فتح نافذة الطباعة', 'error'); }
        setTimeout(() => fr.remove(), 60000);
      }, 900);
    };
    showToast('في نافذة الطباعة اختر «حفظ كـ PDF»', 'info');
  };

  // ---------- الجدولة الذكية: مسافة + منطقة + أسابيع ----------
  window.runIntelligentDistribution = function () {
    if (!supervisors.length || !schools.length) { showToast('الرجاء توفير بيانات الموجهين والمدارس أولاً', 'error'); return; }
    const sd = fromIso(document.getElementById('plan-start-date').value), ed = fromIso(document.getElementById('plan-end-date').value);
    if (!sd || !ed) { showToast('يرجى تحديد تاريخ بداية ونهاية خطة التوزيع!', 'error'); return; }
    if (ed < sd) { showToast('تاريخ النهاية يجب أن يكون بعد تاريخ البداية!', 'error'); return; }
    const tv = document.getElementById('team-size-per-school').value, allMode = tv === 'ALL', teamSize = allMode ? Infinity : (parseInt(tv) || 2);

    const work = []; for (let d = new Date(sd); d <= ed; d = addDays(d, 1)) if (d.getDay() <= 4) work.push(new Date(d)); // الأحد..الخميس
    if (!work.length) { showToast('لا توجد أيام دوام (الأحد–الخميس) ضمن الفترة المحددة', 'error'); return; }
    const pick = days => { const x = work.filter(d => days.includes(d.getDay())); return x.length ? x : work; };
    const DAYSET = { far: pick([0]), city: pick([1]), near: pick([2, 3, 4]) };

    const manual = distributionResults.filter(d => d.manual), manualIds = new Set(manual.map(d => d.schoolId));
    const cat = { far: [], city: [], near: [] }; let unknown = 0;
    schools.filter(s => !manualIds.has(s.id)).forEach(s => {
      const info = schoolInfo(s.name, s.zone); if (!info.known) unknown++;
      cat[info.kind].push({ s, info });
    });

    // مدرسة بنين + مدرسة بنات من المنطقة نفسها؛ والفائض يُزاوَج حسب أقرب مسافة
    const byName = (a, b) => COLL.compare(a.s.name, b.s.name);
    const byKm = (a, b) => a.info.km - b.info.km || COLL.compare(a.info.zone, b.info.zone) || byName(a, b);
    function makeSlots(items) {
      const z = {}; items.forEach(x => { (z[x.info.zone] = z[x.info.zone] || { B: [], G: [] })[x.s.gender === 'ذكور' ? 'B' : 'G'].push(x); });
      const mk = o => Math.min(...o.B.concat(o.G).map(x => x.info.km));
      const zones = Object.keys(z).sort((a, b) => mk(z[a]) - mk(z[b]) || COLL.compare(a, b));
      const slots = [], lb = [], lg = [];
      zones.forEach(k => {
        const { B, G } = z[k]; B.sort(byName); G.sort(byName);
        const c = Math.min(B.length, G.length);
        for (let i = 0; i < c; i++) slots.push({ m: B[i], f: G[i] });
        lb.push(...B.slice(c)); lg.push(...G.slice(c));
      });
      lb.sort(byKm); lg.sort(byKm);
      for (let i = 0; i < Math.max(lb.length, lg.length); i++) slots.push({ m: lb[i], f: lg[i] });
      return slots;
    }

    const load = {}; supervisors.forEach(s => load[s.id] = 0);
    const busy = {};
    const hasSpec = (sup, school) => window.supHasSpec ? supHasSpec(sup, school.targetSpecialty) : sup.specialty === school.targetSpecialty;
    function teamFor(school, dateStr) {
      const pool = supervisors.filter(s => s.gender === (school.gender === 'ذكور' ? 'ذكر' : 'أنثى'));
      let t = [];
      if (allMode) t = pool.slice().sort((a, b) => hasSpec(b, school) - hasSpec(a, school));
      else {
        const b = busy[dateStr] = busy[dateStr] || new Set(), used = new Set();
        const lightest = arr => arr.sort((x, y) => load[x.id] - load[y.id])[0];
        const p = lightest(pool.filter(s => hasSpec(s, school) && !b.has(s.id)));
        if (p) { t.push(p); used.add(p.specialty); }
        while (t.length < teamSize) {
          const n = lightest(pool.filter(s => !t.includes(s) && !b.has(s.id) && !used.has(s.specialty))) || lightest(pool.filter(s => !t.includes(s) && !b.has(s.id)));
          if (!n) break; t.push(n); used.add(n.specialty);
        }
        t.forEach(s => b.add(s.id));
      }
      t.forEach(s => load[s.id]++);
      return t.map(s => ({ id: s.id, name: s.name, specialty: s.specialty, phone: s.phone, gender: s.gender }));
    }

    // زيارة واحدة فقط لكل فريق في اليوم: الأحد=بعيد، الإثنين=مدينة، الثلاثاء-الخميس=قريب
    // (إن فرغت قائمة نوعٍ ما تُملأ أيامه بما تبقى من الأنواع الأخرى كي لا تبقى أيام فارغة)
    const Q = { far: makeSlots(cat.far), city: makeSlots(cat.city), near: makeSlots(cat.near) };
    const own = d => d === 0 ? 'far' : d === 1 ? 'city' : 'near';
    const left = () => Q.far.length + Q.city.length + Q.near.length;
    const out = []; let slotNo = 0, last = null, spilled = 0;
    for (let d = new Date(sd), guard = 0; left() > 0 && guard < 1500; d = addDays(d, 1), guard++) {
      if (d.getDay() > 4) continue;
      let k = own(d.getDay());
      if (!Q[k].length) { k = ['far', 'near', 'city'].find(x => Q[x].length); spilled++; }
      const sl = Q[k].shift(), date = iso(d); slotNo++; last = new Date(d);
      [sl.m, sl.f].forEach(x => {
        if (!x) return;
        out.push({ schoolId: x.s.id, schoolName: x.s.name, gender: x.s.gender, stage: x.s.stage, sections: x.s.sections, zone: x.info.zone, km: x.info.km,
          visitDate: date, slot: slotNo, supervisors: teamFor(x.s, date) });
      });
    }

    distributionResults = manual.concat(out);
    persistData(); renderAllViews(); switchView('distribution');
    const weeks = new Set(out.map(r => parseVisit(r.visitDate).week)).size;
    let msg = `تمت جدولة ${out.length} مدرسة على ${weeks} أسابيع بزيارة واحدة في اليوم: الأحد للريف البعيد، الإثنين لمدينة حلب، الثلاثاء–الخميس للريف القريب`;
    if (last && last > ed) msg += ` — امتدت الخطة حتى ${dmy(last)} بعد تاريخ النهاية لأن الزيارة واحدة في اليوم`;
    if (spilled) msg += ` — في ${spilled} يوم زيارة من نوعٍ آخر لأن قائمة نوع ذلك اليوم انتهت`;
    if (unknown) msg += ` — ${unknown} مدرسة بلا منطقة/مسافة في config.js عُدّت ريفاً قريباً`;
    showToast(msg, (unknown || spilled || (last && last > ed)) ? 'info' : 'success');
  };

  window.addEventListener('DOMContentLoaded', () => { try { ensureUi(); window.renderDistributionTable(); } catch (e) { console.error(e); } });
})();
