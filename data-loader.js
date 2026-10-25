// تحميل البيانات من Google Sheets + تبويب فلترة برنامج الدوام
(function () {
  const CFG = window.APP_CONFIG || {};
  const DAYS = ['الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'السبت'];
  const STAGES = ['السابع', 'الثامن', 'التاسع', 'العاشر', 'الحادي عشر', 'البكالوريا'];
  let FLAT = [], FILTERED = [], LIMIT = 300;

  const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const n = s => String(s ?? '').trim().replace(/[أإآ]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي').replace(/\s+/g, ' ');
  const bare = s => n(s).replace(/(^|\s)ال/g, '$1');
  const ord = t => /اول/.test(t) ? 1 : /ثاني/.test(t) ? 2 : /ثالث/.test(t) ? 3 : /رابع/.test(t) ? 4 : /خامس/.test(t) ? 5 : /سادس/.test(t) ? 6 : /سابع/.test(t) ? 7 : 0;

  function normDay(v) {
    const t = n(v);
    return /احد/.test(t) ? DAYS[0] : /اثنين/.test(t) ? DAYS[1] : /ثلاثاء/.test(t) ? DAYS[2] :
      /اربعاء/.test(t) ? DAYS[3] : /خميس/.test(t) ? DAYS[4] : /سبت/.test(t) ? DAYS[5] : String(v || '').trim();
  }
  function normPeriod(v) {
    const m = String(v ?? '').match(/\d+/);
    return m ? +m[0] : ord(n(v));
  }
  function normGrade(gradeRaw, sectionRaw) {
    const g = bare(gradeRaw), sRaw = String(sectionRaw ?? '').trim();
    let stage = '';
    if (/^\d+$/.test(g)) stage = { 7: STAGES[0], 8: STAGES[1], 9: STAGES[2], 10: STAGES[3], 11: STAGES[4], 12: STAGES[5] }[+g] || '';
    else if (/بك.{0,2}ور|ثاني عشر|ثالث ثانوي/.test(g)) stage = STAGES[5];
    else if (/حادي عشر|ثاني ثانوي/.test(g)) stage = STAGES[4];
    else if (/عاشر|اول ثانوي/.test(g)) stage = STAGES[3];
    else if (/تاسع|ثالث اعدادي/.test(g)) stage = STAGES[2];
    else if (/ثامن|ثاني اعدادي/.test(g)) stage = STAGES[1];
    else if (/سابع|اول اعدادي/.test(g)) stage = STAGES[0];
    const all = g + ' ' + bare(sRaw);
    const track = /ادبي|\sد$/.test(all) ? 'أدبي' : /علمي|\sع$/.test(all) ? 'علمي' : '';
    let sec = 0;
    if (/^\d+$/.test(sRaw) && +sRaw > 0) sec = +sRaw;
    else if (!/وحيد|ادبي|علمي/.test(bare(sRaw))) sec = ord(bare(sRaw));
    if (!sec && stage && !/^\d+$/.test(g)) { const m = g.match(/(\d)\s*$/); if (m) sec = +m[1]; else if (/شعبه/.test(g)) sec = ord(g); }
    const label = stage ? [stage, track, sec ? 'شعبة ' + sec : ''].filter(Boolean).join(' ') : (String(gradeRaw || '').trim() || String(sectionRaw || '').trim() || '—');
    return { stage, label, sec, track };
  }

  // مفاتيح سريعة: مقارن عربي واحد بدل localeCompare (أبطأ بعشرات المرات)
  const COLL = new Intl.Collator('ar');
  const IS_MOB = () => window.matchMedia('(max-width: 767px)').matches;

  // محلل CSV خفيف (بديل عن SheetJS الثقيل على الجوال)
  function parseCSV(t) {
    if (t.charCodeAt(0) === 0xFEFF) t = t.slice(1);
    const rows = []; let row = [], f = '', q = false, i = 0; const L = t.length;
    while (i < L) {
      const c = t[i];
      if (q) {
        if (c === '"') { if (t[i + 1] === '"') { f += '"'; i += 2; continue; } q = false; i++; continue; }
        f += c; i++; continue;
      }
      if (c === '"') { q = true; i++; }
      else if (c === ',') { row.push(f); f = ''; i++; }
      else if (c === '\n' || c === '\r') { if (c === '\r' && t[i + 1] === '\n') i++; row.push(f); f = ''; rows.push(row); row = []; i++; }
      else { f += c; i++; }
    }
    if (f !== '' || row.length) { row.push(f); rows.push(row); }
    return rows;
  }
  function csvToObjects(text, label) {
    const all = parseCSV(text).filter(r => r.some(x => x !== ''));
    if (all.length < 2) throw new Error(label + ' — الورقة فارغة');
    const seen = {};
    const head = all[0].map(h => {
      let k = String(h).replace(/[\u200B-\u200D\uFEFF]/g, '').trim() || 'عمود';
      if (seen[k]) k = k + '_' + (seen[k]++); else seen[k] = 1;
      return k;
    });
    const out = new Array(all.length - 1);
    for (let r = 1; r < all.length; r++) {
      const o = {}, row = all[r];
      for (let c = 0; c < head.length; c++) o[head[c]] = row[c] === undefined ? '' : row[c].trim();
      out[r - 1] = o;
    }
    return out;
  }

  async function fetchText(url, label) {
    let text;
    try {
      const res = await fetch(url, { cache: 'no-store' });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      text = await res.text();
    } catch (e) { throw new Error(label + ' — فشل الاتصال بالرابط (' + e.message + ')'); }
    if (/^\s*<(!doctype|html)/i.test(text)) throw new Error(label + ' — الرابط يعيد صفحة ويب لا CSV (غير منشور أو ليس رابط النشر على الويب)');
    return text;
  }

  // ---- كاش محلي: يعرض آخر بيانات فوراً ثم يحدّثها في الخلفية ----
  const CK = 'sup_app_data_v1';
  const readCache = () => { try { const j = JSON.parse(localStorage.getItem(CK) || 'null'); return j && j.sup && j.sch && j.u === CFG.SUPERVISORS_CSV + '|' + CFG.SCHEDULE_CSV ? j : null; } catch (e) { return null; } };
  const saveCache = d => { try { localStorage.setItem(CK, JSON.stringify({ sup: d.sup, sch: d.sch, u: CFG.SUPERVISORS_CSV + '|' + CFG.SCHEDULE_CSV })); } catch (e) { /* الكاش اختياري */ } };
  async function fetchBoth() {
    const res = await Promise.allSettled([fetchText(CFG.SUPERVISORS_CSV, 'ورقة الموجهين'), fetchText(CFG.SCHEDULE_CSV, 'ورقة الجدول')]);
    const errs = res.filter(r => r.status === 'rejected').map(r => r.reason.message);
    if (errs.length) throw new Error(errs.join(' | '));
    return { sup: res[0].value, sch: res[1].value };
  }

  function buildData(d) {
    const supRows = csvToObjects(d.sup, 'ورقة الموجهين'), schRows = csvToObjects(d.sch, 'ورقة الجدول');
    supervisors = supRows.filter(r => r['اسم الموجه']).map((r, i) => ({
      id: 'sup-' + i, name: String(r['اسم الموجه']).trim(),
      gender: /أنث|انث/.test(String(r['الجنس'])) ? 'أنثى' : 'ذكر',
      specialty: String(r['التخصص'] || '').trim(), phone: String(r['الهاتف'] || '-').trim()
    }));
    if (!schRows[0] || !('المدرسة' in schRows[0])) throw new Error('ورقة الجدول: لا يوجد عمود «المدرسة» (الأعمدة الموجودة: ' + Object.keys(schRows[0] || {}).join('، ') + ')');
    if (!supRows[0] || !('اسم الموجه' in supRows[0])) throw new Error('ورقة الموجهين: لا يوجد عمود «اسم الموجه» (الأعمدة الموجودة: ' + Object.keys(supRows[0] || {}).join('، ') + ')');
    const names = [...new Set(schRows.map(r => String(r['المدرسة'] || '').trim()).filter(Boolean))].sort(COLL.compare);
    const idOf = {}; schools = []; schedules = []; schedules = {}; FLAT = [];
    names.forEach((nm, i) => {
      const ov = (CFG.SCHOOL_GENDER_OVERRIDES || {})[nm];
      idOf[nm] = 'sch-' + i;
      schools.push({ id: 'sch-' + i, name: nm, gender: ov || (/بنات|اناث|إناث/.test(nm) ? 'إناث' : 'ذكور'),
        stage: 'مشترك (7 - بكالوريا)', sections: 0, zone: 'قطاع المركز', targetSpecialty: 'الفقه وأصوله' });
      schedules['sch-' + i] = [];
    });
    schRows.forEach(r => {
      const nm = String(r['المدرسة'] || '').trim(); if (!nm) return;
      const { stage, label, sec, track } = normGrade(r['الصف'], r['الشعبة']);
      const row = { grade: label, stage, sec, track, day: normDay(r['اليوم']), period: normPeriod(r['الحصة']),
        subject: String(r['المادة'] || '').trim(), teacher: String(r['اسم المدرس'] || '').trim() };
      schedules[idOf[nm]].push(row);
      FLAT.push({ school: nm, ...row, _k: n(nm + ' ' + row.day + ' ' + label + ' ' + row.subject + ' ' + row.teacher), _s: n(row.subject), _t: n(row.teacher) });
    });
    schools.forEach(s => { s.sections = new Set(schedules[s.id].map(x => x.grade)).size; });
    distributionResults = distributionResults.filter(d => idOf[d.schoolName] === d.schoolId);
    window.SHEETS_MODE = true;
    // الترتيب مرة واحدة هنا بدل كل عملية فلترة
    FLAT.sort((a, b) => COLL.compare(a.school, b.school) || DAYS.indexOf(a.day) - DAYS.indexOf(b.day) || a.period - b.period);
    buildFilterOptions(); applyFilters();
  }
  window.loadFromSheets = async function () { const d = await fetchBoth(); buildData(d); saveCache(d); };

  // ---------- تبويب الفلترة ----------
  const inp = 'w-full text-xs border border-slate-300 rounded-xl px-3 py-2 bg-slate-50';
  const field = (l, el) => `<div><label class="block text-[11px] font-bold text-slate-600 mb-1">${l}</label>${el}</div>`;
  function injectTab() {
    document.getElementById('tab-btn-import').insertAdjacentHTML('beforebegin',
      `<button onclick="switchView('filter')" id="tab-btn-filter" class="px-3.5 py-2 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition flex items-center gap-2"><i class="fa-solid fa-filter"></i><span>فلترة برنامج الدوام</span></button>`);
    document.getElementById('view-schedules').insertAdjacentHTML('afterend', `
    <div id="view-filter" class="hidden space-y-4">
      <div class="bg-white p-4 rounded-2xl border border-slate-200 space-y-3">
        <div class="flex flex-wrap items-center justify-between gap-2">
          <h3 class="text-base font-bold text-slate-900">فلترة برنامج الدوام الموحد</h3>
          <div class="flex gap-2">
            <button onclick="resetFilters()" class="text-xs bg-slate-100 hover:bg-slate-200 px-3 py-2 rounded-xl font-bold">مسح الفلاتر</button>
            <button onclick="exportFiltered()" class="text-xs bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-2 rounded-xl font-bold"><i class="fa-solid fa-file-excel ml-1"></i>تصدير النتائج</button>
          </div>
        </div>
        <div class="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-3">
          ${field('المدرسة', `<select id="f-school" class="${inp}"></select>`)}
          ${field('اليوم', `<select id="f-day" class="${inp}"></select>`)}
          ${field('الحصة', `<select id="f-period" class="${inp}"></select>`)}
          ${field('الصف', `<select id="f-stage" class="${inp}"></select>`)}
          ${field('الشعبة / الفرع', `<select id="f-section" class="${inp}"></select>`)}
          ${field('المادة', `<input id="f-subject" list="dl-subject" placeholder="ابحث..." class="${inp}"><datalist id="dl-subject"></datalist>`)}
          ${field('المدرس', `<input id="f-teacher" list="dl-teacher" placeholder="ابحث..." class="${inp}"><datalist id="dl-teacher"></datalist>`)}
          ${field('بحث نصي', `<input id="f-text" placeholder="أي كلمة..." class="${inp}">`)}
        </div>
        <label class="inline-flex items-center gap-2 text-xs text-slate-600"><input type="checkbox" id="f-noteacher"> الحصص التي بلا اسم مدرس فقط</label>
      </div>
      <div id="filter-summary" class="text-xs font-bold text-emerald-800 bg-emerald-50 border border-emerald-200 rounded-xl px-4 py-2"></div>
      <div id="filter-cards" class="md:hidden space-y-2"></div>
      <div class="hidden md:block bg-white rounded-2xl border border-slate-200 overflow-x-auto custom-scrollbar">
        <table class="w-full text-xs text-right">
          <thead class="bg-slate-50 text-slate-600"><tr><th class="p-3">المدرسة</th><th class="p-3">اليوم</th><th class="p-3">الحصة</th><th class="p-3">الصف</th><th class="p-3">المادة</th><th class="p-3">المدرس</th></tr></thead>
          <tbody id="filter-body" class="divide-y divide-slate-100"></tbody>
        </table>
      </div>
      <div class="text-center"><button id="filter-more" onclick="LIMIT_MORE()" class="hidden text-xs bg-slate-800 text-white px-4 py-2 rounded-xl font-bold">عرض المزيد</button></div>
    </div>`);
    const grid = document.querySelector('#view-filter .grid'), kids = [...grid.children];
    [0, 5].forEach(i => kids[i].classList.add('col-span-2', 'md:col-span-1'));
    [6, 7].forEach(i => kids[i].classList.add('f-extra', 'hidden', 'md:block', 'col-span-2', 'md:col-span-1'));
    const nt = document.getElementById('f-noteacher').parentElement; nt.classList.add('f-extra', 'hidden', 'md:inline-flex');
    grid.insertAdjacentHTML('afterend', '<button id="f-toggle" class="md:hidden w-full text-xs font-bold text-emerald-700 bg-emerald-50 rounded-xl py-2.5">فلاتر إضافية (المدرس، بحث نصي...) ▾</button>');
    document.getElementById('f-toggle').addEventListener('click', e => {
      document.querySelectorAll('#view-filter .f-extra').forEach(x => x.classList.toggle('hidden'));
      e.target.textContent = document.querySelector('#view-filter .f-extra').classList.contains('hidden') ? 'فلاتر إضافية (المدرس، بحث نصي...) ▾' : 'إخفاء الفلاتر الإضافية ▴';
    });
    ['f-school', 'f-day', 'f-period', 'f-stage', 'f-section', 'f-noteacher'].forEach(id => document.getElementById(id).addEventListener('change', applyFilters));
    document.getElementById('f-school').addEventListener('change', refreshLists);
    let t; ['f-subject', 'f-teacher', 'f-text'].forEach(id => document.getElementById(id).addEventListener('input', () => { clearTimeout(t); t = setTimeout(applyFilters, 200); }));
  }
  const opts = (first, arr) => `<option value="">${first}</option>` + arr.map(v => `<option>${esc(v)}</option>`).join('');
  function buildFilterOptions() {
    const uniq = k => [...new Set(FLAT.map(r => r[k]).filter(Boolean))].sort((a, b) => COLL.compare(String(a), String(b)));
    document.getElementById('f-school').innerHTML = opts('كل المدارس', uniq('school'));
    document.getElementById('f-day').innerHTML = opts('كل الأيام', DAYS);
    document.getElementById('f-period').innerHTML = opts('كل الحصص', [1, 2, 3, 4, 5, 6, 7]);
    document.getElementById('f-stage').innerHTML = opts('كل الصفوف', STAGES);
    document.getElementById('f-section').innerHTML = '<option value="">كل الشعب</option>' + [1, 2, 3, 4, 5].map(i => `<option value="${i}">شعبة ${i}</option>`).join('') + '<option>أدبي</option><option>علمي</option>';
    refreshLists();
  }
  function refreshLists() {
    const sc = document.getElementById('f-school').value, rows = sc ? FLAT.filter(r => r.school === sc) : FLAT;
    const uniq = k => [...new Set(rows.map(r => r[k]).filter(Boolean))].sort((a, b) => COLL.compare(String(a), String(b)));
    document.getElementById('dl-subject').innerHTML = uniq('subject').map(v => `<option value="${esc(v)}">`).join('');
    document.getElementById('dl-teacher').innerHTML = uniq('teacher').map(v => `<option value="${esc(v)}">`).join('');
  }
  window.applyFilters = function () {
    LIMIT = IS_MOB() ? 40 : 300;
    const v = id => document.getElementById(id).value.trim();
    const sc = v('f-school'), dy = v('f-day'), pr = v('f-period'), st = v('f-stage'), se = v('f-section'), tx = n(v('f-text')), sb = n(v('f-subject')), tc = n(v('f-teacher'));
    const no = document.getElementById('f-noteacher').checked;
    FILTERED = FLAT.filter(r => (!sc || r.school === sc) && (!dy || r.day === dy) && (!pr || r.period == pr) && (!st || r.stage === st) && (!se || (/^\d$/.test(se) ? r.sec == se : r.track === se)) &&
      (!tx || r._k.includes(tx)) &&
      (!sb || r._s.includes(sb)) && (!tc || r._t.includes(tc)) && (!no || !r.teacher));
    renderFilter();
  };
  function renderFilter() {
    const shown = FILTERED.slice(0, LIMIT), mob = IS_MOB();
    document.getElementById('filter-body').innerHTML = mob ? '' : shown.map(r =>
      `<tr class="hover:bg-slate-50"><td class="p-3 font-semibold">${esc(r.school)}</td><td class="p-3">${esc(r.day)}</td><td class="p-3">${r.period || '—'}</td><td class="p-3">${esc(r.grade)}</td><td class="p-3">${esc(r.subject)}</td><td class="p-3">${esc(r.teacher) || '<span class="text-rose-400">—</span>'}</td></tr>`).join('')
      || '<tr><td colspan="6" class="p-6 text-center text-slate-400">لا توجد نتائج مطابقة</td></tr>';
    const oneSchool = document.getElementById('f-school').value;
    document.getElementById('filter-cards').innerHTML = !mob ? '' : shown.map(r => `
      <div class="bg-white rounded-2xl border border-slate-200 p-3 shadow-xs active:scale-[.99] transition">
        <div class="flex items-center justify-between gap-2 mb-1.5">
          <div class="flex items-center gap-1.5">
            <span class="bg-emerald-600 text-white text-[11px] font-bold rounded-lg px-2 py-1">الحصة ${r.period || '—'}</span>
            <span class="bg-slate-100 text-slate-700 text-[11px] font-bold rounded-lg px-2 py-1">${esc(r.day)}</span>
          </div>
          <span class="text-[11px] text-indigo-700 bg-indigo-50 rounded-lg px-2 py-1 font-semibold">${esc(r.grade)}</span>
        </div>
        <div class="text-sm font-bold text-slate-900">${esc(r.subject) || '—'}</div>
        <div class="text-xs mt-1 ${r.teacher ? 'text-slate-600' : 'text-rose-500'}"><i class="fa-solid fa-user-tie ml-1"></i>${esc(r.teacher) || 'بلا اسم مدرس'}</div>
        ${oneSchool ? '' : `<div class="text-[11px] text-slate-400 mt-1.5 border-t border-slate-100 pt-1.5"><i class="fa-solid fa-school ml-1"></i>${esc(r.school)}</div>`}
      </div>`).join('') || '<div class="p-8 text-center text-slate-400 text-xs bg-white rounded-2xl border border-slate-200">لا توجد نتائج مطابقة</div>';
    const sc = new Set(FILTERED.map(r => r.school)).size, tc = new Set(FILTERED.map(r => r.teacher).filter(Boolean)).size;
    document.getElementById('filter-summary').textContent = `${FILTERED.length.toLocaleString('ar')} حصة — ${sc} مدرسة — ${tc} مدرس (من أصل ${FLAT.length.toLocaleString('ar')})`;
    document.getElementById('filter-more').classList.toggle('hidden', FILTERED.length <= LIMIT);
  }
  window.LIMIT_MORE = () => { LIMIT += IS_MOB() ? 60 : 500; renderFilter(); };
  window.resetFilters = function () {
    ['f-school', 'f-day', 'f-period', 'f-stage', 'f-section', 'f-subject', 'f-teacher', 'f-text'].forEach(id => document.getElementById(id).value = '');
    refreshLists();
    document.getElementById('f-noteacher').checked = false; applyFilters();
  };
  window.exportFiltered = function () {
    if (!FILTERED.length) return showToast('لا توجد نتائج للتصدير', 'error');
    const ws = XLSX.utils.json_to_sheet(FILTERED.map(r => ({ 'المدرسة': r.school, 'اليوم': r.day, 'الحصة': r.period, 'الصف': r.grade, 'المادة': r.subject, 'اسم المدرس': r.teacher })));
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'نتائج الفلترة');
    XLSX.writeFile(wb, 'برنامج_الدوام_المفلتر.xlsx');
  };

  // ---------- تكوين فريق الزيارة يدوياً ----------
  const SPEC = [
    [/انكليز|انجليز/, /انكليز|انجليز|انكيز/], [/عربي/, /عربي|نحو|ادب|بلاغ|املاء|قواعد|نصوص|تعبير|قراءه|صرف|عروض|مطالعه/],
    [/رياضيات/, /رياضيات|جبر|هندس|تفاضل|احصاء|حساب/], [/احياء/, /احياء|حياتيه|بيولوج|علوم/],
    [/فيزياء|كيمياء/, /فيزياء|كيمياء|كمياء|كيميا|فيزيا|علوم/], [/فلسف/, /فلسف|منطق|نفس|اجتماعيات/],
    [/جغراف/, /جغراف|اجتماعيات/], [/تاريخ/, /تاريخ|اجتماعيات/],
    [/شريع|شرع|فقه|عقيد|قران|حديث/, /عقيد|فقه|حديث|تفسير|قران|سيره|تجويد|تجو|دعوه|خطابه|اصول|شرع|شريع|اسلامي|توحيد|شمائل|استحفاظ|فرائض|تلاوه|مواريث|مصطلح/]
  ];
  const subjRx = spec => { const t = n(spec); const hit = SPEC.find(s => s[0].test(t)); return hit ? hit[1] : new RegExp(t.split(' ').filter(w => w.length > 2).join('|') || '$^'); };
  const dayOf = d => d ? DAYS[new Date(d + 'T00:00:00').getDay()] || '' : '';
  let TEAM = new Set();

  function injectTeamTab() {
    document.getElementById('tab-btn-filter').insertAdjacentHTML('beforebegin',
      `<button onclick="openTeamBuilder()" id="tab-btn-team" class="px-3.5 py-2 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition flex items-center gap-2"><i class="fa-solid fa-people-group"></i><span>تكوين فريق الزيارة</span></button>`);
    document.getElementById('view-filter').insertAdjacentHTML('beforebegin', `
    <div id="view-team" class="hidden space-y-4">
      <div class="bg-white p-4 rounded-2xl border border-slate-200 space-y-3">
        <h3 class="text-base font-bold text-slate-900">تكوين فريق الزيارة (تحدده أنت حسب المتاح في المدرسة)</h3>
        <div class="grid grid-cols-1 md:grid-cols-4 gap-3">
          ${field('المدرسة', `<select id="t-school" class="${inp}"></select>`)}
          ${field('يوم الزيارة', `<select id="t-day" class="${inp}">${'<option value="">— اختر اليوم —</option>' + DAYS.slice(0, 5).map(d => `<option>${d}</option>`).join('')}</select>`)}
          ${field('تاريخ الزيارة (اختياري)', `<input type="date" id="t-date" class="${inp}">`)}
          ${field('عرض الموجهين', `<label class="flex items-center gap-2 text-xs pt-2"><input type="checkbox" id="t-all"> كل الموجهين (وليس نفس جنس المدرسة فقط)</label>`)}
        </div>
      </div>
      <div id="t-avail" class="bg-white p-3 rounded-2xl border border-slate-200 text-xs"></div>
      <div id="t-alert"></div>
      <div id="t-cards" class="md:hidden space-y-2"></div>
      <div class="hidden md:block bg-white rounded-2xl border border-slate-200 overflow-x-auto custom-scrollbar">
        <table class="w-full text-xs text-right"><thead class="bg-slate-50 text-slate-600"><tr><th class="p-3 w-10"></th><th class="p-3">الموجه</th><th class="p-3">الاختصاص</th><th class="p-3">حصصه في المدرسة (أسبوعياً)</th><th class="p-3">حصصه يوم الزيارة</th><th class="p-3">الحالة</th></tr></thead>
        <tbody id="t-body" class="divide-y divide-slate-100"></tbody></table>
      </div>
      <div class="sticky bottom-28 md:bottom-2 z-30 shadow-lg flex flex-wrap items-center justify-between gap-2 bg-emerald-50 border border-emerald-300 rounded-2xl px-4 py-3">
        <span class="text-xs font-bold text-emerald-900">عدد أعضاء الفريق: <span id="t-count">0</span></span>
        <div class="flex gap-2"><button onclick="clearTeam()" class="text-xs bg-white border border-slate-300 px-3 py-2.5 rounded-xl font-bold">إلغاء الفريق</button>
        <button onclick="saveTeam()" class="text-sm bg-emerald-600 hover:bg-emerald-700 text-white px-5 py-2.5 rounded-xl font-bold">حفظ الفريق</button></div>
      </div>
    </div>`);
    document.getElementById('t-school').addEventListener('change', () => loadTeam());
    document.getElementById('t-date').addEventListener('change', () => { const d = dayOf(document.getElementById('t-date').value); if (DAYS.slice(0, 5).includes(d)) document.getElementById('t-day').value = d; renderTeam(); });
    document.getElementById('t-day').addEventListener('change', renderTeam);
    document.getElementById('t-all').addEventListener('change', renderTeam);
    ['t-body', 't-cards'].forEach(id => document.getElementById(id).addEventListener('change', e => {
      if (e.target.dataset.sup) { e.target.checked ? TEAM.add(e.target.dataset.sup) : TEAM.delete(e.target.dataset.sup); renderTeam(); }
    }));
  }
  function curSchool() { return schools.find(s => s.id === document.getElementById('t-school').value); }
  window.openTeamBuilder = function (id) {
    const sel = document.getElementById('t-school');
    sel.innerHTML = schools.map(s => `<option value="${s.id}">${esc(s.name)}</option>`).join('');
    if (id) sel.value = id;
    switchView('team'); loadTeam();
  };
  function loadTeam() {
    const e = distributionResults.find(d => d.schoolId === document.getElementById('t-school').value);
    TEAM = new Set((e?.supervisors || []).map(s => s.id));
    const vd = e?.visitDate || '', isDay = DAYS.includes(vd);
    document.getElementById('t-date').value = isDay ? '' : vd;
    const d = isDay ? vd : dayOf(vd); document.getElementById('t-day').value = DAYS.slice(0, 5).includes(d) ? d : '';
    renderTeam();
  }
  function supStats(sup, school) {
    const rx = subjRx(sup.specialty), rows = (schedules[school.id] || []).filter(r => rx.test(n(r.subject)));
    const dy = document.getElementById('t-day').value;
    const weekly = rows.length, onDay = dy ? rows.filter(r => r.day === dy).length : null;
    const days = [...new Set(rows.map(r => r.day))];
    let warn = '';
    if (!weekly) warn = 'لا توجد حصص لاختصاصه في هذه المدرسة';
    else if (onDay === 0) warn = `لا حصص لاختصاصه يوم ${dy} (يتوفر: ${days.join('، ')})`;
    return { weekly, onDay, warn };
  }
  function renderTeam() {
    const sc = curSchool(); if (!sc) return;
    const all = document.getElementById('t-all').checked;
    const list = supervisors.filter(s => all || (sc.gender === 'ذكور') === (s.gender === 'ذكر'));
    const warns = [];
    document.getElementById('t-body').innerHTML = list.map(s => {
      const st = supStats(s, sc), on = TEAM.has(s.id);
      if (on && st.warn) warns.push(`${s.name} (${s.specialty}): ${st.warn}`);
      const badge = st.warn ? `<span class="text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-2 py-0.5">⚠ ${esc(st.warn)}</span>` : '<span class="text-emerald-700">✓ متاح</span>';
      return `<tr class="${on ? 'bg-emerald-50/60' : ''}"><td class="p-3"><input type="checkbox" data-sup="${s.id}" ${on ? 'checked' : ''}></td><td class="p-3 font-semibold">${esc(s.name)}</td><td class="p-3">${esc(s.specialty)}</td><td class="p-3">${st.weekly}</td><td class="p-3">${st.onDay ?? '—'}</td><td class="p-3">${badge}</td></tr>`;
    }).join('') || '<tr><td colspan="6" class="p-6 text-center text-slate-400">لا يوجد موجهون</td></tr>';
    document.getElementById('t-cards').innerHTML = list.map(s => {
      const st = supStats(s, sc), on = TEAM.has(s.id), bad = !!st.warn;
      return `<label class="block bg-white rounded-2xl border-2 p-3 cursor-pointer transition ${on ? (bad ? 'border-amber-400 bg-amber-50/60' : 'border-emerald-500 bg-emerald-50/60') : 'border-slate-200'}">
        <div class="flex items-start gap-3">
          <input type="checkbox" data-sup="${s.id}" ${on ? 'checked' : ''} class="mt-1 w-6 h-6 accent-emerald-600 shrink-0">
          <div class="flex-1 min-w-0">
            <div class="font-bold text-sm text-slate-900">${esc(s.name)}</div>
            <div class="text-xs text-emerald-700 font-semibold">${esc(s.specialty)}</div>
            <div class="flex flex-wrap gap-1.5 mt-2 text-[11px]">
              <span class="bg-slate-100 rounded-lg px-2 py-0.5">أسبوعياً: <b>${st.weekly}</b></span>
              ${st.onDay === null ? '' : `<span class="bg-slate-100 rounded-lg px-2 py-0.5">يوم الزيارة: <b>${st.onDay}</b></span>`}
            </div>
            <div class="mt-2 text-[11px] ${bad ? 'text-amber-700 font-bold' : 'text-emerald-700'}">${bad ? '⚠ ' + esc(st.warn) : '✓ لديه حصص متاحة'}</div>
          </div>
        </div>
      </label>`;
    }).join('');
    document.getElementById('t-count').textContent = TEAM.size;
    const dy = document.getElementById('t-day').value, rows = (schedules[sc.id] || []).filter(r => !dy || r.day === dy);
    const specs = [...new Set(supervisors.map(s => s.specialty))].map(sp => [sp, rows.filter(r => subjRx(sp).test(n(r.subject))).length]);
    document.getElementById('t-avail').innerHTML = `<b>المتاح في المدرسة ${dy ? 'يوم ' + dy : 'خلال الأسبوع'}:</b> ` + specs.map(([sp, c]) =>
      `<span class="inline-block m-0.5 px-2 py-0.5 rounded-lg border ${c ? 'bg-emerald-50 border-emerald-200 text-emerald-800' : 'bg-rose-50 border-rose-200 text-rose-700'}">${esc(sp)}: ${c} حصة</span>`).join('');
    document.getElementById('t-alert').innerHTML = warns.length ? `<div class="bg-amber-50 border border-amber-300 text-amber-900 rounded-xl p-3 text-xs space-y-1"><b>تنبيه: أعضاء مختارون بلا حصص من اختصاصهم</b>${warns.map(w => `<div>• ${esc(w)}</div>`).join('')}</div>` : '';
  }
  window.saveTeam = function () {
    const sc = curSchool(); if (!sc) return;
    const team = supervisors.filter(s => TEAM.has(s.id)).map(s => ({ id: s.id, name: s.name, specialty: s.specialty, phone: s.phone, gender: s.gender, warn: supStats(s, sc).warn }));
    const entry = { schoolId: sc.id, schoolName: sc.name, gender: sc.gender, stage: sc.stage, sections: sc.sections, zone: sc.zone,
      visitDate: document.getElementById('t-date').value || document.getElementById('t-day').value, supervisors: team, manual: true };
    const i = distributionResults.findIndex(d => d.schoolId === sc.id);
    i >= 0 ? distributionResults[i] = entry : distributionResults.push(entry);
    persistData(); renderAllViews();
    const w = team.filter(t => t.warn).length;
    showToast(w ? `تم الحفظ (${team.length} أعضاء) مع ${w} تنبيه` : `تم حفظ فريق الزيارة (${team.length} أعضاء)`, w ? 'info' : 'success');
  };
  window.clearTeam = function () {
    const sc = curSchool(); if (!sc) return;
    distributionResults = distributionResults.filter(d => d.schoolId !== sc.id);
    TEAM = new Set(); persistData(); renderAllViews(); renderTeam(); showToast('أُلغي فريق المدرسة');
  };
  const _auto = window.runIntelligentDistribution;
  window.runIntelligentDistribution = function () {
    const keep = distributionResults.filter(d => d.manual);
    _auto();
    keep.forEach(m => { const i = distributionResults.findIndex(d => d.schoolId === m.schoolId); i >= 0 ? distributionResults[i] = m : distributionResults.push(m); });
    persistData(); renderAllViews();
  };

  // ---------- بطاقات جدول الزيارات للجوال ----------
  const _renderDist = window.renderDistributionTable;
  window.renderDistributionTable = function () {
    _renderDist();
    const box = document.getElementById('distribution-cards'); if (!box) return;
    const g = document.getElementById('filter-dist-gender').value, q = n(document.getElementById('search-dist-school').value);
    const rows = distributionResults.filter(r => (g === 'ALL' || r.gender === g) && (!q || n(r.schoolName + ' ' + r.zone).includes(q)));
    box.innerHTML = rows.map(r => `
      <div class="bg-white rounded-2xl border border-slate-200 p-3 shadow-xs">
        <div class="flex items-start justify-between gap-2">
          <div class="font-bold text-sm text-slate-900">${esc(r.schoolName)}</div>
          <span class="shrink-0 text-[10px] font-bold px-2 py-0.5 rounded-full ${r.gender === 'ذكور' ? 'bg-blue-100 text-blue-800' : 'bg-rose-100 text-rose-800'}">${r.gender === 'ذكور' ? 'بنين' : 'بنات'}</span>
        </div>
        <div class="text-xs text-emerald-800 font-semibold mt-1"><i class="fa-solid fa-calendar-day ml-1"></i>${esc(r.visitDate || 'زيارة دورية مفتوحة')}</div>
        <div class="mt-2 flex flex-wrap gap-1">${(r.supervisors || []).map(s => `<span class="bg-slate-50 border border-slate-200 rounded-lg px-2 py-1 text-[11px]"><b>${esc(s.name)}</b> <span class="text-emerald-700">(${esc(s.specialty)})</span>${s.warn ? ' <span class="text-amber-500" title="' + esc(s.warn) + '">⚠</span>' : ''}</span>`).join('') || '<span class="text-amber-600 text-[11px]">لم يُحدَّد فريق</span>'}</div>
        <div class="grid grid-cols-3 gap-2 mt-3">
          <button onclick="openTeamBuilder('${r.schoolId}')" class="bg-emerald-600 text-white text-xs font-bold rounded-xl py-2.5"><i class="fa-solid fa-people-group ml-1"></i>تحديد الفريق</button>
          <button onclick="viewSchoolScheduleModal('${r.schoolId}')" class="bg-slate-100 text-slate-700 text-xs font-bold rounded-xl py-2.5"><i class="fa-solid fa-calendar-check ml-1"></i>الحصص</button>
          <button onclick="openWaReport('${r.schoolId}')" class="bg-emerald-50 text-emerald-700 border border-emerald-200 text-xs font-bold rounded-xl py-2.5"><i class="fa-brands fa-whatsapp ml-1"></i>تقرير</button>
        </div>
      </div>`).join('') || '<div class="p-8 text-center text-slate-400 text-xs bg-white rounded-2xl border border-slate-200">لا يوجد جدول زيارات حالياً</div>';
  };


  // ---------- تقرير واتساب ----------
  const WA_KEY = 'sup_wa_phone';
  const NUM = ['1️⃣', '2️⃣', '3️⃣', '4️⃣', '5️⃣', '6️⃣', '7️⃣', '8️⃣', '9️⃣'];
  const fmtDate = v => {
    const t = String(v || '').trim(); if (!t) return 'زيارة دورية مفتوحة';
    const m = t.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return t;
    const d = new Date(+m[1], +m[2] - 1, +m[3]);
    return ['الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'][d.getDay()] + ' ' + m[3] + '/' + m[2] + '/' + m[1];
  };
  function buildWaReport(opt) {
    const rows = (typeof distributionResults !== 'undefined' ? distributionResults : []).filter(r => (opt.schoolId ? r.schoolId === opt.schoolId : true));
    let list = rows.filter(r => opt.gender === 'ALL' || r.gender === opt.gender);
    const scheduled = list.filter(r => r.supervisors && r.supervisors.length);
    if (!opt.unscheduled) list = scheduled;
    list = list.slice().sort((a, b) => (a.gender === b.gender ? 0 : a.gender === 'ذكور' ? -1 : 1) || String(a.visitDate || '9').localeCompare(String(b.visitDate || '9')) || COLL.compare(a.schoolName, b.schoolName));
    const today = new Date().toLocaleDateString('ar', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
    const out = [];
    out.push(opt.schoolId ? '*📋 تقرير زيارة مدرسة*' : '*📋 خطة الزيارات الإشرافية*');
    out.push('دائرة الإشراف والتوجيه التربوي');
    out.push('🗓 ' + today);
    if (!opt.schoolId) out.push(`📌 المجدولة: ${scheduled.length} من ${(typeof schools !== 'undefined' ? schools : []).length} مدرسة`);
    if (!list.length) { out.push('', 'لا توجد زيارات مجدولة ضمن هذا الاختيار.'); return out.join('\n'); }
    let lastG = null, i = 0;
    list.forEach(r => {
      if (r.gender !== lastG && !opt.schoolId) { lastG = r.gender; i = 0; out.push('', r.gender === 'ذكور' ? '*— مدارس البنين —*' : '*— مدارس البنات —*'); }
      i++;
      out.push('', `${NUM[i - 1] || i + '.'} *${r.schoolName}*`);
      out.push('📅 ' + fmtDate(r.visitDate));
      const team = (r.supervisors || []).map(x => x.name + (x.specialty ? ' (' + x.specialty + ')' : '')).join('، ');
      out.push('👥 ' + (team || 'لم يُحدَّد الفريق بعد'));
      const warns = (r.supervisors || []).filter(x => x.warn).length;
      if (warns) out.push('⚠️ يوجد تنبيه على الفريق، يُراجع قبل الزيارة');
    });
    out.push('', '_أُعدّ عبر منظومة الإشراف_');
    return out.join('\n');
  }
  function ensureWaModal() {
    if (document.getElementById('wa-modal')) return;
    document.body.insertAdjacentHTML('beforeend', `
    <div id="wa-modal" class="fixed inset-0 z-50 no-print" style="display:none">
      <div class="absolute inset-0 bg-slate-900/50" data-wa="close"></div>
      <div class="absolute bottom-0 inset-x-0 md:inset-x-auto md:left-1/2 md:-translate-x-1/2 md:bottom-auto md:top-10 md:w-[560px] bg-white rounded-t-3xl md:rounded-3xl p-4 space-y-3 max-h-[92vh] overflow-y-auto" style="padding-bottom:calc(16px + env(safe-area-inset-bottom))">
        <div class="flex items-center justify-between">
          <h3 class="text-base font-bold text-slate-900"><i class="fa-brands fa-whatsapp text-emerald-600 ml-1"></i>تقرير واتساب</h3>
          <button data-wa="close" class="w-9 h-9 rounded-xl bg-slate-100 text-slate-600"><i class="fa-solid fa-xmark"></i></button>
        </div>
        <div class="grid grid-cols-2 gap-2">
          <select id="wa-gender" class="text-sm border border-slate-200 rounded-xl px-3 py-2 bg-slate-50"><option value="ALL">كل المدارس</option><option value="ذكور">البنين</option><option value="إناث">البنات</option></select>
          <label class="flex items-center gap-2 text-xs font-bold text-slate-600 bg-slate-50 border border-slate-200 rounded-xl px-3"><input type="checkbox" id="wa-unsch"> مع غير المجدولة</label>
        </div>
        <textarea id="wa-text" rows="12" class="w-full text-sm border border-slate-200 rounded-2xl p-3 bg-slate-50 leading-relaxed" style="-webkit-user-select:text;user-select:text"></textarea>
        <input id="wa-phone" inputmode="tel" dir="ltr" placeholder="رقم المستلم بصيغة دولية (اختياري) مثل 963912345678" class="w-full text-sm border border-slate-200 rounded-xl px-3 py-2 bg-slate-50">
        <div id="wa-note" class="text-[11px] text-amber-600 hidden"></div>
        <div class="grid grid-cols-3 gap-2">
          <button data-wa="send" class="col-span-3 bg-emerald-600 text-white text-sm font-bold rounded-2xl py-3"><i class="fa-brands fa-whatsapp ml-1"></i>إرسال عبر واتساب</button>
          <button data-wa="share" class="bg-slate-100 text-slate-700 text-xs font-bold rounded-xl py-2.5"><i class="fa-solid fa-share-nodes ml-1"></i>مشاركة</button>
          <button data-wa="copy" class="bg-slate-100 text-slate-700 text-xs font-bold rounded-xl py-2.5"><i class="fa-regular fa-copy ml-1"></i>نسخ</button>
          <button data-wa="refresh" class="bg-slate-100 text-slate-700 text-xs font-bold rounded-xl py-2.5"><i class="fa-solid fa-rotate ml-1"></i>إعادة توليد</button>
        </div>
        <p class="text-[11px] text-slate-400">النص قابل للتعديل قبل الإرسال. لا تُضاف أرقام هواتف الموجهين إلى التقرير.</p>
      </div>
    </div>`);
    const m = document.getElementById('wa-modal');
    const regen = () => { m.dataset.fresh = '1'; document.getElementById('wa-text').value = buildWaReport({ gender: document.getElementById('wa-gender').value, unscheduled: document.getElementById('wa-unsch').checked, schoolId: m.dataset.school || '' }); checkLen(); };
    const checkLen = () => {
      const n = document.getElementById('wa-note'), len = encodeURIComponent(document.getElementById('wa-text').value).length;
      n.classList.toggle('hidden', len < 5500);
      n.textContent = 'التقرير طويل؛ قد لا يستقبله رابط واتساب كاملاً. استخدم «مشاركة» أو «نسخ» ثم الصقه في المحادثة.';
    };
    document.getElementById('wa-gender').addEventListener('change', regen);
    document.getElementById('wa-unsch').addEventListener('change', regen);
    document.getElementById('wa-text').addEventListener('input', checkLen);
    document.getElementById('wa-phone').addEventListener('change', e => { try { localStorage.setItem(WA_KEY, e.target.value.replace(/\D/g, '')); } catch (x) { } });
    m.addEventListener('click', async e => {
      const b = e.target.closest('[data-wa]'); if (!b) return;
      const act = b.dataset.wa, txt = document.getElementById('wa-text').value;
      if (act === 'close') m.style.display = 'none';
      else if (act === 'refresh') regen();
      else if (act === 'send') {
        const num = document.getElementById('wa-phone').value.replace(/\D/g, '');
        window.open('https://wa.me/' + num + '?text=' + encodeURIComponent(txt), '_blank');
      } else if (act === 'share') {
        if (navigator.share) { try { await navigator.share({ text: txt }); } catch (x) { } } else showToast('المشاركة غير مدعومة هنا، استخدم النسخ', 'info');
      } else if (act === 'copy') {
        try { await navigator.clipboard.writeText(txt); showToast('تم نسخ التقرير'); }
        catch (x) { const t = document.getElementById('wa-text'); t.select(); document.execCommand('copy'); showToast('تم نسخ التقرير'); }
      }
    });
    m._regen = regen;
  }
  window.openWaReport = function (schoolId) {
    ensureWaModal();
    const m = document.getElementById('wa-modal');
    m.dataset.school = schoolId || '';
    document.getElementById('wa-gender').value = 'ALL';
    document.getElementById('wa-unsch').checked = !!schoolId;
    try { document.getElementById('wa-phone').value = localStorage.getItem(WA_KEY) || ''; } catch (x) { }
    m._regen(); m.style.display = 'block';
  };

  // ---------- شريط التنقل السفلي للجوال ----------
  function injectBottomNav() {
    const items = [['distribution', 'fa-calendar-days', 'الزيارات'], ['team', 'fa-people-group', 'الفريق'], ['filter', 'fa-filter', 'الفلترة'], ['schedules', 'fa-table-cells', 'الحصص']];
    document.body.insertAdjacentHTML('beforeend', `
    <nav id="bottom-nav" class="md:hidden fixed bottom-0 inset-x-0 z-40 grid grid-cols-5 no-print">
      ${[['distribution', 'fa-calendar-days', 'الزيارات'], ['filter', 'fa-filter', 'الفلترة']].map(([v, ic, l]) => `<button data-nav="${v}" class="flex flex-col items-center justify-center gap-0.5"><i class="fa-solid ${ic} text-lg"></i><span class="text-[11px] font-bold">${l}</span></button>`).join('')}
      <button data-nav="team" aria-label="تكوين فريق"><i class="fa-solid fa-plus text-2xl"></i></button>
      ${[['schedules', 'fa-table-cells', 'الحصص'], ['more', 'fa-ellipsis', 'المزيد']].map(([v, ic, l]) => `<button data-nav="${v}" class="flex flex-col items-center justify-center gap-0.5"><i class="fa-solid ${ic} text-lg"></i><span class="text-[11px] font-bold">${l}</span></button>`).join('')}
    </nav>
    <div id="more-sheet" class="md:hidden fixed inset-0 z-50 hidden">
      <div class="absolute inset-0 bg-slate-900/50" data-nav="close"></div>
      <div class="absolute bottom-0 inset-x-0 bg-white rounded-t-3xl p-4 space-y-2" style="padding-bottom:calc(16px + env(safe-area-inset-bottom))">
        <div class="w-10 h-1.5 bg-slate-300 rounded-full mx-auto mb-2"></div>
        ${[['supervisors', 'fa-user-tie', 'كادر الموجهين'], ['schools', 'fa-mosque', 'سجل المدارس'], ['import', 'fa-cloud-arrow-down', 'تصدير ونسخ احتياطي']].map(([v, ic, l]) =>
          `<button data-nav="${v}" class="w-full flex items-center gap-3 bg-slate-50 rounded-2xl px-4 py-3.5 text-sm font-bold text-slate-700"><i class="fa-solid ${ic} text-emerald-600 w-5 text-center"></i>${l}</button>`).join('')}
        <button data-nav="wa" class="w-full flex items-center gap-3 bg-slate-50 rounded-2xl px-4 py-3.5 text-sm font-bold text-slate-700"><i class="fa-brands fa-whatsapp text-emerald-600 w-5 text-center"></i>تقرير واتساب</button>
        <button data-nav="hero" class="w-full flex items-center gap-3 bg-slate-50 rounded-2xl px-4 py-3.5 text-sm font-bold text-slate-700"><i class="fa-solid fa-gear text-emerald-600 w-5 text-center"></i>إعدادات الجولة والجدولة التلقائية</button>
      </div>
    </div>`);
    const sheet = document.getElementById('more-sheet');
    // ملاحظة أداء: Tailwind CDN يعيد بناء الأنماط عند كل تغيير في class،
    // لذلك على الجوال نغيّر style/attribute فقط (لا يراقبها) بدل classList.
    const hide = el => el && el.style.setProperty('display', 'none', 'important');
    const show = (el, d) => el && (d ? el.style.setProperty('display', d, 'important') : el.style.removeProperty('display'));
    const mob = () => window.matchMedia('(max-width: 767px)').matches;
    const VIEWS = ['distribution', 'supervisors', 'schools', 'schedules', 'team', 'filter', 'import'];
    const sheetOpen = on => on ? show(sheet, 'block') : hide(sheet);
    hide(sheet);
    document.body.addEventListener('click', e => {
      const b = e.target.closest('[data-nav]'); if (!b) return;
      const v = b.dataset.nav;
      if (v === 'more') sheetOpen(true);
      else if (v === 'close') sheetOpen(false);
      else if (v === 'wa') { sheetOpen(false); openWaReport(); }
      else if (v === 'hero') { document.getElementById('app-hero').classList.toggle('mobile-show'); sheetOpen(false); window.scrollTo({ top: 0, behavior: 'smooth' }); }
      else { sheetOpen(false); v === 'team' ? openTeamBuilder() : switchView(v); }
    });
    const _sv = window.switchView;
    window.switchView = function (v) {
      if (mob()) {
        VIEWS.forEach(x => { const el = document.getElementById('view-' + x); if (el) x === v ? show(el, 'block') : hide(el); });
      } else { _sv(v); }
      window.scrollTo(0, 0);
      const home = document.getElementById('m-home'), stats = document.getElementById('app-stats');
      v === 'distribution' ? show(home) : hide(home);
      if (mob()) { v === 'distribution' ? show(stats) : hide(stats); }
      document.querySelectorAll('#bottom-nav [data-nav]').forEach(b => {
        const on = b.dataset.nav === v || (b.dataset.nav === 'more' && ['supervisors', 'schools', 'import'].includes(v));
        on ? b.setAttribute('data-active', '1') : b.removeAttribute('data-active');
      });
    };
    switchView('distribution');
  }

  // ---------- الصفحة الرئيسية للجوال ----------
  function buildMobileHome() {
    document.querySelector('main').insertAdjacentHTML('afterbegin', `
    <div id="m-home" class="md:hidden space-y-4 no-print">
      <input id="m-search" placeholder="🔍 ابحث عن مدرسة أو مدرّس أو مادة..." class="w-full px-4 py-3 text-sm shadow-sm">
      <div id="m-hero"></div>
      <div id="m-dots" class="m-dots flex justify-center gap-1.5"></div>
      <div class="flex items-center justify-between pt-1"><h2 class="m-title">الإنجازات والمؤشرات</h2><button class="m-pill" onclick="switchView('filter')">فلترة الدوام</button></div>
    </div>`);
    let t; document.getElementById('m-search').addEventListener('input', e => {
      clearTimeout(t); t = setTimeout(() => {
        const v = e.target.value.trim(); if (!v) return;
        switchView('filter'); document.getElementById('f-text').value = v; applyFilters();
      }, 400);
    });
    const hero = document.getElementById('m-hero');
    let raf = 0;
    hero.addEventListener('scroll', () => {
      if (raf) return;
      raf = requestAnimationFrame(() => {
        raf = 0;
        const i = Math.round(Math.abs(hero.scrollLeft) / hero.clientWidth);
        document.querySelectorAll('#m-dots span').forEach((d, k) => k === i ? d.setAttribute('data-on', '1') : d.removeAttribute('data-on'));
      });
    }, { passive: true });
    refreshHome();
  }
  function refreshHome() {
    const hero = document.getElementById('m-hero'); if (!hero) return;
    const sched = distributionResults.filter(d => d.supervisors && d.supervisors.length).length;
    const warn = distributionResults.filter(d => (d.supervisors || []).some(s => s.warn)).length;
    const periods = Object.values(schedules).reduce((a, l) => a + l.length, 0);
    const slide = (eyebrow, title, sub, btn, fn) => `<div class="m-slide"><div><small>${eyebrow}</small><h3>${title}</h3><p>${sub}</p></div><div><button class="m-gold-btn" onclick="${fn}">${btn}</button></div></div>`;
    hero.innerHTML =
      slide('خطة الزيارات', `المجدولة: ${sched} من ${schools.length} مدرسة`, warn ? `فرق عليها تنبيهات: ${warn}` : 'لا تنبيهات على الفرق', 'عرض جدول الزيارات', "switchView('distribution')") +
      slide('فريق الزيارة', 'حدّد فريقك بنفسك', 'حسب حصص المدرسة واليوم المتاحين', 'تكوين فريق', 'openTeamBuilder()') +
      slide('برنامج الدوام', `${periods.toLocaleString('ar')} حصة`, `${schools.length} مدرسة · ${supervisors.length} موجه`, 'فلترة الدوام', "switchView('filter')");
    document.getElementById('m-dots').innerHTML = '<span data-on="1"></span><span></span><span></span>';
  }
  const _ra = window.renderAllViews;
  window.renderAllViews = function () { _ra(); refreshHome(); };

  // ---------- بطاقات بدل الجداول على الجوال (بدون تمرير أفقي) ----------
  function mobileCards(tbodyId) {
    const tb = document.getElementById(tbodyId); if (!tb) return;
    const table = tb.closest('table'), wrap = table.parentElement.classList.contains('overflow-x-auto') ? table.parentElement.parentElement : table.parentElement;
    let box = document.getElementById(tbodyId + '-m');
    if (!box) { wrap.insertAdjacentHTML('beforebegin', `<div id="${tbodyId}-m" class="md:hidden space-y-2"></div>`); box = document.getElementById(tbodyId + '-m'); wrap.classList.add('hidden', 'md:block'); }
    const heads = [...table.querySelectorAll('thead th')].map(t => t.textContent.trim());
    box.innerHTML = [...tb.rows].map(tr => {
      const cells = [...tr.cells];
      if (cells.length === 1) return `<div class="bg-white rounded-2xl border border-slate-200 p-6 text-center text-xs text-slate-400">${cells[0].innerHTML}</div>`;
      let title = '', acts = '', rows = '';
      cells.forEach((td, i) => {
        if (heads[i] === '#') return;
        if (td.querySelector('button')) { acts += td.innerHTML; return; }
        if (!title) { title = td.innerHTML; return; }
        if (!td.textContent.trim()) return;
        rows += `<div class="flex items-start justify-between gap-3 text-xs py-1.5 border-t border-slate-100"><span class="text-slate-500 shrink-0">${esc(heads[i])}</span><span class="text-slate-800 font-semibold text-left">${td.innerHTML}</span></div>`;
      });
      return `<div class="bg-white rounded-2xl border border-slate-200 p-3.5 shadow-xs"><div class="text-sm font-bold text-slate-900 mb-1">${title}</div>${rows}${acts ? `<div class="flex flex-wrap gap-2 justify-end mt-2 pt-2 border-t border-slate-100 text-base">${acts}</div>` : ''}</div>`;
    }).join('');
  }
  [['renderSupervisorsTable', 'supervisors-table-body'], ['renderSchoolsTable', 'schools-table-body']].forEach(([fn, id]) => {
    const o = window[fn]; window[fn] = function () { o.apply(this, arguments); mobileCards(id); };
  });

  let SM_DAY = 'الأحد';
  function renderScheduleMobile() {
    const box = document.getElementById('schedule-mobile'); if (!box) return;
    const id = document.getElementById('schedule-school-selector').value, list = schedules[id] || [];
    if (!document.getElementById('sm-body')) {
      box.innerHTML = `<div class="flex gap-1.5 overflow-x-auto pb-1" id="sm-days"></div><select id="sm-grade" class="w-full text-sm px-3 py-2"></select><div id="sm-body" class="space-y-3"></div>`;
      document.getElementById('sm-days').addEventListener('click', e => { const b = e.target.closest('[data-day]'); if (b) { SM_DAY = b.dataset.day; renderScheduleMobile(); } });
      document.getElementById('sm-grade').addEventListener('change', renderScheduleMobile);
    }
    document.getElementById('sm-days').innerHTML = DAYS.slice(0, 5).map(d => `<button data-day="${d}" class="shrink-0 px-4 py-2 rounded-xl text-xs font-bold ${d === SM_DAY ? 'bg-emerald-600 text-white' : 'bg-white border border-slate-200 text-slate-600'}">${d}</button>`).join('');
    const sel = document.getElementById('sm-grade'), keep = sel.value, grades = [...new Set(list.map(x => x.grade))];
    sel.innerHTML = '<option value="">كل الصفوف</option>' + grades.map(g => `<option>${esc(g)}</option>`).join('');
    if (grades.includes(keep)) sel.value = keep;
    const gs = sel.value ? [sel.value] : grades;
    document.getElementById('sm-body').innerHTML = gs.map(g => {
      const rows = [1, 2, 3, 4, 5, 6, 7].map(p => {
        const it = list.find(x => x.grade === g && normalizeDayName(x.day) === normalizeDayName(SM_DAY) && parseInt(x.period) === p);
        return `<div class="flex items-center gap-3 py-2 ${p > 1 ? 'border-t border-slate-100' : ''}"><span class="w-8 h-8 shrink-0 rounded-xl bg-emerald-600 text-white text-xs font-bold flex items-center justify-center">${p}</span>${it ? `<div class="min-w-0"><div class="text-sm font-bold text-slate-900">${esc(it.subject)}</div>${it.teacher ? `<div class="text-[11px] text-slate-500">${esc(it.teacher)}</div>` : ''}</div>` : '<span class="text-xs text-slate-300">—</span>'}</div>`;
      }).join('');
      return `<div class="bg-white rounded-2xl border border-slate-200 p-3.5 shadow-xs"><div class="text-sm font-bold text-emerald-800 mb-1"><i class="fa-solid fa-graduation-cap ml-1"></i>${esc(g)}</div>${rows}</div>`;
    }).join('') || '<div class="p-8 text-center text-slate-400 text-xs">لا يوجد جدول لهذه المدرسة</div>';
  }
  const IS_MOBILE = () => window.matchMedia('(max-width: 767px)').matches;
  const _rs = window.renderSchoolScheduleView;
  // على الجوال لا نبني جدول سطح المكتب المخفي (يُبطئ الصفحة)
  window.renderSchoolScheduleView = function () { if (!IS_MOBILE()) _rs(); renderScheduleMobile(); };
  // تقرير الطباعة لا يُبنى على الجوال إلا عند الطباعة فعلاً
  const _pp = window.prepareOfficialPrintReport;
  if (typeof _pp === 'function') {
    window.prepareOfficialPrintReport = function () { if (!IS_MOBILE()) _pp(); };
    window.addEventListener('beforeprint', () => { try { _pp(); } catch (e) { console.error(e); } });
  }
  function injectMobileTables() {
    const c = document.getElementById('schedule-display-container');
    c.classList.add('hidden', 'md:block');
    c.insertAdjacentHTML('afterend', '<div id="schedule-mobile" class="md:hidden space-y-3"></div>');
    renderScheduleMobile(); mobileCards('supervisors-table-body'); mobileCards('schools-table-body');
  }

  // ---------- الإقلاع ----------
  window.addEventListener('DOMContentLoaded', async () => {
    injectTab(); injectTeamTab(); injectBottomNav(); buildMobileHome(); injectMobileTables();
    if (!CFG.SCHEDULE_CSV || !CFG.SUPERVISORS_CSV) return;
    window.matchMedia('(max-width: 767px)').addEventListener('change', () => { try { renderFilter(); } catch (e) { } });
    // 1) عرض آخر بيانات محفوظة فوراً  2) جلب الجديد في الخلفية وتحديث الشاشة إن تغيّر
    const cached = readCache(); let shownCache = false;
    if (cached) {
      try { buildData(cached); renderAllViews(); shownCache = true; } catch (e) { console.error(e); }
    }
    try {
      const fresh = await fetchBoth();
      if (!shownCache || fresh.sup !== cached.sup || fresh.sch !== cached.sch) {
        buildData(fresh); renderAllViews(); saveCache(fresh);
        showToast(shownCache ? 'تم تحديث البيانات' : 'تم تحميل البيانات من Google Sheets');
      }
    } catch (e) {
      console.error(e);
      showToast(shownCache ? 'تعذّر التحديث، تُعرض آخر بيانات محفوظة' : 'تعذر تحميل البيانات: ' + e.message, 'error');
    }
  });
})();
