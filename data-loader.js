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

  async function fetchRows(url, label) {
    let text;
    try {
      const res = await fetch(url, { cache: 'no-store' });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      text = await res.text();
    } catch (e) { throw new Error(label + ' — فشل الاتصال بالرابط (' + e.message + ')'); }
    if (/^\s*<(!doctype|html)/i.test(text)) throw new Error(label + ' — الرابط يعيد صفحة ويب لا CSV (غير منشور أو ليس رابط النشر على الويب)');
    if (typeof XLSX === 'undefined') throw new Error('مكتبة XLSX لم تُحمَّل (تحقق من الإنترنت)');
    const wb = XLSX.read(text, { type: 'string', raw: true });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' }).map(sanitizeObjectKeys);
    if (!rows.length) throw new Error(label + ' — الورقة فارغة');
    return rows;
  }

  window.loadFromSheets = async function () {
    const res = await Promise.allSettled([fetchRows(CFG.SUPERVISORS_CSV, 'ورقة الموجهين'), fetchRows(CFG.SCHEDULE_CSV, 'ورقة الجدول')]);
    const errs = res.filter(r => r.status === 'rejected').map(r => r.reason.message);
    if (errs.length) throw new Error(errs.join(' | '));
    const [supRows, schRows] = res.map(r => r.value);
    supervisors = supRows.filter(r => r['اسم الموجه']).map((r, i) => ({
      id: 'sup-' + i, name: String(r['اسم الموجه']).trim(),
      gender: /أنث|انث/.test(String(r['الجنس'])) ? 'أنثى' : 'ذكر',
      specialty: String(r['التخصص'] || '').trim(), phone: String(r['الهاتف'] || '-').trim()
    }));
    if (!schRows[0] || !('المدرسة' in schRows[0])) throw new Error('ورقة الجدول: لا يوجد عمود «المدرسة» (الأعمدة الموجودة: ' + Object.keys(schRows[0] || {}).join('، ') + ')');
    if (!supRows[0] || !('اسم الموجه' in supRows[0])) throw new Error('ورقة الموجهين: لا يوجد عمود «اسم الموجه» (الأعمدة الموجودة: ' + Object.keys(supRows[0] || {}).join('، ') + ')');
    const names = [...new Set(schRows.map(r => String(r['المدرسة'] || '').trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'ar'));
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
      schedules[idOf[nm]].push(row); FLAT.push({ school: nm, ...row });
    });
    schools.forEach(s => { s.sections = new Set(schedules[s.id].map(x => x.grade)).size; });
    distributionResults = distributionResults.filter(d => idOf[d.schoolName] === d.schoolId);
    window.SHEETS_MODE = true;
    buildFilterOptions(); applyFilters();
  };

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
    const uniq = k => [...new Set(FLAT.map(r => r[k]).filter(Boolean))].sort((a, b) => String(a).localeCompare(String(b), 'ar'));
    document.getElementById('f-school').innerHTML = opts('كل المدارس', uniq('school'));
    document.getElementById('f-day').innerHTML = opts('كل الأيام', DAYS);
    document.getElementById('f-period').innerHTML = opts('كل الحصص', [1, 2, 3, 4, 5, 6, 7]);
    document.getElementById('f-stage').innerHTML = opts('كل الصفوف', STAGES);
    document.getElementById('f-section').innerHTML = '<option value="">كل الشعب</option>' + [1, 2, 3, 4, 5].map(i => `<option value="${i}">شعبة ${i}</option>`).join('') + '<option>أدبي</option><option>علمي</option>';
    refreshLists();
  }
  function refreshLists() {
    const sc = document.getElementById('f-school').value, rows = sc ? FLAT.filter(r => r.school === sc) : FLAT;
    const uniq = k => [...new Set(rows.map(r => r[k]).filter(Boolean))].sort((a, b) => String(a).localeCompare(String(b), 'ar'));
    document.getElementById('dl-subject').innerHTML = uniq('subject').map(v => `<option value="${esc(v)}">`).join('');
    document.getElementById('dl-teacher').innerHTML = uniq('teacher').map(v => `<option value="${esc(v)}">`).join('');
  }
  window.applyFilters = function () {
    LIMIT = 300;
    const v = id => document.getElementById(id).value.trim();
    const sc = v('f-school'), dy = v('f-day'), pr = v('f-period'), st = v('f-stage'), se = v('f-section'), tx = n(v('f-text')), sb = n(v('f-subject')), tc = n(v('f-teacher'));
    const no = document.getElementById('f-noteacher').checked;
    FILTERED = FLAT.filter(r => (!sc || r.school === sc) && (!dy || r.day === dy) && (!pr || r.period == pr) && (!st || r.stage === st) && (!se || (/^\d$/.test(se) ? r.sec == se : r.track === se)) &&
      (!tx || n([r.school, r.day, r.grade, r.subject, r.teacher].join(' ')).includes(tx)) &&
      (!sb || n(r.subject).includes(sb)) && (!tc || n(r.teacher).includes(tc)) && (!no || !r.teacher));
    FILTERED.sort((a, b) => a.school.localeCompare(b.school, 'ar') || DAYS.indexOf(a.day) - DAYS.indexOf(b.day) || a.period - b.period);
    renderFilter();
  };
  function renderFilter() {
    const shown = FILTERED.slice(0, LIMIT);
    document.getElementById('filter-body').innerHTML = shown.map(r =>
      `<tr class="hover:bg-slate-50"><td class="p-3 font-semibold">${esc(r.school)}</td><td class="p-3">${esc(r.day)}</td><td class="p-3">${r.period || '—'}</td><td class="p-3">${esc(r.grade)}</td><td class="p-3">${esc(r.subject)}</td><td class="p-3">${esc(r.teacher) || '<span class="text-rose-400">—</span>'}</td></tr>`).join('')
      || '<tr><td colspan="6" class="p-6 text-center text-slate-400">لا توجد نتائج مطابقة</td></tr>';
    const oneSchool = document.getElementById('f-school').value;
    document.getElementById('filter-cards').innerHTML = shown.map(r => `
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
  window.LIMIT_MORE = () => { LIMIT += 500; renderFilter(); };
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
        <div class="grid grid-cols-2 gap-2 mt-3">
          <button onclick="openTeamBuilder('${r.schoolId}')" class="bg-emerald-600 text-white text-xs font-bold rounded-xl py-2.5"><i class="fa-solid fa-people-group ml-1"></i>تحديد الفريق</button>
          <button onclick="viewSchoolScheduleModal('${r.schoolId}')" class="bg-slate-100 text-slate-700 text-xs font-bold rounded-xl py-2.5"><i class="fa-solid fa-calendar-check ml-1"></i>الحصص</button>
        </div>
      </div>`).join('') || '<div class="p-8 text-center text-slate-400 text-xs bg-white rounded-2xl border border-slate-200">لا يوجد جدول زيارات حالياً</div>';
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
        ${[['supervisors', 'fa-user-tie', 'كادر الموجهين'], ['schools', 'fa-mosque', 'سجل المدارس'], ['import', 'fa-file-import', 'استيراد وتصدير Excel']].map(([v, ic, l]) =>
          `<button data-nav="${v}" class="w-full flex items-center gap-3 bg-slate-50 rounded-2xl px-4 py-3.5 text-sm font-bold text-slate-700"><i class="fa-solid ${ic} text-emerald-600 w-5 text-center"></i>${l}</button>`).join('')}
        <button data-nav="hero" class="w-full flex items-center gap-3 bg-slate-50 rounded-2xl px-4 py-3.5 text-sm font-bold text-slate-700"><i class="fa-solid fa-gear text-emerald-600 w-5 text-center"></i>إعدادات الجولة والجدولة التلقائية</button>
      </div>
    </div>`);
    const sheet = document.getElementById('more-sheet');
    document.body.addEventListener('click', e => {
      const b = e.target.closest('[data-nav]'); if (!b) return;
      const v = b.dataset.nav;
      if (v === 'more') sheet.classList.remove('hidden');
      else if (v === 'close') sheet.classList.add('hidden');
      else if (v === 'hero') { document.getElementById('app-hero').classList.toggle('mobile-show'); sheet.classList.add('hidden'); window.scrollTo({ top: 0, behavior: 'smooth' }); }
      else { sheet.classList.add('hidden'); v === 'team' ? openTeamBuilder() : switchView(v); }
    });
    const _sv = window.switchView;
    window.switchView = function (v) {
      _sv(v); window.scrollTo(0, 0);
      document.getElementById('m-home')?.classList.toggle('hidden', v !== 'distribution');
      document.getElementById('app-stats')?.classList.toggle('max-md:hidden', v !== 'distribution');
      document.querySelectorAll('#bottom-nav [data-nav]').forEach(b => {
        const on = b.dataset.nav === v || (b.dataset.nav === 'more' && ['supervisors', 'schools', 'import'].includes(v));
        b.classList.toggle('active', on);
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
    hero.addEventListener('scroll', () => {
      const i = Math.round(Math.abs(hero.scrollLeft) / hero.clientWidth);
      document.querySelectorAll('#m-dots span').forEach((d, k) => d.classList.toggle('on', k === i));
    });
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
    document.getElementById('m-dots').innerHTML = '<span class="on"></span><span></span><span></span>';
  }
  const _ra = window.renderAllViews;
  window.renderAllViews = function () { _ra(); refreshHome(); };

  // ---------- الإقلاع ----------
  window.addEventListener('DOMContentLoaded', async () => {
    injectTab(); injectTeamTab(); injectBottomNav(); buildMobileHome();
    if (!CFG.SCHEDULE_CSV || !CFG.SUPERVISORS_CSV) return;
    try { await loadFromSheets(); renderAllViews(); showToast('تم تحميل البيانات من Google Sheets'); }
    catch (e) { console.error(e); showToast('تعذر تحميل البيانات: ' + e.message, 'error'); }
  });
})();
