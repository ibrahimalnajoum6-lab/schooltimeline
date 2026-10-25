// ===== إعدادات الاتصال بقاعدة البيانات (Google Sheets) =====
// الصق هنا روابط النشر CSV (ملف ← مشاركة ← النشر على الويب ← اختر الورقة ← CSV)
window.APP_CONFIG = {
  // رمز الدولة لتحويل أرقام الهواتف المحلية (09xxxxxxxx) إلى صيغة واتساب الدولية. سوريا: 963
  COUNTRY_CODE: "963",
  SCHEDULE_CSV:    "https://docs.google.com/spreadsheets/d/e/2PACX-1vTusK9IqO9yjxcjJ8ooHoYJYyIVjxN9Oc0CE2dDNkzy2TmKvCUiH8rr9gUE_dB-tN2xcbsjZ-vnpB_u/pub?gid=1218088545&single=true&output=csv",   // ورقة "قالب جدول الحصص"
  SUPERVISORS_CSV: "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ1JyPenkA-moCb3jcZ3Fbdk9nhZqvpm_4tEPkYmBRLLRl8ZF3IXXL6GVN7vTPkPxbXYudXoYn5KtPz/pub?gid=1778888823&single=true&output=csv",   // ورقة الموجهين (الصف الأول عناوين)
  // لمدرسة لا يدل اسمها على الجنس (بنين/بنات) اكتب جنسها هنا: 'ذكور' أو 'إناث'
  SCHOOL_GENDER_OVERRIDES: {
    // "الخسروية": "إناث",
    // "الثانوية الشرعية النعمانية": "إناث",
  }
};
