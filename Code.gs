/****************************************************************
 * نظام "قناة المؤسسة" - نظام إشهاري شبيه بواتساب للأولياء
 * يعمل فوق Google Sheets - Google Apps Script
 * ---------------------------------------------------------------
 * أسماء الشيتات المستخدمة (تُنشأ تلقائياً بدالة setupSystem):
 *   1) "الطالبات"   : قاعدة بيانات الطالبات وأولياء الأمور
 *   2) "المنشورات"  : كل ما يُنشر (تعليمات / أخبار / صور / فيديوهات)
 *   3) "القراءات"   : سجل من قرأ كل منشور (لإحصائية القراءة)
 *   4) "المسؤولين"  : حسابات لوحة تحكم الإدارة
 *   5) "الإعدادات"  : إعدادات عامة (التعيين / الموقع):
 *        - اللوقو            : رابط أو معرف ملف صورة في Drive
 *        - المستوى الدراسي   : قائمة المستويات مفصولة بفواصل
 *        - الإعلان           : نص إعلان يظهر في شاشة الدخول وواجهة الولي
 *        - ملف الدرايف       : رابط أو معرف مجلد Drive لحفظ الصور والفيديوهات
 ****************************************************************/

// ------------------------- الإعدادات العامة -------------------------
const APP_NAME          = 'قناة مؤسستي';
const SHEET_STUDENTS     = 'الطالبات';
const SHEET_POSTS        = 'المنشورات';
const SHEET_READS        = 'القراءات';
const SHEET_ADMINS       = 'المسؤولين';
const SHEET_SETTINGS     = 'الإعدادات';
const SHEET_TRACKING     = 'المتابعة';
const SHEET_COLLECTION   = 'التحصيل';
const SHEET_POLL_OPTIONS = 'خيارات_الاستبيان';
const SHEET_POLL_ANSWERS = 'إجابات_الاستبيان';
const DRIVE_FOLDER_NAME  = 'مرفقات_قناة_المؤسسة';
// إن أردت اختيار مجلد Drive معيّن بنفسك بدل الإنشاء التلقائي:
// افتح المجلد في Drive، انسخ المعرف من الرابط
// مثال رابط: https://drive.google.com/drive/folders/1AbCDEfGhIJKLmnoPQRstuVWXyz
// المعرف هو الجزء بعد /folders/  →  الصقه هنا بين علامتي التنصيص:
const DRIVE_FOLDER_ID    = ''; // اتركه فارغاً ليُنشئ النظام مجلداً تلقائياً باسم DRIVE_FOLDER_NAME
const SESSION_DURATION   = 60 * 60 * 6; // 6 ساعات بالثواني (أقصى مدة لـ CacheService)

// ------------------------- إعدادات الإشعارات (OneSignal) -------------------------
// تُملأ القيمتان من لوحة OneSignal: Settings > Keys & IDs
const ONESIGNAL_APP_ID       = ''; // ضع هنا App ID
const ONESIGNAL_REST_API_KEY = ''; // ضع هنا REST API Key
// رابط التطبيق على GitHub Pages (يُفتح عند الضغط على الإشعار)
const PWA_APP_URL            = ''; // مثال: https://username.github.io/hamzat-wasl/

// رأس أعمدة شيت الطالبات (لاحظ إضافة عمود "الفوج" بعد "المستوى"، وعمود "جديدة" في النهاية)
const STUDENT_HEADERS = ['ر.التسجيل', 'اسم الطالبة', 'المستوى', 'الفوج', 'هاتف الأب1', 'هاتف الأب2', 'هاتف الأم', 'كلمة المرور', 'جديدة', 'التربص'];
const POST_HEADERS    = ['معرف المنشور', 'التاريخ والوقت', 'المستوى المستهدف', 'الفوج المستهدف', 'النوع', 'النص', 'رابط الوسائط', 'اسم الملف', 'الناشر', 'مدة الظهور (أيام)', 'بعد الانتهاء', 'تاريخ بدء الظهور', 'مثبت', 'الطالبة المستهدفة'];
const READ_HEADERS    = ['معرف المنشور', 'ر.التسجيل', 'تاريخ القراءة'];
const ADMIN_HEADERS   = ['اسم المستخدم', 'كلمة المرور', 'الاسم الكامل', 'الدور'];
const SETTINGS_HEADERS = ['التعيين', 'الموقع'];
const TRACKING_HEADERS   = ['ر.التسجيل', 'اسم الطالبة', 'الفصل', 'التأخرات', 'الغيابات', 'التكريمات'];
const COLLECTION_HEADERS = ['ر.التسجيل', 'اسم الطالبة', 'القسط', 'الصنف', 'الفترة', 'المطلوب', 'التحصيل'];
const POLL_OPTION_HEADERS = ['معرف المنشور', 'الخيار'];
const POLL_ANSWER_HEADERS = ['معرف المنشور', 'ر.التسجيل', 'الخيار', 'تاريخ الإجابة'];

const TARGET_ALL = 'الكل';           // قيمة خاصة تعني "جميع المستويات/الأفواج"
const NEW_GROUP  = 'طالبات جدد';     // (لم يعد يُعرض في القوائم — حل محله صنف "الجديدات")
const TARGET_NEW = 'الجديدات';       // صنف خاص في "المستوى المستهدف": يستهدف كل طالبة معلَّمة في عمود "جديدة" (عبر كل المستويات)
const TARGET_STUDENT = 'طالبة محددة'; // صنف خاص: يستهدف طالبة واحدة بعينها (رقم تسجيلها في عمود "الطالبة المستهدفة")
const TARGET_INTERNSHIP_PREFIX = 'التربص: '; // بادئة صنف خاص: يستهدف طالبات عمود "التربص" عبر كل المستويات، بصرف النظر عن المستوى/الفوج

const ROLE_FULL  = 'مشرف عام';            // صلاحية كاملة: يمكنه اختيار أي "مستوى مستهدف"
const ROLE_QURAN = 'مشرف القرآن الكريم';  // صلاحية محدودة: لا يمكنه النشر إلا لفئتي "التربص: داخلي" و"التربص: نصف داخلي"

const INTERNSHIP_OPTIONS = ['داخلي', 'نصف داخلي', 'غير معنية']; // قيم عمود "التربص"

// هل الطالبة جديدة؟ أي قيمة غير فارغة في عمود "جديدة" تُعتبر تعليماً
// (كلمة "جديدة"، نعم، ✓، 1، أو صح من مربع اختيار) — عدا FALSE الناتجة عن مربع غير مؤشر
function isNewStudent_(st) {
  const v = norm_(st['جديدة']).toLowerCase();
  return v !== '' && v !== 'false';
}

/* ===================== نقطة الدخول للويب ===================== */
function doGet(e) {
  // الرابط الأساسي = واجهة ولي الأمر فقط، والرابط مع المعامل ?view=admin = واجهة الإدارة فقط
  const view = (e && e.parameter && e.parameter.view) ? norm_(e.parameter.view).toLowerCase() : '';
  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.appMode = (view === 'admin') ? 'admin' : 'parent';
  return tpl.evaluate()
      .setTitle(APP_NAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ===================== الإعداد الأولي (يُشغَّل يدوياً مرة واحدة) ===================== */
function setupSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheet_(ss, SHEET_STUDENTS, STUDENT_HEADERS);
  ensureStudentColumns_();
  ensureSheet_(ss, SHEET_POSTS, POST_HEADERS);
  ensureSheet_(ss, SHEET_READS, READ_HEADERS);
  ensureSheet_(ss, SHEET_TRACKING, TRACKING_HEADERS);
  ensureSheet_(ss, SHEET_COLLECTION, COLLECTION_HEADERS);
  ensureSheet_(ss, SHEET_POLL_OPTIONS, POLL_OPTION_HEADERS);
  ensureSheet_(ss, SHEET_POLL_ANSWERS, POLL_ANSWER_HEADERS);
  const adminSheet = ensureSheet_(ss, SHEET_ADMINS, ADMIN_HEADERS);
  ensureAdminColumns_();

  // شيت الإعدادات: التعيين / الموقع
  const settingsSheet = ensureSheet_(ss, SHEET_SETTINGS, SETTINGS_HEADERS);
  if (settingsSheet.getLastRow() < 2) {
    settingsSheet.appendRow(['اللوقو', '']);
    settingsSheet.appendRow(['الموسم الدراسي', '']);
    settingsSheet.appendRow(['المستوى الدراسي', '']);
    settingsSheet.appendRow(['الإعلان', '']);
    settingsSheet.appendRow(['ملف الدرايف: لحفظ الصور والفيديوهات', '']);
  }

  // إنشاء حساب إدارة افتراضي إن لم يوجد أي حساب
  if (adminSheet.getLastRow() < 2) {
    adminSheet.appendRow(['admin', sha256_('admin123'), 'المسؤول الرئيسي']);
  }

  // بيانات تجريبية إن كان الشيت فارغاً (يمكن حذفها لاحقاً)
  const stSheet = ss.getSheetByName(SHEET_STUDENTS);
  if (stSheet.getLastRow() < 2) {
    stSheet.appendRow(['1001', 'مثال: نور الهدى', 'المستوى الأول', 'الفوج (أ)', '0550000001', '', '0660000001', '0550000001', '']);
  }

  SpreadsheetApp.getUi().alert('تم إعداد النظام بنجاح ✅\nحساب الإدارة الافتراضي:\nاسم المستخدم: admin\nكلمة المرور: admin123\n\n⚠️ يرجى تغييرها فوراً من شيت "المسؤولين".');
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0B7A6B').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/* ===================== أدوات مساعدة عامة ===================== */
function sha256_(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function getSheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) {
    throw new Error('الشيت "' + name + '" غير موجود. الرجاء تشغيل الدالة setupSystem مرة واحدة من محرر Apps Script (القائمة المنسدلة أعلى الصفحة > اختر setupSystem > تشغيل ▶️).');
  }
  return sheet;
}

function sheetToObjects_(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    headers.forEach((h, idx) => obj[h] = data[i][idx]);
    obj.__row = i + 1; // رقم السطر الفعلي في الشيت (1-indexed)
    rows.push(obj);
  }
  return rows;
}

function norm_(v) {
  return String(v === null || v === undefined ? '' : v).trim();
}

function newId_() {
  return Utilities.getUuid();
}

// توليد كلمة مرور عشوائية من 6 أرقام (الرقم الأول 1-9 حتى لا يسقط الصفر البادئ في الشيت)
function generatePassword_() {
  let p = String(Math.floor(Math.random() * 9) + 1);
  for (let i = 0; i < 5; i++) p += Math.floor(Math.random() * 10);
  return p;
}

// تحويل قيمة التاريخ إلى نص قبل إرجاعها للمتصفح
// (google.script.run لا يقبل كائنات Date في القيم المرجعة، وإلا وصل الرد null)
function toIsoStr_(v) {
  if (v instanceof Date) return v.toISOString();
  return String(v === null || v === undefined ? '' : v);
}

// تحويل رابط الوسائط المخزن إلى رابط قابل للعرض داخل الصفحة
// (صيغة uc?export=view أوقفت Google دعمها للتضمين المباشر وتعيد 403)
function toDisplayMediaUrl_(url, type) {
  const u = norm_(url);
  if (!u) return '';
  const id = extractDriveId_(u);
  if (id === u) return u; // ليس رابط درايف معروفاً: يُعاد كما هو
  if (norm_(type) === 'فيديو') {
    return 'https://drive.google.com/file/d/' + id + '/preview';
  }
  return 'https://drive.google.com/thumbnail?id=' + id + '&sz=w2000';
}

/* ===================== مدة ظهور المنشور والحجب/الحذف بعد الانتهاء ===================== */
const AFTER_HIDE    = 'حجب';             // بعد انتهاء المدة: حجب عن الأولياء فقط
const AFTER_DELETE  = 'حذف من الدرايف';  // بعد انتهاء المدة: حجب + حذف الوسائط من الدرايف
const AFTER_DELETED = 'تم الحذف';        // حالة تُكتب تلقائياً بعد تنفيذ الحذف (حتى لا يتكرر)

// ضمان وجود عمودي "مدة الظهور (أيام)" و"بعد الانتهاء" في شيت المنشورات القائم
// (يضيفهما تلقائياً في نهاية سطر الرأس إن لم يوجدا — لا حاجة لإضافتهما يدوياً)
function ensurePostColumns_() {
  const sheet = getSheet_(SHEET_POSTS);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(norm_);
  let col = lastCol;
  ['مدة الظهور (أيام)', 'بعد الانتهاء', 'تاريخ بدء الظهور', 'مثبت', 'الطالبة المستهدفة'].forEach(h => {
    if (headers.indexOf(h) === -1) { col++; sheet.getRange(1, col).setValue(h); }
  });
}

// ضمان وجود عمود "التربص" في شيت الطالبات القائم (يُضاف تلقائياً في نهاية سطر الرأس إن لم يوجد)
function ensureStudentColumns_() {
  const sheet = getSheet_(SHEET_STUDENTS);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(norm_);
  let col = lastCol;
  ['التربص'].forEach(h => {
    if (headers.indexOf(h) === -1) { col++; sheet.getRange(1, col).setValue(h); }
  });
}

// ضمان وجود عمود "الدور" في شيت المسؤولين القائم؛ الحسابات القديمة بلا دور تُعامَل كـ"مشرف عام" افتراضياً
function ensureAdminColumns_() {
  const sheet = getSheet_(SHEET_ADMINS);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(norm_);
  let col = lastCol;
  ['الدور'].forEach(h => {
    if (headers.indexOf(h) === -1) { col++; sheet.getRange(1, col).setValue(h); }
  });
}

// لحظة بدء ظهور المنشور بالمللي ثانية إن حُدد "تاريخ بدء الظهور"، وإلا NaN
function postStartMs_(p) {
  const v = p['تاريخ بدء الظهور'];
  if (v === undefined || v === null || norm_(v) === '') return NaN;
  const d = (v instanceof Date) ? v : new Date(norm_(v));
  const t = d.getTime();
  return isNaN(t) ? NaN : t;
}

// هل المنشور مجدول لموعد لم يحن بعد؟
function isPostScheduledFuture_(p, nowMs) {
  const st = postStartMs_(p);
  return !isNaN(st) && st > nowMs;
}

// هل المنشور مثبت؟ أي قيمة غير فارغة في عمود "مثبت" (عدا FALSE من مربع اختيار)
function isPinned_(p) {
  const v = norm_(p['مثبت']).toLowerCase();
  return v !== '' && v !== 'false';
}

// هل انتهت مدة ظهور المنشور؟ (خانة فارغة أو صفر = ظهور دائم بلا انتهاء)
// المدة تُحسب من "تاريخ بدء الظهور" إن حُدد، وإلا من تاريخ الإنشاء
function isPostExpired_(p, nowMs) {
  const days = parseFloat(String(p['مدة الظهور (أيام)'] === undefined ? '' : p['مدة الظهور (أيام)']).replace(',', '.'));
  if (isNaN(days) || days <= 0) return false;
  const scheduled = postStartMs_(p);
  const dt = p['التاريخ والوقت'];
  const start = !isNaN(scheduled) ? scheduled : ((dt instanceof Date) ? dt.getTime() : new Date(dt).getTime());
  if (isNaN(start)) return false;
  return nowMs > start + days * 86400000;
}

// كنس المنشورات المنتهية: من كان خيارها "حذف من الدرايف" تُنقل وسائطها إلى سلة مهملات الدرايف
// وتُكتب حالتها "تم الحذف" حتى لا يعاد الحذف. يُستدعى الكنس تلقائياً عند كل تحميل للمنشورات.
function sweepExpiredPosts_() {
  const sheet = getSheet_(SHEET_POSTS);
  const posts = sheetToObjects_(sheet);
  const now = Date.now();
  const toDelete = posts.filter(p => isPostExpired_(p, now) && norm_(p['بعد الانتهاء']) === AFTER_DELETE);
  if (!toDelete.length) return;
  ensurePostColumns_();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(norm_);
  const afterCol = headers.indexOf('بعد الانتهاء') + 1;
  toDelete.forEach(p => {
    const url = norm_(p['رابط الوسائط']);
    if (url) {
      try {
        const id = extractDriveId_(url);
        if (id !== url) DriveApp.getFileById(id).setTrashed(true);
      } catch (e) { /* الملف محذوف مسبقاً أو غير متاح: نكمل دون كسر */ }
    }
    sheet.getRange(p.__row, afterCol).setValue(AFTER_DELETED);
  });
}

/* ===================== شيت الإعدادات (التعيين / الموقع) ===================== */
// يقرأ شيت "الإعدادات" ويعيد خريطة: التعيين -> الموقع
function getSettings_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  if (!sheet || sheet.getLastRow() < 1) return {};
  const data = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 0; i < data.length; i++) {
    const key = norm_(data[i][0]);
    if (!key || key === 'التعيين') continue; // تجاوز سطر الرأس إن وُجد
    map[key] = norm_(data[i][1]);
  }
  return map;
}

// بحث متسامح عن قيمة إعداد: يكفي أن يحتوي التعيين على الكلمة المفتاحية
// (مثلاً "ملف الدرايف: لحفظ الصور والفيديوهات" يطابق الكلمة "الدرايف")
function settingByKeyword_(settings, keywords) {
  const keys = Object.keys(settings);
  for (let i = 0; i < keywords.length; i++) {
    const found = keys.find(k => k.indexOf(keywords[i]) !== -1);
    if (found) return settings[found];
  }
  return '';
}

// استخراج معرف Drive من رابط مجلد أو ملف، أو إرجاع القيمة كما هي إن كانت معرفاً
function extractDriveId_(value) {
  const v = norm_(value);
  let m = v.match(/\/folders\/([A-Za-z0-9_-]+)/);
  if (m) return m[1];
  m = v.match(/\/d\/([A-Za-z0-9_-]+)/);
  if (m) return m[1];
  m = v.match(/[?&]id=([A-Za-z0-9_-]+)/);
  if (m) return m[1];
  return v;
}

// إعدادات عامة تُعرض قبل تسجيل الدخول (اللوقو + الإعلان)
function getPublicSettings() {
  const settings = getSettings_();
  const announcement = norm_(settingByKeyword_(settings, ['الإعلان']));
  const season = norm_(settingByKeyword_(settings, ['الموسم']));

  let logoData = '';
  const logoRaw = norm_(settingByKeyword_(settings, ['اللوقو', 'اللوغو', 'الشعار']));
  if (logoRaw) {
    try {
      const blob = DriveApp.getFileById(extractDriveId_(logoRaw)).getBlob();
      logoData = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
    } catch (e) {
      logoData = ''; // لوقو غير صالح: نتجاهله دون كسر الواجهة
    }
  }

  return { announcement: announcement, season: season, logoData: logoData };
}

/* ===================== الجلسات (Sessions) ===================== */
function createSession_(payload) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(token, JSON.stringify(payload), SESSION_DURATION);
  return token;
}

function getSession_(token) {
  if (!token) return null;
  const raw = CacheService.getScriptCache().get(token);
  return raw ? JSON.parse(raw) : null;
}

function requireSession_(token, type) {
  const s = getSession_(token);
  if (!s || s.type !== type) {
    throw new Error('انتهت صلاحية الجلسة، الرجاء تسجيل الدخول من جديد');
  }
  return s;
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove(token);
  return true;
}

/* ===================== تسجيل دخول الأولياء ===================== */
// اسم المستخدم: ر.التسجيل | كلمة المرور: تبدأ برقم الهاتف ثم قابلة للتعديل من الولي
function parentLogin(regId, password) {
  regId = norm_(regId);
  password = norm_(password);
  const students = sheetToObjects_(getSheet_(SHEET_STUDENTS));
  const student = students.find(s => norm_(s['ر.التسجيل']) === regId);

  if (!student) throw new Error('رقم التسجيل غير موجود');

  const storedPass = norm_(student['كلمة المرور']);
  if (storedPass !== password) throw new Error('كلمة المرور غير صحيحة');

  const token = createSession_({
    type: 'parent',
    regId: student['ر.التسجيل'],
    name: student['اسم الطالبة'],
    level: student['المستوى'],
    group: student['الفوج']
  });

  return {
    token: token,
    studentName: student['اسم الطالبة'],
    level: student['المستوى'],
    group: student['الفوج']
  };
}

function changeParentPassword(token, newPassword) {
  const s = requireSession_(token, 'parent');
  newPassword = String(newPassword || '').trim();
  if (newPassword.length < 4) throw new Error('كلمة المرور قصيرة جداً (4 أحرف على الأقل)');

  const sheet = getSheet_(SHEET_STUDENTS);
  const students = sheetToObjects_(sheet);
  const student = students.find(x => String(x['ر.التسجيل']).trim() === String(s.regId).trim());
  if (!student) throw new Error('تعذر إيجاد الطالبة');

  const colIndex = STUDENT_HEADERS.indexOf('كلمة المرور') + 1;
  sheet.getRange(student.__row, colIndex).setValue(newPassword);
  return true;
}

/* ===================== تسجيل دخول الإدارة ===================== */
// قائمة أسماء المستخدمين فقط (بلا كلمات مرور ولا أدوار) لتعبئة القائمة المنسدلة في شاشة الدخول
// ملاحظة أمنية: هذه الدالة عامة (بلا جلسة) لأنها تُستدعى قبل تسجيل الدخول، وتكشف فقط أسماء المستخدمين لا كلمات المرور
function getAdminUsernamesPublic() {
  ensureAdminColumns_();
  const admins = sheetToObjects_(getSheet_(SHEET_ADMINS));
  return admins.map(a => ({
    username: a['اسم المستخدم'],
    fullName: a['الاسم الكامل'] || a['اسم المستخدم']
  })).filter(a => norm_(a.username) !== '');
}

// يغيّر المشرف (أياً كان دوره) كلمة مروره الخاصة بنفسه
function adminChangePassword(token, newPassword) {
  const s = requireSession_(token, 'admin');
  newPassword = String(newPassword || '').trim();
  if (newPassword.length < 4) throw new Error('كلمة المرور قصيرة جداً (4 أحرف على الأقل)');

  const sheet = getSheet_(SHEET_ADMINS);
  const admins = sheetToObjects_(sheet);
  const admin = admins.find(a => String(a['اسم المستخدم']).trim() === String(s.username).trim());
  if (!admin) throw new Error('تعذر إيجاد الحساب');

  const colIndex = ADMIN_HEADERS.indexOf('كلمة المرور') + 1;
  sheet.getRange(admin.__row, colIndex).setValue(sha256_(newPassword));
  return true;
}

function adminLogin(username, password) {
  username = String(username || '').trim();
  password = String(password || '').trim();
  ensureAdminColumns_();
  const admins = sheetToObjects_(getSheet_(SHEET_ADMINS));
  const admin = admins.find(a => String(a['اسم المستخدم']).trim() === username);
  if (!admin) throw new Error('اسم المستخدم غير صحيح');
  if (String(admin['كلمة المرور']) !== sha256_(password)) throw new Error('كلمة المرور غير صحيحة');

  const role = norm_(admin['الدور']) || ROLE_FULL; // حساب بلا دور محدَّد يُعامَل كمشرف عام (توافقاً مع الحسابات القديمة)
  const token = createSession_({ type: 'admin', username: admin['اسم المستخدم'], fullName: admin['الاسم الكامل'], role: role });
  return { token: token, fullName: admin['الاسم الكامل'], role: role };
}

/* ===================== تغذية المنشورات لولي الأمر ===================== */
// هل الطالبة الموصوفة بـ d = {regId, level, group, isNew, internship} مستهدَفة بالمنشور p؟
// (يوحّد منطق الاستهداف: الكل / مستوى-فوج / الجديدات / التربص / طالبة محددة)
function postTargetsStudent_(p, d) {
  const postLevel = norm_(p['المستوى المستهدف']);
  const postGroup = norm_(p['الفوج المستهدف']);
  if (postLevel === TARGET_STUDENT) {
    return String(d.regId).trim() === String(p['الطالبة المستهدفة']).trim();
  }
  if (postLevel === TARGET_NEW) return !!d.isNew;
  if (postLevel.indexOf(TARGET_INTERNSHIP_PREFIX) === 0) {
    const wanted = postLevel.slice(TARGET_INTERNSHIP_PREFIX.length);
    return norm_(d.internship || 'غير معنية') === wanted;
  }
  const levelOk = postLevel === TARGET_ALL || postLevel === norm_(d.level);
  const groupOk = postGroup === TARGET_ALL || postGroup === norm_(d.group);
  return levelOk && groupOk;
}

function getParentFeed(token) {
  const s = requireSession_(token, 'parent');
  try { sweepExpiredPosts_(); } catch (e) { /* لا نكسر الفيد إن تعذر الكنس */ }
  const posts = sheetToObjects_(getSheet_(SHEET_POSTS));
  const reads = sheetToObjects_(getSheet_(SHEET_READS));

  // قراءة علامة "جديدة" مباشرة من الشيت (لا من الجلسة) حتى يسري أي تعديل فوراً دون إعادة دخول
  const students = sheetToObjects_(getSheet_(SHEET_STUDENTS));
  const me = students.find(x => norm_(x['ر.التسجيل']) === norm_(s.regId));
  const iAmNew = me ? isNewStudent_(me) : false;
  const myInternship = me ? (norm_(me['التربص']) || 'غير معنية') : 'غير معنية';

  const myReadIds = new Set(reads.filter(r => String(r['ر.التسجيل']).trim() === String(s.regId).trim())
                                  .map(r => r['معرف المنشور']));

  // بيانات الاستبيانات: الخيارات المتاحة لكل منشور، وإجابات طالبتي تحديداً
  const pollOptionsAll = sheetToObjects_(getSheet_(SHEET_POLL_OPTIONS));
  const myPollAnswers = sheetToObjects_(getSheet_(SHEET_POLL_ANSWERS))
    .filter(a => String(a['ر.التسجيل']).trim() === String(s.regId).trim());
  const myAnsweredPollIds = new Set(myPollAnswers.map(a => a['معرف المنشور']));

  const nowMs = Date.now();
  const feed = posts.filter(p => {
    if (isPostScheduledFuture_(p, nowMs)) return false; // مجدول لموعد لم يحن بعد
    if (isPostExpired_(p, nowMs)) return false; // انتهت مدة ظهوره: يُحجب عن الأولياء
    return postTargetsStudent_(p, { regId: s.regId, level: s.level, group: s.group, isNew: iAmNew, internship: myInternship });
  }).map(p => ({
    id: p['معرف المنشور'],
    // التاريخ المعروض: تاريخ بدء الظهور إن حُدد (منشور مجدول)، وإلا تاريخ الإنشاء
    datetime: toIsoStr_(!isNaN(postStartMs_(p)) ? new Date(postStartMs_(p)) : p['التاريخ والوقت']),
    type: p['النوع'],
    text: p['النص'],
    mediaUrl: toDisplayMediaUrl_(p['رابط الوسائط'], p['النوع']),
    fileName: p['اسم الملف'],
    publisher: p['الناشر'],
    pinned: isPinned_(p),
    read: myReadIds.has(p['معرف المنشور']),
    options: (p['النوع'] === 'استبيان') ? pollOptionsAll.filter(o => o['معرف المنشور'] === p['معرف المنشور']).map(o => o['الخيار']) : undefined,
    answered: (p['النوع'] === 'استبيان') ? myAnsweredPollIds.has(p['معرف المنشور']) : undefined
  }));

  // المثبتة أولاً، ثم الأحدث فالأقدم
  feed.sort((a, b) => (b.pinned - a.pinned) || (new Date(b.datetime) - new Date(a.datetime)));
  const result = { student: { name: s.name, level: s.level, group: s.group }, posts: feed };
  return result;
}

function markPostAsRead(token, postId) {
  const s = requireSession_(token, 'parent');
  const lock = LockService.getScriptLock();
  const locked = lock.tryLock(5000);
  try {
    const sheet = getSheet_(SHEET_READS);
    const reads = sheetToObjects_(sheet);
    const already = reads.some(r => r['معرف المنشور'] === postId && String(r['ر.التسجيل']).trim() === String(s.regId).trim());
    if (!already) {
      sheet.appendRow([postId, s.regId, new Date()]);
    }
    return true;
  } finally {
    if (locked) lock.releaseLock();
  }
}

/* ===================== المتابعة والتحصيل (اطّلاع الولي) ===================== */
// إرجاع سطور شيت "المتابعة" الخاصة بطالبة الولي المسجّل دخوله فقط (بحسب ر.التسجيل)
function getParentTracking(token) {
  const s = requireSession_(token, 'parent');
  const rows = sheetToObjects_(getSheet_(SHEET_TRACKING));
  return rows
    .filter(r => norm_(r['ر.التسجيل']) === norm_(s.regId))
    .map(r => ({
      className: norm_(r['الفصل']),
      late: norm_(r['التأخرات']),
      absence: norm_(r['الغيابات']),
      honors: norm_(r['التكريمات'])
    }));
}

// إرجاع سطور شيت "التحصيل" الخاصة بطالبة الولي المسجّل دخوله فقط (بحسب ر.التسجيل)
function getParentCollection(token) {
  const s = requireSession_(token, 'parent');
  const rows = sheetToObjects_(getSheet_(SHEET_COLLECTION));
  return rows
    .filter(r => norm_(r['ر.التسجيل']) === norm_(s.regId))
    .map(r => ({
      installment: norm_(r['القسط']),
      category: norm_(r['الصنف']),
      period: norm_(r['الفترة']),
      required: norm_(r['المطلوب']),
      collected: norm_(r['التحصيل'])
    }));
}

/* ===================== لوحة تحكم الإدارة: القوائم المنسدلة ===================== */
// كل شيء هنا مبنيّ حصراً من شيت "الطالبات" الفعلي — لا شيء يُقرأ من شيت الإعدادات
function adminGetLevelsAndGroups(token) {
  requireSession_(token, 'admin');
  ensureStudentColumns_();
  const students = sheetToObjects_(getSheet_(SHEET_STUDENTS));
  const levels = {};
  let hasNew = false;
  let hasInternal = false;      // توجد طالبة واحدة على الأقل "داخلي"
  let hasHalfInternal = false;  // توجد طالبة واحدة على الأقل "نصف داخلي"

  students.forEach(s => {
    const lvl = String(s['المستوى'] || '').trim();
    const grp = String(s['الفوج'] || '').trim();
    if (lvl) {
      if (!levels[lvl]) levels[lvl] = new Set();
      if (grp) levels[lvl].add(grp);
    }
    if (isNewStudent_(s)) hasNew = true;
    const internship = norm_(s['التربص']);
    if (internship === 'داخلي') hasInternal = true;
    if (internship === 'نصف داخلي') hasHalfInternal = true;
  });

  const levelsResult = {};
  Object.keys(levels).forEach(l => levelsResult[l] = Array.from(levels[l]));
  return { levels: levelsResult, hasNew: hasNew, hasInternal: hasInternal, hasHalfInternal: hasHalfInternal };
}

/* ===================== لوحة تحكم الإدارة: المنشورات ===================== */
function adminGetPosts(token) {
  requireSession_(token, 'admin');
  try { sweepExpiredPosts_(); } catch (e) { /* لا نكسر القائمة إن تعذر الكنس */ }
  const posts = sheetToObjects_(getSheet_(SHEET_POSTS));
  posts.sort((a, b) => new Date(b['التاريخ والوقت']) - new Date(a['التاريخ والوقت']));
  const nowMs = Date.now();
  const studentsById = {};
  sheetToObjects_(getSheet_(SHEET_STUDENTS)).forEach(st => { studentsById[String(st['ر.التسجيل']).trim()] = st['اسم الطالبة']; });
  return posts.map(p => {
    const mediaDeleted = norm_(p['بعد الانتهاء']) === AFTER_DELETED;
    const targetStudentId = norm_(p['الطالبة المستهدفة']);
    return {
      id: p['معرف المنشور'],
      datetime: toIsoStr_(p['التاريخ والوقت']), // Date غير مسموح في الرد، وإلا وصل null للمتصفح
      level: p['المستوى المستهدف'],
      group: p['الفوج المستهدف'],
      targetStudentId: targetStudentId,
      targetStudentName: targetStudentId ? (studentsById[targetStudentId] || targetStudentId) : '',
      type: p['النوع'],
      text: p['النص'],
      mediaUrl: mediaDeleted ? '' : toDisplayMediaUrl_(p['رابط الوسائط'], p['النوع']),
      fileName: p['اسم الملف'],
      publisher: p['الناشر'],
      durationDays: String(p['مدة الظهور (أيام)'] === undefined ? '' : p['مدة الظهور (أيام)']),
      afterExpiry: norm_(p['بعد الانتهاء'] === undefined ? '' : p['بعد الانتهاء']),
      expired: isPostExpired_(p, nowMs),
      pinned: isPinned_(p),
      scheduled: isPostScheduledFuture_(p, nowMs),
      startDate: isNaN(postStartMs_(p)) ? '' : new Date(postStartMs_(p)).toISOString()
    };
  });
}

// تثبيت منشور أو إلغاء تثبيته من قائمة الإدارة
// تعديل منشور موجود: النص، الاستهداف، مدة الظهور، التثبيت — النوع والوسائط المرفقة لا يتغيران
function adminUpdatePost(token, postId, payload) {
  const s = requireSession_(token, 'admin');
  ensurePostColumns_();
  const sheet = getSheet_(SHEET_POSTS);
  const posts = sheetToObjects_(sheet);
  const post = posts.find(p => p['معرف المنشور'] === postId);
  if (!post) throw new Error('المنشور غير موجود');

  let level = payload.level || TARGET_ALL;
  let group = payload.group || TARGET_ALL;

  let targetStudentId = '';
  if (norm_(payload.targetStudent) !== '') {
    const student = sheetToObjects_(getSheet_(SHEET_STUDENTS))
      .find(st => String(st['ر.التسجيل']).trim() === String(payload.targetStudent).trim());
    if (!student) throw new Error('الطالبة المستهدَفة غير موجودة');
    targetStudentId = norm_(payload.targetStudent);
    level = TARGET_STUDENT;
    group = '';
  }

  if (s.role === ROLE_QURAN) {
    const allowed = [TARGET_INTERNSHIP_PREFIX + 'داخلي', TARGET_INTERNSHIP_PREFIX + 'نصف داخلي'];
    if (allowed.indexOf(level) === -1) {
      throw new Error('ليست لديك صلاحية التعديل لهذا الاستهداف. يمكنك الاختيار فقط بين: التربص: داخلي / التربص: نصف داخلي');
    }
  }

  const text = payload.text !== undefined ? String(payload.text) : post['النص'];
  const parsedDays = parseFloat(payload.durationDays);
  const duration = (!isNaN(parsedDays) && parsedDays > 0) ? parsedDays : '';
  const afterExpiry = duration === '' ? '' : (norm_(payload.afterExpiry) === AFTER_DELETE ? AFTER_DELETE : AFTER_HIDE);
  const startDate = norm_(payload.startDate);
  const pinned = payload.pinned ? 'نعم' : '';

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(norm_);
  const setCol = (headerName, value) => {
    const idx = headers.indexOf(headerName);
    if (idx !== -1) sheet.getRange(post.__row, idx + 1).setValue(value);
  };
  setCol('المستوى المستهدف', level);
  setCol('الفوج المستهدف', group);
  setCol('النص', text);
  setCol('مدة الظهور (أيام)', duration);
  setCol('بعد الانتهاء', afterExpiry);
  setCol('تاريخ بدء الظهور', startDate);
  setCol('مثبت', pinned);
  setCol('الطالبة المستهدفة', targetStudentId);
  return true;
}

function adminTogglePin(token, postId) {
  requireSession_(token, 'admin');
  ensurePostColumns_();
  const sheet = getSheet_(SHEET_POSTS);
  const posts = sheetToObjects_(sheet);
  const post = posts.find(p => p['معرف المنشور'] === postId);
  if (!post) throw new Error('المنشور غير موجود');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(norm_);
  const col = headers.indexOf('مثبت') + 1;
  const newVal = isPinned_(post) ? '' : 'نعم';
  sheet.getRange(post.__row, col).setValue(newVal);
  return newVal === 'نعم';
}

function adminCreatePost(token, payload) {
  const s = requireSession_(token, 'admin');
  let level = payload.level || TARGET_ALL;
  let group = payload.group || TARGET_ALL;
  const type = payload.type; // نص | صورة | فيديو | تعليمات | استبيان
  const text = payload.text || '';

  // استهداف طالبة واحدة بعينها: يتجاوز المستوى/الفوج المختارين
  let targetStudentId = '';
  if (norm_(payload.targetStudent) !== '') {
    const student = sheetToObjects_(getSheet_(SHEET_STUDENTS))
      .find(st => String(st['ر.التسجيل']).trim() === String(payload.targetStudent).trim());
    if (!student) throw new Error('الطالبة المستهدَفة غير موجودة');
    targetStudentId = norm_(payload.targetStudent);
    level = TARGET_STUDENT;
    group = '';
  }

  // فرض صلاحية النشر بحسب دور المشرف: "مشرف القرآن الكريم" مقيَّد بفئتي التربص فقط
  if (s.role === ROLE_QURAN) {
    const allowed = [TARGET_INTERNSHIP_PREFIX + 'داخلي', TARGET_INTERNSHIP_PREFIX + 'نصف داخلي'];
    if (allowed.indexOf(level) === -1) {
      throw new Error('ليست لديك صلاحية النشر بهذا الاستهداف. يمكنك النشر فقط لفئتي: التربص: داخلي / التربص: نصف داخلي');
    }
  }

  let mediaUrl = '';
  let fileName = '';
  if (payload.fileData && payload.fileName) {
    const uploaded = uploadMediaToDrive_(payload.fileData, payload.fileName, payload.mimeType);
    mediaUrl = uploaded.url;
    fileName = uploaded.name;
  }

  const id = newId_();

  // إن كان المنشور استبياناً: نتحقق من وجود خيارين على الأقل (بلا تكرار وبلا خانات فارغة)
  let pollOptions = [];
  if (norm_(type) === 'استبيان') {
    pollOptions = Array.isArray(payload.options) ? payload.options.map(norm_).filter(o => o !== '') : [];
    pollOptions = [...new Set(pollOptions)]; // إزالة التكرار
    if (pollOptions.length < 2) throw new Error('الاستبيان يحتاج خيارين مختلفين على الأقل');
    if (!text) throw new Error('الرجاء كتابة نص سؤال الاستبيان');
  }

  // مدة الظهور بالأيام (فارغة أو 0 = دائم) وخيار ما بعد الانتهاء (حجب / حذف من الدرايف)
  const parsedDays = parseFloat(payload.durationDays);
  const duration = (!isNaN(parsedDays) && parsedDays > 0) ? parsedDays : '';
  const afterExpiry = duration === '' ? '' : (norm_(payload.afterExpiry) === AFTER_DELETE ? AFTER_DELETE : AFTER_HIDE);

  // تاريخ بدء الظهور (منشور مجدول) والتثبيت أعلى الصفحة
  const startDate = norm_(payload.startDate); // بصيغة datetime-local من المتصفح، أو فارغ = فوراً
  const pinned = payload.pinned ? 'نعم' : '';

  ensurePostColumns_(); // يضيف الأعمدة تلقائياً للشيتات القائمة قبل الإصدار الحالي
  const lock = LockService.getScriptLock();
  const locked = lock.tryLock(5000);
  try {
    getSheet_(SHEET_POSTS).appendRow([id, new Date(), level, group, type, text, mediaUrl, fileName, s.fullName || s.username, duration, afterExpiry, startDate, pinned, targetStudentId]);
    if (pollOptions.length) {
      const optSheet = getSheet_(SHEET_POLL_OPTIONS);
      pollOptions.forEach(opt => optSheet.appendRow([id, opt]));
    }
  } finally {
    if (locked) lock.releaseLock();
  }

  // إشعار الهواتف: يُرسل فوراً للمنشورات العادية فقط
  // (لا يُرسل للمنشور المجدول مستقبلاً، ولا لمنشور طالبة محددة حفاظاً على الخصوصية)
  var isScheduledFuture = startDate !== '' && postStartMs_({ 'تاريخ بدء الظهور': startDate }) > Date.now();
  if (!isScheduledFuture && level !== TARGET_STUDENT) {
    var notifText = text ? String(text).substring(0, 80) : ('منشور جديد: ' + type);
    sendPushNotification_('مؤسسة النور — همزة وصل', notifText);
  }

  return { id: id };
}

/* ===================== إرسال إشعار للهواتف (OneSignal) ===================== */
// يُرسل إشعاراً لجميع الهواتف المفعّلة للإشعارات في تطبيق همزة وصل.
// لا يوقف النشر إطلاقاً إذا فشل الإرسال (خطأ شبكة أو مفاتيح فارغة).
function sendPushNotification_(title, message) {
  if (!ONESIGNAL_APP_ID || !ONESIGNAL_REST_API_KEY) return; // المفاتيح غير مهيأة بعد
  try {
    var body = {
      app_id: ONESIGNAL_APP_ID,
      included_segments: ['Total Subscriptions'],
      headings: { en: title, ar: title },
      contents: { en: message, ar: message }
    };
    if (PWA_APP_URL) body.url = PWA_APP_URL;
    UrlFetchApp.fetch('https://api.onesignal.com/notifications', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Basic ' + ONESIGNAL_REST_API_KEY },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
  } catch (err) {
    // نتجاهل أي خطأ: الإشعار خدمة إضافية ولا يجوز أن يكسر النشر
  }
}

function adminDeletePost(token, postId) {
  requireSession_(token, 'admin');
  const sheet = getSheet_(SHEET_POSTS);
  const posts = sheetToObjects_(sheet);
  const target = posts.find(p => p['معرف المنشور'] === postId);
  if (target) sheet.deleteRow(target.__row);

  // حذف سجلات القراءة المرتبطة أيضاً
  const readSheet = getSheet_(SHEET_READS);
  const reads = sheetToObjects_(readSheet);
  const rowsToDelete = reads.filter(r => r['معرف المنشور'] === postId).map(r => r.__row).sort((a, b) => b - a);
  rowsToDelete.forEach(r => readSheet.deleteRow(r));

  // حذف خيارات وإجابات الاستبيان المرتبطة (إن وُجدت)
  const optSheet = getSheet_(SHEET_POLL_OPTIONS);
  const optRows = sheetToObjects_(optSheet).filter(o => o['معرف المنشور'] === postId).map(o => o.__row).sort((a, b) => b - a);
  optRows.forEach(r => optSheet.deleteRow(r));

  const ansSheet = getSheet_(SHEET_POLL_ANSWERS);
  const ansRows = sheetToObjects_(ansSheet).filter(a => a['معرف المنشور'] === postId).map(a => a.__row).sort((a, b) => b - a);
  ansRows.forEach(r => ansSheet.deleteRow(r));

  return true;
}

// إحصائية القراءة: من قرأ ومن لم يقرأ بعد لمنشور معيّن
function adminGetReadStats(token, postId) {
  requireSession_(token, 'admin');
  const posts = sheetToObjects_(getSheet_(SHEET_POSTS));
  const post = posts.find(p => p['معرف المنشور'] === postId);
  if (!post) throw new Error('المنشور غير موجود');

  const students = sheetToObjects_(getSheet_(SHEET_STUDENTS)).filter(st =>
    postTargetsStudent_(post, { regId: st['ر.التسجيل'], level: st['المستوى'], group: st['الفوج'], isNew: isNewStudent_(st), internship: norm_(st['التربص']) || 'غير معنية' })
  );

  const reads = sheetToObjects_(getSheet_(SHEET_READS)).filter(r => r['معرف المنشور'] === postId);
  const readIds = new Set(reads.map(r => norm_(r['ر.التسجيل'])));

  const readList = [];
  const unreadList = [];
  students.forEach(st => {
    const item = {
      regId: st['ر.التسجيل'],
      name: st['اسم الطالبة'],
      phone: String(st['هاتف الأب1'] || st['هاتف الأم'] || st['هاتف الأب2'] || '')
    };
    if (readIds.has(String(st['ر.التسجيل']).trim())) readList.push(item);
    else unreadList.push(item);
  });

  return { total: students.length, readCount: readList.length, unreadCount: unreadList.length, readList: readList, unreadList: unreadList };
}

// إرسال إجابة الولي على استبيان معيّن (إجابة نهائية، غير قابلة للتعديل، اختيار متعدد)
function parentSubmitPollAnswer(token, postId, selectedOptions) {
  const s = requireSession_(token, 'parent');
  if (!Array.isArray(selectedOptions) || !selectedOptions.length) throw new Error('الرجاء اختيار إجابة واحدة على الأقل');

  const posts = sheetToObjects_(getSheet_(SHEET_POSTS));
  const post = posts.find(p => p['معرف المنشور'] === postId);
  if (!post) throw new Error('المنشور غير موجود');
  if (norm_(post['النوع']) !== 'استبيان') throw new Error('هذا المنشور ليس استبياناً');

  const validOptions = new Set(
    sheetToObjects_(getSheet_(SHEET_POLL_OPTIONS)).filter(o => o['معرف المنشور'] === postId).map(o => norm_(o['الخيار']))
  );
  const clean = [...new Set(selectedOptions.map(norm_))].filter(o => o !== '');
  if (!clean.length || clean.some(o => !validOptions.has(o))) throw new Error('خيار غير صالح');

  const lock = LockService.getScriptLock();
  const locked = lock.tryLock(5000);
  try {
    const ansSheet = getSheet_(SHEET_POLL_ANSWERS);
    const already = sheetToObjects_(ansSheet).some(a =>
      a['معرف المنشور'] === postId && String(a['ر.التسجيل']).trim() === String(s.regId).trim());
    if (already) throw new Error('لقد أجبتَ مسبقاً على هذا الاستبيان، ولا يمكن تعديل الإجابة');

    const now = new Date();
    clean.forEach(opt => ansSheet.appendRow([postId, s.regId, opt, now]));

    // نعتبر الإجابة على الاستبيان بمثابة قراءته أيضاً
    const readSheet = getSheet_(SHEET_READS);
    const alreadyRead = sheetToObjects_(readSheet).some(r =>
      r['معرف المنشور'] === postId && String(r['ر.التسجيل']).trim() === String(s.regId).trim());
    if (!alreadyRead) readSheet.appendRow([postId, s.regId, now]);
  } finally {
    if (locked) lock.releaseLock();
  }
  return true;
}

// نتائج الاستبيان: عدد الأصوات لكل خيار (للإدارة فقط، لا يراها الولي)
function adminGetPollResults(token, postId) {
  requireSession_(token, 'admin');
  const posts = sheetToObjects_(getSheet_(SHEET_POSTS));
  const post = posts.find(p => p['معرف المنشور'] === postId);
  if (!post) throw new Error('المنشور غير موجود');
  if (norm_(post['النوع']) !== 'استبيان') throw new Error('هذا المنشور ليس استبياناً');

  const options = sheetToObjects_(getSheet_(SHEET_POLL_OPTIONS)).filter(o => o['معرف المنشور'] === postId).map(o => o['الخيار']);
  const answers = sheetToObjects_(getSheet_(SHEET_POLL_ANSWERS)).filter(a => a['معرف المنشور'] === postId);

  const results = options.map(opt => ({
    option: opt,
    count: answers.filter(a => norm_(a['الخيار']) === norm_(opt)).length
  }));
  const totalVoters = new Set(answers.map(a => String(a['ر.التسجيل']).trim())).size;

  // عدد الطالبات المستهدَفات بالمنشور، لحساب نسبة المشاركة
  const totalStudents = sheetToObjects_(getSheet_(SHEET_STUDENTS)).filter(st =>
    postTargetsStudent_(post, { regId: st['ر.التسجيل'], level: st['المستوى'], group: st['الفوج'], isNew: isNewStudent_(st), internship: norm_(st['التربص']) || 'غير معنية' })
  ).length;

  return { question: post['النص'], results: results, totalVoters: totalVoters, totalStudents: totalStudents };
}

/* ===================== رفع الوسائط إلى Drive ===================== */
function getAttachmentsFolder_() {
  // الأولوية 1: قيمة "ملف الدرايف" في شيت الإعدادات (رابط المجلد أو معرفه)
  const driveSetting = norm_(settingByKeyword_(getSettings_(), ['الدرايف', 'درايف', 'Drive']));
  if (driveSetting) {
    try {
      return DriveApp.getFolderById(extractDriveId_(driveSetting));
    } catch (e) {
      throw new Error('قيمة "ملف الدرايف" في شيت الإعدادات غير صحيحة أو لا تملك صلاحية الوصول للمجلد. ضع رابط المجلد أو معرفه في عمود "الموقع".');
    }
  }
  // الأولوية 2: الثابت DRIVE_FOLDER_ID في أعلى الملف
  if (DRIVE_FOLDER_ID) {
    try {
      return DriveApp.getFolderById(DRIVE_FOLDER_ID);
    } catch (e) {
      throw new Error('معرف مجلد Drive (DRIVE_FOLDER_ID) غير صحيح أو لا تملك صلاحية الوصول إليه. تحقق من القيمة في أعلى ملف Code.gs.');
    }
  }
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function uploadMediaToDrive_(base64Data, fileName, mimeType) {
  const folder = getAttachmentsFolder_();
  const cleanBase64 = base64Data.indexOf(',') > -1 ? base64Data.split(',')[1] : base64Data;
  const bytes = Utilities.base64Decode(cleanBase64);
  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return { url: 'https://drive.google.com/uc?export=view&id=' + file.getId(), fileId: file.getId(), name: fileName };
}

// تُستخدم من لوحة الإدارة لعرض رابط مجلد المرفقات الحالي (مفيد لمعرفة أين تُحفظ الملفات)
function adminGetAttachmentsFolderLink(token) {
  requireSession_(token, 'admin');
  const folder = getAttachmentsFolder_();
  return { name: folder.getName(), url: folder.getUrl() };
}

/* ===================== لوحة تحكم الإدارة: إدارة الطالبات ===================== */
function adminGetStudents(token) {
  requireSession_(token, 'admin');
  ensureStudentColumns_();
  const students = sheetToObjects_(getSheet_(SHEET_STUDENTS));
  return students.map(s => ({
    regId: s['ر.التسجيل'],
    name: s['اسم الطالبة'],
    level: s['المستوى'],
    group: s['الفوج'],
    phone1: s['هاتف الأب1'],
    phone2: s['هاتف الأب2'],
    phone3: s['هاتف الأم'],
    internship: norm_(s['التربص']) || 'غير معنية'
  }));
}

// إحصائية عدد الطالبات في كل مستوى دراسي (تُبنى من شيت الطالبات مباشرة)
function adminGetLevelStats(token) {
  requireSession_(token, 'admin');
  const students = sheetToObjects_(getSheet_(SHEET_STUDENTS));
  const counts = {};
  students.forEach(s => {
    const lvl = norm_(s['المستوى']) || 'بلا مستوى';
    counts[lvl] = (counts[lvl] || 0) + 1;
  });
  const rows = Object.keys(counts).map(lvl => ({ level: lvl, count: counts[lvl] }))
    .sort((a, b) => b.count - a.count);
  return { rows: rows, total: students.length };
}

// إحصائية عدد الطالبات: الجديدات + كل صنف من "التربص" عدا "غير معنية" (لا تحمل دلالة إحصائية)
function adminGetInternshipStats(token) {
  requireSession_(token, 'admin');
  ensureStudentColumns_();
  const students = sheetToObjects_(getSheet_(SHEET_STUDENTS));
  const counts = { 'داخلي': 0, 'نصف داخلي': 0 };
  let newCount = 0;
  students.forEach(s => {
    const v = norm_(s['التربص']);
    if (counts.hasOwnProperty(v)) counts[v]++;
    if (isNewStudent_(s)) newCount++;
  });
  return { counts: counts, newCount: newCount, total: students.length };
}

function adminAddStudent(token, data) {
  requireSession_(token, 'admin');
  ensureStudentColumns_();
  const lock = LockService.getScriptLock();
  const locked = lock.tryLock(5000);
  try {
    const sheet = getSheet_(SHEET_STUDENTS);
    const students = sheetToObjects_(sheet);
    if (students.some(s => String(s['ر.التسجيل']).trim() === String(data.regId).trim())) {
      throw new Error('رقم التسجيل مستخدم من قبل');
    }
    // كلمة مرور خاصة توزع على الولي: يكتبها المشرف أو تُولَّد عشوائياً إن تُركت فارغة
    const pass = norm_(data.password) || generatePassword_();
    const internship = INTERNSHIP_OPTIONS.indexOf(norm_(data.internship)) !== -1 ? norm_(data.internship) : 'غير معنية';
    sheet.appendRow([data.regId, data.name, data.level, data.group, data.phone1 || '', data.phone2 || '', data.phone3 || '', pass, '', internship]);
    return { password: pass };
  } finally {
    if (locked) lock.releaseLock();
  }
}

function adminUpdateStudent(token, regId, data) {
  requireSession_(token, 'admin');
  ensureStudentColumns_();
  const sheet = getSheet_(SHEET_STUDENTS);
  const students = sheetToObjects_(sheet);
  const student = students.find(s => String(s['ر.التسجيل']).trim() === String(regId).trim());
  if (!student) throw new Error('الطالبة غير موجودة');

  const internship = INTERNSHIP_OPTIONS.indexOf(norm_(data.internship)) !== -1 ? norm_(data.internship) : (norm_(student['التربص']) || 'غير معنية');
  const rowRange = sheet.getRange(student.__row, 1, 1, STUDENT_HEADERS.length);
  const values = [regId, data.name, data.level, data.group, data.phone1 || '', data.phone2 || '', data.phone3 || '', student['كلمة المرور'], student['جديدة'] || '', internship];
  rowRange.setValues([values]);
  return true;
}

function adminDeleteStudent(token, regId) {
  requireSession_(token, 'admin');
  const sheet = getSheet_(SHEET_STUDENTS);
  const students = sheetToObjects_(sheet);
  const student = students.find(s => String(s['ر.التسجيل']).trim() === String(regId).trim());
  if (student) sheet.deleteRow(student.__row);
  return true;
}

function adminResetStudentPassword(token, regId) {
  requireSession_(token, 'admin');
  const sheet = getSheet_(SHEET_STUDENTS);
  const students = sheetToObjects_(sheet);
  const student = students.find(s => String(s['ر.التسجيل']).trim() === String(regId).trim());
  if (!student) throw new Error('الطالبة غير موجودة');
  const newPass = generatePassword_(); // كلمة عشوائية جديدة تُسلَّم للولي (لم يعد رقم الهاتف يُستخدم)
  const colIndex = STUDENT_HEADERS.indexOf('كلمة المرور') + 1;
  sheet.getRange(student.__row, colIndex).setValue(newPass);
  return newPass;
}