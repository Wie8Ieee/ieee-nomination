/**
 * IEEE Al-Quds University Student Branch Elections 2026 - Interactive Nomination Web App
 * الملفات: Code.gs (هاد الملف) + Index.html
 */

// ===== إعدادات =====
const CFG = {
  SHEET_ID: '1Omg9XVNECpcX6yKF8LmPJUfv8Mj3iQH1EKap4-IjxRM', // شيت الردود اللي انعمل قبل
  SHEET_NAME: 'Nominations',                                   // التبويب الجديد اللي رح ينعمل
  CONTACT_EMAIL: 'drsafa@staff.alquds.edu',                                           // مثال: 'ieee@students.alquds.edu' (اتركيه فاضي لإخفائه)
  NOTIFY_EMAILS: ['drsafa@staff.alquds.edu'],                  // كل ترشيح جديد بيوصلهم على الإيميل مع كل التفاصيل
  MEMBERSHIP_NUMBER_REQUIRED: true,                           // true إذا د. صفاء طلبت رقم العضوية إجباري
  SEND_CONFIRMATION_EMAIL: true,                               // إيميل تأكيد للمرشح
  CLOSED: false                                                // true لإغلاق باب الترشح (بعدها New version)
};

const POSITIONS = [
  'Chairperson', 'Vice Chairperson', 'Secretary', 'Chair for WIE',
  'Events & Activities Coordinator', 'PR & Marketing Team', 'Webmaster', 'Technical Team'
];

const HEADERS = [
  'Timestamp', 'Full Name', 'Student ID', 'University Email', 'Phone', 'Major', 'Academic Year',
  'Active IEEE Member', 'IEEE Membership Number', 'Position', 'Why this position', 'Skills',
  'Previous experience', 'Contribution', 'Initiative idea', 'Commit to term', 'Attend events', 'Confirmed'
];

// ===== صفحة الويب =====
function doGet(e) {
  // لنسخة GitHub: ?config=1 بيرجّع الإعدادات كـ JSON
  if (e && e.parameter && e.parameter.config) {
    return json_({ ok: true, config: getConfig() });
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('IEEE Al-Quds SB Elections 2026 – Nomination')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// لنسخة GitHub: الصفحة بترسل الترشيح هون
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    return json_(submitNomination(d));
  } catch (err) {
    console.error(err);
    return json_({ ok: false, code: 'ERROR' });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// الصفحة بتطلب الإعدادات من هون عند التحميل
function getConfig() {
  return {
    deadline: CFG.DEADLINE,
    contact: CFG.CONTACT_EMAIL,
    membershipRequired: CFG.MEMBERSHIP_NUMBER_REQUIRED,
    closed: CFG.CLOSED
  };
}

// ===== استقبال الترشيح =====
function submitNomination(d) {
  if (CFG.CLOSED) return { ok: false, code: 'CLOSED' };

  const err = validate_(d);
  if (err) return { ok: false, code: 'INVALID', field: err };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (e) {
    return { ok: false, code: 'BUSY' };
  }

  try {
    const sheet = getSheet_();
    const studentId = String(d.studentId).trim();

    // منع تكرار الرقم الجامعي
    const last = sheet.getLastRow();
    if (last > 1) {
      const ids = sheet.getRange(2, 3, last - 1, 1).getValues().flat().map(v => String(v).replace(/^'/, '').trim());
      if (ids.indexOf(studentId) !== -1) return { ok: false, code: 'DUPLICATE' };
    }

    sheet.appendRow([
      new Date(),
      safe_(d.fullName), safe_(studentId), safe_(d.email), safe_(d.phone), safe_(d.major), safe_(d.year),
      safe_(d.active), safe_(d.membershipNo), safe_(d.position), safe_(d.why),
      safe_((d.skills || []).join(', ')), safe_(d.experience), safe_(d.contribution), safe_(d.idea),
      safe_(d.commitTerm), safe_(d.commitAttend), 'Yes'
    ]);

    sendNotification_(d);
    if (CFG.SEND_CONFIRMATION_EMAIL) sendConfirmation_(d);
    return { ok: true };
  } catch (e) {
    console.error(e);
    return { ok: false, code: 'ERROR' };
  } finally {
    lock.releaseLock();
  }
}

// ===== دوال مساعدة =====
function getSheet_() {
  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  let sh = ss.getSheetByName(CFG.SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CFG.SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sh;
}

function validate_(d) {
  if (!d) return 'all';
  const req = ['fullName', 'studentId', 'email', 'phone', 'major', 'year', 'active', 'position',
               'why', 'contribution', 'idea', 'commitTerm', 'commitAttend'];
  for (const k of req) {
    if (!d[k] || !String(d[k]).trim()) return k;
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(d.email).trim())) return 'email';
  if (POSITIONS.indexOf(d.position) === -1) return 'position';
  if (!Array.isArray(d.skills) || d.skills.length === 0) return 'skills';
  if (CFG.MEMBERSHIP_NUMBER_REQUIRED && !String(d.membershipNo || '').trim()) return 'membershipNo';
  if (d.confirm !== true) return 'confirm';
  for (const k in d) {
    if (typeof d[k] === 'string' && d[k].length > 3000) return k;
  }
  return null;
}

// يمنع حقن الصيغ في الشيت (=, +, -, @)
function safe_(v) {
  const s = String(v == null ? '' : v).trim();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

// إيميل للمراجِعة (د. صفاء) بكل ترشيح جديد وبكل التفاصيل
function sendNotification_(d) {
  try {
    if (!CFG.NOTIFY_EMAILS || !CFG.NOTIFY_EMAILS.length) return;
    const line = (k, v) => k + ': ' + (v && String(v).trim() ? v : '—');
    const body = [
      'New nomination received / ترشيح جديد',
      '',
      line('Position / المنصب', d.position),
      line('Full name / الاسم', d.fullName),
      line('Student ID / الرقم الجامعي', d.studentId),
      line('Email / البريد', d.email),
      line('Phone / الهاتف', d.phone),
      line('Major / التخصص', d.major),
      line('Academic year / السنة', d.year),
      line('Active IEEE member / العضوية مفعّلة', d.active),
      line('IEEE membership no. / رقم العضوية', d.membershipNo),
      '',
      line('Why this position / سبب الترشح', d.why),
      '',
      line('Skills / المهارات', (d.skills || []).join(', ')),
      '',
      line('Previous experience / الخبرة السابقة', d.experience),
      '',
      line('Contribution / ما سيقدمه', d.contribution),
      '',
      line('Initiative idea / المبادرة', d.idea),
      '',
      line('Commit to term / الالتزام بالفترة', d.commitTerm),
      line('Attend events / حضور الفعاليات', d.commitAttend),
      '',
      'All nominations: ' + SpreadsheetApp.openById(CFG.SHEET_ID).getUrl()
    ].join('\n');

    MailApp.sendEmail({
      to: CFG.NOTIFY_EMAILS.join(','),
      replyTo: String(d.email).trim(),
      subject: 'IEEE Nomination – ' + d.position + ' – ' + d.fullName,
      body: body
    });
  } catch (e) {
    console.error('Notify failed: ' + e);
  }
}

// شغّليها مرة وحدة من المحرر إذا بدك د. صفاء تقدر تفتح شيت الردود (صلاحية عرض فقط)
function shareSheetWithReviewers() {
  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  CFG.NOTIFY_EMAILS.forEach(e => ss.addViewer(e));
  Logger.log('Shared (view only) with: ' + CFG.NOTIFY_EMAILS.join(', '));
}

function sendConfirmation_(d) {
  try {
    const contact = CFG.CONTACT_EMAIL ? '\n\nللاستفسار / For questions: ' + CFG.CONTACT_EMAIL : '';
    MailApp.sendEmail({
      to: String(d.email).trim(),
      subject: 'IEEE Al-Quds SB Elections 2026 – تم استلام طلب الترشح',
      body:
        'مرحبًا ' + d.fullName + '،\n\n' +
        'تم استلام طلب ترشحك لمنصب: ' + d.position + '.\n' +
        'سيتم مراجعة الطلبات وفق الإجراءات المعتمدة، وسيتم التواصل معك بخصوص الخطوات القادمة.\n' +
        'ملاحظة: استلام الطلب لا يعني الأهلية النهائية للترشح.\n\n' +
        'Hello ' + d.fullName + ',\n\n' +
        'We received your nomination for: ' + d.position + '.\n' +
        'Nominations will be reviewed according to the approved procedures, and we will contact you about next steps.\n' +
        'Note: receiving your application does not confirm final eligibility.' + contact +
        '\n\nIEEE Al-Quds University Student Branch'
    });
  } catch (e) {
    console.error('Email failed: ' + e);
  }
}
