/**
 * אוטומציה: סיכום פגישה מ-Timeless  ->  קובץ Word  ->  מייל "טיוטה לאישורך"
 *
 * רץ כל 10 דקות (טריגר זמן), מחפש מיילים חדשים מ-Timeless שמכילים את סיכום הפגישה המלא,
 * ממיר את הסיכום לקובץ DOCX בעברית (RTL), ושולח אותו במייל עם הנדון:
 *     "<שם הפגישה> – טיוטה לאישורך"
 * כל מייל שטופל מסומן בתווית כדי שלא יישלח פעמיים.
 *
 * התקנה: ראו README.md שבאותה תיקייה.
 */

// ===================== הגדרות =====================
const CONFIG = {
  // לאן לשלוח את הטיוטה
  RECIPIENT: 'sigal@sigalaw.co.il',

  // סיומת הנדון (אחרי שם הפגישה)
  SUBJECT_SUFFIX: 'טיוטה לאישורך',

  // שאילתת החיפוש ב-Gmail: מיילים של Timeless שמכילים את הסיכום המלא
  SEARCH_QUERY: 'from:hey@magical.team subject:"shared meeting notes with you"',

  // תווית שמסמנת מייל שכבר טופל
  PROCESSED_LABEL: 'Timeless/נשלח כטיוטה',

  // לא לטפל במיילים ישנים מזה (מונע הצפה בהפעלה הראשונה)
  MAX_AGE: '3d',

  // גופן המסמך
  FONT: 'David',
};
// ==================================================

/** הפונקציה הראשית – מופעלת ע"י הטריגר. */
function processTimelessEmails() {
  const label = getOrCreateLabel_(CONFIG.PROCESSED_LABEL);
  const query = `${CONFIG.SEARCH_QUERY} newer_than:${CONFIG.MAX_AGE} -label:"${CONFIG.PROCESSED_LABEL}"`;
  const threads = GmailApp.search(query, 0, 20);

  threads.forEach((thread) => {
    try {
      thread.getMessages().forEach((msg) => {
        if (!/magical\.team/i.test(msg.getFrom())) return;
        handleMessage_(msg);
      });
      thread.addLabel(label);
    } catch (err) {
      console.error(`שגיאה בטיפול בשרשור ${thread.getId()}: ${err && err.stack || err}`);
    }
  });
}

/** טיפול בהודעה בודדת: חילוץ, יצירת DOCX ושליחה. */
function handleMessage_(msg) {
  const html = msg.getBody();
  const title = extractTitle_(html, msg.getSubject());
  const summaryHtml = extractSummary_(html);
  if (!summaryHtml) {
    console.warn(`לא נמצא סיכום בהודעה "${msg.getSubject()}" – דילוג`);
    return;
  }

  const docHtml = buildDocHtml_(title, summaryHtml);
  const fileName = `${sanitizeFileName_(title)} - טיוטה.docx`;
  const docx = htmlToDocx_(docHtml, fileName);

  const subject = `${title} – ${CONFIG.SUBJECT_SUFFIX}`;
  const body =
    'שלום סיגל,\n\n' +
    `מצורפת טיוטת סיכום הפגישה "${title}" לאישורך.\n\n` +
    '(נוצר אוטומטית מתוך Timeless)';
  GmailApp.sendEmail(CONFIG.RECIPIENT, subject, body, {
    attachments: [docx],
    name: 'Sigalaw – סיכומי פגישות',
  });
  console.log(`נשלח: ${subject}`);
}

// ---------------- חילוץ תוכן ----------------

function extractTitle_(html, subject) {
  const m = html.match(/<div class="event-title">([\s\S]*?)<\/div>/i);
  if (m) return decodeEntities_(stripTags_(m[1])).trim();
  // גיבוי: "<שם הפגישה> - <שם> shared meeting notes with you"
  return subject.replace(/\s+-\s+[^-]*shared meeting notes with you\s*$/i, '').trim();
}

function extractSummary_(html) {
  const m = html.match(/<div class="summary">([\s\S]*?)<\/div>\s*<\/td>/i);
  if (!m) return null;
  let s = m[1];
  s = s.replace(/<floating-media[^>]*>\s*<\/floating-media>/gi, ''); // תגיות זמן של Timeless
  s = s.replace(/<h1>[\s\S]*?<\/h1>/i, ''); // הכותרת תתווסף בנפרד
  // משימות: תיבת סימון
  s = s.replace(/<li data-type="taskItem" data-checked="true">/gi, '<li>☑ ');
  s = s.replace(/<li data-type="taskItem" data-checked="false">/gi, '<li>☐ ');
  s = s.replace(/<ul data-type="taskList">/gi, '<ul style="list-style:none">');
  return s;
}

function buildDocHtml_(title, summaryHtml) {
  const rtl = `direction:rtl;text-align:right;font-family:${CONFIG.FONT};`;
  // הוספת RTL לכל בלוק
  const body = summaryHtml.replace(
    /<(p|h1|h2|h3|ul|ol|li)(\s[^>]*)?>/gi,
    (all, tag, attrs) => {
      attrs = attrs || '';
      if (/style="/i.test(attrs)) {
        attrs = attrs.replace(/style="/i, `dir="rtl" style="${rtl}`);
      } else {
        attrs += ` dir="rtl" style="${rtl}"`;
      }
      return `<${tag}${attrs}>`;
    }
  );
  return (
    '<html dir="rtl"><head><meta charset="UTF-8"></head>' +
    `<body dir="rtl" style="${rtl}font-size:12pt;">` +
    `<p dir="rtl" style="${rtl}color:#C00000;font-weight:bold;">טיוטה לאישורך</p>` +
    `<h1 dir="rtl" style="${rtl}">${escapeHtml_(title)}</h1>` +
    body +
    '</body></html>'
  );
}

// ---------------- המרה ל-DOCX דרך Google Drive ----------------

function htmlToDocx_(html, fileName) {
  const token = ScriptApp.getOAuthToken();
  const boundary = 'sigalaw' + Date.now();
  const metadata = {
    name: fileName.replace(/\.docx$/, ''),
    mimeType: 'application/vnd.google-apps.document',
  };
  const payload =
    `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n` +
    JSON.stringify(metadata) +
    `\r\n--${boundary}\r\nContent-Type: text/html; charset=UTF-8\r\n\r\n` +
    html +
    `\r\n--${boundary}--`;

  // 1. העלאה + המרה ל-Google Doc
  const createRes = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart',
    {
      method: 'post',
      contentType: `multipart/related; boundary=${boundary}`,
      payload: Utilities.newBlob(payload, 'text/plain').getBytes(),
      headers: { Authorization: `Bearer ${token}` },
    }
  );
  const fileId = JSON.parse(createRes.getContentText()).id;

  try {
    // 2. ייצוא כ-DOCX
    const exportRes = UrlFetchApp.fetch(
      `https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=` +
        encodeURIComponent('application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
      { headers: { Authorization: `Bearer ${token}` } }
    );
    return exportRes.getBlob().setName(fileName);
  } finally {
    // 3. מחיקת קובץ הביניים מ-Drive
    UrlFetchApp.fetch(`https://www.googleapis.com/drive/v3/files/${fileId}`, {
      method: 'delete',
      headers: { Authorization: `Bearer ${token}` },
      muteHttpExceptions: true,
    });
  }
}

// ---------------- התקנה ----------------

/** הפעילו פעם אחת ידנית: יוצר טריגר שרץ כל 10 דקות. */
function installTrigger() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'processTimelessEmails')
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processTimelessEmails').timeBased().everyMinutes(10).create();
  getOrCreateLabel_(CONFIG.PROCESSED_LABEL);
  console.log('הטריגר הותקן – הסקריפט ירוץ כל 10 דקות.');
}

/** בדיקה: שולח טיוטה עבור מייל ה-Timeless האחרון (גם אם כבר טופל). */
function testOnLatest() {
  const threads = GmailApp.search(CONFIG.SEARCH_QUERY, 0, 1);
  if (!threads.length) throw new Error('לא נמצא מייל Timeless מתאים');
  const msgs = threads[0].getMessages().filter((m) => /magical\.team/i.test(m.getFrom()));
  handleMessage_(msgs[msgs.length - 1]);
}

// ---------------- עזר ----------------

function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function stripTags_(s) {
  return s.replace(/<[^>]+>/g, '');
}

function decodeEntities_(s) {
  return s
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&');
}

function escapeHtml_(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function sanitizeFileName_(s) {
  return s.replace(/[\\/:*?"<>|]/g, '').replace(/\s+/g, ' ').trim().slice(0, 120);
}
