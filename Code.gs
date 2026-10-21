/**
 * Charlene Dellosa – Agent Recruitment Landing Page
 * Google Apps Script + Google Sheets
 */
// Short cache keeps page loads fast while allowing quiz edits to appear quickly.
const QUIZ_CACHE_KEY = 'quiz_json_v1';
const QUIZ_CACHE_SECONDS = 60;
let spreadsheetInstance_;

const CONFIG = {
  SPREADSHEET_ID: '',      // Leave empty if this script is bound to your Sheet (Extensions > Apps Script)
  SHEET_NAME: 'Applicants',
  NOTIFY_EMAIL: '',        // Type the email that should receive applicant alerts (recommended). Empty = the account that owns this script
  SEND_CONFIRMATION: true
};

/**
 * Quiz content lives in the spreadsheet tabs below.
 * Run setupQuizSheets() once from the Apps Script editor to create and seed them.
 */
const QUIZ_QUESTIONS_SHEET = 'Quiz Questions';
const QUIZ_OPTIONS_SHEET = 'Quiz Options';
const QUIZ_QUESTION_HEADERS = [
  'Active', 'Order', 'Question ID', 'Short label', 'Question (English)', 'Question (Tagalog)'
];
const QUIZ_OPTION_HEADERS = [
  'Question ID', 'Order', 'Answer (English)', 'Answer (Tagalog)'
];
const APPLICANT_FIXED_HEADERS = ['Timestamp', 'Full name', 'Mobile', 'Email', 'City'];
const APPLICANT_FINAL_HEADERS = ['Review', 'Follow-up path', 'Status', 'Language', 'Consent'];

// Starter content used only when setupQuizSheets() creates the quiz tabs for the first time.
// After setup, edit these values in the spreadsheet; the live quiz reads from the tabs.
const DEFAULT_QUIZ_QUESTIONS = [
  [true, 1, 'license', 'License', 'Do you have a PRC real estate license?', 'May PRC real estate license ka ba?'],
  [true, 2, 'experience', 'Experience', 'Have you worked in sales, customer service, or another client-facing role?', 'Nakapagtrabaho ka na ba sa sales, customer service, o ibang trabahong may kliyente?'],
  [true, 3, 'hours', 'Weekly hours', 'How many hours per week can you commit to real estate?', 'Ilang oras kada linggo ang maibibigay mo sa real estate?'],
  [true, 4, 'commission', 'Commission', 'Real estate income is commission-based. How do you feel about that?', 'Ang kita sa real estate ay batay sa komisyon. Ano ang pakiramdam mo rito?'],
  [true, 5, 'marketing', 'Marketing', 'Are you ready to use your phone and social media to market listings and follow up leads?', 'Handa ka bang gumamit ng phone at social media para i-market ang listings at i-follow up ang leads?'],
  [true, 6, 'start_date', 'Start date', 'When can you start?', 'Kailan ka puwedeng magsimula?']
];
const DEFAULT_QUIZ_OPTIONS = [
  ['license', 1, 'Yes, I am licensed', 'Oo, lisensyado ako'],
  ['license', 2, 'I am studying or applying for it', 'Nag-aaral o nag-a-apply ako nito'],
  ['license', 3, 'Not yet', 'Wala pa'],
  ['experience', 1, 'Yes, 2 years or more', 'Oo, 2 taon o higit pa'],
  ['experience', 2, 'Yes, less than 2 years', 'Oo, wala pang 2 taon'],
  ['experience', 3, 'No experience yet', 'Wala pang experience'],
  ['hours', 1, 'Full-time (40+ hours a week)', 'Full-time (40+ oras kada linggo)'],
  ['hours', 2, 'Part-time (15 to 39 hours a week)', 'Part-time (15 hanggang 39 oras kada linggo)'],
  ['hours', 3, 'Fewer than 15 hours a week', 'Wala pang 15 oras kada linggo'],
  ['commission', 1, 'Comfortable. I can cover my needs while I build my sales', 'Komportable ako. Kaya kong tustusan ang pangangailangan ko habang bumubuo ng benta'],
  ['commission', 2, 'Okay with it, but I need income soon', 'Okay ako rito, pero kailangan ko ng kita agad'],
  ['commission', 3, 'I need a fixed salary', 'Kailangan ko ng fixed na sahod'],
  ['marketing', 1, 'Yes, I already do this', 'Oo, ginagawa ko na ito'],
  ['marketing', 2, 'Yes, I am willing to learn', 'Oo, handa akong matuto'],
  ['marketing', 3, 'No', 'Hindi'],
  ['start_date', 1, 'Within 2 weeks', 'Sa loob ng 2 linggo'],
  ['start_date', 2, 'Within a month', 'Sa loob ng isang buwan'],
  ['start_date', 3, 'More than a month from now', 'Mahigit isang buwan mula ngayon']
];

function getSpreadsheet_() {
  if (spreadsheetInstance_) return spreadsheetInstance_;

  spreadsheetInstance_ = CONFIG.SPREADSHEET_ID
    ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheetInstance_) {
    throw new Error('No active spreadsheet. Set CONFIG.SPREADSHEET_ID to your Google Sheet ID.');
  }
  return spreadsheetInstance_;
}

function styleHeader_(sheet, headers) {
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#0B3B2E').setFontColor('#FAF5E6');
  sheet.setFrozenRows(1);
}

function ensureQuizSheet_(ss, name, headers, starterRows) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    styleHeader_(sheet, headers);
    if (starterRows.length) sheet.getRange(2, 1, starterRows.length, headers.length).setValues(starterRows);
    sheet.autoResizeColumns(1, headers.length);
    return sheet;
  }

  if (sheet.getLastRow() === 0) {
    styleHeader_(sheet, headers);
    if (starterRows.length) sheet.getRange(2, 1, starterRows.length, headers.length).setValues(starterRows);
    sheet.autoResizeColumns(1, headers.length);
    return sheet;
  }

  const found = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0].map(function (v) { return String(v).trim(); });
  if (found.join('\u0001') !== headers.join('\u0001')) {
    throw new Error('The headers in "' + name + '" do not match the expected template. Keep the header row unchanged.');
  }
  styleHeader_(sheet, headers);
  return sheet;
}

function setupQuizSheets() {
  const ss = getSpreadsheet_();
  ensureQuizSheet_(ss, QUIZ_QUESTIONS_SHEET, QUIZ_QUESTION_HEADERS, DEFAULT_QUIZ_QUESTIONS);
  ensureQuizSheet_(ss, QUIZ_OPTIONS_SHEET, QUIZ_OPTION_HEADERS, DEFAULT_QUIZ_OPTIONS);
  console.log('Quiz sheets are ready: ' + ss.getUrl());
  return ss.getUrl();
}

function sheetRows_(sheet, width) {
  const lastRow = sheet.getLastRow();
  return lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, width).getValues() : [];
}

function cellText_(value) {
  return String(value == null ? '' : value).trim();
}

function orderValue_(value, fallback) {
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function getQuiz_() {
  const ss = getSpreadsheet_();
  const questionsSheet = ss.getSheetByName(QUIZ_QUESTIONS_SHEET);
  const optionsSheet = ss.getSheetByName(QUIZ_OPTIONS_SHEET);
  if (!questionsSheet || !optionsSheet) {
    throw new Error('Quiz sheets are missing. Run setupQuizSheets() once from the Apps Script editor.');
  }
  const qRows = sheetRows_(questionsSheet, QUIZ_QUESTION_HEADERS.length);
  const optionRows = sheetRows_(optionsSheet, QUIZ_OPTION_HEADERS.length);
  const questions = [];
  const byId = Object.create(null);
  const ids = Object.create(null);
  const labels = Object.create(null);

  qRows.forEach(function (row, index) {
    const id = cellText_(row[2]).toLowerCase();
    const label = clean_(row[3], 80);
    const q = clean_(row[4], 800);
    const tl = clean_(row[5], 800);
    if (!id && !label && !q && !tl) return; // ignore blank rows
    // Only explicit active values are live. Blank rows are treated as drafts.
    const activeText = cellText_(row[0]).toLowerCase();
    const isActive = ['true', 'yes', 'on', '1', 'active'].indexOf(activeText) >= 0;
    if (!isActive) return;
    if (!id || !label || !q || !tl) {
      throw new Error(
        'Incomplete active quiz row ' + (index + 2) + ' in "' + QUIZ_QUESTIONS_SHEET + '". ' +
        'Fill in Active, Question ID, Short label, Question (English), and Question (Tagalog), ' +
        'or set Active to FALSE for a draft row.'
      );
    }
    if (ids[id]) throw new Error('Question IDs must be unique. Duplicate ID: ' + id);
    const labelKey = label.toLowerCase();
    if (labels[labelKey]) throw new Error('Short labels must be unique because they are used as applicant-sheet columns: ' + label);
    ids[id] = true;
    labels[labelKey] = true;
    const question = { id: id, label: label, q: q, tl: tl, order: orderValue_(row[1], index + 1), sourceOrder: index, options: [] };
    byId[id] = question;
    questions.push(question);
  });

  if (!questions.length) throw new Error('No active quiz questions found in the "' + QUIZ_QUESTIONS_SHEET + '" tab.');

  optionRows.forEach(function (row, index) {
    const id = cellText_(row[0]).toLowerCase();
    const en = clean_(row[2], 500);
    const tl = clean_(row[3], 500);
    if (!id && !en && !tl) return; // ignore blank rows
    if (!byId[id]) return; // ignore options belonging to inactive or deleted questions
    if (!en || !tl) throw new Error('Every active answer choice needs English and Tagalog text (question ID: ' + id + ').');
    byId[id].options.push({ t: en, tl: tl, order: orderValue_(row[1], index + 1), sourceOrder: index });
  });

  questions.forEach(function (question) {
    question.options.sort(function (a, b) { return a.order - b.order || a.sourceOrder - b.sourceOrder; });
    if (question.options.length < 2) throw new Error('Each active question needs at least two answer choices: ' + question.label);
    delete question.order;
    delete question.sourceOrder;
    question.options.forEach(function (option) { delete option.order; delete option.sourceOrder; });
  });
  // Sort by the order shown in the sheet, preserving sheet order for ties.
  const orderById = Object.create(null);
  qRows.forEach(function (row, index) {
    const id = cellText_(row[2]).toLowerCase();
    if (id) orderById[id] = { order: orderValue_(row[1], index + 1), row: index };
  });
  questions.sort(function (a, b) {
    return orderById[a.id].order - orderById[b.id].order || orderById[a.id].row - orderById[b.id].row;
  });
  return questions;
}

function quizForBrowser_(quiz) {
  return quiz.map(function (x) {
    return { q: { en: x.q, tl: x.tl }, options: x.options.map(function (o) { return { en: o.t, tl: o.tl }; }) };
  });
}

function getQuizJson_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(QUIZ_CACHE_KEY);
  if (cached) return cached;

  const json = JSON.stringify(quizForBrowser_(getQuiz_()));
  cache.put(QUIZ_CACHE_KEY, json, QUIZ_CACHE_SECONDS);
  return json;
}

/** Run after editing Quiz Questions or Quiz Options to make changes appear immediately. */
function clearQuizCache() {
  CacheService.getScriptCache().remove(QUIZ_CACHE_KEY);
  return 'Quiz cache cleared.';
}

function applicantHeaders_(quiz) {
  return APPLICANT_FIXED_HEADERS.concat(quiz.map(function (x) { return x.label; })).concat(APPLICANT_FINAL_HEADERS);
}

/**
 * Keep old applications intact when the question labels/order change by starting a new Applicants vN tab.
 */
function getApplicantSheet_(quiz) {
  const ss = getSpreadsheet_();
  const headers = applicantHeaders_(quiz);
  for (let version = 1; version <= 100; version++) {
    const name = version === 1 ? CONFIG.SHEET_NAME : CONFIG.SHEET_NAME + ' v' + version;
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      styleHeader_(sheet, headers);
      return sheet;
    }
    if (sheet.getLastRow() <= 1) {
      const width = Math.max(headers.length, sheet.getLastColumn(), 1);
      sheet.getRange(1, 1, 1, width).clearContent();
      styleHeader_(sheet, headers);
      return sheet;
    }
    const width = sheet.getLastColumn();
    const existing = width ? sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(function (v) { return String(v).trim(); }) : [];
    if (existing.length === headers.length && existing.join('\u0001') === headers.join('\u0001')) return sheet;
  }
  throw new Error('Could not create an applicant sheet for the current quiz version.');
}

function doGet() {
  const t = HtmlService.createTemplateFromFile('Index');
  t.quizJson = getQuizJson_();
  return t.evaluate()
    .setTitle('Join Charlene Dellosa | Real Estate Agent Recruitment')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** JSON endpoint used by the Vercel /api/submit proxy. */
function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    return ContentService
      .createTextOutput(JSON.stringify(submitApplication(data)))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error(err);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: 'Invalid request.' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function clean_(v, max) {
  let s = String(v || '').trim().slice(0, max || 200);
  if (/^[=+\-@]/.test(s)) s = "'" + s;   // block spreadsheet formula injection
  return s;
}

/* ===================== EMAIL ===================== */
const BRAND = {
  sender: 'Charlene Dellosa Real Estate',
  dark: '#0B3B2E', emerald: '#0F7B5A', gold: '#B8902F', pale: '#E9D9A6', cream: '#FAF5E6', ink: '#22262A',
  photoId: '17VwMmZ45dHHYRfofzJBGn9shBdC1SaNP',   // Google Drive file ID of the email banner image
  bondingPhotoId: '1BxxzWUM3kLBy117C81sSQJxAahsSHis-', // Agent-bonding image for confirmation emails
  agentPortal: 'https://innersparcagentportal.vercel.app/',
  socials: [
    ['Facebook', 'https://www.facebook.com/CharleneDellosaProperties'],
    ['Instagram', 'https://www.instagram.com/chadlls13'],
    ['TikTok', 'https://www.tiktok.com/@dynamicpropertyagent'],
    ['YouTube', 'https://www.youtube.com/@chadellosa']
  ]
};


const MAIL_TEXT = {
  en: {
    greet: 'Hi {name},',
    subject: "You're One Step Closer to Your Real Estate Journey!",
    title: 'Your real estate journey starts here',
    intro: "Thank you for taking the time to share a little about yourself. It is always exciting to meet people who are curious about building a career in real estate, and we are happy to learn more about your goals.",
    intro2: "Real estate is not only about selling properties. It is also about meeting people, understanding what they need, answering questions, following up, and helping clients make important decisions with confidence. You do not have to know everything on day one. With the right training, guidance, and consistent effort, you can continue learning and grow step by step.",
    perksH: 'What you can look forward to with the team',
    perks: [
      'Practical training and mentorship to help you build confidence in real client conversations.',
      'A supportive agent community where you can learn, share ideas, and celebrate progress together.',
      'Access to current team programs, recognition, and incentive opportunities based on eligibility and performance.',
      'Possible team experiences, perks, and travel opportunities when included in eligible programs.',
      'A flexible path that can fit your goals while you build your own real estate business.'
    ],
    earningQ: 'Are you ready to start earning in real estate?',
    earningNote: 'Real estate income is commission-based, so earnings depend on completed transactions, current team terms, and your performance. The team will explain the details before you decide.',
    guideH: 'A few helpful tips as you get started',
    guide: [
      'Learn the basic requirements for becoming a licensed real estate professional in the Philippines and make a plan for completing the steps that apply to you.',
      'Be open to training and mentorship. The more you learn about properties, the market, documentation, and client service, the more prepared you will feel.',
      'Build your communication and follow-up habits. Many opportunities begin with a simple conversation and grow through timely, respectful follow-up.',
      'Use your phone and social media wisely. Sharing useful property information and staying in touch with your network can help you build visibility over time.',
      'Be patient and consistent. Real estate income is commission-based, so results may take time while you build your skills, relationships, and pipeline.'
    ],
    nextH: 'What happens next',
    next: [
      'A member of the team will reach out to schedule a friendly conversation and learn more about your goals.',
      'We will talk through the training, mentorship, schedule, and opportunities that may fit your current experience and availability.',
      'You can ask questions openly. Whether you are already licensed, still preparing, or completely new to real estate, we will help you understand the possible next steps.',
      'After the conversation, you can take your time, review your options, and decide whether this opportunity is right for you.'
    ],
    reminderH: 'A quick reminder',
    reminder: "There is no need to have everything figured out right away. A strong start comes from being willing to learn, communicate well, show up consistently, and take care of people. We look forward to getting to know you better.",
    btn: 'Explore the InnerSparc Agent Portal',
    sign: 'Warm regards,',
    priv: 'Your information is handled in line with the Data Privacy Act of 2012 (RA 10173).'
  },
  tl: {
    greet: 'Hi {name},',
    subject: 'Ang Bagong Career Mo sa Real Estate, Dito Nagsisimula!',
    title: 'Dito nagsisimula ang journey mo sa real estate',
    intro: 'Salamat sa paglalaan ng oras para magbahagi tungkol sa iyong sarili. Nakakatuwang makilala ang mga taong interesado at curious na magkaroon ng career sa real estate, at gusto naming mas makilala pa ang iyong mga goals.',
    intro2: 'Ang pagiging real estate agent ay hindi lang tungkol sa pagbebenta ng property. Kasama rito ang pakikipag-usap sa mga tao, pag-intindi sa kailangan nila, pagsagot sa mga tanong, pag-follow up, at pagtulong sa clients na makapagdesisyon nang maayos. Hindi mo kailangang alam na agad ang lahat. Sa tamang training, gabay, at tuloy-tuloy na effort, matututo ka at lalago step by step.',
    perksH: 'Ano ang puwedeng ma-enjoy kasama ang team',
    perks: [
      'Practical training at mentorship para mas maging confident ka sa pakikipag-usap sa clients.',
      'Supportive na community ng agents kung saan puwedeng matuto, mag-share ng ideas, at mag-celebrate ng progress.',
      'Access sa current team programs, recognition, at incentive opportunities depende sa eligibility at performance.',
      'Puwedeng magkaroon ng team experiences, perks, at travel opportunities kung kasama sa eligible programs.',
      'Flexible na path na puwedeng i-fit sa goals mo habang binubuo mo ang sarili mong real estate business.'
    ],
    earningQ: 'Ready ka na bang kumita sa real estate?',
    earningNote: 'Commission-based ang income sa real estate, kaya nakadepende ang earnings sa completed transactions, current team terms, at performance mo. Ipapaliwanag ng team ang mga detalye bago ka magdesisyon.',
    guideH: 'Ilang tips para makapagsimula',
    guide: [
      'Alamin ang mga basic requirement para maging licensed real estate professional sa Pilipinas, at gumawa ng simpleng plan para sa mga hakbang na kailangan mong gawin.',
      'Maging open sa training at mentorship. Habang mas natututo ka tungkol sa properties, market, documents, at pag-aalaga sa clients, mas magiging ready ka.',
      'Sanayin ang communication at follow-up. Maraming opportunities ang nagsisimula sa simpleng usapan at lumalaki dahil sa maayos at tamang follow-up.',
      'Gamitin nang maayos ang phone at social media. Makakatulong ang pag-share ng useful property information at pag-maintain ng connection sa iyong network.',
      'Maging patient at consistent. Commission-based ang income sa real estate, kaya maaaring kailangan ng panahon habang binubuo mo ang iyong skills, relationships, at leads.'
    ],
    nextH: 'Ano ang susunod na mangyayari',
    next: [
      'Kokontakin ka ng isang member ng team para mag-schedule ng friendly na usapan at malaman ang iyong goals.',
      'Pag-uusapan natin ang training, mentorship, schedule, at mga oportunidad na maaaring bagay sa iyong experience at availability.',
      'Puwede kang magtanong nang malaya. Licensed ka man, naghahanda pa lang, o wala pang experience sa real estate, tutulungan ka naming maintindihan ang possible next steps.',
      'Pagkatapos ng usapan, puwede mong pag-isipan ang options mo at magdesisyon kung tama para sa iyo ang opportunity na ito.'
    ],
    reminderH: 'Isang paalala',
    reminder: 'Hindi mo kailangang alam na agad ang lahat. Ang magandang simula ay ang pagiging willing matuto, marunong makipag-usap, consistent, at maalaga sa clients. Excited kaming mas makilala ka.',
    btn: 'Tingnan ang InnerSparc Agent Portal',
    sign: 'Maraming salamat,',
    priv: 'Ang iyong impormasyon ay pinangangasiwaan alinsunod sa Data Privacy Act of 2012 (RA 10173).'
  }
};

function esc_(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function notifyAddress_() {
  return CONFIG.NOTIFY_EMAIL || Session.getEffectiveUser().getEmail() || Session.getActiveUser().getEmail() || '';
}

function sheetUrl_() {
  try { return getSpreadsheet_().getUrl(); } catch (e) { return ''; }
}

/** Charlene's photo as an inline email image. Returns null if it can't be loaded (emails then still send without it). */
function photoBlob_() {
  const id = BRAND.photoId;
  // 1) The public Drive thumbnail (small and fast)
  try {
    const r = UrlFetchApp.fetch('https://drive.google.com/thumbnail?id=' + id + '&sz=w1000', { muteHttpExceptions: true });
    const b = r.getBlob(), type = b.getContentType() || '';
    if (r.getResponseCode() === 200 && /^image\//.test(type)) return b.setName('charlene');
    console.warn('Thumbnail fetch returned ' + r.getResponseCode() + ' (' + type + '). Is the photo shared as "Anyone with the link"?');
  } catch (e) { console.warn('Thumbnail fetch failed: ' + e); }
  // 2) Read the file straight from Drive (works even if it is not public, as long as this account can open it)
  try {
    const f = DriveApp.getFileById(id);
    let b = f.getBlob();
    if (b.getBytes().length > 3000000) b = f.getThumbnail();   // keep emails light
    return b.setName('charlene');
  } catch (e) { console.error('Drive read failed: ' + e); }
  return null;
}

/** Run this from the editor to see whether the photo loads and, if not, why (check the Execution log). */
function checkPhoto() {
  const b = photoBlob_();
  console.log(b ? 'Photo OK: ' + b.getContentType() + ', ' + b.getBytes().length + ' bytes' : 'Photo FAILED. Read the warnings and errors above this line.');
}

/** Agent-bonding image used in applicant confirmation emails. */
function bondingPhotoBlob_() {
  const id = BRAND.bondingPhotoId;
  try {
    const r = UrlFetchApp.fetch('https://drive.google.com/thumbnail?id=' + id + '&sz=w1200', { muteHttpExceptions: true });
    const b = r.getBlob(), type = b.getContentType() || '';
    if (r.getResponseCode() === 200 && /^image\//.test(type)) return b.setName('agent-bonding');
    console.warn('Bonding image thumbnail returned ' + r.getResponseCode() + ' (' + type + '). Check Drive sharing.');
  } catch (e) { console.warn('Bonding image fetch failed: ' + e); }
  try {
    const f = DriveApp.getFileById(id);
    let b = f.getBlob();
    if (b.getBytes().length > 3000000) b = f.getThumbnail();
    return b.setName('agent-bonding');
  } catch (e) { console.error('Bonding image Drive read failed: ' + e); }
  return null;
}

function checkBondingPhoto() {
  const b = bondingPhotoBlob_();
  console.log(b ? 'Bonding photo OK: ' + b.getContentType() + ', ' + b.getBytes().length + ' bytes' : 'Bonding photo FAILED. Check the Drive file ID and sharing.');
}

function btn_(text, url, ghost) {
  const B = BRAND;
  const style = ghost
    ? 'background:#ffffff;color:' + B.dark + ';border:2px solid ' + B.dark + ';'
    : 'background:' + B.gold + ';color:' + B.dark + ';border:2px solid ' + B.gold + ';';
  return '<a href="' + esc_(url) + '" style="display:inline-block;' + style + 'font-weight:bold;font-size:16px;text-decoration:none;padding:12px 24px;border-radius:4px;margin:4px;">' + esc_(text) + '</a>';
}

function list_(items, ordered) {
  const tag = ordered ? 'ol' : 'ul';
  return '<' + tag + ' style="margin:0;padding-left:22px;">' + items.map(function (x) {
    return '<li style="margin:6px 0;">' + esc_(x) + '</li>';
  }).join('') + '</' + tag + '>';
}

function h1_(text) {
  return '<h1 style="font-family:Georgia,\'Times New Roman\',serif;font-size:28px;line-height:1.2;color:' + BRAND.dark + ';margin:0 0 16px;">' + esc_(text) + '</h1>';
}

function emailImage_(cid, alt) {
  return '<div style="margin:0 0 24px;border-radius:6px;overflow:hidden;background:#E9D9A6;line-height:0;">' +
    '<img src="cid:' + cid + '" alt="' + esc_(alt) + '" width="600" style="display:block;width:100%;max-width:600px;height:auto;border:0;">' +
    '</div>';
}

function layout_(o) {
  const B = BRAND;
  const soc = B.socials.map(function (s) {
    return '<a href="' + s[1] + '" style="color:' + B.pale + ';text-decoration:none;font-weight:bold;margin:0 8px;">' + s[0] + '</a>';
  }).join(' ');
  const band = o.photo
    ? '<tr><td style="background:' + B.dark + ';line-height:0;font-size:0;"><img src="cid:charlene" alt="Charlene Dellosa – Join Our Team" width="600" style="display:block;width:100%;max-width:600px;height:auto;border:0;"></td></tr>'
    : '<tr><td style="background:' + B.gold + ';height:6px;line-height:6px;font-size:0;">&nbsp;</td></tr>';
  return '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
    '<body style="margin:0;padding:0;background:#EFE8D2;">' +
    '<div style="display:none;max-height:0;overflow:hidden;opacity:0;">' + esc_(o.pre) + '</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#EFE8D2;"><tr><td align="center" style="padding:24px 12px;">' +
    '<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="width:100%;max-width:600px;background:' + B.cream + ';border-radius:8px;overflow:hidden;font-family:Arial,Helvetica,sans-serif;color:' + B.ink + ';">' +
    '<tr><td style="background:' + B.dark + ';padding:22px 28px;"><div style="font-family:Georgia,\'Times New Roman\',serif;font-size:24px;font-weight:bold;color:' + B.cream + ';">Charlene Dellosa</div>' +
    '<div style="font-size:12px;letter-spacing:2px;text-transform:uppercase;color:' + B.pale + ';margin-top:4px;">Real Estate</div></td></tr>' +
    band +
    '<tr><td style="padding:30px 28px 26px;font-size:16px;line-height:1.6;">' + o.body + '</td></tr>' +
    '<tr><td align="center" style="background:' + B.dark + ';padding:22px 20px;font-size:14px;line-height:1.6;color:' + B.pale + ';">' +
    '<div style="margin-bottom:10px;">' + soc + '</div><div style="font-size:12px;opacity:.8;">' + esc_(o.foot) + '</div></td></tr>' +
    '</table></td></tr></table></body></html>';
}

/** Warm follow-up email to every applicant (there is no pass/fail outcome). */
function applicantEmail_(lang, name, hasPhoto, hasBondingPhoto) {
  const B = BRAND, T = MAIL_TEXT[lang === 'tl' ? 'tl' : 'en'];
  const first = String(name).trim().split(/\s+/)[0] || 'there';
  const greet = T.greet.replace('{name}', first);

  const html = layout_({
    pre: T.intro,
    photo: hasPhoto,
    foot: T.priv,
    body:
      h1_(T.title) +
      '<p style="margin:0 0 12px;">' + esc_(greet) + '</p>' +
      '<p style="margin:0 0 16px;">' + esc_(T.intro) + '</p>' +
      '<p style="margin:0 0 24px;">' + esc_(T.intro2) + '</p>' +
      (hasBondingPhoto ? emailImage_('agentBonding', 'Real estate agents bonding and building a supportive team') : '') +

      '<div style="background:#ffffff;border:1px solid #E2D9B5;border-left:4px solid ' + B.gold + ';border-radius:4px;padding:16px 20px;margin:0 0 26px;">' +
      '<div style="font-family:Georgia,\'Times New Roman\',serif;font-size:19px;font-weight:bold;color:' + B.dark + ';margin-bottom:8px;">' + esc_(T.perksH) + '</div>' +
      list_(T.perks, false) +
      '</div>' +

      '<div style="background:' + B.dark + ';color:' + B.cream + ';border-radius:4px;padding:18px 20px;margin:0 0 26px;">' +
      '<div style="font-family:Georgia,\'Times New Roman\',serif;font-size:21px;font-weight:bold;color:' + B.pale + ';margin-bottom:8px;">' + esc_(T.earningQ) + '</div>' +
      '<div style="font-size:14px;line-height:1.6;color:' + B.cream + ';">' + esc_(T.earningNote) + '</div>' +
      '</div>' +

      '<div style="background:#ffffff;border:1px solid #E2D9B5;border-left:4px solid ' + B.emerald + ';border-radius:4px;padding:16px 20px;margin:0 0 26px;">' +
      '<div style="font-family:Georgia,\'Times New Roman\',serif;font-size:19px;font-weight:bold;color:' + B.dark + ';margin-bottom:8px;">' + esc_(T.guideH) + '</div>' +
      list_(T.guide, false) +
      '</div>' +

      '<div style="background:#ffffff;border:1px solid #E2D9B5;border-left:4px solid ' + B.gold + ';border-radius:4px;padding:16px 20px;margin:0 0 26px;">' +
      '<div style="font-family:Georgia,\'Times New Roman\',serif;font-size:19px;font-weight:bold;color:' + B.dark + ';margin-bottom:8px;">' + esc_(T.nextH) + '</div>' +
      list_(T.next, true) +
      '</div>' +

      '<p style="margin:0 0 24px;">' + esc_(T.reminder) + '</p>' +
      '<div style="text-align:center;margin:0 0 26px;">' + btn_(T.btn, B.agentPortal) + '</div>' +
      '<p style="margin:0;">' + esc_(T.sign) + '<br><strong style="color:' + B.dark + ';">' + esc_(B.sender) + '</strong></p>'
  });

  const text =
    T.title + '\n\n' +
    greet + '\n' +
    T.intro + '\n\n' +
    T.intro2 + '\n\n' +
    T.perksH + ':\n- ' + T.perks.join('\n- ') + '\n\n' +
    T.earningQ + '\n' + T.earningNote + '\n\n' +
    T.guideH + ':\n- ' + T.guide.join('\n- ') + '\n\n' +
    T.nextH + ':\n- ' + T.next.join('\n- ') + '\n\n' +
    T.reminderH + ':\n' + T.reminder + '\n\n' +
    B.agentPortal + '\n\n' +
    T.sign + '\n' + B.sender;

  return { subject: T.subject, html: html, text: text };
}
function row_(label, valueHtml) {
  return '<tr><td style="padding:9px 0;border-bottom:1px solid #E2D9B5;width:38%;color:#5a5f62;font-size:14px;vertical-align:top;">' + esc_(label) +
    '</td><td style="padding:9px 0;border-bottom:1px solid #E2D9B5;font-weight:bold;color:' + BRAND.dark + ';vertical-align:top;">' + valueHtml + '</td></tr>';
}

/** Email to Charlene about a new applicant and training follow-up. */
function ownerEmail_(d, hasPhoto) {
  const B = BRAND;
  const badgeText = 'NEW APPLICANT · TRAINING FOLLOW-UP';
  const details =
    row_('Mobile', '<a href="tel:' + esc_(d.phone.replace(/[^\d+]/g, '')) + '" style="color:' + B.emerald + ';">' + esc_(d.phone) + '</a>') +
    row_('Email', '<a href="mailto:' + esc_(d.email) + '" style="color:' + B.emerald + ';">' + esc_(d.email) + '</a>') +
    row_('City', esc_(d.city || '-')) +
    row_('Language', d.lang === 'tl' ? 'Tagalog' : 'English') +
    row_('Consent', esc_(d.consent));
  const answers = d.answers.map(function (a) { return row_(a.label, esc_(a.text)); }).join('');
  const intro = 'This applicant is ready for a conversation. Please reach out warmly, learn more about their goals, and guide them toward the training and support that best fits their starting point.';
  const html = layout_({
    pre: 'New applicant ready for a supportive follow-up: ' + d.name, photo: hasPhoto, foot: 'New applicant notification',
    body: '<div style="margin:0 0 14px;"><span style="display:inline-block;background:' + B.emerald + ';color:' + B.cream + ';font-size:12px;font-weight:bold;letter-spacing:1.2px;padding:6px 12px;border-radius:999px;">' + badgeText + '</span></div>' +
      h1_(d.name) + '<p style="margin:0 0 22px;">' + esc_(intro) + '</p>' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:0 0 22px;">' + details + '</table>' +
      '<div style="font-family:Georgia,\'Times New Roman\',serif;font-size:19px;font-weight:bold;color:' + B.dark + ';margin:0 0 4px;">Questionnaire answers</div>' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:0 0 26px;">' + answers + '</table>' +
      '<div style="text-align:center;">' + btn_('Call ' + d.phone, 'tel:' + d.phone.replace(/[^\d+]/g, '')) + btn_('Reply by email', 'mailto:' + d.email, true) + (d.sheetUrl ? btn_('Open applicants sheet', d.sheetUrl, true) : '') + '</div>'
  });
  const text = badgeText + '\n\n' + intro + '\n\nName: ' + d.name + '\nMobile: ' + d.phone + '\nEmail: ' + d.email + '\nCity: ' + d.city + '\nLanguage: ' + d.lang + '\nConsent: ' + d.consent + '\n\n' + d.answers.map(function (a) { return a.label + ': ' + a.text; }).join('\n') + (d.sheetUrl ? '\n\nSheet: ' + d.sheetUrl : '');
  return { subject: 'New applicant · Training follow-up: ' + d.name, html: html, text: text };
}

function sendMail_(to, m, photo, replyTo, bondingPhoto) {
  const opts = { to: to, subject: m.subject, body: m.text, htmlBody: m.html, name: BRAND.sender };
  if (replyTo) opts.replyTo = replyTo;
  if (photo || bondingPhoto) {
    opts.inlineImages = {};
    if (photo) opts.inlineImages.charlene = photo;
    if (bondingPhoto) opts.inlineImages.agentBonding = bondingPhoto;
  }
  MailApp.sendEmail(opts);
}

/**
 * An email problem must never break a saved application, so failures are only logged.
 * Return a boolean so the frontend can distinguish a saved application from a
 * successfully delivered notification.
 */
function safeSend_(to, build, photo, replyTo, bondingPhoto) {
  try {
    if (!to) throw new Error('No recipient address');
    sendMail_(to, build(), photo, replyTo, bondingPhoto);
    return true;
  } catch (e) {
    console.error('Email to ' + to + ' failed: ' + e);
    return false;
  }
}

/**
 * Test one language only.
 * Change TEST_LANG to 'en' for English or 'tl' for Tagalog before running.
 * This sends one owner notification and one applicant confirmation email.
 */
function sendTestEmails() {
  const to = notifyAddress_();
  const quiz = getQuiz_();

  // Use 'en' for English or 'tl' for Tagalog.
  const TEST_LANG = 'en';

  const photo = photoBlob_();
  const hasPhoto = !!photo;
  const bondingPhoto = bondingPhotoBlob_();
  const hasBondingPhoto = !!bondingPhoto;

  const sample = {
    name: 'Maria Santos',
    phone: '0917 123 4567',
    email: to,
    city: 'Bacoor, Cavite',
    lang: TEST_LANG,
    consent: 'Yes (test)',
    sheetUrl: sheetUrl_(),
    answers: quiz.map(function (x) {
      return {
        label: x.label,
        text: TEST_LANG === 'tl' ? x.options[0].tl : x.options[0].t
      };
    })
  };

  // Owner notification.
  sendMail_(to, ownerEmail_(sample, hasPhoto), photo, null);

  // One applicant confirmation email in the selected test language.
  sendMail_(
    to,
    applicantEmail_(TEST_LANG, 'Maria Santos', hasPhoto, hasBondingPhoto),
    photo,
    null,
    bondingPhoto
  );

  console.log(
    'Sent one ' + (TEST_LANG === 'tl' ? 'Tagalog' : 'English') +
    ' applicant test email to ' + to +
    (hasPhoto ? '' : '. The photo could not be loaded: share it as "Anyone with the link" in Google Drive.')
  );
}

function submitApplication(data) {
  try {
    if (data.website) return { ok: true, qualified: false, tips: [] }; // honeypot

    const quiz = getQuiz_();
    const name = clean_(data.name, 100);
    const phone = clean_(data.phone, 30);
    const email = clean_(data.email, 120);
    const city = clean_(data.city, 80);
    const lang = data.lang === 'tl' ? 'tl' : 'en';
    if (!name || !phone || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
      return { ok: false, error: 'Please enter your name, mobile number, and a valid email.' };
    }
    if (data.consent !== true) {
      return { ok: false, error: 'Consent is required to submit.' };
    }

    const picks = data.answers;
    if (!Array.isArray(picks) || picks.length !== quiz.length) {
      return { ok: false, error: 'Please answer all questions.' };
    }
    const answerText = [];
    for (let i = 0; i < quiz.length; i++) {
      const pick = picks[i];
      if (typeof pick === 'number' && quiz[i].options[pick]) answerText.push(quiz[i].options[pick].t);
      else if (typeof pick === 'string' && pick.trim()) answerText.push(clean_(pick, 800));
      else return { ok: false, error: 'Please answer all questions.' };
    }
    const consentNote = 'Yes (' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') + ')';

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      getApplicantSheet_(quiz).appendRow([new Date(), name, phone, email, city]
        .concat(answerText, ['No grading', 'Training follow-up', 'New', lang === 'tl' ? 'Tagalog' : 'English', consentNote]));
    } finally {
      lock.releaseLock();
    }

    const notify = notifyAddress_();
    const photo = photoBlob_(), has = !!photo;
    const bondingPhoto = bondingPhotoBlob_(), hasBondingPhoto = !!bondingPhoto;
    const notificationSent = safeSend_(notify, function () {
      return ownerEmail_({
        name: name, phone: phone, email: email, city: city, lang: lang, consent: consentNote,
        sheetUrl: sheetUrl_(),
        answers: quiz.map(function (x, i) { return { label: x.label, text: answerText[i] }; })
      }, has);
    }, photo, email);
    let confirmationSent = false;
    if (CONFIG.SEND_CONFIRMATION) {
      confirmationSent = safeSend_(email, function () {
        return applicantEmail_(lang, name, has, hasBondingPhoto);
      }, photo, notify, bondingPhoto);
    }
    return {
      ok: true,
      qualified: true,
      tips: [],
      notificationSent: notificationSent,
      confirmationSent: confirmationSent,
      // The application is saved even if an email provider temporarily fails.
      saved: true
    }; // retained for compatibility with the current Index.html
  } catch (err) {
    console.error(err);
    return { ok: false, error: 'Something went wrong. Please try again in a moment.' };
  }
}
