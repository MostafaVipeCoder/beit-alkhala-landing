/**
 * Beit Alkhala — landing page + pre-booking (Google Apps Script web app).
 *
 * Parents read about the nursery and its subscriptions, then send a
 * pre-booking request. Requests land in a Google Sheet; the nursery staff
 * call the family and register the child manually in the main system.
 *
 * First time: run setup() once from the editor (creates the sheet), then
 * Deploy > New deployment > Web app (Execute as: Me, Access: Anyone).
 */

// ------------------------------------------------------------------ edit me

const CONFIG = {
  name: 'بيت الخالة',
  tagline: 'مكان آمن.. ينمو فيه طفلك',
  opening: 'قريبًا',
  days: 'من الأحد للخميس',
  hours: 'فترة صباحية 7 ساعات',
  // Fill these in; empty ones are hidden on the page.
  phone: '',
  whatsapp: '', // e.g. 201012345678 (country code, no +)
  address: '',
  mapUrl: '',
  ageRange: '',
  capacity: 35,
  siblingDiscount: 10,
  dayRate: 350,
  packages: [
    { id: 'monthly', name: 'شهري', price: 3000, days: 20, note: 'الأكثر طلبًا' },
    { id: 'half_month', name: 'نصف شهري', price: 1800, days: 10 },
    { id: 'weekly', name: 'أسبوعي', price: 1000, days: 5 },
    { id: 'three_months', name: 'عرض 3 شهور', price: 7500, days: 60, note: 'أماكن محدودة: 6 بس', offer: true },
  ],
  extras: [
    { id: 'daily', name: 'يومي', price: 350, unit: 'لليوم (7 ساعات)' },
    { id: 'hourly', name: 'بالساعة', price: 50, unit: 'للساعة' },
  ],
  installments: [60, 20, 20],
  sources: ['فيسبوك', 'إنستجرام', 'ترشيح من ولي أمر', 'لافتة', 'جوجل', 'أخرى'],
  // The web app /exec link. The GitHub Pages copy of the page sends bookings here.
  bookingUrl: 'https://script.google.com/macros/s/AKfycbxBq91jYvfBCbY078Vq0Tf-skfl2T8Ws1y51LqxdaXdX0nzDIigz9C4lGKhk9u5T7XtsA/exec',
};

const SHEET_NAME = 'الحجوزات';
const STATUSES = ['جديد', 'اتواصلنا', 'اتسجل', 'مش مهتم'];
const HEADERS = [
  'رقم الطلب',
  'التاريخ والوقت',
  'اسم ولي الأمر',
  'التليفون',
  'اسم الطفل',
  'سن الطفل (شهور)',
  'الباقة المهتمين بيها',
  'البداية المتوقعة',
  'عرفونا منين',
  'ملاحظات الأهل',
  'رقم مكرر؟',
  'الحالة',
  'ملاحظات الحضانة',
];

// ------------------------------------------------------------------ web app

function doGet() {
  const page = HtmlService.createTemplateFromFile('Index');
  page.config = CONFIG;
  return page
    .evaluate()
    .setTitle(CONFIG.name + ' — احجزوا مكانكم')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Bookings from the page hosted outside Google (GitHub Pages). The page posts
 * JSON as text/plain without cookies, so the visitor's Google accounts never
 * matter and the "unable to open file" multi-account error cannot happen.
 */
function doPost(e) {
  let result;
  try {
    result = submitBooking(JSON.parse((e && e.postData && e.postData.contents) || '{}'));
  } catch (err) {
    result = { ok: false, error: 'failed' };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

/** Used by the template to inline other HTML files (the logo). */
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// ------------------------------------------------------------------ setup

const SUMMARY_NAME = 'ملخص';

/** Adds a "بيت الخالة" menu to the sheet (script opened from the sheet). */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(CONFIG.name)
    .addItem('تجهيز الشيت (الورق والأعمدة)', 'setup')
    .addToUi();
}

/**
 * Builds the tabs and columns. Safe to run again: it never deletes rows.
 * Opened from a Google Sheet (Extensions > Apps Script), the bookings go
 * into that sheet; a standalone script creates a new one.
 */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss && props.getProperty('SHEET_ID')) ss = SpreadsheetApp.openById(props.getProperty('SHEET_ID'));
  if (!ss) ss = SpreadsheetApp.create(CONFIG.name + ' — حجوزات مسبقة');
  props.setProperty('SHEET_ID', ss.getId());
  buildBookings_(ss);
  buildSummary_(ss);
  Logger.log('Bookings sheet: ' + ss.getUrl());
  return ss.getUrl();
}

function buildBookings_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    // A fresh spreadsheet has one empty "Sheet1": reuse it instead of leaving it around.
    const first = ss.getSheets()[0];
    sheet = ss.getSheets().length === 1 && first.getLastRow() === 0 ? first.setName(SHEET_NAME) : ss.insertSheet(SHEET_NAME, 0);
  }
  sheet.setRightToLeft(true);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#FFE9F1');
  sheet.setFrozenRows(1);
  // Phones keep their leading zero.
  sheet.getRange('D:D').setNumberFormat('@');
  sheet.getRange('B2:B').setNumberFormat('yyyy-mm-dd hh:mm');
  const status = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build();
  sheet.getRange(2, 12, sheet.getMaxRows() - 1, 1).setDataValidation(status);
  // Status colours: new = pink, contacted = blue, registered = green, not interested = grey.
  const range = sheet.getRange(2, 12, sheet.getMaxRows() - 1, 1);
  const colors = { 'جديد': '#FFE9F1', 'اتواصلنا': '#EAF4FF', 'اتسجل': '#E3F7EE', 'مش مهتم': '#EEEEEE' };
  const rules = sheet.getConditionalFormatRules().filter((r) => {
    const rs = r.getRanges();
    return !(rs.length === 1 && rs[0].getColumn() === 12);
  });
  Object.keys(colors).forEach((s) => {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(s).setBackground(colors[s]).setRanges([range]).build());
  });
  sheet.setConditionalFormatRules(rules);
  sheet.autoResizeColumns(1, HEADERS.length);
  return sheet;
}

/** Counts by status, package and source; formulas only, so it updates by itself. */
function buildSummary_(ss) {
  const sheet = ss.getSheetByName(SUMMARY_NAME) || ss.insertSheet(SUMMARY_NAME);
  sheet.clear();
  sheet.setRightToLeft(true);
  const src = "'" + SHEET_NAME + "'!";
  const rows = [['البند', 'العدد'], ['إجمالي الطلبات', '=COUNTA(' + src + 'A2:A)'], ['أرقام مكررة', '=COUNTIF(' + src + 'K2:K,"نعم")'], ['', '']];
  const block = (title, values, col) => {
    rows.push([title, '']);
    values.forEach((v) => rows.push([v, '=COUNTIF(' + src + col + '2:' + col + ',"' + v + '")']));
    rows.push(['', '']);
  };
  block('حسب الحالة', STATUSES, 'L');
  block('حسب الباقة', ['لسه مش متأكدين'].concat(CONFIG.packages.concat(CONFIG.extras).map((p) => p.name)), 'G');
  block('عرفونا منين', CONFIG.sources, 'I');
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#FFE9F1');
  rows.forEach((r, i) => {
    if (r[0] && r[1] === '') sheet.getRange(i + 1, 1, 1, 2).setFontWeight('bold').setBackground('#EAF4FF');
  });
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 2);
  return sheet;
}

/** The bookings tab; rebuilt if someone renamed or removed it. */
function bookingsSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('setup() was not run');
  if (!id) PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  return ss.getSheetByName(SHEET_NAME) || buildBookings_(ss);
}

// ------------------------------------------------------------------ booking

function clean_(value, max) {
  return String(value == null ? '' : value)
    .replace(/[<>]/g, '')
    .replace(/^[=+\-@]+/, '') // never let a cell start a formula
    .trim()
    .slice(0, max);
}

function latinDigits_(s) {
  return String(s).replace(/[٠-٩]/g, (d) => '٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/[۰-۹]/g, (d) => '۰۱۲۳۴۵۶۷۸۹'.indexOf(d));
}

function normalizePhone_(s) {
  let p = latinDigits_(s).replace(/[^\d+]/g, '');
  if (p.indexOf('+20') === 0) p = '0' + p.slice(3);
  else if (p.indexOf('0020') === 0) p = '0' + p.slice(4);
  return p.replace(/\+/g, '');
}

/**
 * Called from the page. Returns { ok, ref } or { ok: false, error }.
 * Errors are short codes; the page shows the Arabic text.
 */
function submitBooking(form) {
  form = form || {};
  // Bots fill the hidden field; humans do not.
  if (form.website) return { ok: true, ref: '—' };

  const phone = normalizePhone_(form.phone);
  const data = {
    parent: clean_(form.parent, 80),
    child: clean_(form.child, 80),
    ageMonths: Number(form.ageMonths),
    pkg: clean_(form.pkg, 40),
    start: clean_(form.start, 20),
    source: clean_(form.source, 40),
    notes: clean_(form.notes, 500),
  };
  if (data.parent.length < 2 || data.child.length < 2) return { ok: false, error: 'missing' };
  if (!/^01\d{9}$/.test(phone)) return { ok: false, error: 'phone' };
  if (!(data.ageMonths >= 0 && data.ageMonths <= 120)) return { ok: false, error: 'age' };
  // Optional, but if given it must be after today (Cairo time).
  const today = Utilities.formatDate(new Date(), 'Africa/Cairo', 'yyyy-MM-dd');
  if (data.start && !(/^\d{4}-\d{2}-\d{2}$/.test(data.start) && data.start > today)) return { ok: false, error: 'start' };
  if (!form.consent) return { ok: false, error: 'consent' };

  // Same phone again within 10 minutes = double click / resend.
  const cache = CacheService.getScriptCache();
  const recent = cache.get('b:' + phone);
  if (recent) return { ok: true, ref: recent };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = bookingsSheet_();
    const last = sheet.getLastRow();
    const phones = last > 1 ? sheet.getRange(2, 4, last - 1, 1).getValues().map((r) => String(r[0])) : [];
    const duplicate = phones.indexOf(phone) >= 0;
    const ref = 'BK-' + String(last).padStart(4, '0');
    sheet.appendRow([
      ref,
      new Date(),
      data.parent,
      phone,
      data.child,
      data.ageMonths,
      data.pkg,
      data.start,
      data.source,
      data.notes,
      duplicate ? 'نعم' : '',
      'جديد',
      '',
    ]);
    sheet.getRange(sheet.getLastRow(), 4).setNumberFormat('@').setValue(phone);
    cache.put('b:' + phone, ref, 600);
    return { ok: true, ref: ref };
  } finally {
    lock.releaseLock();
  }
}
