// ============================================================
//  pea-repair-tracking | Code.gs
//  ระบบตรวจงานแก้ไฟฟ้าขัดข้อง กฟต.2  ปี 2569
// ============================================================

const SS_ID           = '10dIThrq9fZy2MEpAfrJmd5fNzoZ0AYufTZTb9X9Cra0';
const DRIVE_FOLDER_ID = '1tvEJv23Fjta-dCWBOlnYC0p2F1ZExCNK';

const SH = {
  CONFIG      : 'CONFIG',
  STATIONS    : 'STATIONS',
  CHECKLIST   : 'CHECKLIST',
  INSPECTIONS : 'INSPECTIONS',
  SCORES      : 'SCORES',
  LOG         : 'LOG',
  COMMITTEES  : 'COMMITTEES'
};

function doGet(e) {
  const page    = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'index';
  const station = (e && e.parameter && e.parameter.station) ? e.parameter.station : '';
  const allowed = ['index','form','report','settings','spare_parts','vehicle_check'];

  // form.html — inject PRE_STATION_ID
  if (page === 'form') {
    var content = HtmlService.createHtmlOutputFromFile('form').getContent();
    content = content.replace(
      "var PRE_STATION_ID = '';",
      "var PRE_STATION_ID = '" + station + "';"
    );
    return HtmlService.createHtmlOutput(content)
      .setTitle('ระบบตรวจงานแก้ไฟฟ้าขัดข้อง กฟต.2')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }

  // settings.html
  if (page === 'settings') {
    return HtmlService.createHtmlOutputFromFile('settings')
      .setTitle('ตั้งค่าคณะกรรมการตรวจ')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }

  // spare_parts.html
  if (page === 'spare_parts') {
    return HtmlService.createHtmlOutputFromFile('spare_parts')
      .setTitle('รายการพัสดุสำรองแก้ไฟ 2569')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }

  // vehicle_check.html — ใช้ template เพื่อให้ GAS render mobile ได้ถูกต้อง
  if (page === 'vehicle_check') {
    var tplVC = HtmlService.createTemplateFromFile('vehicle_check');
    return tplVC.evaluate()
      .setTitle('แบบฟอร์มตรวจยานพาหนะ FM1-SOH3-S04-6103')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }

  // index / report — template
  const tpl = HtmlService.createTemplateFromFile(allowed.includes(page) ? page : 'index');
  tpl.page = page;
  return tpl.evaluate()
    .setTitle('ระบบตรวจงานแก้ไฟฟ้าขัดข้อง กฟต.2')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSheet(name) {
  return SpreadsheetApp.openById(SS_ID).getSheetByName(name);
}

function getSheetData(name) {
  const sh = getSheet(name);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function appendRow(sheetName, rowObj) {
  const sh = getSheet(sheetName);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const row = headers.map(h => rowObj[h] !== undefined ? rowObj[h] : '');
  sh.appendRow(row);
}

function nowTH() {
  return Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
}

function genId(prefix) {
  return prefix + '_' + new Date().getTime();
}

function writeLog(action, detail, user) {
  try {
    appendRow(SH.LOG, {
      timestamp : nowTH(),
      user      : user || Session.getActiveUser().getEmail() || 'system',
      action    : action,
      detail    : detail
    });
  } catch(e) {}
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}