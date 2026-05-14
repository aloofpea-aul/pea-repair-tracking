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
  const allowed = ['index','form','report','settings'];

  // form.html — serve ตรงๆ แล้ว inject station_id
  if (page === 'form') {
    var content = HtmlService.createHtmlOutputFromFile('form').getContent();
    content = content.replace(
      "var PRE_STATION_ID = '';",
      "var PRE_STATION_ID = '" + station + "';"
    );
    return HtmlService.createHtmlOutput(content)
      .setTitle('ระบบตรวจงานแก้ไฟฟ้าขัดข้อง กฟต.2')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // settings.html — serve ตรงๆ
  if (page === 'settings') {
    return HtmlService.createHtmlOutputFromFile('settings')
      .setTitle('ตั้งค่าคณะกรรมการตรวจ')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // index / report — ใช้ template
  const tpl = HtmlService.createTemplateFromFile(allowed.includes(page) ? page : 'index');
  tpl.page = page;
  return tpl.evaluate()
    .setTitle('ระบบตรวจงานแก้ไฟฟ้าขัดข้อง กฟต.2')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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