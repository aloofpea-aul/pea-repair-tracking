// ============================================================
//  Setup.gs — ตั้งค่าระบบ & นำเข้าหน่วยงาน
//  pea-repair-tracking | กฟต.2 ปี 2569
// ============================================================

// ============================================================
//  setupAll() — รันครั้งแรกเพื่อสร้าง sheet headers ทั้งหมด
// ============================================================
function setupAll() {
  setupStationsSheet();
  setupChecklistSheet();
  setupInspectionsSheet();
  setupScoresSheet();
  setupLogSheet();
  setupConfigSheet();
  importAllStations();
  Logger.log('setupAll complete');
}

// ============================================================
//  setupStationsSheet
// ============================================================
function setupStationsSheet() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SH.STATIONS);
  if (!sh) sh = ss.insertSheet(SH.STATIONS);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['station_id','name','short_name','type','province','active']);
    sh.getRange(1,1,1,6).setFontWeight('bold').setBackground('#1a3c5e').setFontColor('#ffffff');
  }
}

// ============================================================
//  setupInspectionsSheet
// ============================================================
function setupInspectionsSheet() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SH.INSPECTIONS);
  if (!sh) sh = ss.insertSheet(SH.INSPECTIONS);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['inspect_id','station_id','station_name','inspect_date','inspector_name','inspector_email','total_score','note','created_at']);
    sh.getRange(1,1,1,9).setFontWeight('bold').setBackground('#1a3c5e').setFontColor('#ffffff');
  }
}

// ============================================================
//  setupScoresSheet
// ============================================================
function setupScoresSheet() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SH.SCORES);
  if (!sh) sh = ss.insertSheet(SH.SCORES);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['score_id','inspect_id','station_id','item_id','item_no','score','pct','note']);
    sh.getRange(1,1,1,8).setFontWeight('bold').setBackground('#1a3c5e').setFontColor('#ffffff');
  }
}

// ============================================================
//  setupLogSheet
// ============================================================
function setupLogSheet() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SH.LOG);
  if (!sh) sh = ss.insertSheet(SH.LOG);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['timestamp','user','action','detail']);
    sh.getRange(1,1,1,4).setFontWeight('bold').setBackground('#1a3c5e').setFontColor('#ffffff');
  }
}

// ============================================================
//  setupConfigSheet
// ============================================================
function setupConfigSheet() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SH.CONFIG);
  if (!sh) sh = ss.insertSheet(SH.CONFIG);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['key','value']);
    sh.appendRow(['system_name','ระบบตรวจงานแก้ไฟฟ้าขัดข้อง กฟต.2']);
    sh.appendRow(['year','2569']);
    sh.getRange(1,1,1,2).setFontWeight('bold').setBackground('#1a3c5e').setFontColor('#ffffff');
  }
}

// ============================================================
//  setupChecklistSheet — หัวข้อ checklist
// ============================================================
function setupChecklistSheet() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SH.CHECKLIST);
  if (!sh) sh = ss.insertSheet(SH.CHECKLIST);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['item_id','item_no','category','category_max','topic','max_score','criteria']);
    sh.getRange(1,1,1,7).setFontWeight('bold').setBackground('#1a3c5e').setFontColor('#ffffff');
  }
}

// ============================================================
//  importAllStations() — นำเข้าหน่วยงานทั้ง 79 แห่ง
//  รันครั้งเดียว หรือรันซ้ำได้ (ไม่ duplicate)
// ============================================================
function importAllStations() {
  var STATION_LIST = [
    // ===== กฟส.L (กฟจ./กฟส.L-M) =====
    {station_id:'ST_LM_นศ',   name:'กฟฟ.นครศรีธรรมราช', short_name:'นศ',  type:'L', province:'นครศรีธรรมราช'},
    {station_id:'ST_LM_สฎ',   name:'กฟฟ.สุราษฎร์ธานี',  short_name:'สฎ',  type:'L', province:'สุราษฎร์ธานี'},
    {station_id:'ST_LM_ตง',   name:'กฟฟ.ตรัง',           short_name:'ตง',  type:'L', province:'ตรัง'},
    {station_id:'ST_LM_กบ',   name:'กฟฟ.กระบี่',         short_name:'กบ',  type:'L', province:'กระบี่'},
    {station_id:'ST_LM_พง',   name:'กฟฟ.พังงา',          short_name:'พง',  type:'L', province:'พังงา'},
    {station_id:'ST_LM_ภก',   name:'กฟฟ.ภูเก็ต',         short_name:'ภก',  type:'L', province:'ภูเก็ต'},
    {station_id:'ST_LM_ทส',   name:'กฟฟ.ท่าศาลา',        short_name:'ทส',  type:'M', province:'นครศรีธรรมราช'},
    {station_id:'ST_LM_กม',   name:'กฟฟ.กาญจนดิษฐ์',    short_name:'กม',  type:'M', province:'สุราษฎร์ธานี'},
    {station_id:'ST_LM_พพ',   name:'กฟฟ.พระแสง',         short_name:'พพ',  type:'M', province:'สุราษฎร์ธานี'},
    {station_id:'ST_LM_ปพน',  name:'กฟฟ.ปากพนัง',        short_name:'ปพน', type:'M', province:'นครศรีธรรมราช'},
    {station_id:'ST_LM_ถล',   name:'กฟฟ.ถลาง',           short_name:'ถล',  type:'M', province:'ภูเก็ต'},
    {station_id:'ST_LM_ตป',   name:'กฟฟ.ตะกั่วป่า',      short_name:'ตป',  type:'M', province:'พังงา'},
    {station_id:'ST_LM_ปต',   name:'กฟฟ.ปลายพระยา',      short_name:'ปต',  type:'M', province:'กระบี่'},
    {station_id:'ST_LM_กญ',   name:'กฟฟ.กาญจนบุรี',      short_name:'กญ',  type:'M', province:'กระบี่'},
    {station_id:'ST_LM_วงส',  name:'กฟฟ.วังวิเศษ',       short_name:'วงส', type:'M', province:'ตรัง'},
    {station_id:'ST_LM_ทศ',   name:'กฟฟ.ทุ่งสง',         short_name:'ทศ',  type:'M', province:'นครศรีธรรมราช'},
    {station_id:'ST_LM_หย',   name:'กฟฟ.หัวยาว',         short_name:'หย',  type:'M', province:'นครศรีธรรมราช'},
    // ===== กฟส.S =====
    {station_id:'ST_S_รพ',    name:'กฟฟ.ร่อนพิบูลย์',   short_name:'รพ',   type:'S', province:'นครศรีธรรมราช'},
    {station_id:'ST_S_ชอด',   name:'กฟฟ.ชะอวด',          short_name:'ชอด',  type:'S', province:'นครศรีธรรมราช'},
    {station_id:'ST_S_บนส',   name:'กฟฟ.บางนาค',         short_name:'บนส',  type:'S', province:'นครศรีธรรมราช'},
    {station_id:'ST_S_กต',    name:'กฟฟ.กระแสสินธุ์',   short_name:'กต',   type:'S', province:'สงขลา'},
    {station_id:'ST_S_ยข',    name:'กฟฟ.ยะขอ',           short_name:'ยข',   type:'S', province:'นครศรีธรรมราช'},
    {station_id:'ST_S_ปล',    name:'กฟฟ.ปากพะยูน',       short_name:'ปล',   type:'S', province:'พัทลุง'},
    {station_id:'ST_S_อล',    name:'กฟฟ.อ่าวลึก',        short_name:'อล',   type:'S', province:'กระบี่'},
    {station_id:'ST_S_คท',    name:'กฟฟ.คลองท่อม',       short_name:'คท',   type:'S', province:'กระบี่'},
    {station_id:'ST_S_กลต',   name:'กฟฟ.กลันตัน',        short_name:'กลต',  type:'S', province:'กระบี่'},
    {station_id:'ST_S_หคง',   name:'กฟฟ.หัวคง',          short_name:'หคง',  type:'S', province:'สุราษฎร์ธานี'},
    {station_id:'ST_S_ขพ',    name:'กฟฟ.ขนอม',           short_name:'ขพ',   type:'S', province:'นครศรีธรรมราช'},
    {station_id:'ST_S_ทหม',   name:'กฟฟ.ทุ่งหมื่น',     short_name:'ทหม',  type:'S', province:'นครศรีธรรมราช'},
    {station_id:'ST_S_ตม',    name:'กฟฟ.ตะโหมด',         short_name:'ตม',   type:'S', province:'พัทลุง'},
    {station_id:'ST_S_ฉล',    name:'กฟฟ.ฉลอง',           short_name:'ฉล',   type:'S', province:'ภูเก็ต'},
    {station_id:'ST_S_ฉว',    name:'กฟฟ.ฉวาง',           short_name:'ฉว',   type:'S', province:'นครศรีธรรมราช'},
    {station_id:'ST_S_ชก',    name:'กฟฟ.ชะกาง',          short_name:'ชก',   type:'S', province:'ตรัง'},
    {station_id:'ST_S_ทญ',    name:'กฟฟ.ท่าญวน',         short_name:'ทญ',   type:'S', province:'สุราษฎร์ธานี'},
    {station_id:'ST_S_กพง',   name:'กฟฟ.กะปง',           short_name:'กพง',  type:'S', province:'พังงา'},
    {station_id:'ST_S_ชย',    name:'กฟฟ.ชัยบุรี',        short_name:'ชย',   type:'S', province:'สุราษฎร์ธานี'},
    {station_id:'ST_S_บตข',   name:'กฟฟ.บ้านตาขุน',     short_name:'บตข',  type:'S', province:'สุราษฎร์ธานี'},
    {station_id:'ST_S_ทน',    name:'กฟฟ.ท่าแนะ',         short_name:'ทน',   type:'S', province:'กระบี่'},
    {station_id:'ST_S_คร',    name:'กฟฟ.คลองหอยโข่ง',  short_name:'คร',   type:'S', province:'สงขลา'},
    {station_id:'ST_S_หท',    name:'กฟฟ.หนองทะเล',       short_name:'หท',   type:'S', province:'กระบี่'},
    {station_id:'ST_S_พสง',   name:'กฟฟ.พระสิงห์',      short_name:'พสง',  type:'S', province:'นครศรีธรรมราช'},
    {station_id:'ST_S_สช',    name:'กฟฟ.สิชล',           short_name:'สช',   type:'S', province:'นครศรีธรรมราช'},
    {station_id:'ST_S_ขอ',    name:'กฟฟ.ขนอน',           short_name:'ขอ',   type:'S', province:'สุราษฎร์ธานี'},
    // ===== กฟส.XS =====
    {station_id:'ST_XS_ลานสกา',          name:'กฟฟ.ลานสกา',          short_name:'ลานสกา',          type:'XS', province:'นครศรีธรรมราช'},
    {station_id:'ST_XS_พรหมคีรี',        name:'กฟฟ.พรหมคีรี',        short_name:'พรหมคีรี',        type:'XS', province:'นครศรีธรรมราช'},
    {station_id:'ST_XS_พระพรหม',         name:'กฟฟ.พระพรหม',         short_name:'พระพรหม',         type:'XS', province:'นครศรีธรรมราช'},
    {station_id:'ST_XS_เฉลิมพระเกียรติ', name:'กฟฟ.เฉลิมพระเกียรติ', short_name:'เฉลิมพระเกียรติ', type:'XS', province:'นครศรีธรรมราช'},
    {station_id:'ST_XS_จุฬาภรณ์',        name:'กฟฟ.จุฬาภรณ์',        short_name:'จุฬาภรณ์',        type:'XS', province:'นครศรีธรรมราช'},
    {station_id:'ST_XS_บ้านนาเดิม',      name:'กฟฟ.บ้านนาเดิม',      short_name:'บ้านนาเดิม',      type:'XS', province:'สุราษฎร์ธานี'},
    {station_id:'ST_XS_สิเกา',           name:'กฟฟ.สิเกา',           short_name:'สิเกา',           type:'XS', province:'ตรัง'},
    {station_id:'ST_XS_นาโยง',           name:'กฟฟ.นาโยง',           short_name:'นาโยง',           type:'XS', province:'ตรัง'},
    {station_id:'ST_XS_เกาะลิบง',        name:'กฟฟ.เกาะลิบง',        short_name:'เกาะลิบง',        type:'XS', province:'ตรัง'},
    {station_id:'ST_XS_เกาะมุกต์',       name:'กฟฟ.เกาะมุกต์',       short_name:'เกาะมุกต์',       type:'XS', province:'ตรัง'},
    {station_id:'ST_XS_เกาะสุกร',        name:'กฟฟ.เกาะสุกร',        short_name:'เกาะสุกร',        type:'XS', province:'ตรัง'},
    {station_id:'ST_XS_หาดสำราญ',        name:'กฟฟ.หาดสำราญ',        short_name:'หาดสำราญ',        type:'XS', province:'ตรัง'},
    {station_id:'ST_XS_บ้านเกาะพีพี',   name:'กฟฟ.บ้านเกาะพีพี',   short_name:'บ้านเกาะพีพี',   type:'XS', province:'กระบี่'},
    {station_id:'ST_XS_อ่าวนาง',         name:'กฟฟ.อ่าวนาง',         short_name:'อ่าวนาง',         type:'XS', province:'กระบี่'},
    {station_id:'ST_XS_ปลายพระยา',       name:'กฟฟ.ปลายพระยา',       short_name:'ปลายพระยา',       type:'XS', province:'กระบี่'},
    {station_id:'ST_XS_ลำทับ',           name:'กฟฟ.ลำทับ',           short_name:'ลำทับ',           type:'XS', province:'กระบี่'},
    {station_id:'ST_XS_ทับปุด',          name:'กฟฟ.ทับปุด',          short_name:'ทับปุด',          type:'XS', province:'พังงา'},
    {station_id:'ST_XS_นาบอน',           name:'กฟฟ.นาบอน',           short_name:'นาบอน',           type:'XS', province:'นครศรีธรรมราช'},
    {station_id:'ST_XS_บางขัน',          name:'กฟฟ.บางขัน',          short_name:'บางขัน',          type:'XS', province:'นครศรีธรรมราช'},
    {station_id:'ST_XS_พิปูน',           name:'กฟฟ.พิปูน',           short_name:'พิปูน',           type:'XS', province:'นครศรีธรรมราช'},
    {station_id:'ST_XS_ถ้ำพรรณรา',       name:'กฟฟ.ถ้ำพรรณรา',       short_name:'ถ้ำพรรณรา',       type:'XS', province:'นครศรีธรรมราช'},
    {station_id:'ST_XS_เกาะเต่า',        name:'กฟฟ.เกาะเต่า',        short_name:'เกาะเต่า',        type:'XS', province:'สุราษฎร์ธานี'},
    {station_id:'ST_XS_เคียนชา',         name:'กฟฟ.เคียนชา',         short_name:'เคียนชา',         type:'XS', province:'สุราษฎร์ธานี'},
    {station_id:'ST_XS_ท่าฉาง',          name:'กฟฟ.ท่าฉาง',          short_name:'ท่าฉาง',          type:'XS', province:'สุราษฎร์ธานี'},
    {station_id:'ST_XS_พนม',             name:'กฟฟ.พนม',             short_name:'พนม',             type:'XS', province:'สุราษฎร์ธานี'},
    {station_id:'ST_XS_วิภาวดี',         name:'กฟฟ.วิภาวดี',         short_name:'วิภาวดี',         type:'XS', province:'สุราษฎร์ธานี'},
    {station_id:'ST_XS_เชียรใหญ่',       name:'กฟฟ.เชียรใหญ่',       short_name:'เชียรใหญ่',       type:'XS', province:'นครศรีธรรมราช'},
    {station_id:'ST_XS_เกาะยาว',         name:'กฟฟ.เกาะยาว',         short_name:'เกาะยาว',         type:'XS', province:'พังงา'},
    {station_id:'ST_XS_กะปง',            name:'กฟฟ.กะปง',            short_name:'กะปง',            type:'XS', province:'พังงา'},
    {station_id:'ST_XS_คุระบุรี',        name:'กฟฟ.คุระบุรี',        short_name:'คุระบุรี',        type:'XS', province:'พังงา'},
    {station_id:'ST_XS_บ้านเขาหลัก',    name:'กฟฟ.บ้านเขาหลัก',    short_name:'บ้านเขาหลัก',    type:'XS', province:'พังงา'},
    {station_id:'ST_XS_ดอนสัก',          name:'กฟฟ.ดอนสัก',          short_name:'ดอนสัก',          type:'XS', province:'สุราษฎร์ธานี'},
    {station_id:'ST_XS_ชัยบุรี',         name:'กฟฟ.ชัยบุรี',         short_name:'ชัยบุรี',         type:'XS', province:'สุราษฎร์ธานี'},
    {station_id:'ST_XS_นบพิตำ',          name:'กฟฟ.นบพิตำ',          short_name:'นบพิตำ',          type:'XS', province:'นครศรีธรรมราช'},
    {station_id:'ST_XS_วังวิเศษ',        name:'กฟฟ.วังวิเศษ',        short_name:'วังวิเศษ',        type:'XS', province:'ตรัง'},
    {station_id:'ST_XS_รัษฎา',           name:'กฟฟ.รัษฎา',           short_name:'รัษฎา',           type:'XS', province:'ตรัง'}
  ];

  try {
    // อ่านหน่วยงานที่มีอยู่แล้ว
    var existing = getSheetData(SH.STATIONS);
    var existingIds = existing.map(function(s){ return s.station_id; });

    var added = 0;
    STATION_LIST.forEach(function(s) {
      if (existingIds.indexOf(s.station_id) === -1) {
        appendRow(SH.STATIONS, {
          station_id : s.station_id,
          name       : s.name,
          short_name : s.short_name,
          type       : s.type,
          province   : s.province,
          active     : 'Y'
        });
        added++;
      }
    });

    var msg = 'นำเข้าหน่วยงานสำเร็จ: เพิ่มใหม่ ' + added + ' รายการ (ทั้งหมด 79 หน่วยงาน)';
    Logger.log(msg);
    return msg;

  } catch(e) {
    Logger.log('importAllStations error: ' + e.message);
    return 'Error: ' + e.message;
  }
}

// ============================================================
//  clearAndReimportStations() — ล้างและนำเข้าใหม่ทั้งหมด
//  ⚠️ ใช้เมื่อต้องการ reset stations เท่านั้น
// ============================================================
function clearAndReimportStations() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sh = ss.getSheetByName(SH.STATIONS);
  if (!sh) { setupStationsSheet(); sh = ss.getSheetByName(SH.STATIONS); }

  // เก็บ header row ไว้
  var header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];

  // ลบข้อมูลทั้งหมดยกเว้น header
  if (sh.getLastRow() > 1) {
    sh.deleteRows(2, sh.getLastRow() - 1);
  }

  return importAllStations();
}


// ============================================================
//  getStations — สำหรับ form.html
// ============================================================
function getStations() {
  return getSheetData(SH.STATIONS).filter(function(s){
    return s.active !== 'N';
  });
}