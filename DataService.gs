// ============================================================
//  DataService.gs  –  API functions called from HTML frontend
//  pea-repair-tracking | กฟต.2 ปี 2569
// ============================================================

function getConfig() {
  var rows = getSheetData(SH.CONFIG);
  var cfg = {};
  rows.forEach(function(r) { cfg[r['key']] = r['value']; });
  return cfg;
}

// ============================================================
//  getStations — active !== 'N'
// ============================================================
function getStations() {
  return getSheetData(SH.STATIONS).filter(function(s) {
    return s.active !== 'N';
  });
}

function getChecklist() { return getSheetData(SH.CHECKLIST); }
function getInspections() { return getSheetData(SH.INSPECTIONS); }

function getInspectionById(id) {
  var insp   = getSheetData(SH.INSPECTIONS).find(function(r){ return r.inspection_id === id; });
  var scores = getSheetData(SH.SCORES).filter(function(r){ return r.inspection_id === id; });
  return { inspection: insp, scores: scores };
}

function getGrade(percent) {
  if (percent >= 90) return 'A (ดีเยี่ยม)';
  if (percent >= 80) return 'B (ดี)';
  if (percent >= 70) return 'C (พอใช้)';
  if (percent >= 60) return 'D (ต้องปรับปรุง)';
  return 'F (ไม่ผ่าน)';
}

function deleteInspection(inspId) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    deleteRowById(ss.getSheetByName(SH.INSPECTIONS), 'inspection_id', inspId);
    deleteRowsByField(ss.getSheetByName(SH.SCORES), 'inspection_id', inspId);
    writeLog('DELETE_INSPECTION', inspId, '');
    return { success: true };
  } catch(err) { return { success: false, error: err.message }; }
}

function deleteRowById(sh, keyCol, val) {
  var data = sh.getDataRange().getValues();
  var hIdx = data[0].indexOf(keyCol);
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][hIdx]) === String(val)) { sh.deleteRow(i + 1); break; }
  }
}

function deleteRowsByField(sh, keyCol, val) {
  var data = sh.getDataRange().getValues();
  var hIdx = data[0].indexOf(keyCol);
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][hIdx]) === String(val)) sh.deleteRow(i + 1);
  }
}

// ============================================================
//  mapping item_id -> row ในชีตสรุป
// ============================================================
var ITEM_ROW_MAP = {
  1:4,  2:5,  3:6,  4:7,  5:8,  6:9,  7:10, 8:11, 9:12,
  10:14,11:15,12:16,13:17,
  14:19,15:20,16:21,
  17:23,18:24,19:25,
  20:27,21:28,22:29,
  23:31,24:32,25:33,26:34,27:35,28:36,
  29:38,30:39,31:40,32:41,
  33:43
};

function getSummarySheetName_(type) {
  if (type === 'L' || type === 'M') return 'รวมคะแนน กฟจ.กฟส.L-M';
  if (type === 'S')  return 'รวมคะแนน กฟส.S';
  if (type === 'XS') return 'รวมคะแนน กฟส.XS';
  return null;
}

function updateSummarySheet_(stationShortName, stationType, scores) {
  try {
    var shName = getSummarySheetName_(stationType);
    if (!shName) return;
    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = ss.getSheetByName(shName);
    if (!sh) return;
    var headerRow = sh.getRange(2, 1, 1, sh.getLastColumn()).getValues()[0];
    var col = -1;
    for (var i = 3; i < headerRow.length; i++) {
      if (String(headerRow[i]).trim() === String(stationShortName).trim()) { col = i + 1; break; }
    }
    if (col === -1) return;
    scores.forEach(function(s) {
      var row = ITEM_ROW_MAP[parseInt(s.item_id)];
      if (row) sh.getRange(row, col).setValue(s.score);
    });
    var total = scores.reduce(function(sum, s){ return sum + (parseFloat(s.score)||0); }, 0);
    sh.getRange(44, col).setValue(Math.round(total * 10) / 10);
  } catch(e) { Logger.log('updateSummarySheet error: ' + e.message); }
}

// ============================================================
//  saveInspection — บันทึกการตรวจ + อัพเดทชีตสรุป
// ============================================================
function saveInspection(payload) {
  try {
    var ss       = SpreadsheetApp.openById(SS_ID);
    var stations = getSheetData(SH.STATIONS);
    var station  = stations.find(function(s){ return s.station_id === payload.station_id; });

    var inspectId = genId('INS');
    var now  = nowTH();
    var user = payload.inspector_email || Session.getActiveUser().getEmail() || 'unknown';

    var maxScore   = 100;
    var totalScore = Math.round((payload.total_score || 0) * 10) / 10;
    var pct        = Math.round(totalScore / maxScore * 100);
    var grade      = totalScore >= 90 ? 'A (ดีเยี่ยม)' :
                     totalScore >= 80 ? 'B (ดี)'        :
                     totalScore >= 70 ? 'C (พอใช้)'     :
                     totalScore >= 60 ? 'D (ต้องปรับปรุง)' : 'F (ไม่ผ่าน)';

    appendRow(SH.INSPECTIONS, {
      inspect_id     : inspectId,
      station_id     : payload.station_id,
      station_name   : payload.station_name,
      inspect_date   : payload.inspect_date,
      inspector_name : payload.inspector_name,
      inspector_email: payload.inspector_email || '',
      committee_id   : payload.committee_id || '',
      total_score    : totalScore,
      max_score      : maxScore,
      percent        : pct,
      grade          : grade,
      note           : payload.note || '',
      created_at     : now
    });

    (payload.scores || []).forEach(function(s) {
      appendRow(SH.SCORES, {
        score_id   : genId('SC'),
        inspect_id : inspectId,
        station_id : payload.station_id,
        item_id    : s.item_id,
        item_no    : s.item_no || '',
        score      : s.score,
        pct        : s.pct || '',
        note       : s.note || ''
      });
    });

    if (station) updateSummarySheet_(station.short_name, station.type, payload.scores || []);

    writeLog('saveInspection', inspectId + ' | ' + payload.station_name + ' | ' + totalScore, user);

    return { success: true, inspect_id: inspectId, total_score: totalScore, max_score: maxScore, percent: pct, grade: grade };
  } catch(e) { return { success: false, error: e.message }; }
}

// ============================================================
//  getReportData — ดึงข้อมูลรายงานแยกกลุ่ม
// ============================================================
function getReportData() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var result = {};
    var groups = [
      { key: 'lm', sheetName: 'รวมคะแนน กฟจ.กฟส.L-M', title: 'กฟจ. / กฟส.L-M' },
      { key: 's',  sheetName: 'รวมคะแนน กฟส.S',        title: 'กฟส.S'           },
      { key: 'xs', sheetName: 'รวมคะแนน กฟส.XS',       title: 'กฟส.XS'          }
    ];
    groups.forEach(function(g) {
      var sh = ss.getSheetByName(g.sheetName);
      if (!sh) { result[g.key] = null; return; }
      var data = sh.getDataRange().getValues();
      var headerRow = data[1];
      var stationCols = [];
      for (var c = 3; c < headerRow.length; c++) {
        if (headerRow[c] !== null && String(headerRow[c]).trim() !== '')
          stationCols.push({ name: String(headerRow[c]).trim(), col: c });
      }
      var items = [], currentCat = '', totals = {};
      for (var r = 2; r < data.length; r++) {
        var rowA = data[r][0], rowB = data[r][1], rowC = data[r][2];
        if (rowA === null && rowB !== null && String(rowB).trim() !== '') {
          currentCat = String(rowB).trim();
        } else if (rowA !== null && rowB !== null && String(rowA).trim() !== '') {
          var no = String(rowA).trim();
          if (no === 'รวมคะแนน') {
            stationCols.forEach(function(st){ totals[st.name] = parseFloat(data[r][st.col]) || 0; });
            result[g.key + '_totals'] = totals;
          } else {
            var scoresObj = {};
            stationCols.forEach(function(st){
              var v = data[r][st.col];
              scoresObj[st.name] = (v !== null && v !== '') ? parseFloat(v) : null;
            });
            items.push({ no: no, topic: String(rowB||'').replace(/\n/g,' ').trim(), max: parseFloat(rowC)||0, cat: currentCat, scores: scoresObj });
          }
        }
      }
      result[g.key] = { key: g.key, title: g.title, stations: stationCols.map(function(s){ return s.name; }), items: items };
    });
    return result;
  } catch(e) { return { error: e.message }; }
}

// ============================================================
//  Committee functions
// ============================================================
function getCommittees() {
  return getSheetData('COMMITTEES').filter(function(c){ return c.active !== 'N'; });
}

function getCommitteeById(id) {
  return getSheetData('COMMITTEES').find(function(c){ return c.committee_id === id; }) || null;
}

function getCommitteesByStation(stationId) {
  return getSheetData('COMMITTEES').filter(function(c){
    if (c.active === 'N') return false;
    return (c.station_ids || '').split('|').indexOf(stationId) > -1;
  });
}

function saveCommittee(payload) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = ss.getSheetByName('COMMITTEES');
    if (!sh) {
      sh = ss.insertSheet('COMMITTEES');
      sh.appendRow(['committee_id','name','members','station_ids','active']);
      sh.getRange(1,1,1,5).setFontWeight('bold').setBackground('#1a3c5e').setFontColor('#ffffff');
    }
    if (payload.committee_id) {
      var data = sh.getDataRange().getValues();
      var headers = data[0];
      var idIdx = headers.indexOf('committee_id');
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][idIdx]) === String(payload.committee_id)) {
          headers.forEach(function(h, j){ if (payload[h] !== undefined) sh.getRange(i+1, j+1).setValue(payload[h]); });
          break;
        }
      }
    } else {
      payload.committee_id = genId('COM');
      appendRow('COMMITTEES', payload);
    }
    writeLog('saveCommittee', payload.name, '');
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function deleteCommittee(id) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = ss.getSheetByName('COMMITTEES');
    if (!sh) return { success: false, error: 'ไม่พบ sheet' };
    var data = sh.getDataRange().getValues();
    var idIdx = data[0].indexOf('committee_id');
    var activeIdx = data[0].indexOf('active');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]) === String(id)) { sh.getRange(i+1, activeIdx+1).setValue('N'); break; }
    }
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}