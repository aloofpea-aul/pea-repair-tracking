// ============================================================
//  DataService.gs  –  API functions called from HTML frontend
// ============================================================

function getConfig() {
  const rows = getSheetData(SH.CONFIG);
  const cfg = {};
  rows.forEach(r => { cfg[r['KEY']] = r['VALUE']; });
  return cfg;
}

function getStations() {
  return getSheetData(SH.STATIONS).filter(s => s.active === true || s.active === 'TRUE');
}

function getChecklist() {
  return getSheetData(SH.CHECKLIST);
}

function getInspections() {
  return getSheetData(SH.INSPECTIONS);
}

function getInspectionById(id) {
  const insp   = getSheetData(SH.INSPECTIONS).find(r => r.inspection_id === id);
  const scores = getSheetData(SH.SCORES).filter(r => r.inspection_id === id);
  return { inspection: insp, scores: scores };
}

function getDashboardData() {
  const stations    = getStations();
  const inspections = getSheetData(SH.INSPECTIONS);
  const checklist   = getSheetData(SH.CHECKLIST);

  const stationMap = {};
  stations.forEach(s => {
    stationMap[s.station_id] = { ...s, last_inspect: null, last_score: null, last_percent: null, last_grade: null, inspect_count: 0 };
  });

  inspections.forEach(ins => {
    const st = stationMap[ins.station_id];
    if (!st) return;
    st.inspect_count++;
    if (!st.last_inspect || String(ins.inspect_date) > String(st.last_inspect)) {
      st.last_inspect = ins.inspect_date;
      st.last_score   = ins.total_score;
      st.last_percent = ins.percent;
      st.last_grade   = ins.grade;
    }
  });

  const categories = [];
  const seen = {};
  checklist.forEach(c => {
    if (!seen[c.category]) {
      seen[c.category] = true;
      categories.push({
        category     : c.category,
        category_max : c.category_max,
        item_count   : checklist.filter(x => x.category === c.category).length
      });
    }
  });

  return {
    stations      : Object.values(stationMap),
    catSummary    : categories,
    totalInspect  : inspections.length,
    totalStations : stations.length,
    inspectedCount: Object.values(stationMap).filter(s => s.last_inspect).length
  };
}

function saveInspection(payload) {
  try {
    const ss        = SpreadsheetApp.openById(SS_ID);
    const inspSh    = ss.getSheetByName(SH.INSPECTIONS);
    const scoresSh  = ss.getSheetByName(SH.SCORES);
    const checklist = getSheetData(SH.CHECKLIST);

    const inspId    = genId('INS');
    const createdAt = nowTH();

    let totalScore = 0;
    let maxScore   = 0;
    checklist.forEach(item => { maxScore += Number(item.max_score) || 0; });
    payload.scores.forEach(s  => { totalScore += Number(s.score)  || 0; });
    totalScore = Math.round(totalScore * 10) / 10;

    const percent = maxScore > 0 ? Math.round((totalScore / maxScore) * 100) : 0;
    const grade   = getGrade(percent);

    const inspHeaders = inspSh.getRange(1,1,1,inspSh.getLastColumn()).getValues()[0];
    const inspRow = {
      inspection_id  : inspId,
      station_id     : payload.station_id,
      station_name   : payload.station_name,
      inspect_date   : payload.inspect_date,
      inspector_name : payload.inspector_name,
      inspector_email: payload.inspector_email || '',
      status         : 'เสร็จสิ้น',
      total_score    : totalScore,
      max_score      : maxScore,
      percent        : percent,
      grade          : grade,
      note           : payload.note || '',
      created_at     : createdAt
    };
    inspSh.appendRow(inspHeaders.map(h => inspRow[h] !== undefined ? inspRow[h] : ''));

    const scoreHeaders = scoresSh.getRange(1,1,1,scoresSh.getLastColumn()).getValues()[0];
    const checkMap = {};
    checklist.forEach(c => { checkMap[c.item_id] = c; });

    payload.scores.forEach(s => {
      const item = checkMap[s.item_id] || {};
      const scoreRow = {
        score_id     : genId('SC'),
        inspection_id: inspId,
        item_id      : s.item_id,
        category     : item.category  || '',
        item_no      : item.item_no   || '',
        topic        : item.topic     || '',
        max_score    : item.max_score || 0,
        score        : s.score,
        note         : s.note || ''
      };
      scoresSh.appendRow(scoreHeaders.map(h => scoreRow[h] !== undefined ? scoreRow[h] : ''));
    });

    writeLog('SAVE_INSPECTION', `${inspId} | ${payload.station_name} | ${totalScore}/${maxScore}`, payload.inspector_email);
    return { success: true, inspection_id: inspId, total_score: totalScore, max_score: maxScore, percent: percent, grade: grade };
  } catch(err) {
    return { success: false, error: err.message };
  }
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
    const ss = SpreadsheetApp.openById(SS_ID);
    deleteRowById(ss.getSheetByName(SH.INSPECTIONS), 'inspection_id', inspId);
    deleteRowsByField(ss.getSheetByName(SH.SCORES), 'inspection_id', inspId);
    writeLog('DELETE_INSPECTION', inspId, '');
    return { success: true };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function deleteRowById(sh, keyCol, val) {
  const data = sh.getDataRange().getValues();
  const hIdx = data[0].indexOf(keyCol);
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][hIdx]) === String(val)) { sh.deleteRow(i + 1); break; }
  }
}

function deleteRowsByField(sh, keyCol, val) {
  const data = sh.getDataRange().getValues();
  const hIdx = data[0].indexOf(keyCol);
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][hIdx]) === String(val)) sh.deleteRow(i + 1);
  }
}

// ============================================================
//  เพิ่มเติมใน DataService.gs
//  วางต่อท้ายไฟล์ DataService.gs เดิม
// ============================================================

// mapping item_id -> row number ในชีตสรุป
var ITEM_ROW_MAP = {
  1:4,2:5,3:6,4:7,5:8,6:9,7:10,8:11,9:12,
  10:14,11:15,12:16,13:17,
  14:19,15:20,16:21,
  17:23,18:24,19:25,
  20:27,21:28,22:29,
  23:31,24:32,25:33,26:34,27:35,28:36,
  29:38,30:39,31:40,32:41,
  33:43
};

// ดึงชื่อชีตสรุปตามประเภทหน่วยงาน
function getSummarySheetName_(type) {
  if (type === 'L' || type === 'M') return 'รวมคะแนน กฟจ.กฟส.L-M';
  if (type === 'S')  return 'รวมคะแนน กฟส.S';
  if (type === 'XS') return 'รวมคะแนน กฟส.XS';
  return null;
}

// อัพเดทคะแนนลงชีตสรุป
function updateSummarySheet_(stationShortName, stationType, scores) {
  try {
    var shName = getSummarySheetName_(stationType);
    if (!shName) return;
    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = ss.getSheetByName(shName);
    if (!sh) return;

    // หาคอลัมน์ของหน่วยงาน (row 2 = index 2, col D+ = col 4+)
    var headerRow = sh.getRange(2, 1, 1, sh.getLastColumn()).getValues()[0];
    var col = -1;
    for (var i = 3; i < headerRow.length; i++) {
      if (String(headerRow[i]).trim() === String(stationShortName).trim()) {
        col = i + 1; // 1-indexed
        break;
      }
    }
    if (col === -1) return; // ไม่พบหน่วยงาน

    // เขียนคะแนนแต่ละข้อ
    scores.forEach(function(s) {
      var row = ITEM_ROW_MAP[parseInt(s.item_id)];
      if (row) {
        sh.getRange(row, col).setValue(s.score);
      }
    });

    // คำนวณรวมคะแนน (row 44)
    var total = scores.reduce(function(sum, s) { return sum + (parseFloat(s.score) || 0); }, 0);
    sh.getRange(44, col).setValue(Math.round(total * 10) / 10);

  } catch(e) {
    Logger.log('updateSummarySheet error: ' + e.message);
  }
}

// ====================================================
// saveInspection — บันทึกการตรวจ + อัพเดทชีตสรุป
// ====================================================
function saveInspection(payload) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);

    // หาข้อมูลสถานี
    var stations = getSheetData(SH.STATIONS);
    var station = stations.find(function(s) {
      return s.station_id === payload.station_id;
    });

    var inspectId = genId('INS');
    var now = nowTH();
    var user = payload.inspector_email || Session.getActiveUser().getEmail() || 'unknown';

    // บันทึก INSPECTIONS
    appendRow(SH.INSPECTIONS, {
      inspect_id    : inspectId,
      station_id    : payload.station_id,
      station_name  : payload.station_name,
      inspect_date  : payload.inspect_date,
      inspector_name: payload.inspector_name,
      inspector_email: payload.inspector_email || '',
      total_score   : payload.total_score,
      note          : payload.note || '',
      created_at    : now
    });

    // บันทึก SCORES
    var scoreSheet = ss.getSheetByName(SH.SCORES);
    (payload.scores || []).forEach(function(s) {
      appendRow(SH.SCORES, {
        score_id   : genId('SCR'),
        inspect_id : inspectId,
        station_id : payload.station_id,
        item_id    : s.item_id,
        item_no    : s.item_no,
        score      : s.score,
        pct        : s.pct || '',
        note       : s.note || ''
      });
    });

    // อัพเดทชีตสรุป
    if (station) {
      updateSummarySheet_(station.short_name, station.type, payload.scores || []);
    }

    // LOG
    writeLog('saveInspection', 'station=' + payload.station_id + ' score=' + payload.total_score, user);

    // คำนวณ percent / grade
    var maxScore = 100;
    var pct = Math.round(payload.total_score / maxScore * 100);
    var grade = payload.total_score >= 90 ? 'A' :
                payload.total_score >= 80 ? 'B' :
                payload.total_score >= 70 ? 'C' :
                payload.total_score >= 60 ? 'D' : 'F';

    return {
      success     : true,
      inspect_id  : inspectId,
      total_score : payload.total_score,
      max_score   : maxScore,
      percent     : pct,
      grade       : grade
    };

  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ====================================================
// getReportData — ดึงข้อมูลรายงานแยกกลุ่ม
// ====================================================
function getReportData() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var result = {};

    var groups = [
      { key: 'lm', sheetName: 'รวมคะแนน กฟจ.กฟส.L-M', title: 'กฟจ. / กฟส.L-M' },
      { key: 's',  sheetName: 'รวมคะแนน กฟส.S',        title: 'กฟส.S' },
      { key: 'xs', sheetName: 'รวมคะแนน กฟส.XS',       title: 'กฟส.XS' }
    ];

    groups.forEach(function(g) {
      var sh = ss.getSheetByName(g.sheetName);
      if (!sh) { result[g.key] = null; return; }

      var data = sh.getDataRange().getValues();
      var headerRow = data[1]; // row 2 (index 1)

      // ดึงชื่อหน่วยงาน col D+
      var stations = [];
      for (var c = 3; c < headerRow.length; c++) {
        if (headerRow[c] !== null && headerRow[c] !== '') {
          stations.push({ name: String(headerRow[c]), col: c });
        }
      }

      // ดึงรายการข้อ (rows ที่มีเลขในคอลัมน์ A)
      var items = [];
      var currentCat = '';
      for (var r = 2; r < data.length; r++) {
        var rowA = data[r][0];
        var rowB = data[r][1];
        var rowC = data[r][2];

        if (rowA === null && rowB !== null && rowB !== '') {
          // แถวหมวด
          currentCat = String(rowB);
        } else if (rowA !== null && rowB !== null) {
          var no = String(rowA).trim();
          if (no === 'รวมคะแนน') {
            // แถวรวม
            var totals = {};
            stations.forEach(function(st) {
              totals[st.name] = parseFloat(data[r][st.col]) || 0;
            });
            result[g.key + '_totals'] = totals;
          } else {
            var scores = {};
            stations.forEach(function(st) {
              scores[st.name] = data[r][st.col] !== null && data[r][st.col] !== '' 
                ? parseFloat(data[r][st.col]) : null;
            });
            items.push({
              no    : no,
              topic : String(rowB || '').replace(/\n/g, ' ').trim(),
              max   : parseFloat(rowC) || 0,
              cat   : currentCat,
              scores: scores
            });
          }
        }
      }

      result[g.key] = {
        key      : g.key,
        title    : g.title,
        stations : stations.map(function(s){ return s.name; }),
        items    : items
      };
    });

    return result;

  } catch(e) {
    return { error: e.message };
  }
}