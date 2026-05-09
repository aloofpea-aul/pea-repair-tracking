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