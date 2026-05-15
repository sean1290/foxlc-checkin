var SPREADSHEET_ID = '17JJYwp1rHUSf2j7DXI2wKC84hXJ2ghOPDZdmQALeyo0';

/* ═══════════════════════════════════════════════════════════════
 * WEB APP ENDPOINTS  (기존 코드 - 변경 없음)
 * ═══════════════════════════════════════════════════════════════ */

function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  if (action === 'getStudents') {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var names = getStudentList(ss);
    return out({ status: 'ok', names: names });
  }
  return out({ status: 'ok', message: 'FoxLC running' });
}

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if      (d.session === 'morning') { appendMorning(ss, d); }
    else if (d.session === 'evening') { appendEvening(ss, d); }
    else if (d.session === 'mission') { appendMission(ss, d); }
    return out({ status: 'ok' });
  } catch (err) {
    return out({ status: 'error', message: err.toString() });
  }
}

function appendMorning(ss, d) {
  var sheet = getSheet(ss, '입실체크', [
    '날짜','이름','입실/퇴실',
    '감정','감정 신체부위','감정 원인/의도',
    '에너지 레벨',
    '걱정 혹은 불안한 일',
    '걱정의 크기 + 이유 (걱정 있을 때)',
    '오늘 집중하고 싶은 목표',
    '목표를 이루기 위한 계획',
    '마음이 어렵다면',
    '오늘 친구들에게 좋은 태도',
    '오늘_기대되는_것'
  ]);
  var anx = '';
  if (d.axis1_anxiety === 'yes') {
    var p = [];
    if (d.axis1_anxiety_level)  { p.push('크기: ' + d.axis1_anxiety_level); }
    if (d.axis1_anxiety_reason) { p.push('이유: ' + d.axis1_anxiety_reason); }
    anx = p.join(' / ');
  }
  sheet.appendRow([
    formatDate(d.date),              // A 날짜
    d.student_name          || '',   // B 이름
    '입실',                          // C 입실/퇴실
    d.axis1_emotion_followup
      ? (d.axis1_emotion || '') + ' // ' + d.axis1_emotion_followup
      : (d.axis1_emotion  || ''),    // D 감정
    d.axis1_emotion_body    || '',   // E 감정 신체부위
    d.axis1_emotion_followup|| '',   // F 감정 원인/의도
    d.axis1_energy          || '',   // G 에너지 레벨
    d.axis1_anxiety         || '',   // F 걱정 여부
    anx,                             // G 걱정 크기+이유
    d.axis2_goal            || '',   // H 오늘 목표
    d.axis1_difficulty      || '',   // I 목표 계획
    d.axis1_coping          || '',   // J 마음이 어렵다면
    d.axis3_relation_intent || '',   // K 좋은 태도
    d['오늘_기대되는_것']   || ''    // L 오늘 기대되는 것
  ]);
}

function appendEvening(ss, d) {
  var sheet = getSheet(ss, '퇴실체크', [
    '날짜','이름','마음','감정 원인/의도',
    '활동 만족도 (1~5)',
    '오늘 친구들과의 관계',
    '어떤 점이 어려웠나요? (관계 1~2점 시)',
    '오늘 보인 좋은 태도',
    '감정 조절을 잘 했나요?',
    '어떤 방법을 썼는지 (감정조절 예 선택시)',
    '오늘 누군가와 갈등이 있었나요?',
    '어떻게 해결했나요? (갈등 예 선택시)',
    '오늘 활동 집중도 (1~5)',
    '오늘 가장 자랑스러운 것',
    '내일 실천할 한 가지',
    '오늘의 질문'
  ]);
  sheet.appendRow([
    formatDate(d.date),
    d.student_name                    || '',
    d.axis1_emotion_change            || '',
    d.axis1_emotion_change_followup   || '',
    d.axis2_goal_score                || '',
    d.axis3_relation          || '',
    d.axis3_relation_reason   || '',
    d.axis3_prosocial         || '',
    d.axis1_selfcontrol       || '',
    d.axis1_selfcontrol_type  || '',
    d.axis3_conflict          || '',
    d.axis3_conflict_strategy || '',
    d.axis4_focus             || '',
    d.self_praise             || '',
    d.next_action             || '',
    d.daily_rotation          || ''
  ]);
}

function appendMission(ss, d) {
  var sheet = getSheet(ss, '특별미션', [
    '날짜','이름','미션ID',
    '선택한답변','내생각(직접입력)','선택한보상'
  ]);
  sheet.appendRow([
    formatDate(d.date), d.student_name || '',
    d.mission_id || '', d.selected_option || '',
    d.reflection || '', d.selected_reward || ''
  ]);
}

function getSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    var r = sheet.getRange(1, 1, 1, headers.length);
    r.setValues([headers]);
    r.setFontWeight('bold');
    r.setBackground('#1A73E8');
    r.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function formatDate(dateStr) {
  if (!dateStr) return '';
  try {
    var p = dateStr.split('-');
    var d = new Date(parseInt(p[0]), parseInt(p[1])-1, parseInt(p[2]));
    var days = ['일','월','화','수','목','금','토'];
    return (d.getMonth()+1)+'/'+d.getDate()+' ('+days[d.getDay()]+')';
  } catch(e) { return dateStr; }
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}


/* ═══════════════════════════════════════════════════════════════
 * DASHBOARD
 * - B2: 학생 이름 드롭다운 → 선택하면 자동 갱신
 * - C2: 🔄 새로고침 버튼 (체크박스) → 클릭하면 즉시 갱신
 * - 첫 사용 시: setupDashboard() 한 번 실행
 * ═══════════════════════════════════════════════════════════════ */

var DASHBOARD_NAME = 'Dashboard';
var SEARCH_CELL    = 'B2';
var REFRESH_CELL   = 'C2';

var SECTIONS = [
  { key:'morning', tab:'입실체크', label:'🌅 입실 체크', titleBg:'#FF6B35', headerBg:'#FFE8DF', bandBg:'#FFF8F4' },
  { key:'evening', tab:'퇴실체크', label:'🌙 퇴실 체크', titleBg:'#5BAD8A', headerBg:'#E0F7F5', bandBg:'#F4FBFA' },
  { key:'mission', tab:'특별미션', label:'✨ 특별 미션', titleBg:'#A78BFA', headerBg:'#EDE9FE', bandBg:'#F8F5FF' }
];

/* ── Simple trigger ─────────────────────────────────────────── */

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    var shName = sh.getName();

    if (shName === DASHBOARD_NAME) {
      var cell = e.range.getA1Notation();
      if (cell === SEARCH_CELL) {
        refreshDashboard(e.source);
      } else if (cell === REFRESH_CELL && e.value === 'TRUE') {
        refreshDashboard(e.source);
        sh.getRange(REFRESH_CELL).setValue(false);
      }
    } else if (shName === '학생 관리') {
      // 학생 명단 변경 시 Dashboard 드롭다운 자동 갱신
      buildNameDropdown(e.source);
    }
  } catch (err) {
    Logger.log('onEdit error: ' + err);
  }
}

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('🦊 FoxLC')
      .addItem('Dashboard 새로고침', 'refreshDashboard')
      .addItem('학생 이름 드롭다운 갱신', 'buildNameDropdown')
      .addItem('Dashboard 초기 셋업', 'setupDashboard')
      .addSeparator()
      .addItem('학생 관리 탭 초기 셋업 (최초 1회)', 'setupStudentSheet')
      .addToUi();
  } catch(e) {}
}

/* ── Setup ──────────────────────────────────────────────────── */

function setupDashboard() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dash = ss.getSheetByName(DASHBOARD_NAME);
  if (!dash) {
    dash = ss.insertSheet(DASHBOARD_NAME);
  } else {
    dash.clear();
    dash.clearConditionalFormatRules();
    dash.getRange(1, 1, dash.getMaxRows(), dash.getMaxColumns()).clearDataValidations();
  }

  // Row 1 — title
  dash.getRange('A1:O1').merge()
    .setValue('🦊 Fox Learning Center · 학생 기록 대시보드')
    .setBackground('#1A1A2E').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(16).setHorizontalAlignment('center').setVerticalAlignment('middle');
  dash.setRowHeight(1, 44);

  // Row 2 — search label
  dash.getRange('A2')
    .setValue('🔍 학생 이름:')
    .setBackground('#FFF8F0').setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('right').setVerticalAlignment('middle');

  // B2 — name dropdown (styled, validation added by buildNameDropdown)
  dash.getRange(SEARCH_CELL)
    .setBackground('#FFFFFF').setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#FF6B35', SpreadsheetApp.BorderStyle.SOLID_THICK);

  // C2 — refresh checkbox button
  dash.getRange(REFRESH_CELL)
    .setValue(false)
    .setBackground('#FF6B35').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setNote('체크하면 즉시 새로고침')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());

  // D2:O2 — hint text
  dash.getRange('D2:O2').merge()
    .setValue('🔄 새로고침     ← 이름 선택 후 C2 체크박스를 누르거나, 이름 선택만 해도 자동 갱신됩니다')
    .setBackground('#FFF8F0').setFontColor('#6B7280').setFontStyle('italic')
    .setFontSize(11).setVerticalAlignment('middle');

  dash.setColumnWidth(1, 190);   // A
  dash.setColumnWidth(2, 220);   // B dropdown
  dash.setColumnWidth(3, 60);    // C button
  for (var i = 4; i <= 16; i++) dash.setColumnWidth(i, 165);
  dash.setRowHeight(2, 40);
  dash.setFrozenRows(2);
  dash.setHiddenGridlines(true);

  buildNameDropdown(ss);
  refreshDashboard(ss);
  try {
    SpreadsheetApp.getUi().alert('Dashboard 준비 완료!\n\nB2 드롭다운으로 이름 선택 → 자동 갱신\nC2 체크박스 클릭 → 수동 새로고침');
  } catch(e) {}
}

/* ── Student list ───────────────────────────────────────────── */

// '학생 관리' 탭에서 이름 목록 읽기 (없으면 체크인 탭에서 수집)
function getStudentList(ss) {
  ss = ss || SpreadsheetApp.openById(SPREADSHEET_ID);
  var mgmt = ss.getSheetByName('학생 관리');
  if (mgmt) {
    var data = mgmt.getDataRange().getValues();
    var names = [];
    for (var r = 1; r < data.length; r++) {  // row 1부터 (0행은 헤더)
      var nm = String(data[r][0] || '').trim();
      if (nm) names.push(nm);
    }
    if (names.length > 0) return names;
  }
  // fallback: 체크인 탭에서 수집
  var nameSet = {};
  var tabs = ['입실체크', '퇴실체크', '특별미션'];
  for (var t = 0; t < tabs.length; t++) {
    var src = ss.getSheetByName(tabs[t]);
    if (!src) continue;
    var rows = src.getDataRange().getValues();
    if (rows.length < 2) continue;
    var nameIdx = findColumnIndex(rows[0], ['이름', 'student_name', 'name']);
    if (nameIdx === -1) continue;
    for (var r = 1; r < rows.length; r++) {
      var nm = String(rows[r][nameIdx] || '').trim();
      if (nm) nameSet[nm] = true;
    }
  }
  return Object.keys(nameSet).sort();
}

// '학생 관리' 탭 최초 생성 + 현재 학생 목록 입력
function setupStudentSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('학생 관리');
  if (!sheet) {
    sheet = ss.insertSheet('학생 관리');
  } else {
    var ui = SpreadsheetApp.getUi();
    var res = ui.alert('⚠️ 이미 존재합니다', '"학생 관리" 탭이 이미 있습니다. 덮어쓸까요?', ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) return;
    sheet.clear();
  }

  // 헤더
  sheet.getRange('A1').setValue('이름')
    .setBackground('#1A73E8').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(13);
  sheet.getRange('B1').setValue('메모')
    .setBackground('#1A73E8').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(13);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 280);

  // 현재 학생 목록
  var names = [
    '강설현','고윤아','김민준','김소은','김승운','김윤성','김유진','김재윤','김재이',
    '김제인','김태웅','김폭스','김현서','김현우','박래승','박영찬','서재현','성시윤',
    '손아율','손태율','송서후','송예온','송지호','신연재','신유준','여소현','여승현',
    '여창현','오엘리','오하성','유가영','이디아나','이서호','이시호','이윤호','이이린',
    '이재이','이지오','이지호','이서호','정윤우','조아인','조윤재','조은혁','조유나',
    '조이서','조하은','최서하','최윤서','최윤재','최주원','홍근준'
  ];

  var rows = names.map(function(n) { return [n, '']; });
  sheet.getRange(2, 1, rows.length, 2).setValues(rows).setFontSize(12);

  // 안내 메시지
  sheet.getRange('D1').setValue('ℹ️  A열에 이름을 추가/삭제하면 앱에 자동 반영됩니다.');

  buildNameDropdown(ss);
  SpreadsheetApp.getUi().alert('완료! "학생 관리" 탭이 생성되었습니다.\n\n이제 이 탭에서 이름을 추가/삭제하면\n앱과 Dashboard 드롭다운에 자동 반영됩니다.');
}

/* ── Name dropdown ──────────────────────────────────────────── */

function buildNameDropdown(ss) {
  ss = ss || SpreadsheetApp.openById(SPREADSHEET_ID);
  var dash = ss.getSheetByName(DASHBOARD_NAME);
  if (!dash) return;

  var names = getStudentList(ss);
  if (names.length === 0) return;

  dash.getRange(SEARCH_CELL).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(names, true)
      .setAllowInvalid(true)
      .build()
  );
}

/* ── Refresh ────────────────────────────────────────────────── */

function refreshDashboard(ss) {
  ss = ss || SpreadsheetApp.openById(SPREADSHEET_ID);
  var dash = ss.getSheetByName(DASHBOARD_NAME);
  if (!dash) { setupDashboard(); return; }

  var searchName = String(dash.getRange(SEARCH_CELL).getValue() || '').trim();

  // Clear rows 3+
  var lastRow = dash.getLastRow();
  if (lastRow >= 3) {
    var cr = dash.getRange(3, 1, lastRow - 2, dash.getMaxColumns());
    cr.breakApart();
    cr.clear();
  }

  if (!searchName) {
    dash.getRange('A4:O4').merge()
      .setValue('👆 B2 드롭다운에서 학생 이름을 선택해주세요')
      .setFontColor('#9CA3AF').setFontStyle('italic').setFontSize(13)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    dash.setRowHeight(4, 60);
    return;
  }

  var row = 4;
  var totalRecords = 0;

  for (var s = 0; s < SECTIONS.length; s++) {
    var sec = SECTIONS[s];
    var srcSheet = ss.getSheetByName(sec.tab);

    // Section title bar
    dash.getRange(row, 1, 1, 15).merge()
      .setValue(sec.label + '  ·  ' + sec.tab)
      .setBackground(sec.titleBg).setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(14).setHorizontalAlignment('left').setVerticalAlignment('middle');
    dash.setRowHeight(row, 34);
    row++;

    if (!srcSheet) {
      writeNoticeRow(dash, row, '⚠️ "' + sec.tab + '" 탭을 찾을 수 없습니다.', '#DC2626');
      row += 2; continue;
    }

    var data = srcSheet.getDataRange().getValues();
    if (data.length < 2) {
      writeNoticeRow(dash, row, '기록 없음', '#9CA3AF');
      row += 2; continue;
    }

    var headers = data[0];
    var nameIdx = findColumnIndex(headers, ['이름', 'student_name', 'name']);
    if (nameIdx === -1) {
      writeNoticeRow(dash, row, '⚠️ "' + sec.tab + '" 탭에서 이름 컬럼을 찾지 못했습니다.', '#DC2626');
      row += 2; continue;
    }

    var matches = [];
    for (var r = 1; r < data.length; r++) {
      var nm = String(data[r][nameIdx] || '').trim();
      if (nm && (nm === searchName || nm.indexOf(searchName) !== -1 || searchName.indexOf(nm) !== -1)) {
        matches.push(data[r]);
      }
    }
    matches.reverse(); // 최신순

    // Count badge
    dash.getRange(row, 1, 1, 15).merge()
      .setValue('  총 ' + matches.length + '건')
      .setBackground('#F9FAFB').setFontColor('#374151').setFontWeight('bold')
      .setFontSize(11).setHorizontalAlignment('left').setVerticalAlignment('middle');
    dash.setRowHeight(row, 22);
    row++;

    if (matches.length === 0) {
      writeNoticeRow(dash, row, '「' + searchName + '」 학생의 기록이 없습니다.', '#9CA3AF');
      row += 2; continue;
    }

    totalRecords += matches.length;

    // Header row
    dash.getRange(row, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground(sec.headerBg).setFontColor('#1A1A2E')
      .setFontSize(11).setVerticalAlignment('middle').setHorizontalAlignment('center')
      .setBorder(true, true, true, true, true, true, '#9CA3AF', SpreadsheetApp.BorderStyle.SOLID);
    dash.setRowHeight(row, 32);
    row++;

    // Data rows
    dash.getRange(row, 1, matches.length, headers.length).setValues(matches)
      .setFontSize(10).setVerticalAlignment('top').setWrap(true)
      .setBorder(true, true, true, true, true, true, '#E5E7EB', SpreadsheetApp.BorderStyle.SOLID);

    for (var i = 0; i < matches.length; i++) {
      dash.getRange(row + i, 1, 1, headers.length)
        .setBackground(i % 2 === 1 ? sec.bandBg : '#FFFFFF');
      dash.setRowHeight(row + i, 32);
    }

    row += matches.length + 1; // +1 spacer
  }

  // Summary bar
  var summary = dash.getRange(row, 1, 1, 15).merge();
  if (totalRecords > 0) {
    summary.setValue('🎉 「' + searchName + '」 학생의 전체 기록: ' + totalRecords + '건')
      .setBackground('#1A1A2E').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(12).setHorizontalAlignment('center').setVerticalAlignment('middle');
  } else {
    summary.setValue('「' + searchName + '」 학생의 기록이 전혀 없습니다. 이름을 다시 확인해 주세요.')
      .setBackground('#FEF2F2').setFontColor('#991B1B').setFontWeight('bold')
      .setFontSize(12).setHorizontalAlignment('center').setVerticalAlignment('middle');
  }
  dash.setRowHeight(row, 36);
}

/* ── Helpers ────────────────────────────────────────────────── */

function findColumnIndex(headers, candidates) {
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || '').toLowerCase().trim();
    for (var c = 0; c < candidates.length; c++) {
      if (h === candidates[c].toLowerCase()) return i;
    }
  }
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || '').toLowerCase().trim();
    for (var c = 0; c < candidates.length; c++) {
      if (h.indexOf(candidates[c].toLowerCase()) !== -1) return i;
    }
  }
  return -1;
}

function writeNoticeRow(dash, row, msg, color) {
  dash.getRange(row, 1, 1, 15).merge()
    .setValue(msg).setFontColor(color || '#6B7280').setFontStyle('italic')
    .setFontSize(11).setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setBackground('#FFFFFF');
  dash.setRowHeight(row, 28);
}
