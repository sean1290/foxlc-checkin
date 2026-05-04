var SPREADSHEET_ID = '17JJYwp1rHUSf2j7DXI2wKC84hXJ2ghOPDZdmQALeyo0';

/* ═══════════════════════════════════════════════════════════════
 * WEB APP ENDPOINTS  (기존 코드 - 변경 없음)
 * ═══════════════════════════════════════════════════════════════ */

function doGet(e) {
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
    '감정',
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
    d.axis1_energy          || '',   // E 에너지 레벨
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
    '날짜','이름','마음',
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
    d.student_name            || '',
    d.axis1_emotion_change    || '',
    d.axis2_goal_score        || '',
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
 * DASHBOARD  (신규)
 * - Dashboard 탭의 B2 셀에 학생 이름을 입력하면 자동으로 갱신
 * - 입실/퇴실/특별미션 기록을 별도 섹션으로 표시
 * - 첫 사용 시: setupDashboard() 한 번 실행
 * ═══════════════════════════════════════════════════════════════ */

var DASHBOARD_NAME = 'Dashboard';
var SEARCH_CELL = 'B2';

var SECTIONS = [
  {
    key: 'morning',
    tab: '입실체크',
    label: '🌅 입실 체크',
    titleBg: '#FF6B35',
    headerBg: '#FFE8DF',
    bandBg: '#FFF8F4'
  },
  {
    key: 'evening',
    tab: '퇴실체크',
    label: '🌙 퇴실 체크',
    titleBg: '#5BAD8A',
    headerBg: '#E0F7F5',
    bandBg: '#F4FBFA'
  },
  {
    key: 'mission',
    tab: '특별미션',
    label: '✨ 특별 미션',
    titleBg: '#A78BFA',
    headerBg: '#EDE9FE',
    bandBg: '#F8F5FF'
  }
];

/**
 * Dashboard 탭을 처음 만들거나 초기화.
 * Apps Script 편집기에서 한 번 실행 후 권한 승인.
 */
function setupDashboard() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dash = ss.getSheetByName(DASHBOARD_NAME);
  if (!dash) {
    dash = ss.insertSheet(DASHBOARD_NAME);
  } else {
    dash.clear();
    dash.clearConditionalFormatRules();
    var dataValidations = dash.getRange(1, 1, dash.getMaxRows(), dash.getMaxColumns());
    dataValidations.clearDataValidations();
  }

  // ── Title bar (row 1)
  dash.getRange('A1:O1').merge()
    .setValue('🦊 Fox Learning Center · 학생 기록 대시보드')
    .setBackground('#1A1A2E')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  dash.setRowHeight(1, 44);

  // ── Search row (row 2)
  dash.getRange('A2').setValue('🔍 학생 이름 검색:')
    .setBackground('#FFF8F0')
    .setFontWeight('bold')
    .setFontSize(13)
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle');

  dash.getRange(SEARCH_CELL)
    .setBackground('#FFFFFF')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#FF6B35', SpreadsheetApp.BorderStyle.SOLID_THICK);

  dash.getRange('C2:O2').merge()
    .setValue('  ⬅ 드롭다운에서 이름을 선택하면 자동으로 모든 기록이 표시됩니다')
    .setBackground('#FFF8F0')
    .setFontColor('#6B7280')
    .setFontStyle('italic')
    .setFontSize(11)
    .setVerticalAlignment('middle');

  dash.setColumnWidth(1, 200);
  dash.setColumnWidth(2, 240);
  for (var i = 3; i <= 16; i++) dash.setColumnWidth(i, 170);
  dash.setRowHeight(2, 38);
  dash.setFrozenRows(2);

  // ── Hide gridlines for cleaner look
  dash.setHiddenGridlines(true);

  buildNameDropdown();
  refreshDashboard();
  try {
    SpreadsheetApp.getUi().alert('Dashboard 준비 완료!\n\nB2 셀의 드롭다운에서 학생 이름을 선택하세요.');
  } catch(e) {}
}

/**
 * Simple trigger — Dashboard B2 편집 시 자동 갱신
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== DASHBOARD_NAME) return;
    if (e.range.getA1Notation() !== SEARCH_CELL) return;
    refreshDashboard();
  } catch (err) {
    Logger.log('onEdit error: ' + err);
  }
}

/**
 * 메뉴 등록
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('🦊 FoxLC')
      .addItem('Dashboard 새로고침', 'refreshDashboard')
      .addItem('학생 이름 드롭다운 갱신', 'buildNameDropdown')
      .addItem('Dashboard 초기 셋업', 'setupDashboard')
      .addToUi();
  } catch(e) {}
}

/**
 * 입실체크·퇴실체크·특별미션 탭에서 고유 학생 이름을 수집해
 * Dashboard B2 셀에 드롭다운(데이터 유효성 검사)으로 설정.
 * 새 학생이 추가될 때마다 메뉴에서 다시 실행하면 됨.
 */
function buildNameDropdown() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dash = ss.getSheetByName(DASHBOARD_NAME);
  if (!dash) return;

  var nameSet = {};
  var tabs = ['입실체크', '퇴실체크', '특별미션'];

  for (var t = 0; t < tabs.length; t++) {
    var src = ss.getSheetByName(tabs[t]);
    if (!src) continue;
    var data = src.getDataRange().getValues();
    if (data.length < 2) continue;
    var nameIdx = findColumnIndex(data[0], ['이름', 'student_name', 'name']);
    if (nameIdx === -1) continue;
    for (var r = 1; r < data.length; r++) {
      var nm = String(data[r][nameIdx] || '').trim();
      if (nm) nameSet[nm] = true;
    }
  }

  var names = Object.keys(nameSet).sort();
  if (names.length === 0) return;

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(names, true)
    .setAllowInvalid(true)   // 직접 입력도 허용
    .build();

  dash.getRange(SEARCH_CELL).setDataValidation(rule);
}

/**
 * 검색 결과를 그려 넣음
 */
function refreshDashboard() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dash = ss.getSheetByName(DASHBOARD_NAME);
  if (!dash) {
    setupDashboard();
    return;
  }

  var searchName = String(dash.getRange(SEARCH_CELL).getValue() || '').trim();

  // 행 3 이하 모두 청소
  var lastRow = dash.getLastRow();
  var maxCol  = dash.getMaxColumns();
  if (lastRow >= 3) {
    var clearRange = dash.getRange(3, 1, lastRow - 2, maxCol);
    clearRange.breakApart();
    clearRange.clear();
  }

  if (!searchName) {
    dash.getRange('A4:O4').merge()
      .setValue('👆 위의 ' + SEARCH_CELL + ' 셀에 학생 이름을 입력해주세요')
      .setFontColor('#9CA3AF')
      .setFontStyle('italic')
      .setFontSize(13)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    dash.setRowHeight(4, 60);
    return;
  }

  var row = 4;
  var totalRecords = 0;

  for (var s = 0; s < SECTIONS.length; s++) {
    var sec = SECTIONS[s];
    var srcSheet = ss.getSheetByName(sec.tab);

    // ── 섹션 타이틀 바
    var titleRow = dash.getRange(row, 1, 1, 15);
    titleRow.merge()
      .setValue(sec.label + '  ·  ' + sec.tab)
      .setBackground(sec.titleBg)
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(14)
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
    dash.setRowHeight(row, 34);
    row++;

    if (!srcSheet) {
      writeNoticeRow(dash, row, '⚠️ "' + sec.tab + '" 탭을 찾을 수 없습니다.', '#DC2626');
      row += 2;
      continue;
    }

    var data = srcSheet.getDataRange().getValues();
    if (data.length < 2) {
      writeNoticeRow(dash, row, '기록 없음', '#9CA3AF');
      row += 2;
      continue;
    }

    var headers = data[0];
    var nameIdx = findColumnIndex(headers, ['이름', 'student_name', 'name']);
    if (nameIdx === -1) {
      writeNoticeRow(dash, row, '⚠️ "' + sec.tab + '" 탭에서 이름 컬럼을 찾지 못했습니다.', '#DC2626');
      row += 2;
      continue;
    }

    // ── 학생별 필터 (정확 일치 + 부분 포함)
    var matches = [];
    for (var r = 1; r < data.length; r++) {
      var nm = String(data[r][nameIdx] || '').trim();
      if (nm && (nm === searchName || nm.indexOf(searchName) !== -1 || searchName.indexOf(nm) !== -1)) {
        matches.push(data[r]);
      }
    }

    // 최신순 (sheet에 append된 순서를 뒤집음)
    matches.reverse();

    // ── 카운트 배지
    var countText = '총 ' + matches.length + '건';
    dash.getRange(row, 1, 1, 15).merge()
      .setValue('  ' + countText)
      .setBackground('#F9FAFB')
      .setFontColor('#374151')
      .setFontWeight('bold')
      .setFontSize(11)
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
    dash.setRowHeight(row, 22);
    row++;

    if (matches.length === 0) {
      writeNoticeRow(dash, row, '「' + searchName + '」 학생의 기록이 없습니다.', '#9CA3AF');
      row += 2;
      continue;
    }

    totalRecords += matches.length;

    // ── 헤더 행
    var hdrRange = dash.getRange(row, 1, 1, headers.length);
    hdrRange.setValues([headers])
      .setFontWeight('bold')
      .setBackground(sec.headerBg)
      .setFontColor('#1A1A2E')
      .setFontSize(11)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, true, true, '#9CA3AF', SpreadsheetApp.BorderStyle.SOLID);
    dash.setRowHeight(row, 32);
    row++;

    // ── 데이터 행
    var dataRange = dash.getRange(row, 1, matches.length, headers.length);
    dataRange.setValues(matches)
      .setFontSize(10)
      .setVerticalAlignment('top')
      .setWrap(true)
      .setBorder(true, true, true, true, true, true, '#E5E7EB', SpreadsheetApp.BorderStyle.SOLID);

    // 교차 줄무늬
    for (var i = 0; i < matches.length; i++) {
      if (i % 2 === 1) {
        dash.getRange(row + i, 1, 1, headers.length).setBackground(sec.bandBg);
      } else {
        dash.getRange(row + i, 1, 1, headers.length).setBackground('#FFFFFF');
      }
      dash.setRowHeight(row + i, 32);
    }

    row += matches.length;

    // 섹션 사이 빈 줄
    row += 1;
  }

  // 전체 요약
  var summary = dash.getRange(row, 1, 1, 15).merge();
  if (totalRecords > 0) {
    summary.setValue('🎉 「' + searchName + '」 학생의 전체 기록: ' + totalRecords + '건')
      .setBackground('#1A1A2E')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(12)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  } else {
    summary.setValue('「' + searchName + '」 학생의 기록이 전혀 없습니다. 이름을 다시 확인해 주세요.')
      .setBackground('#FEF2F2')
      .setFontColor('#991B1B')
      .setFontWeight('bold')
      .setFontSize(12)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  }
  dash.setRowHeight(row, 36);
}

/* ── Helpers ───────────────────────────────────── */

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
    .setValue(msg)
    .setFontColor(color || '#6B7280')
    .setFontStyle('italic')
    .setFontSize(11)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setBackground('#FFFFFF');
  dash.setRowHeight(row, 28);
}
