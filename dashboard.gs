/**
 * Fox Learning Center - Dashboard for Google Sheet
 *
 * 사용법:
 * 1. Google Sheet 열기 → Extensions → Apps Script
 * 2. 기존 Dashboard 관련 코드 삭제하고 이 코드로 교체
 * 3. 저장 후, setupDashboard() 한 번 실행 (권한 승인)
 * 4. Dashboard 탭의 B1 셀에 학생 이름 입력하면 자동 갱신
 *
 * ⚠️ TAB_NAMES를 본인의 시트 탭 이름에 맞게 수정하세요.
 */

const DASHBOARD_NAME = 'Dashboard';

// ⚠️ 여기를 실제 시트 탭 이름과 일치시키세요
const TAB_NAMES = {
  morning: '입실체크',   // 입실 데이터가 저장된 탭 이름
  evening: '퇴실체크',   // 퇴실 데이터가 저장된 탭 이름
  mission: '특별미션'    // 특별미션 데이터가 저장된 탭 이름
};

const SECTIONS = [
  { key: 'morning', label: '🌅 입실 체크 기록', titleBg: '#FF6B35', titleFg: '#FFFFFF', headerBg: '#FFE8DF' },
  { key: 'evening', label: '🌙 퇴실 체크 기록', titleBg: '#5BAD8A', titleFg: '#FFFFFF', headerBg: '#E0F7F5' },
  { key: 'mission', label: '✨ 특별 미션 기록', titleBg: '#A78BFA', titleFg: '#FFFFFF', headerBg: '#EDE9FE' }
];

const NAME_HEADER_CANDIDATES = ['student_name', '학생이름', '이름', 'name', '학생 이름'];
const DATE_HEADER_CANDIDATES = ['date', '날짜'];
const TIME_HEADER_CANDIDATES = ['time', '시간'];

/* ──────────────────────────────────────────────
 * Setup — 한 번만 실행
 * ────────────────────────────────────────────── */
function setupDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dash = ss.getSheetByName(DASHBOARD_NAME);
  if (!dash) {
    dash = ss.insertSheet(DASHBOARD_NAME);
  } else {
    dash.clear();
    dash.clearConditionalFormatRules();
  }

  // Header row
  dash.getRange('A1').setValue('🔍 학생 이름 검색').setFontWeight('bold').setFontSize(16).setFontColor('#1A1A2E');
  dash.getRange('B1')
    .setBackground('#FFF8F0')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBorder(true, true, true, true, false, false, '#FF6B35', SpreadsheetApp.BorderStyle.SOLID_THICK)
    .setHorizontalAlignment('center');
  dash.getRange('A2:Z2').merge();
  dash.getRange('A2').setValue('이름을 입력하면 자동으로 모든 체크 기록이 표시됩니다.')
    .setFontColor('#6B7280')
    .setFontStyle('italic')
    .setFontSize(11);

  dash.setColumnWidth(1, 220);
  dash.setColumnWidth(2, 280);
  for (let i = 3; i <= 26; i++) dash.setColumnWidth(i, 160);
  dash.setRowHeight(1, 40);

  dash.setFrozenRows(2);

  refreshDashboard();
  SpreadsheetApp.getUi().alert('Dashboard 준비 완료!\nB1 셀에 학생 이름을 입력하세요.');
}

/* ──────────────────────────────────────────────
 * Auto refresh on B1 edit
 * ────────────────────────────────────────────── */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== DASHBOARD_NAME) return;
    if (e.range.getA1Notation() !== 'B1') return;
    refreshDashboard();
  } catch (err) {
    Logger.log('onEdit error: ' + err);
  }
}

/* ──────────────────────────────────────────────
 * Refresh — 검색 결과 채우기
 * ────────────────────────────────────────────── */
function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName(DASHBOARD_NAME);
  if (!dash) return;

  const searchName = String(dash.getRange('B1').getValue() || '').trim();

  // Clear below row 2
  const lastRow = dash.getLastRow();
  const maxCols = dash.getMaxColumns();
  if (lastRow > 2) {
    dash.getRange(3, 1, lastRow - 2, maxCols).clear({ contentsOnly: false, formatOnly: false });
    dash.getRange(3, 1, lastRow - 2, maxCols).breakApart();
  }

  if (!searchName) {
    const cell = dash.getRange('A4');
    cell.setValue('👈 B1 셀에 학생 이름을 입력하세요').setFontSize(13).setFontColor('#9CA3AF').setFontStyle('italic');
    return;
  }

  let row = 4;

  for (const sec of SECTIONS) {
    const tabName = TAB_NAMES[sec.key];
    const srcSheet = ss.getSheetByName(tabName);

    // Section title bar
    dash.getRange(row, 1, 1, 12).merge()
      .setValue(sec.label)
      .setBackground(sec.titleBg)
      .setFontColor(sec.titleFg)
      .setFontWeight('bold')
      .setFontSize(14)
      .setHorizontalAlignment('center');
    dash.setRowHeight(row, 32);
    row++;

    if (!srcSheet) {
      dash.getRange(row, 1).setValue(`⚠️ "${tabName}" 탭을 찾을 수 없습니다. TAB_NAMES 설정을 확인하세요.`)
        .setFontColor('#DC2626').setFontStyle('italic');
      row += 2;
      continue;
    }

    const data = srcSheet.getDataRange().getValues();
    if (data.length < 2) {
      dash.getRange(row, 1).setValue('기록 없음').setFontColor('#9CA3AF').setFontStyle('italic');
      row += 2;
      continue;
    }

    const headers = data[0];
    const nameIdx = findHeaderIndex(headers, NAME_HEADER_CANDIDATES);
    if (nameIdx === -1) {
      dash.getRange(row, 1).setValue(`⚠️ "${tabName}"에서 이름 열을 찾지 못했습니다.`).setFontColor('#DC2626');
      row += 2;
      continue;
    }
    const dateIdx = findHeaderIndex(headers, DATE_HEADER_CANDIDATES);

    // Filter rows by name match (정확히 일치 OR 부분 포함)
    const matches = data.slice(1).filter(r => {
      const v = String(r[nameIdx] || '').trim();
      return v === searchName || v.includes(searchName);
    });

    // Sort newest first by date column if present
    if (dateIdx !== -1) {
      matches.sort((a, b) => {
        const da = String(a[dateIdx] || '');
        const db = String(b[dateIdx] || '');
        return db.localeCompare(da);
      });
    }

    if (matches.length === 0) {
      dash.getRange(row, 1).setValue(`「${searchName}」 학생의 ${sec.label.replace(/[🌅🌙✨]/g,'').trim()} 기록이 없습니다.`)
        .setFontColor('#9CA3AF')
        .setFontStyle('italic');
      row += 2;
      continue;
    }

    // Match count badge
    dash.getRange(row, 1, 1, 12).merge()
      .setValue(`총 ${matches.length}건`)
      .setBackground('#F9FAFB')
      .setFontColor('#6B7280')
      .setFontWeight('bold')
      .setFontSize(11)
      .setHorizontalAlignment('right');
    row++;

    // Header row
    const headerRange = dash.getRange(row, 1, 1, headers.length);
    headerRange.setValues([headers])
      .setFontWeight('bold')
      .setBackground(sec.headerBg)
      .setFontColor('#1A1A2E')
      .setFontSize(11)
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false, '#9CA3AF', SpreadsheetApp.BorderStyle.SOLID);
    dash.setRowHeight(row, 28);
    row++;

    // Data rows
    const dataRange = dash.getRange(row, 1, matches.length, headers.length);
    dataRange.setValues(matches)
      .setFontSize(10)
      .setVerticalAlignment('top')
      .setWrap(true)
      .setBorder(true, true, true, true, true, true, '#E5E7EB', SpreadsheetApp.BorderStyle.SOLID);

    // Alternate row banding
    for (let i = 0; i < matches.length; i++) {
      if (i % 2 === 1) {
        dash.getRange(row + i, 1, 1, headers.length).setBackground('#FAFAFA');
      }
    }

    row += matches.length + 2; // blank spacer row
  }
}

/* ──────────────────────────────────────────────
 * Helpers
 * ────────────────────────────────────────────── */
function findHeaderIndex(headers, candidates) {
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || '').toLowerCase().trim();
    for (const c of candidates) {
      if (h === c.toLowerCase()) return i;
    }
  }
  // fuzzy fallback
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || '').toLowerCase().trim();
    for (const c of candidates) {
      if (h.indexOf(c.toLowerCase()) !== -1) return i;
    }
  }
  return -1;
}

/* ──────────────────────────────────────────────
 * Manual menu
 * ────────────────────────────────────────────── */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🦊 FoxLC')
    .addItem('Dashboard 새로고침', 'refreshDashboard')
    .addItem('Dashboard 초기 셋업', 'setupDashboard')
    .addToUi();
}
