/**
 * Code.gs — 메인 변환 엔진 (자동 변환)
 *
 * 네이버주문 탭에 붙여넣기만 하면 → 한진송장 탭이 자동으로 채워집니다.
 * 버튼을 누를 필요가 없습니다. (onEdit 자동 트리거)
 * 준비된 결과는 메뉴에서 한진양식 CSV로 내려받아 원클릭 대량접수에 업로드합니다.
 * 설정(컬럼 매핑)은 Config.gs 에서 조정합니다.
 */

var SHEET_CONFIG = '_설정';   // 자동변환 ON/OFF 체크박스가 있는 탭

// ── 시트 열릴 때 커스텀 메뉴 생성 ──────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('택배자동화')
    .addItem('① 초기 시트 세팅', 'setupSheets')
    .addItem('한진양식 CSV 다운로드', 'downloadHanjinCsv')
    .addSeparator()
    .addItem('지금 즉시 다시 변환', 'convertNaverToHanjin')
    .addItem('한진송장 탭 비우기', 'clearHanjinSheet')
    .addToUi();
}

// ── ① 초기 시트 세팅: 탭·헤더·자동변환 스위치 생성 ────────────────────
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var naver = ss.getSheetByName(SHEET_NAVER) || ss.insertSheet(SHEET_NAVER);
  if (naver.getLastRow() === 0) {
    var naverHeaders = Object.keys(NAVER_COLS).map(function (k) { return NAVER_COLS[k]; });
    naver.getRange(1, 1, 1, naverHeaders.length).setValues([naverHeaders]).setFontWeight('bold');
    naver.setFrozenRows(1);
  }

  var hanjin = ss.getSheetByName(SHEET_HANJIN) || ss.insertSheet(SHEET_HANJIN);
  hanjin.clear();
  hanjin.getRange(1, 1, 1, HANJIN_COLS.length).setValues([HANJIN_COLS]).setFontWeight('bold');
  hanjin.setFrozenRows(1);

  // 자동변환 스위치 탭 (체크박스). 체크돼 있으면 붙여넣는 즉시 자동 변환.
  var cfg = ss.getSheetByName(SHEET_CONFIG) || ss.insertSheet(SHEET_CONFIG);
  cfg.clear();
  cfg.getRange('A1').setValue('자동변환(붙여넣으면 바로)').setFontWeight('bold');
  cfg.getRange('B1').insertCheckboxes().check();
  cfg.getRange('A3').setValue('※ 이 탭은 건드리지 마세요. B1 체크를 끄면 자동변환이 멈춥니다.')
    .setFontColor('#888888');
  cfg.setColumnWidth(1, 220);

  SpreadsheetApp.getUi().alert(
    '세팅 완료 ✅\n\n' +
    '이제 "' + SHEET_NAVER + '" 탭에 네이버 주문 엑셀을 붙여넣기만 하면\n' +
    '"' + SHEET_HANJIN + '" 탭이 자동으로 채워집니다. (버튼 불필요)\n\n' +
    '출력할 때: 메뉴 "택배자동화 > 한진양식 CSV 다운로드".'
  );
}

// ── 자동 변환 트리거: 네이버주문 탭이 바뀌면 스스로 실행 ───────────────
// 단순 onEdit 트리거 — 사용자의 붙여넣기/입력에만 반응하고,
// 스크립트가 한진송장에 쓰는 것에는 반응하지 않아 무한루프가 없습니다.
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var edited = e.range.getSheet();
    if (edited.getName() !== SHEET_NAVER) return;       // 주문 탭 변경만 반응
    if (!isAutoOn(e.source)) return;                    // 스위치 꺼져 있으면 패스
    runConversion(false);                               // 조용히 변환(알림 없음)
  } catch (err) {
    // 자동 트리거에서는 알림을 띄울 수 없으므로 로그만 남김
    console.error('자동변환 오류: ' + err);
  }
}

function isAutoOn(ss) {
  var cfg = ss.getSheetByName(SHEET_CONFIG);
  if (!cfg) return true;                 // 스위치 탭이 없으면 기본 ON
  return cfg.getRange('B1').getValue() === true;
}

// ── 메뉴: 지금 즉시 다시 변환(알림 표시) ──────────────────────────────
function convertNaverToHanjin() {
  var n = runConversion(true);
  if (n < 0) {
    SpreadsheetApp.getUi().alert('"' + SHEET_NAVER + '" 탭에 주문 데이터가 없습니다.');
    return;
  }
  SpreadsheetApp.getUi().alert('변환 완료 ✅  송장 ' + n + '건 생성됨.\n"' +
    SHEET_HANJIN + '" 탭을 확인하고 CSV로 내려받으세요.');
}

// ── 변환 코어(자동/수동 공용) : 성공 시 송장 건수, 데이터 없으면 -1 ────
function runConversion(showResize) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var naver = ss.getSheetByName(SHEET_NAVER);
  if (!naver || naver.getLastRow() < 2) return -1;

  var rows = readAsObjects(naver);          // 헤더 기준으로 각 행을 {키:값}으로
  if (COMBINE_BY_RECEIVER) rows = combineByReceiver(rows);

  var out = rows.map(buildHanjinRow).filter(function (r) { return r !== null; });

  var hanjin = ss.getSheetByName(SHEET_HANJIN) || ss.insertSheet(SHEET_HANJIN);
  // 기존 데이터 영역만 지우고 다시 씀(전체 clear 대신 → 서식/틀고정 유지)
  if (hanjin.getLastRow() > 1) {
    hanjin.getRange(2, 1, hanjin.getLastRow() - 1, HANJIN_COLS.length).clearContent();
  }
  hanjin.getRange(1, 1, 1, HANJIN_COLS.length).setValues([HANJIN_COLS]).setFontWeight('bold');
  if (out.length > 0) {
    hanjin.getRange(2, 1, out.length, HANJIN_COLS.length).setValues(out);
  }
  hanjin.setFrozenRows(1);
  if (showResize) hanjin.autoResizeColumns(1, HANJIN_COLS.length);
  return out.length;
}

// ── ③ 한진송장 탭을 CSV로 다운로드(링크 제공) ─────────────────────────
function downloadHanjinCsv() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hanjin = ss.getSheetByName(SHEET_HANJIN);
  if (!hanjin || hanjin.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('먼저 변환을 실행하세요. (한진송장 데이터 없음)');
    return;
  }
  var data = hanjin.getDataRange().getValues();
  var csv = data.map(function (row) {
    return row.map(csvEscape).join(',');
  }).join('\r\n');

  // 한글 깨짐 방지를 위해 UTF-8 BOM 추가.
  // Apps Script 바이트는 부호형(-128~127)이라 BOM도 부호값으로 붙인다.
  var bytes = [-17, -69, -65].concat(Utilities.newBlob(csv).getBytes()); // EF BB BF
  var blob = Utilities.newBlob(bytes, 'text/csv', '한진송장_' + dateStamp() + '.csv');
  var file = DriveApp.createFile(blob);

  var html = '<p>CSV 파일이 구글 드라이브에 생성되었습니다.</p>' +
    '<p><a href="' + file.getUrl() + '" target="_blank">👉 파일 열기 / 다운로드</a></p>' +
    '<p style="color:#888;font-size:12px">다운로드 후 한진 원클릭 택배 &gt; 대량접수에 업로드하세요.</p>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(380).setHeight(160),
    '한진양식 CSV 다운로드'
  );
}

function clearHanjinSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hanjin = ss.getSheetByName(SHEET_HANJIN);
  if (!hanjin) return;
  hanjin.clear();
  hanjin.getRange(1, 1, 1, HANJIN_COLS.length).setValues([HANJIN_COLS]).setFontWeight('bold');
  hanjin.setFrozenRows(1);
}

/* ══════════════════════════ 내부 로직 ══════════════════════════ */

// 한진 한 행 만들기: MAPPING 규칙대로 채움
function buildHanjinRow(nrow) {
  // 필수값(수취인/주소) 없으면 스킵
  var receiver = pick(nrow, NAVER_COLS.receiver);
  var address  = pick(nrow, NAVER_COLS.address);
  if (!receiver && !address) return null;

  return HANJIN_COLS.map(function (col) {
    var rule = MAPPING[col];
    if (!rule) return '';
    if (rule.const !== undefined) return rule.const;
    if (rule.from) return pick(nrow, NAVER_COLS[rule.from]) || '';
    if (rule.fn && COMPUTERS[rule.fn]) return COMPUTERS[rule.fn](nrow);
    return '';
  });
}

// 계산 함수들 (MAPPING의 fn: 에서 참조)
var COMPUTERS = {
  phone: function (r) {
    return normalizePhone(pick(r, NAVER_COLS.receiverPhone));
  },
  phone2: function (r) {
    return normalizePhone(pick(r, NAVER_COLS.receiverPhone2));
  },
  // 상품명 + 옵션 + (수량>1이면 x수량)  → 한 필드로
  itemName: function (r) {
    var name = (pick(r, NAVER_COLS.productName) || '').toString().trim();
    var opt  = (pick(r, NAVER_COLS.option) || '').toString().trim();
    var qty  = parseInt(pick(r, NAVER_COLS.quantity), 10) || 1;
    var s = name;
    if (opt) s += ' (' + opt + ')';
    if (qty > 1) s += ' x' + qty;
    return s.substring(0, 100); // 한진 품목명 길이 안전선
  }
};

// 같은 수취인+주소+연락처를 한 송장으로 합치고 품목 합산
function combineByReceiver(rows) {
  var map = {};
  var order = [];
  rows.forEach(function (r) {
    var key = [
      (pick(r, NAVER_COLS.receiver) || '').trim(),
      (pick(r, NAVER_COLS.address) || '').replace(/\s+/g, ''),
      normalizePhone(pick(r, NAVER_COLS.receiverPhone))
    ].join('|');

    if (!map[key]) {
      map[key] = Object.assign({}, r);
      map[key].__items = [];
      map[key].__qty = 0;
      order.push(key);
    }
    var name = (pick(r, NAVER_COLS.productName) || '').trim();
    var opt  = (pick(r, NAVER_COLS.option) || '').trim();
    var qty  = parseInt(pick(r, NAVER_COLS.quantity), 10) || 1;
    var label = name + (opt ? ' (' + opt + ')' : '') + (qty > 1 ? ' x' + qty : '');
    map[key].__items.push(label);
    map[key].__qty += qty;
  });

  return order.map(function (key) {
    var r = map[key];
    // 합쳐진 품목/수량을 네이버 컬럼 자리에 다시 넣어 buildHanjinRow가 쓰게 함
    r[NAVER_COLS.productName] = r.__items.join(' + ');
    r[NAVER_COLS.option] = '';
    r[NAVER_COLS.quantity] = r.__qty;
    return r;
  });
}

/* ── 유틸 ── */

// 시트를 [{헤더:값}] 배열로 읽기
function readAsObjects(sheet) {
  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(function (h) { return (h || '').toString().trim(); });
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row.every(function (c) { return c === '' || c === null; })) continue; // 빈 줄
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = row[j];
    out.push(obj);
  }
  return out;
}

// 헤더 이름으로 값 꺼내기 (정확 일치 우선, 없으면 부분 포함)
function pick(obj, colName) {
  if (!colName) return '';
  if (obj[colName] !== undefined) return obj[colName];
  var keys = Object.keys(obj);
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].indexOf(colName) !== -1 || colName.indexOf(keys[i]) !== -1) {
      if (keys[i].indexOf('__') === 0) continue;
      return obj[keys[i]];
    }
  }
  return '';
}

// 전화번호 정규화: 숫자만 뽑아 010-xxxx-xxxx / 지역번호 형태로
function normalizePhone(v) {
  if (v === undefined || v === null) return '';
  var d = v.toString().replace(/[^0-9]/g, '');
  if (!d) return '';
  // 엑셀에서 앞 0이 사라진 경우 보정 (10자리이고 1로 시작하면 010...)
  if (d.length === 9 && d.charAt(0) === '1') d = '0' + d;    // 만일의 경우
  if (d.length === 10 && d.charAt(0) === '1') d = '0' + d;   // 1012345678 → 01012345678
  if (d.length === 11) return d.replace(/(\d{3})(\d{4})(\d{4})/, '$1-$2-$3');
  if (d.length === 10) {
    if (d.charAt(0) === '0' && d.charAt(1) === '2')
      return d.replace(/(\d{2})(\d{4})(\d{4})/, '$1-$2-$3'); // 02
    return d.replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3');
  }
  return d; // 애매하면 숫자 그대로
}

function csvEscape(v) {
  var s = (v === null || v === undefined) ? '' : v.toString();
  if (/[",\r\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

function dateStamp() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd_HHmm');
}
