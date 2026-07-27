/**
 * 전체코드.gs — 도하커피 택배자동화 (설정 + 엔진 한 파일)
 *
 * 사용법: Apps Script 편집기에서 전체 선택(Ctrl+A) → 삭제 → 이 내용 붙여넣기 → 저장(Ctrl+S)
 *         시트 새로고침 후 메뉴 "택배자동화 > ① 초기 시트 세팅" 실행.
 * 이후 "네이버주문" 탭에 주문을 붙여넣으면 "한진송장" 탭이 자동으로 채워집니다.
 *
 * 컬럼 매핑은 아래 NAVER_COLS / HANJIN_COLS / MAPPING 만 고치면 됩니다.
 * (도하커피 스마트스토어 실제 헤더에 맞춰 설정됨)
 */

/* ═══════════════════════ 설정(매핑) ═══════════════════════ */

// 시트 탭 이름
var SHEET_NAVER  = '네이버주문';
var SHEET_HANJIN = '한진송장';
var SHEET_CONFIG = '_설정';

// 네이버 주문 엑셀의 헤더 이름 (열 순서는 상관없음 — 이름으로 찾음)
var NAVER_COLS = {
  productOrderNo: '상품주문번호',
  orderNo:        '주문번호',
  buyer:          '구매자명',
  receiver:       '수취인명',
  receiverPhone:  '수취인연락처1',
  receiverPhone2: '수취인연락처2',
  zipcode:        '우편번호',
  address:        '통합배송지',       // 도로명+상세 합쳐진 전체 주소
  productName:    '상품명',
  option:         '옵션정보',
  quantity:       '수량',
  deliveryMemo:   '배송메세지'
};

// 한진 원클릭 대량접수 양식 헤더 (순서 = 출력 순서)
var HANJIN_COLS = [
  '받는분성명',
  '받는분전화번호',
  '받는분기타연락처',
  '받는분우편번호',
  '받는분주소',
  '품목명',
  '내품수량',
  '박스수량',
  '운임구분',
  '배송메시지'
];

// 매핑 규칙: from=네이버값 그대로, const=고정값, fn=계산
var MAPPING = {
  '받는분성명':      { from: 'receiver' },
  '받는분전화번호':  { fn: 'phone' },
  '받는분기타연락처':{ fn: 'phone2' },
  '받는분우편번호':  { from: 'zipcode' },
  '받는분주소':      { from: 'address' },
  '품목명':          { fn: 'itemName' },
  '내품수량':        { from: 'quantity' },
  '박스수량':        { const: '1' },
  '운임구분':        { const: '신용' },
  '배송메시지':      { from: 'deliveryMemo' }
};

// 같은 수취인+주소+연락처를 한 송장으로 합칠지 여부
var COMBINE_BY_RECEIVER = true;

/* ═══════════════════════ 엔진 (아래는 안 건드려도 됨) ═══════════════════════ */

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

// 자동 변환: 네이버주문 탭이 바뀌면 스스로 실행 (붙여넣기/입력에만 반응)
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    if (e.range.getSheet().getName() !== SHEET_NAVER) return;
    if (!isAutoOn(e.source)) return;
    runConversion(false);
  } catch (err) {
    console.error('자동변환 오류: ' + err);
  }
}

function isAutoOn(ss) {
  var cfg = ss.getSheetByName(SHEET_CONFIG);
  if (!cfg) return true;
  return cfg.getRange('B1').getValue() === true;
}

function convertNaverToHanjin() {
  var n = runConversion(true);
  if (n < 0) {
    SpreadsheetApp.getUi().alert('"' + SHEET_NAVER + '" 탭에 주문 데이터가 없습니다.');
    return;
  }
  SpreadsheetApp.getUi().alert('변환 완료 ✅  송장 ' + n + '건 생성됨.\n"' +
    SHEET_HANJIN + '" 탭을 확인하고 CSV로 내려받으세요.');
}

function runConversion(showResize) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var naver = ss.getSheetByName(SHEET_NAVER);
  if (!naver || naver.getLastRow() < 2) return -1;

  var rows = readAsObjects(naver);
  if (COMBINE_BY_RECEIVER) rows = combineByReceiver(rows);

  var out = rows.map(buildHanjinRow).filter(function (r) { return r !== null; });

  var hanjin = ss.getSheetByName(SHEET_HANJIN) || ss.insertSheet(SHEET_HANJIN);
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

function downloadHanjinCsv() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hanjin = ss.getSheetByName(SHEET_HANJIN);
  if (!hanjin || hanjin.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('먼저 변환을 실행하세요. (한진송장 데이터 없음)');
    return;
  }
  var data = hanjin.getDataRange().getValues();
  var csv = data.map(function (row) { return row.map(csvEscape).join(','); }).join('\r\n');
  var bytes = [-17, -69, -65].concat(Utilities.newBlob(csv).getBytes()); // UTF-8 BOM
  var blob = Utilities.newBlob(bytes, 'text/csv', '한진송장_' + dateStamp() + '.csv');
  var file = DriveApp.createFile(blob);

  var html = '<p>CSV 파일이 구글 드라이브에 생성되었습니다.</p>' +
    '<p><a href="' + file.getUrl() + '" target="_blank">👉 파일 열기 / 다운로드</a></p>' +
    '<p style="color:#888;font-size:12px">다운로드 후 한진 원클릭 택배 &gt; 대량접수에 업로드하세요.</p>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(380).setHeight(160), '한진양식 CSV 다운로드');
}

function clearHanjinSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hanjin = ss.getSheetByName(SHEET_HANJIN);
  if (!hanjin) return;
  hanjin.clear();
  hanjin.getRange(1, 1, 1, HANJIN_COLS.length).setValues([HANJIN_COLS]).setFontWeight('bold');
  hanjin.setFrozenRows(1);
}

function buildHanjinRow(nrow) {
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

var COMPUTERS = {
  phone:  function (r) { return normalizePhone(pick(r, NAVER_COLS.receiverPhone)); },
  phone2: function (r) { return normalizePhone(pick(r, NAVER_COLS.receiverPhone2)); },
  itemName: function (r) {
    var name = (pick(r, NAVER_COLS.productName) || '').toString().trim();
    var opt  = (pick(r, NAVER_COLS.option) || '').toString().trim();
    var qty  = parseInt(pick(r, NAVER_COLS.quantity), 10) || 1;
    var s = name;
    if (opt) s += ' (' + opt + ')';
    if (qty > 1) s += ' x' + qty;
    return s.substring(0, 100);
  }
};

function combineByReceiver(rows) {
  var map = {}, order = [];
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
    map[key].__items.push(name + (opt ? ' (' + opt + ')' : '') + (qty > 1 ? ' x' + qty : ''));
    map[key].__qty += qty;
  });
  return order.map(function (key) {
    var r = map[key];
    r[NAVER_COLS.productName] = r.__items.join(' + ');
    r[NAVER_COLS.option] = '';
    r[NAVER_COLS.quantity] = r.__qty;
    return r;
  });
}

function readAsObjects(sheet) {
  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(function (h) { return (h || '').toString().trim(); });
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row.every(function (c) { return c === '' || c === null; })) continue;
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = row[j];
    out.push(obj);
  }
  return out;
}

function pick(obj, colName) {
  if (!colName) return '';
  if (obj[colName] !== undefined) return obj[colName];
  var keys = Object.keys(obj);
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].indexOf('__') === 0) continue;
    if (keys[i].indexOf(colName) !== -1 || colName.indexOf(keys[i]) !== -1) return obj[keys[i]];
  }
  return '';
}

function normalizePhone(v) {
  if (v === undefined || v === null) return '';
  var d = v.toString().replace(/[^0-9]/g, '');
  if (!d) return '';
  if (d.length === 10 && d.charAt(0) === '1') d = '0' + d;
  if (d.length === 11) return d.replace(/(\d{3})(\d{4})(\d{4})/, '$1-$2-$3');
  if (d.length === 10) {
    if (d.charAt(0) === '0' && d.charAt(1) === '2') return d.replace(/(\d{2})(\d{4})(\d{4})/, '$1-$2-$3');
    return d.replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3');
  }
  return d;
}

function csvEscape(v) {
  var s = (v === null || v === undefined) ? '' : v.toString();
  if (/[",\r\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

function dateStamp() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd_HHmm');
}
