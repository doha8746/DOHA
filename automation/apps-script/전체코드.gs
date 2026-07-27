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
var SHEET_NAVER    = '네이버주문';
var SHEET_HANJIN   = '한진송장';
var SHEET_CONFIG   = '_설정';
var SHEET_HISTORY  = '주문내역';   // 그동안 받은 주문 누적 로그
var SHEET_CUSTOMER = '고객관리';   // 고객별 집계(단골 관리)
var SHEET_MSG      = '문자관리';   // 발송 후 안부/재구매 문자 대상·문구

// 탭 열 구성
var HISTORY_COLS  = ['기록일시','주문번호','주문일시','발송일','수취인명','전화번호','우편번호','주소','상품명','수량','배송메모'];
var CUSTOMER_COLS = ['수취인명','전화번호','최근주소','총주문건수','총수량','구매상품','첫주문일','최근주문일'];
var MSG_COLS      = ['수취인명','전화번호','상품','발송일','안부예정일(+4)','안부상태','안부문구','재구매예정일(+10)','재구매상태','재구매문구'];

// ── 문자 발송 설정 (여기만 고치면 됨) ──
var FOLLOWUP_HELLO_DAYS      = 4;   // 발송 후 며칠에 안부 문자
var FOLLOWUP_REPURCHASE_DAYS = 10;  // 발송 후 며칠에 재구매 문자
// {name} 자리에 수취인명이 들어갑니다. 문구는 자유롭게 바꾸세요.
var MSG_HELLO      = '안녕하세요 {name}님, 도하커피입니다 :) 주문하신 원두 맛있게 즐기고 계신가요? 혹시 불편한 점 있으면 편하게 말씀해 주세요. 오늘도 향기로운 하루 보내세요 ☕';
var MSG_REPURCHASE = '{name}님, 도하커피예요 :) 원두가 슬슬 떨어질 때쯤이죠? 주문해 주시면 그날 볶은 신선한 원두로 정성껏 보내드릴게요. 오늘도 좋은 하루 되세요 ☕';

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
  deliveryMemo:   '배송메세지',
  orderDate:      '주문일시',
  shipDate:       '발송일'
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
  '받는분우편번호':  { fn: 'zip' },
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
    .addItem('주문내역에 누적하기', 'appendOrderHistory')
    .addItem('고객관리 갱신', 'updateCustomers')
    .addItem('문자관리 갱신 (오늘 보낼 문자)', 'updateMessagePlan')
    .addSeparator()
    .addItem('지금 즉시 다시 변환', 'convertNaverToHanjin')
    .addItem('한진송장 탭 비우기', 'clearHanjinSheet')
    .addToUi();
}

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var naver = ss.getSheetByName(SHEET_NAVER) || ss.insertSheet(SHEET_NAVER);
  // 우편번호·전화번호가 날짜/숫자로 자동변환되지 않도록 전체를 텍스트 서식으로
  naver.getRange(1, 1, naver.getMaxRows(), naver.getMaxColumns()).setNumberFormat('@');
  if (naver.getLastRow() === 0) {
    var naverHeaders = Object.keys(NAVER_COLS).map(function (k) { return NAVER_COLS[k]; });
    naver.getRange(1, 1, 1, naverHeaders.length).setValues([naverHeaders]).setFontWeight('bold');
    naver.setFrozenRows(1);
  }

  var hanjin = ss.getSheetByName(SHEET_HANJIN) || ss.insertSheet(SHEET_HANJIN);
  hanjin.clear();
  // 한진송장도 텍스트 서식(우편번호 앞 0 유지, 날짜 변환 방지)
  hanjin.getRange(1, 1, hanjin.getMaxRows(), hanjin.getMaxColumns()).setNumberFormat('@');
  hanjin.getRange(1, 1, 1, HANJIN_COLS.length).setValues([HANJIN_COLS]).setFontWeight('bold');
  hanjin.setFrozenRows(1);

  var cfg = ss.getSheetByName(SHEET_CONFIG) || ss.insertSheet(SHEET_CONFIG);
  cfg.clear();
  cfg.getRange('A1').setValue('자동변환(붙여넣으면 바로)').setFontWeight('bold');
  cfg.getRange('B1').insertCheckboxes().check();
  cfg.getRange('A3').setValue('※ 이 탭은 건드리지 마세요. B1 체크를 끄면 자동변환이 멈춥니다.')
    .setFontColor('#888888');
  cfg.setColumnWidth(1, 220);

  // 주문내역(누적 로그) 탭 — 헤더가 현재 버전과 다르면 자동 초기화(칸 밀림 방지)
  ensureHistorySheet(ss);

  // 고객관리 탭 — 없으면 만들고 헤더 생성.
  var cust = ss.getSheetByName(SHEET_CUSTOMER) || ss.insertSheet(SHEET_CUSTOMER);
  if (cust.getLastRow() === 0) {
    cust.getRange(1, 1, 1, CUSTOMER_COLS.length).setValues([CUSTOMER_COLS]).setFontWeight('bold');
    cust.setFrozenRows(1);
  }

  // 문자관리 탭 — 없으면 만들고 헤더 생성.
  var msg = ss.getSheetByName(SHEET_MSG) || ss.insertSheet(SHEET_MSG);
  if (msg.getLastRow() === 0) {
    msg.getRange(1, 1, 1, MSG_COLS.length).setValues([MSG_COLS]).setFontWeight('bold');
    msg.setFrozenRows(1);
  }

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
  hanjin.getRange(1, 1, hanjin.getMaxRows(), hanjin.getMaxColumns()).setNumberFormat('@');
  hanjin.getRange(1, 1, 1, HANJIN_COLS.length).setValues([HANJIN_COLS]).setFontWeight('bold');
  hanjin.setFrozenRows(1);
}

// 주문내역 탭의 헤더가 현재 버전(HISTORY_COLS)과 다르면 초기화(칸 밀림 자동 복구)
function ensureHistorySheet(ss) {
  var hist = ss.getSheetByName(SHEET_HISTORY) || ss.insertSheet(SHEET_HISTORY);
  var ok = false;
  if (hist.getLastRow() >= 1) {
    var hdr = hist.getRange(1, 1, 1, HISTORY_COLS.length).getValues()[0]
      .map(function (x) { return ('' + x).trim(); });
    ok = hdr.join('|') === HISTORY_COLS.join('|');
  }
  if (!ok) {
    hist.clear();
    hist.getRange(1, 1, 1, HISTORY_COLS.length).setValues([HISTORY_COLS]).setFontWeight('bold');
    hist.setFrozenRows(1);
  }
  hist.getRange(1, 1, hist.getMaxRows(), Math.max(hist.getMaxColumns(), HISTORY_COLS.length)).setNumberFormat('@');
  return hist;
}

// ── 주문내역 누적: 네이버주문의 새 주문(상품주문번호 기준)만 주문내역 탭에 추가 ──
function appendOrderHistory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var naver = ss.getSheetByName(SHEET_NAVER);
  if (!naver || naver.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('"' + SHEET_NAVER + '" 탭에 주문 데이터가 없습니다.');
    return;
  }
  var hist = ensureHistorySheet(ss);

  // 이미 기록된 주문번호 집합(중복 방지)
  var seen = {};
  if (hist.getLastRow() > 1) {
    hist.getRange(2, 2, hist.getLastRow() - 1, 1).getValues().forEach(function (r) {
      seen[('' + r[0]).trim()] = true;
    });
  }

  // 같은 주문번호(=같은 주문자·한 주문)의 여러 상품을 한 줄로 합침
  var groups = {}, order = [];
  readAsObjects(naver).forEach(function (r) {
    var ono = ('' + pick(r, NAVER_COLS.orderNo)).trim();
    if (!ono) return;
    if (!groups[ono]) { groups[ono] = { first: r, items: [], qty: 0 }; order.push(ono); }
    var name = ('' + pick(r, NAVER_COLS.productName)).trim();
    var opt  = ('' + pick(r, NAVER_COLS.option)).trim();
    var qty  = parseInt(pick(r, NAVER_COLS.quantity), 10) || 1;
    groups[ono].items.push(name + (opt ? ' (' + opt + ')' : '') + (qty > 1 ? ' x' + qty : ''));
    groups[ono].qty += qty;
  });

  var stamp = dateStamp();
  var newRows = [];
  order.forEach(function (ono) {
    if (seen[ono]) return;
    seen[ono] = true;
    var g = groups[ono], r = g.first;
    newRows.push([
      stamp, ono,
      pick(r, NAVER_COLS.orderDate),
      pick(r, NAVER_COLS.shipDate),
      pick(r, NAVER_COLS.receiver),
      normalizePhone(pick(r, NAVER_COLS.receiverPhone)),
      normalizeZip(pick(r, NAVER_COLS.zipcode)),
      pick(r, NAVER_COLS.address),
      g.items.join(' + '),
      g.qty,
      pick(r, NAVER_COLS.deliveryMemo)
    ]);
  });
  if (newRows.length) {
    hist.getRange(hist.getLastRow() + 1, 1, newRows.length, HISTORY_COLS.length).setValues(newRows);
  }
  SpreadsheetApp.getUi().alert('주문내역에 ' + newRows.length + '건(주문 단위로 합침, 중복 제외) 추가됨.\n' +
    '"고객관리 갱신"을 누르면 고객별로 집계됩니다.');
}

// ── 고객관리 갱신: 주문내역을 수취인+전화 기준으로 집계 ──
function updateCustomers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hist = ss.getSheetByName(SHEET_HISTORY);
  if (!hist || hist.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('"' + SHEET_HISTORY + '" 탭에 데이터가 없습니다.\n먼저 "주문내역에 누적하기"를 실행하세요.');
    return;
  }
  var map = {}, order = [];
  readAsObjects(hist).forEach(function (r) {
    var name = ('' + (r['수취인명'] || '')).trim();
    var phone = ('' + (r['전화번호'] || '')).trim();
    if (!name && !phone) return;
    var key = name + '|' + phone;
    if (!map[key]) {
      map[key] = { name: name, phone: phone, addr: '', cnt: 0, qty: 0, items: {}, first: '', last: '' };
      order.push(key);
    }
    var c = map[key];
    if (r['주소']) c.addr = r['주소'];
    c.cnt += 1;
    c.qty += parseInt(r['수량'], 10) || 0;
    var item = ('' + (r['상품명'] || '')).trim();
    if (item) c.items[item] = true;
    var d = ('' + (r['주문일시'] || '')).trim();
    if (d) {
      if (!c.first || d < c.first) c.first = d;
      if (!c.last  || d > c.last)  c.last = d;
    }
  });

  var out = order.map(function (key) {
    var c = map[key];
    return [c.name, c.phone, c.addr, c.cnt, c.qty, Object.keys(c.items).join(', '), c.first, c.last];
  });
  out.sort(function (a, b) { return b[3] - a[3]; }); // 총주문건수 많은 순

  var cust = ss.getSheetByName(SHEET_CUSTOMER) || ss.insertSheet(SHEET_CUSTOMER);
  cust.clear();
  cust.getRange(1, 1, 1, CUSTOMER_COLS.length).setValues([CUSTOMER_COLS]).setFontWeight('bold');
  if (out.length) cust.getRange(2, 1, out.length, CUSTOMER_COLS.length).setValues(out);
  cust.setFrozenRows(1);
  SpreadsheetApp.getUi().alert('고객관리 갱신 완료 ✅  고객 ' + out.length + '명 집계.');
}

// ── 문자관리 갱신: 발송일 기준 안부(+4)/재구매(+10) 대상·문구 정리 ──
function updateMessagePlan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hist = ss.getSheetByName(SHEET_HISTORY);
  if (!hist || hist.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('"' + SHEET_HISTORY + '" 탭에 데이터가 없습니다.\n먼저 "주문내역에 누적하기"를 실행하세요.');
    return;
  }
  var today = ymd(new Date());
  var map = {}, order = [];
  readAsObjects(hist).forEach(function (r) {
    var ship = parseDateLoose(r['발송일']);
    if (!ship) return;   // 아직 발송 안 된 주문은 제외
    var name  = ('' + (r['수취인명'] || '')).trim();
    var phone = ('' + (r['전화번호'] || '')).trim();
    var key = name + '|' + phone + '|' + ymd(ship);
    if (!map[key]) { map[key] = { name: name, phone: phone, ship: ship, items: {} }; order.push(key); }
    var item = ('' + (r['상품명'] || '')).trim();
    if (item) map[key].items[item] = true;
  });

  var rows = order.map(function (key) {
    var c = map[key];
    var helloDue = ymd(addDays(c.ship, FOLLOWUP_HELLO_DAYS));
    var repDue   = ymd(addDays(c.ship, FOLLOWUP_REPURCHASE_DAYS));
    var products = Object.keys(c.items).join(', ');
    return {
      todayHit: (helloDue === today || repDue === today),
      row: [
        c.name, c.phone, products, ymd(c.ship),
        helloDue, followStatus(helloDue, today), MSG_HELLO.replace('{name}', c.name),
        repDue,   followStatus(repDue, today),   MSG_REPURCHASE.replace('{name}', c.name)
      ]
    };
  });
  // 오늘 보낼 대상 먼저, 그다음 발송일 최신순
  rows.sort(function (a, b) {
    if (a.todayHit !== b.todayHit) return a.todayHit ? -1 : 1;
    return a.row[3] < b.row[3] ? 1 : -1;
  });
  var out = rows.map(function (x) { return x.row; });
  var todayCount = rows.filter(function (x) { return x.todayHit; }).length;

  var msg = ss.getSheetByName(SHEET_MSG) || ss.insertSheet(SHEET_MSG);
  msg.clear();
  msg.getRange(1, 1, 1, MSG_COLS.length).setValues([MSG_COLS]).setFontWeight('bold');
  if (out.length) msg.getRange(2, 1, out.length, MSG_COLS.length).setValues(out);
  msg.setFrozenRows(1);

  SpreadsheetApp.getUi().alert('문자관리 갱신 완료 ✅\n오늘 보낼 문자 대상: ' + todayCount + '명\n' +
    '(안부상태·재구매상태 칸에 "📮 오늘 보내기"로 표시된 사람에게 문구를 복사해 보내세요.)');
}

// 상태: 오늘이면 보내기, 지났으면 완료(지남), 아직이면 대기
function followStatus(dueYmd, todayYmd) {
  if (dueYmd === todayYmd) return '📮 오늘 보내기';
  return dueYmd < todayYmd ? '완료(지남)' : '대기';
}

// 느슨한 날짜 파싱: "2026/07/27", "2026-07-27", "2026.7.27", Date 모두 처리
function parseDateLoose(v) {
  if (!v && v !== 0) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return v;
  var m = ('' + v).match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  if (!m) return null;
  return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
}

function addDays(d, n) { var x = new Date(d.getTime()); x.setDate(x.getDate() + n); return x; }
function ymd(d) { return Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd'); }

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
  zip:    function (r) { return normalizeZip(pick(r, NAVER_COLS.zipcode)); },
  itemName: function (r) {
    if (r.__itemName) return r.__itemName.substring(0, 100); // 병합된 행: 이미 조립됨(수량 중복 방지)
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
    r.__itemName = r.__items.join(' + ');   // 품목명은 __itemName으로 (수량 중복 방지)
    r[NAVER_COLS.quantity] = r.__qty;        // 내품수량 = 합산 수량
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

// 우편번호 정규화: 숫자만 뽑아 5자리 미만이면 앞에 0을 채움(03968 등 앞자리 0 복구)
function normalizeZip(v) {
  var s = (v === null || v === undefined) ? '' : ('' + v).trim();
  var d = s.replace(/[^0-9]/g, '');
  if (d.length >= 1 && d.length < 5) d = ('00000' + d).slice(-5);
  return d || s;
}

function csvEscape(v) {
  var s = (v === null || v === undefined) ? '' : v.toString();
  if (/[",\r\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

function dateStamp() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd_HHmm');
}
