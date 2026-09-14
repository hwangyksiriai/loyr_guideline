/* ════════════════════════════════════════════════════════════
   로이어(LOYR) 캠페인 신청서 수집 스크립트

   - 신청 내용을 '시트1' 탭에 한 줄씩 쌓습니다.
   - 타입(A~G)·고료가 별도 컬럼으로 들어가고, 헤더에 필터가 걸려 있어
     고료별로 걸러 보거나 정렬할 수 있습니다.
   - 메뉴 [로이어] > '고료별로 정렬' 로 고료 순 정렬도 가능합니다.
   ════════════════════════════════════════════════════════════ */

/* ── 1. 대상 스프레드시트 ──────────────────────────────────
   이 스크립트를 해당 시트에 바인딩(시트 > 확장 프로그램 > Apps Script)했다면
   '' 로 비워둬도 됩니다. ID를 넣어두면 항상 이 파일에 기록합니다. */
const SPREADSHEET_ID = '1gN-5NTg3XEvdTW4HBd5bOReKWeXC6nsACnuv5oXC0xI';

/* ── 2. 기록할 탭 이름 ─────────────────────────────────── */
const SHEET_NAME = '시트1';

/* ── 3. 컬럼 구성 ──────────────────────────────────────────
   순서를 바꾸면 아래 PHONE_COL / ZIP_COL / TIER_COL / FEE_COL 번호도
   같이 바꿔야 합니다. */
const HEADERS = [
  '제출일시', '타입', '고료', '이름', '인스타그램', '휴대폰',
  '이메일', '희망 제품 라인', '우편번호', '배송지 주소', '요청사항'
];
const TIER_COL  = 2;   // 타입
const FEE_COL   = 3;   // 고료
const PHONE_COL = 6;   // 휴대폰
const ZIP_COL   = 9;   // 우편번호

/* ── 4. 고료 정렬 순서 ─────────────────────────────────────
   '고료별로 정렬' 메뉴에서 쓰는 순서입니다. 목록에 없는 값은 맨 뒤로 갑니다. */
const FEE_ORDER = ['5만원', '10만원', '15만원', '20만원', '25만원', '30만원', '고료 조정'];


/* ════════════════════════════════════════════════════════════
   신청서 수신
   ════════════════════════════════════════════════════════════ */
function doPost(e) {
  // 동시에 여러 명이 제출해도 행이 겹치지 않도록 잠금
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const data = JSON.parse(e.postData.contents);

    const sheet = getOrCreateSheet_();

    const values = [
      new Date(),
      formatTier_(data.tier),     // 'A' → 'A Type'
      data.price || '',
      data.name || '',
      data.instagram || '',
      data.phone || '',
      data.email || '',
      data.productLine || '',
      data.zipcode || '',
      data.address || '',
      data.note || ''
    ];

    const row = sheet.getLastRow() + 1;

    // 앞자리 0이 날아가지 않도록 숫자처럼 보이는 칸을 먼저 텍스트 서식으로 고정
    sheet.getRange(row, PHONE_COL).setNumberFormat('@');
    sheet.getRange(row, ZIP_COL).setNumberFormat('@');

    const range = sheet.getRange(row, 1, 1, values.length);
    range.setValues([values]);
    range.setVerticalAlignment('middle');

    // 타입·고료 칸만 옅게 강조해서 한눈에 구분되게 (행 전체는 흰색 유지)
    sheet.getRange(row, TIER_COL, 1, 2).setBackground('#E8EDF4').setFontWeight('bold');

    return json_({ result: 'success', sheet: sheet.getName(), row: row });

  } catch (err) {
    // 실패해도 신청 내용이 사라지지 않도록 로그에 원문을 남김
    console.error('제출 처리 실패: ' + err + ' | payload: ' +
      (e && e.postData ? e.postData.contents : '(없음)'));
    return json_({ result: 'error', message: String(err) });

  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ status: 'ok' });
}


/* ════════════════════════════════════════════════════════════
   내부 함수
   ════════════════════════════════════════════════════════════ */

/** 페이지는 tier를 'A'처럼 한 글자로 보냅니다. 시트에는 'A Type'으로 기록. */
function formatTier_(tier) {
  const t = String(tier || '').trim();
  if (!t) return '';
  return /type/i.test(t) ? t : t + ' Type';
}

function getSpreadsheet_() {
  return SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

/** 탭이 없으면 만들고, 헤더가 없으면 헤더를 세팅한다 */
function getOrCreateSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setValues([HEADERS])
      .setFontWeight('bold')
      .setBackground('#355484')
      .setFontColor('#FFFFFF')
      .setVerticalAlignment('middle');

    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150);   // 제출일시
    sheet.setColumnWidth(2, 90);    // 타입
    sheet.setColumnWidth(3, 90);    // 고료
    sheet.setColumnWidth(5, 230);   // 인스타그램
    sheet.setColumnWidth(8, 180);   // 희망 제품 라인
    sheet.setColumnWidth(10, 300);  // 배송지 주소
    sheet.setColumnWidth(11, 220);  // 요청사항

    if (!sheet.getFilter()) {
      sheet.getRange(1, 1, 1, HEADERS.length).createFilter();
    }
  }

  return sheet;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/* ════════════════════════════════════════════════════════════
   시트에서 직접 쓰는 메뉴 (스크립트가 시트에 바인딩된 경우에만 표시)
   ════════════════════════════════════════════════════════════ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('로이어')
    .addItem('고료별로 정렬', 'sortByFee')
    .addItem('헤더 다시 설정', 'setupHeader')
    .addToUi();
}

/** 고료(5만원 → 30만원 → 고료 조정) → 제출일시 순으로 정렬 */
function sortByFee() {
  const sheet = getOrCreateSheet_();
  const last = sheet.getLastRow();
  if (last < 3) return;

  const range = sheet.getRange(2, 1, last - 1, HEADERS.length);
  const rows = range.getValues();

  rows.sort(function (a, b) {
    const ra = feeRank_(a[FEE_COL - 1]);
    const rb = feeRank_(b[FEE_COL - 1]);
    if (ra !== rb) return ra - rb;
    return new Date(a[0]) - new Date(b[0]);   // 같은 고료면 제출 순
  });

  range.setValues(rows);
}

function feeRank_(fee) {
  const i = FEE_ORDER.indexOf(String(fee || '').trim());
  return i === -1 ? FEE_ORDER.length : i;
}

/** 헤더가 날아갔거나 컬럼을 바꿨을 때 한 번 실행 */
function setupHeader() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  sheet.getRange(1, 1, 1, HEADERS.length)
    .setValues([HEADERS])
    .setFontWeight('bold')
    .setBackground('#355484')
    .setFontColor('#FFFFFF')
    .setVerticalAlignment('middle');

  sheet.setFrozenRows(1);
  if (!sheet.getFilter()) {
    sheet.getRange(1, 1, 1, HEADERS.length).createFilter();
  }
}


/* ════════════════════════════════════════════════════════════
   테스트용 — 편집기에서 실행하면 시트1에 A~G 더미 행 7개가 들어갑니다.
   확인 후 해당 행들은 지우세요.
   ════════════════════════════════════════════════════════════ */
function 테스트_더미행_넣기() {
  const cases = [
    ['A', '5만원'],
    ['B', '10만원'],
    ['C', '15만원'],
    ['D', '20만원'],
    ['E', '25만원'],
    ['F', '30만원'],
    ['G', '고료 조정']
  ];

  cases.forEach(function (c, i) {
    doPost({
      postData: {
        contents: JSON.stringify({
          tier: c[0],
          price: c[1],
          name: '테스트' + (i + 1),
          instagram: 'https://instagram.com/test' + (i + 1),
          phone: '010-0000-000' + (i + 1),
          email: 'test@example.com',
          productLine: i % 2 === 0 ? '블루펩타이드 라인' : '세럼 라인',
          zipcode: '06234',
          address: '서울시 강남구 테헤란로 1 101동 101호',
          note: '테스트 행입니다'
        })
      }
    });
  });
}
