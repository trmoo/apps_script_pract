/**
 * ============================================================================
 *  Code.gs  —  서버에서 실행되는 코드 (.gs 파일)
 * ============================================================================
 *  © 2026 티쳐무 · 모든 권리 보유 — 학교 수업 목적으로만 이용해 주세요.
 *  무단 배포·상업적 이용을 금합니다. 자세한 내용은 LICENSE 파일을 보세요.
 * ============================================================================
 *  주제 ②  학급 문고 대출 관리 웹앱
 *
 *  이 프로젝트의 목표는 CRUD 네 가지를 모두 경험하는 것이다.
 *
 *      C  Create  만들기  → addBook()      appendRow()
 *      R  Read    읽기    → getBooks()     getDataRange().getValues()
 *      U  Update  고치기  → borrowBook()   getRange(...).setValues()
 *      D  Delete  지우기  → deleteBook()   deleteRow()
 *
 *  ⚠ 데이터 윤리: 대출자는 실명 대신 '별칭'을 쓴다.
 *     누가 어떤 책을 읽었는지는 취향에 관한 정보이고,
 *     학급 전체가 볼 수 있는 화면에 실명으로 남길 이유가 없다.
 * ============================================================================
 */


/* ── 1. 설정값 ───────────────────────────────────────────────────────────── */
const CONFIG = {
  SPREADSHEET_ID: '',                 // 비워 두면 이 스크립트가 붙어 있는 스프레드시트 사용
  SHEET_NAME: '도서목록',
  LOG_SHEET_NAME: '대출기록',

  HEADERS:     ['도서ID', '제목', '저자', '상태', '대출자 별칭', '대출일'],
  LOG_HEADERS: ['타임스탬프', '도서ID', '제목', '구분', '별칭'],

  STATUS_FREE: '대출가능',
  STATUS_OUT:  '대출중',

  // 열 번호를 숫자로 흩어 쓰면 나중에 열을 추가할 때 전부 어긋난다.
  // 이름으로 정해 두고 쓰면 열 순서가 바뀌어도 여기만 고치면 된다.
  COL: { ID: 1, TITLE: 2, AUTHOR: 3, STATUS: 4, NICK: 5, DATE: 6 }
};


/* ── 2. 웹앱의 대문 ──────────────────────────────────────────────────────── */
function doGet(e) {
  // 주소 뒤에 ?api=books 를 붙이면 화면 대신 JSON을 돌려준다.
  if (e && e.parameter && e.parameter.api === 'books') {
    return ContentService
        .createTextOutput(JSON.stringify(getBooks('')))
        .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('학급 문고 대출 관리')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/* ── 3. 처음 한 번만 실행 ────────────────────────────────────────────────── */
function setupSheet() {
  const ss = openSpreadsheet_();

  // 도서목록 시트
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.insertSheet(CONFIG.SHEET_NAME);
  sheet.clear();
  sheet.getRange(1, 1, 1, CONFIG.HEADERS.length)
       .setValues([CONFIG.HEADERS]).setFontWeight('bold').setBackground('#e6f4ec');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(2, 240);

  // 연습용 샘플 3권을 한 번에 쓴다.
  // 한 줄씩 appendRow를 3번 부르는 것보다 setValues 한 번이 훨씬 빠르다.
  const samples = [
    ['B001', '아몬드',        '손원평', CONFIG.STATUS_FREE, '', ''],
    ['B002', '완득이',        '김려령', CONFIG.STATUS_FREE, '', ''],
    ['B003', '수레바퀴 아래서', '헤르만 헤세', CONFIG.STATUS_FREE, '', '']
  ];
  sheet.getRange(2, 1, samples.length, CONFIG.HEADERS.length).setValues(samples);

  // 대출기록 시트
  let log = ss.getSheetByName(CONFIG.LOG_SHEET_NAME) || ss.insertSheet(CONFIG.LOG_SHEET_NAME);
  log.clear();
  log.getRange(1, 1, 1, CONFIG.LOG_HEADERS.length)
     .setValues([CONFIG.LOG_HEADERS]).setFontWeight('bold').setBackground('#e6f4ec');
  log.setFrozenRows(1);

  Logger.log('준비 완료: 시트 2개와 샘플 도서 3권을 만들었습니다.');
  return '준비 완료';
}


/* ── 4. R : 읽기 ─────────────────────────────────────────────────────────── */
/**
 * 도서 목록을 돌려준다.
 * @param {string} keyword  검색어(비우면 전체)
 * @return {Object} { books: [...], total, out }
 */
function getBooks(keyword) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1);                       // 머리글 제거

  const key = String(keyword || '').trim().toLowerCase();

  const books = rows
      .filter(function (row) { return row[CONFIG.COL.ID - 1]; })   // 빈 줄 제외
      .map(function (row) {
        // 배열을 객체로 바꿔 두면 화면 코드에서 row[3] 대신 book.status 로 쓸 수 있다.
        return {
          id:     row[CONFIG.COL.ID     - 1],
          title:  row[CONFIG.COL.TITLE  - 1],
          author: row[CONFIG.COL.AUTHOR - 1],
          status: row[CONFIG.COL.STATUS - 1],
          nick:   row[CONFIG.COL.NICK   - 1],
          date:   formatDate_(row[CONFIG.COL.DATE - 1])
        };
      })
      .filter(function (book) {
        if (!key) return true;
        return (book.title + ' ' + book.author).toLowerCase().indexOf(key) !== -1;
      });

  return {
    books: books,
    total: rows.filter(function (r) { return r[0]; }).length,
    out: rows.filter(function (r) { return r[CONFIG.COL.STATUS - 1] === CONFIG.STATUS_OUT; }).length
  };
}


/* ── 5. C : 만들기 ───────────────────────────────────────────────────────── */
function addBook(payload) {
  const title  = String((payload && payload.title)  || '').trim();
  const author = String((payload && payload.author) || '').trim();

  if (!title) throw new Error('책 제목을 입력해 주세요.');
  if (title.length > 60) throw new Error('제목이 너무 깁니다.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10 * 1000);
  try {
    const sheet = getSheet_();
    const id = nextId_(sheet);                       // B004, B005 … 자동으로 만든다

    sheet.appendRow([id, title, author, CONFIG.STATUS_FREE, '', '']);
    writeLog_(id, title, '등록', '');

    return { ok: true, id: id };
  } finally {
    lock.releaseLock();
  }
}


/* ── 6. U : 고치기 ───────────────────────────────────────────────────────── */
/**
 * 책을 대출 처리한다.
 * 핵심: "몇 번째 줄인지"를 미리 알 수 없다. 그래서 ID로 줄을 찾아야 한다.
 */
function borrowBook(id, nickname) {
  const nick = String(nickname || '').trim();
  if (!nick) throw new Error('별칭을 입력해 주세요. (실명은 쓰지 않습니다)');
  if (nick.length > 12) throw new Error('별칭은 12자까지 가능합니다.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10 * 1000);
  try {
    const sheet = getSheet_();
    const row = findRow_(sheet, id);                 // 예: 3 (3번째 줄)

    // 이미 대출 중인지 확인 — 두 사람이 동시에 빌리는 것을 막는다.
    const status = sheet.getRange(row, CONFIG.COL.STATUS).getValue();
    if (status === CONFIG.STATUS_OUT) {
      throw new Error('이미 대출 중인 책입니다. 목록을 새로 고쳐 주세요.');
    }

    const title = sheet.getRange(row, CONFIG.COL.TITLE).getValue();

    // 세 칸을 한 번에 고친다. 4~6열이 붙어 있으므로 range 하나로 처리할 수 있다.
    sheet.getRange(row, CONFIG.COL.STATUS, 1, 3)
         .setValues([[CONFIG.STATUS_OUT, nick, new Date()]]);

    writeLog_(id, title, '대출', nick);
    return { ok: true, message: '「' + title + '」 대출 완료' };
  } finally {
    lock.releaseLock();
  }
}

/** 책을 반납 처리한다. */
function returnBook(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10 * 1000);
  try {
    const sheet = getSheet_();
    const row = findRow_(sheet, id);

    const title = sheet.getRange(row, CONFIG.COL.TITLE).getValue();
    const nick  = sheet.getRange(row, CONFIG.COL.NICK).getValue();

    // 반납하면 대출자·대출일을 지운다.
    // 데이터 윤리: 목록에 남길 필요가 없는 정보는 지운다.
    //             (기록이 필요하면 별도의 대출기록 시트에만 남긴다)
    sheet.getRange(row, CONFIG.COL.STATUS, 1, 3)
         .setValues([[CONFIG.STATUS_FREE, '', '']]);

    writeLog_(id, title, '반납', nick);
    return { ok: true, message: '「' + title + '」 반납 완료' };
  } finally {
    lock.releaseLock();
  }
}


/* ── 7. D : 지우기 ───────────────────────────────────────────────────────── */
function deleteBook(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10 * 1000);
  try {
    const sheet = getSheet_();
    const row = findRow_(sheet, id);

    const status = sheet.getRange(row, CONFIG.COL.STATUS).getValue();
    if (status === CONFIG.STATUS_OUT) {
      throw new Error('대출 중인 책은 삭제할 수 없습니다. 먼저 반납해 주세요.');
    }

    const title = sheet.getRange(row, CONFIG.COL.TITLE).getValue();
    sheet.deleteRow(row);                            // 줄 자체를 없앤다 (되돌릴 수 없다!)

    writeLog_(id, title, '삭제', '');
    return { ok: true, message: '「' + title + '」 삭제 완료' };
  } finally {
    lock.releaseLock();
  }
}


/* ── 8. 내부 도우미 (이름 끝에 _ → 화면에서 부를 수 없다) ───────────────── */

function openSpreadsheet_() {
  return CONFIG.SPREADSHEET_ID
      ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
      : SpreadsheetApp.getActive();
}

function getSheet_() {
  const sheet = openSpreadsheet_().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('"' + CONFIG.SHEET_NAME + '" 시트가 없습니다. setupSheet를 먼저 실행하세요.');
  return sheet;
}

/**
 * 도서ID로 시트의 줄 번호를 찾는다.
 * A열만 읽어 오는 것이 핵심 — 시트 전체를 읽으면 느리다.
 * indexOf가 0을 돌려주면 1행(머리글)이므로 +1을 해서 실제 행 번호로 바꾼다.
 */
function findRow_(sheet, id) {
  const ids = sheet.getRange(1, CONFIG.COL.ID, sheet.getLastRow(), 1)
                   .getValues()
                   .map(function (r) { return String(r[0]); });

  const index = ids.indexOf(String(id));
  if (index === -1) throw new Error('해당 도서를 찾을 수 없습니다. (ID: ' + id + ')');
  return index + 1;                                  // 배열은 0부터, 시트 행은 1부터
}

/** 다음 도서ID를 만든다. B003 다음은 B004 */
function nextId_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return 'B001';

  const ids = sheet.getRange(2, CONFIG.COL.ID, last - 1, 1).getValues();
  let max = 0;
  ids.forEach(function (r) {
    const n = parseInt(String(r[0]).replace('B', ''), 10);
    if (!isNaN(n) && n > max) max = n;
  });

  // 4 → "B004"  ('0'을 붙여 자릿수를 맞추는 흔한 방법)
  return 'B' + ('00' + (max + 1)).slice(-3);
}

/** 대출기록 시트에 한 줄 남긴다. */
function writeLog_(id, title, kind, nick) {
  const log = openSpreadsheet_().getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (log) log.appendRow([new Date(), id, title, kind, nick]);
}

/** 날짜 값을 "9/12" 형태의 문자로 바꾼다. (Date 객체는 화면으로 보낼 수 없다!) */
function formatDate_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'M/d');
  }
  return String(value);
}


/* ── 9. 확장 미션 ────────────────────────────────────────────────────────── */
/*  ① 대출한 지 7일이 지난 책을 목록에서 붉게 표시해 보자.
 *     힌트: 대출일과 오늘의 차이를 계산해 getBooks가 overdue: true 를 함께 보내게 한다.
 *  ② '대출기록' 시트를 읽어 가장 인기 있는 책 3권을 보여 주는 함수를 만들어 보자.
 *  ③ 삭제 대신 '상태'를 '폐기'로 바꾸는 방식으로 고쳐 보자.
 *     실제 서비스는 데이터를 진짜로 지우지 않고 표시만 바꾸는 경우가 많다. 왜 그럴까?
 */
