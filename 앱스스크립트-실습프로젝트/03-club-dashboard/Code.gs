/**
 * ============================================================================
 *  Code.gs  —  서버에서 실행되는 코드 (.gs 파일)
 * ============================================================================
 *  © 2026 티쳐무 · 모든 권리 보유 — 학교 수업 목적으로만 이용해 주세요.
 *  무단 배포·상업적 이용을 금합니다. 자세한 내용은 LICENSE 파일을 보세요.
 * ============================================================================
 *  주제 ③  동아리 활동 기록 & 대시보드 웹앱
 *
 *  이 프로젝트의 목표는 "쌓인 데이터를 요약해서 보여 주기"다.
 *  시트에는 활동 기록이 한 줄씩 쌓이지만, 사람이 보고 싶은 것은 줄이 아니라
 *  "이번 달 몇 시간 했지?", "가장 많이 한 활동이 뭐지?" 같은 요약이다.
 *
 *      시트(원자료)  →  getSummary()에서 계산  →  화면(그래프·표)
 *
 *  ⚠ 데이터 윤리: 참여자 이름을 한 명씩 적지 않고 '참여 인원 수'만 기록한다.
 *     "누가 안 왔는지"는 이 대시보드의 질문에 필요하지 않다.
 * ============================================================================
 */


/* ── 1. 설정값 ───────────────────────────────────────────────────────────── */
const CONFIG = {
  SPREADSHEET_ID: '',
  SHEET_NAME: '활동기록',
  HEADERS: ['날짜', '활동명', '참여 인원', '활동 시간(분)', '메모'],

  MAX_MINUTES: 600,          // 하루 10시간을 넘는 값은 잘못 입력한 것으로 본다
  RECENT_COUNT: 6            // 최근 기록 몇 개를 보여 줄지
};


/* ── 2. 웹앱의 대문 ──────────────────────────────────────────────────────── */
function doGet(e) {
  if (e && e.parameter && e.parameter.api === 'summary') {
    return ContentService
        .createTextOutput(JSON.stringify(getSummary()))
        .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('동아리 활동 대시보드')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/* ── 3. 처음 한 번만 실행 ────────────────────────────────────────────────── */
function setupSheet() {
  const ss = openSpreadsheet_();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.insertSheet(CONFIG.SHEET_NAME);

  sheet.clear();
  sheet.getRange(1, 1, 1, CONFIG.HEADERS.length)
       .setValues([CONFIG.HEADERS]).setFontWeight('bold').setBackground('#f1eafe');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(5, 300);

  // 날짜 열에 서식을 지정해 두면 시트에서 보기 좋다.
  sheet.getRange('A2:A').setNumberFormat('yyyy-mm-dd');

  Logger.log('준비 완료: "%s" 시트를 만들었습니다.', CONFIG.SHEET_NAME);
  return '준비 완료';
}


/* ── 4. C : 활동 기록 추가 ───────────────────────────────────────────────── */
/**
 * @param {Object} payload { date:'2026-08-19', name, members, minutes, memo }
 *   ⚠ 날짜는 문자열로 받는다. google.script.run으로는 Date 객체를 보낼 수 없다!
 */
function addRecord(payload) {
  const clean = validate_(payload);

  const lock = LockService.getScriptLock();
  lock.waitLock(10 * 1000);
  try {
    const sheet = getSheet_();

    sheet.appendRow([
      clean.date,          // 'yyyy-MM-dd' 문자열을 시트가 날짜로 알아서 인식한다
      clean.name,
      clean.members,
      clean.minutes,
      clean.memo
    ]);

    return { ok: true, total: sheet.getLastRow() - 1 };
  } finally {
    lock.releaseLock();
  }
}


/* ── 5. R : 요약 계산 ────────────────────────────────────────────────────── */
/**
 * 시트 전체를 한 번 읽고, 필요한 숫자를 한 번의 반복으로 모두 계산한다.
 * @return {Object} 화면이 그대로 그릴 수 있는 형태로 정리해서 보낸다
 */
function getSummary() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return { count: 0, totalMinutes: 0, avgMembers: 0,
             byMonth: [], topActivities: [], recent: [] };
  }

  const rows = values.slice(1).filter(function (r) { return r[1]; });   // 활동명이 있는 줄만

  let totalMinutes = 0;
  let memberSum = 0;
  const monthMap = {};      // { '2026-08': 240, ... }
  const actMap   = {};      // { '봉사활동': 3, ... }

  rows.forEach(function (row) {
    const date    = row[0];
    const name    = String(row[1]).trim();
    const members = Number(row[2]) || 0;
    const minutes = Number(row[3]) || 0;

    totalMinutes += minutes;
    memberSum    += members;

    // 월별로 묶기 — '2026-08' 같은 열쇠(key)를 만들어 더한다.
    const key = monthKey_(date);
    if (key) monthMap[key] = (monthMap[key] || 0) + minutes;

    // 활동명별 횟수 세기
    actMap[name] = (actMap[name] || 0) + 1;
  });

  // ── 객체를 배열로 바꾼다. 화면에서 순서대로 그리려면 배열이 편하다.
  const byMonth = Object.keys(monthMap).sort().map(function (k) {
    return { month: k, hours: round1_(monthMap[k] / 60) };
  });

  // ── 많이 한 활동 순으로 정렬해 위에서 4개만
  const topActivities = Object.keys(actMap)
      .map(function (k) { return { name: k, count: actMap[k] }; })
      .sort(function (a, b) { return b.count - a.count; })      // 내림차순
      .slice(0, 4);

  // ── 최근 기록 (뒤에서부터 몇 개)
  const recent = rows.slice(-CONFIG.RECENT_COUNT).reverse().map(function (row) {
    return {
      date:    formatDate_(row[0]),
      name:    String(row[1]),
      members: Number(row[2]) || 0,
      minutes: Number(row[3]) || 0,
      memo:    String(row[4] || '')
    };
  });

  return {
    count: rows.length,
    totalMinutes: totalMinutes,
    totalHours: round1_(totalMinutes / 60),
    avgMembers: round1_(memberSum / rows.length),
    byMonth: byMonth,
    topActivities: topActivities,
    recent: recent
  };
}


/* ── 6. 내부 도우미 ──────────────────────────────────────────────────────── */

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

function validate_(payload) {
  if (!payload) throw new Error('보낸 값이 없습니다.');

  const date    = String(payload.date || '').trim();
  const name    = String(payload.name || '').trim();
  const members = Number(payload.members);
  const minutes = Number(payload.minutes);
  const memo    = String(payload.memo || '').trim();

  // 정규식으로 날짜 모양을 확인한다.  \d = 숫자 한 자
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('날짜를 선택해 주세요.');
  if (!name) throw new Error('활동명을 입력해 주세요.');
  if (!(members >= 1 && members <= 100)) throw new Error('참여 인원은 1~100명 사이여야 합니다.');
  if (!(minutes >= 1 && minutes <= CONFIG.MAX_MINUTES)) {
    throw new Error('활동 시간은 1~' + CONFIG.MAX_MINUTES + '분 사이여야 합니다.');
  }
  if (memo.length > 200) throw new Error('메모는 200자까지 쓸 수 있습니다.');

  return { date: date, name: name, members: members, minutes: minutes, memo: memo };
}

/** 날짜 값에서 '2026-08' 같은 월 열쇠를 만든다. */
function monthKey_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  return String(value).slice(0, 7);        // '2026-08-19' → '2026-08'
}

/** 날짜 값을 '8/19' 형태로 바꾼다. */
function formatDate_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'M/d');
  }
  return String(value).slice(5).replace('-', '/');
}

function round1_(n) {
  return Math.round(n * 10) / 10;
}


/* ── 7. 확장 미션 ────────────────────────────────────────────────────────── */
/*  ① 활동 시간이 가장 긴 달을 자동으로 찾아 "최고 기록" 카드로 보여 주자.
 *  ② 활동명을 미리 정해 둔 목록(봉사·연습·회의…)에서 고르게 바꿔 보자.
 *     자유 입력과 목록 선택 중 어느 쪽이 집계에 유리할까? 왜 그럴까?
 *  ③ 매주 월요일 아침에 지난주 활동 시간을 정리해 주는 트리거를 만들어 보자.
 *  ④ getSummary()가 계산을 다 하는 대신, 시트의 함수(SUMIF)를 쓰는 방법도 있다.
 *     두 방법의 장단점을 비교해 보자.
 */
