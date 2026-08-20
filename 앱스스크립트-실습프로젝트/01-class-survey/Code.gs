/**
 * ============================================================================
 *  Code.gs  —  서버에서 실행되는 코드 (.gs 파일)
 * ============================================================================
 *  © 2026 티쳐무 · 모든 권리 보유 — 학교 수업 목적으로만 이용해 주세요.
 *  무단 배포·상업적 이용을 금합니다. 자세한 내용은 LICENSE 파일을 보세요.
 * ============================================================================
 *  주제 ①  학급 설문 웹앱  ─  "우리 학교 급식, 어떻게 생각해?"
 *
 *  .gs 파일은 구글 서버에서 실행된다. 그래서 이 파일 안에서만
 *  스프레드시트를 직접 열고, 읽고, 쓸 수 있다.
 *  브라우저(index.html)는 스프레드시트를 직접 만질 수 없고,
 *  반드시 이 파일의 함수를 "불러서" 부탁해야 한다.
 *
 *      [브라우저] --- google.script.run ---> [Code.gs] ---> [스프레드시트]
 *         화면                 부탁              일꾼            창고
 *
 *  ⚠ 데이터 윤리: 이 설문은 이름·학번을 받지 않는다.
 *     "누가 답했는지"는 이 프로젝트의 질문에 필요하지 않기 때문이다.
 *     필요 없는 항목을 받지 않는 것을 '최소 수집'이라고 한다.
 * ============================================================================
 */


/* ────────────────────────────────────────────────────────────────────────────
 * 1. 설정값 (CONFIG)
 *    바꿔야 할 값을 코드 곳곳에 흩어 두지 않고 맨 위에 모아 둔다.
 *    나중에 시트 이름을 바꿀 때 여기 한 곳만 고치면 된다.
 * ──────────────────────────────────────────────────────────────────────────*/
const CONFIG = {
  // 빈 문자열('')로 두면 "이 스크립트가 붙어 있는 스프레드시트"를 사용한다.
  // 따로 만든 스프레드시트를 쓰려면 주소창 URL에서 ID를 복사해 넣는다.
  //   https://docs.google.com/spreadsheets/d/★이 부분이 ID★/edit
  SPREADSHEET_ID: '',

  SHEET_NAME: '응답',                       // 데이터를 쌓을 시트 이름

  HEADERS: ['타임스탬프', '학년', '만족도', '좋았던 메뉴', '개선 의견'],

  MAX_COMMENT_LENGTH: 200,                  // 서술형 답변 최대 길이
  RECENT_COMMENT_COUNT: 5                   // 결과 화면에 보여줄 의견 개수
};


/* ────────────────────────────────────────────────────────────────────────────
 * 2. doGet — 웹앱의 대문
 *    웹앱 주소로 접속하면 구글이 자동으로 이 함수를 실행한다.
 *    이름을 doGet 말고 다른 것으로 바꾸면 웹앱이 열리지 않는다. (약속된 이름)
 *
 *    반환값은 두 가지 중 하나여야 한다.
 *      · HtmlOutput  → 사람이 보는 화면
 *      · TextOutput  → 프로그램이 읽는 데이터(JSON 등)
 * ──────────────────────────────────────────────────────────────────────────*/
function doGet(e) {
  // ── 2-1. 주소 뒤에 ?api=stats 를 붙이면 화면 대신 JSON을 돌려준다.
  //         웹앱은 '화면'도 되고 '데이터 제공소(API)'도 된다는 것을 보여주는 부분.
  //         테스트: 웹앱주소?api=stats
  if (e && e.parameter && e.parameter.api === 'stats') {
    return ContentService
        .createTextOutput(JSON.stringify(getStats()))
        .setMimeType(ContentService.MimeType.JSON);
  }

  // ── 2-2. 평소에는 index.html을 화면으로 만들어 보낸다.
  //         createTemplateFromFile을 쓰면 <?!= ?> 같은 스크립틀릿이 작동한다.
  //         (createHtmlOutputFromFile을 쓰면 스크립틀릿이 그냥 글자로 나온다)
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('급식 만족도 설문')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


/* ────────────────────────────────────────────────────────────────────────────
 * 3. include — 파일을 조각조각 나눠 쓰기 위한 도우미
 *    index.html 안에서  <?!= include('css') ?>  라고 쓰면
 *    css.html의 내용이 그 자리에 그대로 끼워진다.
 *    덕분에 화면(index) · 모양(css) · 동작(js)을 따로 관리할 수 있다.
 * ──────────────────────────────────────────────────────────────────────────*/
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/* ────────────────────────────────────────────────────────────────────────────
 * 4. setupSheet — 처음 한 번만 실행하는 준비 함수
 *    편집기 위쪽 함수 목록에서 setupSheet를 고르고 ▶실행 을 누른다.
 *    시트를 손으로 만드는 대신 코드로 만들어 두면
 *    친구가 내 프로젝트를 복제할 때도 똑같은 구조가 보장된다.
 * ──────────────────────────────────────────────────────────────────────────*/
function setupSheet() {
  const ss = openSpreadsheet_();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  // 시트가 없으면 새로 만든다.
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }

  // 1행에 머리글을 쓴다.  getRange(행, 열, 행개수, 열개수)
  sheet.getRange(1, 1, 1, CONFIG.HEADERS.length)
       .setValues([CONFIG.HEADERS])          // setValues는 항상 '2차원 배열'을 받는다
       .setFontWeight('bold')
       .setBackground('#eaf0ff');

  sheet.setFrozenRows(1);                    // 1행 고정 → 스크롤해도 머리글이 보인다
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 320);

  Logger.log('준비 완료: "%s" 시트에 머리글을 넣었습니다.', CONFIG.SHEET_NAME);
  return '준비 완료';
}


/* ────────────────────────────────────────────────────────────────────────────
 * 5. saveResponse — 응답 1건을 시트에 저장한다  (C : Create)
 *    브라우저의 js.html이 google.script.run.saveResponse(payload) 로 호출한다.
 *
 *    @param {Object} payload  { grade, score, menu, comment }
 *    @return {Object}         { ok: true, total: 12 }  ← 화면에 돌려줄 결과
 * ──────────────────────────────────────────────────────────────────────────*/
function saveResponse(payload) {
  // ── 5-1. 검증: 화면에서 막아도 서버에서 한 번 더 확인한다.
  //         화면 코드는 사용자가 얼마든지 우회할 수 있기 때문이다. (중요!)
  const clean = validate_(payload);

  // ── 5-2. 잠금: 두 사람이 같은 순간에 제출하면 한 줄이 사라질 수 있다.
  //         LockService로 "한 명씩 차례대로" 쓰게 만든다.
  const lock = LockService.getScriptLock();
  lock.waitLock(10 * 1000);                  // 최대 10초 기다린다

  try {
    const sheet = getSheet_();

    // appendRow: 데이터가 있는 마지막 줄 다음에 한 줄을 붙인다.
    // 순서는 CONFIG.HEADERS와 반드시 같아야 한다.
    sheet.appendRow([
      new Date(),            // 타임스탬프 (시트가 자동으로 날짜 서식으로 보여준다)
      clean.grade,           // 학년
      clean.score,           // 만족도 1~5
      clean.menu,            // 좋았던 메뉴
      clean.comment          // 개선 의견
    ]);

    // 머리글 1줄을 뺀 값이 실제 응답 수
    const total = sheet.getLastRow() - 1;
    return { ok: true, total: total };

  } catch (err) {
    // 서버에서 일부러 오류를 던지면 화면의 withFailureHandler가 받는다.
    Logger.log('저장 실패: ' + err.message);
    throw new Error('저장 중 문제가 생겼습니다: ' + err.message);

  } finally {
    lock.releaseLock();                      // 성공이든 실패든 잠금은 반드시 해제
  }
}


/* ────────────────────────────────────────────────────────────────────────────
 * 6. getStats — 시트를 읽어 통계를 계산한다  (R : Read)
 *    시트 값을 한 번에 다 읽어 자바스크립트 안에서 계산한다.
 *    셀을 하나씩 getValue()로 읽으면 아주 느려진다. 반드시 getValues()로 한 번에!
 * ──────────────────────────────────────────────────────────────────────────*/
function getStats() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();   // [[머리글...], [1행], [2행], ...]

  // 머리글만 있고 응답이 없는 경우
  if (values.length < 2) {
    return { total: 0, average: 0, dist: [0, 0, 0, 0, 0], byGrade: {}, comments: [] };
  }

  const rows = values.slice(1);                      // 머리글 한 줄 잘라내기

  let sum = 0;
  const dist = [0, 0, 0, 0, 0];                      // 1점~5점이 각각 몇 개인지
  const gradeSum = {};                               // 학년별 점수 합
  const gradeCnt = {};                               // 학년별 응답 수
  const comments = [];

  rows.forEach(function (row) {
    const grade   = row[1];
    const score   = Number(row[2]);
    const comment = String(row[4] || '').trim();

    if (score >= 1 && score <= 5) {
      sum += score;
      dist[score - 1]++;
      gradeSum[grade] = (gradeSum[grade] || 0) + score;
      gradeCnt[grade] = (gradeCnt[grade] || 0) + 1;
    }
    if (comment) comments.push(comment);
  });

  // 학년별 평균 만들기
  const byGrade = {};
  Object.keys(gradeSum).sort().forEach(function (g) {
    byGrade[g] = round1_(gradeSum[g] / gradeCnt[g]);
  });

  return {
    total: rows.length,
    average: round1_(sum / rows.length),
    dist: dist,
    byGrade: byGrade,
    // 최근 의견 몇 개만 (원자료 전체를 화면에 뿌리지 않는다 — 데이터 윤리)
    comments: comments.slice(-CONFIG.RECENT_COMMENT_COUNT).reverse()
  };
}


/* ────────────────────────────────────────────────────────────────────────────
 * 7. 내부 도우미 함수들
 *    함수 이름 끝에 밑줄(_)을 붙이면 '이 파일 안에서만 쓰는 함수'가 된다.
 *    밑줄이 붙은 함수는 google.script.run으로 호출할 수 없다 → 안전장치.
 * ──────────────────────────────────────────────────────────────────────────*/

/** 스프레드시트를 연다. */
function openSpreadsheet_() {
  return CONFIG.SPREADSHEET_ID
      ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)   // 따로 만든 시트
      : SpreadsheetApp.getActive();                      // 이 스크립트가 붙어 있는 시트
}

/** 데이터 시트를 가져온다. 없으면 친절한 오류를 낸다. */
function getSheet_() {
  const sheet = openSpreadsheet_().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    throw new Error('"' + CONFIG.SHEET_NAME + '" 시트가 없습니다. setupSheet를 먼저 실행하세요.');
  }
  return sheet;
}

/** 들어온 값이 규칙에 맞는지 확인하고, 깨끗하게 다듬어 돌려준다. */
function validate_(payload) {
  if (!payload) throw new Error('보낸 값이 없습니다.');

  const grade = String(payload.grade || '').trim();
  const score = Number(payload.score);
  const menu  = String(payload.menu || '').trim();
  const comment = String(payload.comment || '').trim();

  if (['1학년', '2학년', '3학년'].indexOf(grade) === -1) {
    throw new Error('학년을 선택해 주세요.');
  }
  if (!(score >= 1 && score <= 5)) {
    throw new Error('만족도를 1~5 중에서 골라 주세요.');
  }
  if (comment.length > CONFIG.MAX_COMMENT_LENGTH) {
    throw new Error('의견은 ' + CONFIG.MAX_COMMENT_LENGTH + '자까지 쓸 수 있습니다.');
  }

  return { grade: grade, score: score, menu: menu, comment: comment };
}

/** 소수 첫째 자리까지 반올림. 12.3456 → 12.3 */
function round1_(n) {
  return Math.round(n * 10) / 10;
}


/* ────────────────────────────────────────────────────────────────────────────
 * 8. 확장 미션 (스스로 또는 AI와 함께 해 보기)
 *    ① 같은 사람이 여러 번 제출하는 것을 막아 보자.
 *       힌트: PropertiesService.getUserProperties() 로 "이미 제출함" 표시를 남긴다.
 *    ② 응답이 들어올 때마다 나에게 메일이 오게 해 보자.
 *       힌트: MailApp.sendEmail(...)  — 단, 응답자 정보는 보내지 않는다.
 *    ③ 결과 화면에 학년별 응답 '수'도 함께 보여 주자.
 *       힌트: getStats()가 돌려주는 객체에 gradeCount를 추가한다.
 *    ④ 오후 6시마다 하루 응답 수를 기록하는 트리거를 만들어 보자.
 *       힌트: 편집기 왼쪽 ⏰ 트리거 메뉴 → 시간 기반 트리거
 * ──────────────────────────────────────────────────────────────────────────*/
