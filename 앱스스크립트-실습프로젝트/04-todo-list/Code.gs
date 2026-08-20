/**
 * ============================================================================
 *  Code.gs  —  서버에서 실행되는 코드 (.gs 파일)
 * ============================================================================
 *  © 2026 티쳐무 · 모든 권리 보유 — 학교 수업 목적으로만 이용해 주세요.
 *  무단 배포·상업적 이용을 금합니다. 자세한 내용은 LICENSE 파일을 보세요.
 * ============================================================================
 *  주제 ④  나의 학습 계획 관리 웹앱  (가장 단순한 입문용)
 *
 *  네 가지 동작만 있으면 '앱'이 된다.
 *      추가 → addTask()      appendRow()
 *      목록 → getTasks()     getValues()
 *      완료 → toggleTask()   setValue()
 *      삭제 → deleteTask()   deleteRow()
 *
 *  이 프로젝트는 남의 데이터를 다루지 않는다. 내가 쓴 내 기록만 저장된다.
 *  그래서 웹앱 배포 시 접근 권한을 '나만'으로 두는 것이 자연스럽다.
 * ============================================================================
 */


/* ── 1. 설정값 ───────────────────────────────────────────────────────────── */
const CONFIG = {
  SPREADSHEET_ID: '',
  SHEET_NAME: '할일',
  HEADERS: ['ID', '등록일', '할 일', '마감일', '중요도', '완료'],
  COL: { ID: 1, CREATED: 2, TITLE: 3, DUE: 4, PRIORITY: 5, DONE: 6 }
};


/* ── 2. 웹앱의 대문 ──────────────────────────────────────────────────────── */
function doGet(e) {
  if (e && e.parameter && e.parameter.api === 'tasks') {
    return ContentService
        .createTextOutput(JSON.stringify(getTasks()))
        .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('나의 학습 계획')
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
       .setValues([CONFIG.HEADERS]).setFontWeight('bold').setBackground('#fdeee4');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(3, 280);

  Logger.log('준비 완료: "%s" 시트를 만들었습니다.', CONFIG.SHEET_NAME);
  return '준비 완료';
}


/* ── 4. R : 목록 읽기 ────────────────────────────────────────────────────── */
function getTasks() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) return { tasks: [], todo: 0, done: 0 };

  const tasks = values.slice(1)
      .filter(function (r) { return r[CONFIG.COL.TITLE - 1]; })
      .map(function (r) {
        return {
          id:       String(r[CONFIG.COL.ID - 1]),
          title:    String(r[CONFIG.COL.TITLE - 1]),
          due:      formatDate_(r[CONFIG.COL.DUE - 1]),
          priority: String(r[CONFIG.COL.PRIORITY - 1] || '보통'),
          done:     r[CONFIG.COL.DONE - 1] === true || r[CONFIG.COL.DONE - 1] === 'TRUE'
        };
      });

  // 정렬: 아직 안 한 일을 위로, 그다음 마감일이 빠른 것을 위로.
  // 비교 함수가 음수를 돌려주면 a가 앞, 양수면 b가 앞으로 간다.
  tasks.sort(function (a, b) {
    if (a.done !== b.done) return a.done ? 1 : -1;
    return String(a.due || '9999').localeCompare(String(b.due || '9999'));
  });

  return {
    tasks: tasks,
    todo: tasks.filter(function (t) { return !t.done; }).length,
    done: tasks.filter(function (t) { return t.done; }).length
  };
}


/* ── 5. C : 추가 ─────────────────────────────────────────────────────────── */
function addTask(payload) {
  const title = String((payload && payload.title) || '').trim();
  const due   = String((payload && payload.due) || '').trim();
  const priority = String((payload && payload.priority) || '보통').trim();

  if (!title) throw new Error('할 일을 입력해 주세요.');
  if (title.length > 80) throw new Error('할 일은 80자까지 쓸 수 있습니다.');
  if (due && !/^\d{4}-\d{2}-\d{2}$/.test(due)) throw new Error('마감일 형식이 올바르지 않습니다.');

  const sheet = getSheet_();

  // ID는 절대 겹치면 안 된다. 겹치면 findRow_가 엉뚱한 줄을 찾아
  // 다른 항목이 완료 표시되거나 삭제된다.
  //   ✗ 'T' + new Date().getTime()  → 같은 밀리초에 두 개가 들어오면 똑같아진다
  //   ✓ Utilities.getUuid()         → 구글이 겹치지 않는 값을 만들어 준다
  const id = 'T' + Utilities.getUuid().slice(0, 8);

  sheet.appendRow([id, new Date(), title, due, priority, false]);
  return { ok: true, id: id };
}


/* ── 6. U : 완료 표시 바꾸기 ─────────────────────────────────────────────── */
function toggleTask(id, done) {
  const sheet = getSheet_();
  const row = findRow_(sheet, id);

  // 셀 하나만 바꿀 때는 setValue (s가 없다), 여러 칸은 setValues (s가 있다)
  sheet.getRange(row, CONFIG.COL.DONE).setValue(done === true);

  return { ok: true };
}


/* ── 7. D : 삭제 ─────────────────────────────────────────────────────────── */
function deleteTask(id) {
  const sheet = getSheet_();
  const row = findRow_(sheet, id);
  const title = sheet.getRange(row, CONFIG.COL.TITLE).getValue();

  sheet.deleteRow(row);
  return { ok: true, message: '「' + title + '」 삭제됨' };
}

/** 완료한 일을 한 번에 모두 지운다. */
function clearDone() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();

  // ⚠ 아래에서 위로 지워야 한다!
  //   위에서부터 지우면 줄이 하나 지워질 때마다 아래 줄들이 위로 올라와
  //   행 번호가 어긋난다. 이건 초보자가 가장 많이 겪는 버그다.
  let count = 0;
  for (let i = values.length - 1; i >= 1; i--) {
    const done = values[i][CONFIG.COL.DONE - 1];
    if (done === true || done === 'TRUE') {
      sheet.deleteRow(i + 1);      // 배열 index → 시트 행 번호
      count++;
    }
  }
  return { ok: true, message: count + '개를 정리했습니다.' };
}


/* ── 8. 내부 도우미 ──────────────────────────────────────────────────────── */

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

function findRow_(sheet, id) {
  const ids = sheet.getRange(1, CONFIG.COL.ID, sheet.getLastRow(), 1)
                   .getValues().map(function (r) { return String(r[0]); });
  const index = ids.indexOf(String(id));
  if (index === -1) throw new Error('항목을 찾을 수 없습니다. 목록을 새로 고쳐 주세요.');
  return index + 1;
}

function formatDate_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value);
}


/* ── 9. 확장 미션 ────────────────────────────────────────────────────────── */
/*  ① 마감일이 지난 항목을 붉게 표시해 보자.
 *  ② 과목별로 분류하는 열을 추가하고, 과목별 완료율을 계산해 보자.
 *  ③ 아침 7시에 오늘 마감인 할 일을 메일로 보내는 트리거를 만들어 보자.
 *     힌트: MailApp.sendEmail(Session.getActiveUser().getEmail(), 제목, 내용)
 */
