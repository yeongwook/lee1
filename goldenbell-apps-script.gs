/**
 * 독도 골든벨 기록 저장용 구글 Apps Script
 *
 * [설정 방법]
 * 1. 구글 시트를 새로 만듭니다. (예: "독도 골든벨 기록")
 * 2. 시트 메뉴 [확장 프로그램] → [Apps Script]를 엽니다.
 * 3. 기본으로 있는 내용을 모두 지우고, 이 파일 내용을 통째로 붙여 넣은 뒤 저장합니다.
 * 4. 오른쪽 위 [배포] → [새 배포] → 유형 선택(톱니바퀴) → [웹 앱]
 *      - 다음 사용자 인증 정보로 실행: 나
 *      - 액세스 권한이 있는 사용자: 모든 사용자
 *    → [배포]를 누르고 권한을 허용합니다.
 * 5. 나오는 "웹 앱 URL"(https://script.google.com/macros/s/.../exec)을 복사해
 *    goldenbell.html 과 goldenbell-board.html 맨 아래 쪽 API_URL 에 붙여 넣습니다.
 *
 * ※ 이 스크립트를 고친 뒤에는 [배포] → [배포 관리] → 연필 → 버전 "새 버전" → [배포]를 해야 반영됩니다.
 * ※ 기록을 지우려면 시트의 해당 행을 직접 삭제하면 됩니다. (1행 제목은 남겨 두세요)
 */

var SHEET_NAME = '골든벨기록';
var HEADERS = ['저장 시각', '학년', '반', '이름', '점수', '정답 수', '도달 문제', '골든벨', '키'];
var QUESTION_COUNT = 20;          // 골든벨 문항 수
var MAX_SCORE = QUESTION_COUNT * 30; // 문항당 최대 30점(정답 10 + 남은 시간 최대 20)

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function normName_(s) {
  return String(s || '').replace(/\s+/g, '').toLowerCase();
}

/** 순위판이 기록을 읽어 갈 때 */
function doGet(e) {
  var sh = getSheet_();
  var last = sh.getLastRow();
  var rows = last > 1 ? sh.getRange(2, 1, last - 1, HEADERS.length).getValues() : [];
  var results = rows
    .filter(function (r) { return r[1] !== '' && r[3] !== ''; })
    .map(function (r) {
      return {
        at: r[0] instanceof Date ? r[0].getTime() : Number(r[0]) || 0,
        g: Number(r[1]), c: Number(r[2]), name: String(r[3]),
        score: Number(r[4]), correct: Number(r[5]), reached: Number(r[6]),
        golden: r[7] === 'O'
      };
    });
  return json_({ ok: true, results: results, serverTime: Date.now() });
}

/** 학생 기록 저장 */
function doPost(e) {
  var data;
  try { data = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'invalid' }); }
  if (data.action !== 'submit') return json_({ ok: false, error: 'invalid' });

  var g = Number(data.g), c = Number(data.c), score = Number(data.score),
      correct = Number(data.correct), reached = Number(data.reached);
  var name = String(data.name || '').replace(/\s+/g, ' ').trim().slice(0, 12);

  // 말이 안 되는 값은 거절
  var valid = [1, 2, 3].indexOf(g) >= 0 && c >= 1 && c <= 15 && name.length >= 2 &&
    score >= 0 && score <= MAX_SCORE && correct >= 0 && correct <= reached &&
    reached >= 1 && reached <= QUESTION_COUNT && score <= correct * 30 && score >= correct * 10;
  if (!valid) return json_({ ok: false, error: 'invalid' });

  var key = g + '|' + c + '|' + normName_(name);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = getSheet_();
    var last = sh.getLastRow();
    if (last > 1) {
      var keys = sh.getRange(2, 9, last - 1, 1).getValues();
      for (var i = 0; i < keys.length; i++) {
        if (keys[i][0] === key) return json_({ ok: true, dup: true });
      }
    }
    var golden = (reached === QUESTION_COUNT && correct >= QUESTION_COUNT - 1) ? 'O' : '';
    sh.appendRow([new Date(), g, c, name, score, correct, reached, golden, key]);
    return json_({ ok: true, dup: false });
  } finally {
    lock.releaseLock();
  }
}
