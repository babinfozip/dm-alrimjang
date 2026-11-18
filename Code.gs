/**
 * 동명 알림장 · 우리 반 공유 서버 (구글 앱스크립트)
 *
 * 설치 방법
 * 1. 구글 시트를 새로 만들고 [확장 프로그램] → [Apps Script]를 눌러요.
 * 2. 원래 있던 코드를 지우고 이 파일 내용을 모두 붙여 넣은 뒤 저장해요.
 * 3. [배포] → [새 배포] → 유형: 웹 앱
 *      - 다음 사용자 인증 정보로 실행: 나
 *      - 액세스 권한이 있는 사용자: 모든 사용자
 *    [배포]를 누르고 권한을 허용해요.
 * 4. 나온 웹 앱 URL(https://script.google.com/macros/s/…/exec)을
 *    config.js의 SHARE_URL 따옴표 안에 붙여 넣어요.
 *
 * 시트에는 'rules'(과목별 준비물)와 'tests'(수행평가) 탭이 자동으로 생겨요.
 * 선생님은 시트에서 내용을 직접 보고 고치거나 지울 수 있어요.
 */

const RULES = 'rules';
const TESTS = 'tests';
const CLASSES = ['1-1', '1-2', '1-3', '2-1', '2-2', '2-3', '3-1', '3-2'];
// 시트에 '1-2'처럼 쓰면 날짜로 바뀌어서 '1학년 2반'으로 저장해요.
const label_ = gc => gc.split('-')[0] + '학년 ' + gc.split('-')[1] + '반';

function sheet_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
  }
  return sh;
}
const rulesSheet_ = () => sheet_(RULES, ['반', '과목', '준비물', '고친 시각']);
const testsSheet_ = () => sheet_(TESTS, ['반', 'id', '날짜', '과목', '내용', '적은 시각']);

function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\r\n\t]/g, ' ').trim().slice(0, max);
}

function read_(gc) {
  const lb = label_(gc);
  const rules = rulesSheet_().getDataRange().getValues().slice(1)
    .filter(r => r[0] === lb)
    .map(r => ({ subj: String(r[1]), items: String(r[2]) }));
  const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd');
  const tests = testsSheet_().getDataRange().getValues().slice(1)
    .filter(r => r[0] === lb && String(r[2]) >= today)
    .map(r => ({ id: String(r[1]), date: String(r[2]), subj: String(r[3]), what: String(r[4]) }))
    .sort((a, b) => a.date.localeCompare(b.date));
  return { ok: true, gc: gc, rules: rules, tests: tests, hasRules: rules.length > 0 };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const gc = String((e && e.parameter && e.parameter.gc) || '');
  if (CLASSES.indexOf(gc) < 0) return json_({ ok: false, error: '반 정보가 올바르지 않아요.' });
  return json_(read_(gc));
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: '잘못된 요청이에요.' }); }
  const gc = String(body.gc || '');
  if (CLASSES.indexOf(gc) < 0) return json_({ ok: false, error: '반 정보가 올바르지 않아요.' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const lb = label_(gc);
    const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');

    if (body.action === 'setRules') {
      const sh = rulesSheet_();
      const values = sh.getDataRange().getValues();
      for (let i = values.length - 1; i >= 1; i--) if (values[i][0] === lb) sh.deleteRow(i + 1);
      const rows = (Array.isArray(body.rules) ? body.rules : []).slice(0, 30)
        .map(r => [lb, clean_(r.subj, 30), clean_(r.items, 80), now])
        .filter(r => r[1] && r[2]);
      if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
    }

    if (body.action === 'addTest') {
      const t = body.test || {};
      const date = clean_(t.date, 8);
      const subj = clean_(t.subj, 30);
      if (/^\d{8}$/.test(date) && subj) {
        testsSheet_().appendRow([lb, 't' + Utilities.getUuid().slice(0, 8), date, subj, clean_(t.what, 40), now]);
      }
    }

    if (body.action === 'delTest') {
      const sh = testsSheet_();
      const values = sh.getDataRange().getValues();
      for (let i = values.length - 1; i >= 1; i--) {
        if (values[i][0] === lb && String(values[i][1]) === String(body.id)) sh.deleteRow(i + 1);
      }
    }
  } finally {
    lock.releaseLock();
  }
  return json_(read_(gc));
}
