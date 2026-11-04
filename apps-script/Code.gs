/**
 * 〈라파엘〉 시사회 참석 여부 회신 — Google Apps Script 웹 앱
 *
 * 역할
 *  - 초대장 페이지(index.html)에서 보낸 회신을 "시사회 회신" 시트에 그대로 기록한다.
 *  - 명단 대조는 하지 않는다. (잘못 들어온 데이터는 시트에서 수기로 거른다)
 *  - 같은 연락처로 다시 보내면 기존 줄을 덮어쓴다. (페이지의 "회신 내용 수정" 기능)
 *
 * 설치
 *  1. 스프레드시트 메뉴 확장 프로그램 > Apps Script 를 열고 이 파일 내용을 Code.gs에 붙여넣고 저장한다.
 *  2. 배포 > 새 배포 > 유형: 웹 앱 / 실행 사용자: 나 / 액세스 권한: 모든 사용자 → 배포.
 *     (이미 배포한 적이 있으면 배포 > 배포 관리 > 연필 > 버전 "새 버전" 으로 다시 배포)
 *  3. 웹 앱 URL(…/exec)을 index.html 의 ENDPOINT 에 넣는다.
 */

const CONFIG = {
  REPLY_SHEET: '시사회 회신',   // 기록 시트 이름 (없으면 자동 생성)
  MAX_COMPANIONS: 2,
};

const REPLY_HEADERS = [
  '회신 시각', '성함', '연락처', '이메일',
  '동행자1 성함', '동행자1 연락처', '동행자2 성함', '동행자2 연락처',
  '동행 인원', '총 인원', '수정 횟수'
];

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (_) { return json_({ ok: false, status: 'bad_request' }); }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    return json_(handle_(body));
  } catch (err) {
    return json_({ ok: false, status: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ ok: true, status: 'alive' });
}

function handle_(body) {
  const name = clean_(body.name);
  const phone = digits_(body.phone);
  const email = clean_(body.email).toLowerCase();
  if (!name || !phone || !email) return { ok: false, status: 'bad_request' };

  // 동행자: 성함·연락처가 모두 있는 경우만 기록 (최대 2명)
  const companions = (Array.isArray(body.companions) ? body.companions : [])
    .map(c => ({ name: clean_(c && c.name), phone: digits_(c && c.phone) }))
    .filter(c => c.name && c.phone)
    .slice(0, CONFIG.MAX_COMPANIONS);

  const sh = getReplySheet_();
  const prev = findByPhone_(sh, phone);
  const edits = prev ? (Number(prev.edits) || 0) + 1 : 0;
  const row = [
    new Date(), name, phone, email,
    companions[0] ? companions[0].name : '', companions[0] ? companions[0].phone : '',
    companions[1] ? companions[1].name : '', companions[1] ? companions[1].phone : '',
    companions.length, 1 + companions.length, edits
  ];
  let rowIndex;
  if (prev) { rowIndex = prev.row; }
  else { sh.appendRow(row); rowIndex = sh.getLastRow(); }
  const range = sh.getRange(rowIndex, 1, 1, row.length);
  // 연락처 열은 텍스트로 고정(앞자리 0이 숫자 변환으로 사라지는 것 방지) 후 값 기록
  [3, 6, 8].forEach(col => sh.getRange(rowIndex, col).setNumberFormat('@'));
  range.setValues([row]);

  return { ok: true, status: 'recorded', invitee: { name, phone, email }, companions, updated: !!prev };
}

/* ---------------- 회신 시트 ---------------- */

function getReplySheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CONFIG.REPLY_SHEET);
  // 이전 버전(명단 대조 버전)이 만든 시트면 이름을 바꿔 보관하고 새로 만든다
  if (sh) {
    const head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(clean_);
    const same = REPLY_HEADERS.every((h, i) => head[i] === h);
    if (!same) {
      let n = 1, nm = CONFIG.REPLY_SHEET + ' (이전)';
      while (ss.getSheetByName(nm)) nm = CONFIG.REPLY_SHEET + ' (이전 ' + (++n) + ')';
      sh.setName(nm);
      sh = null;
    }
  }
  if (!sh) {
    sh = ss.insertSheet(CONFIG.REPLY_SHEET);
    sh.getRange(1, 1, 1, REPLY_HEADERS.length).setValues([REPLY_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
    ['C:C', 'F:F', 'H:H'].forEach(r => sh.getRange(r).setNumberFormat('@'));   // 연락처 열은 텍스트(앞자리 0 유지)
  }
  return sh;
}

function findByPhone_(sh, phone) {
  const last = sh.getLastRow();
  if (last < 2) return null;
  const values = sh.getRange(2, 1, last - 1, REPLY_HEADERS.length).getValues();
  const key = phone.replace(/^0+/, '');   // 숫자로 저장돼 앞자리 0이 빠진 경우도 같은 번호로 취급
  for (let i = 0; i < values.length; i++) {
    const v = digits_(values[i][2]).replace(/^0+/, '');
    if (v && v === key) return { row: i + 2, edits: values[i][10] };
  }
  return null;
}

/* ---------------- 유틸 ---------------- */

function clean_(v) { return v == null ? '' : String(v).replace(/\s+/g, ' ').trim(); }
function digits_(v) {
  let d = String(v == null ? '' : v).replace(/\D/g, '');
  if (d.startsWith('82') && d.length >= 11) d = '0' + d.slice(2);
  return d;
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
