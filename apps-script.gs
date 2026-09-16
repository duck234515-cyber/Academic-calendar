/* ============================================================
   학사일정 의견 접수·반영 스크립트 (Google Apps Script)

   설치 방법 (한 번만):
   1. 연동 구글시트 열기 → 확장 프로그램 → Apps Script
   2. 이 파일 내용을 전부 붙여넣고, 아래 ADMIN_PW 를
      사이트 관리자 비밀번호와 같게 수정
   3. 저장 → [배포] → [새 배포] → 유형: 웹 앱
      - 실행 계정: 나
      - 액세스 권한: 모든 사용자
   4. 배포 후 나오는 웹 앱 URL(…/exec)을 복사해,
      시트에 행을 하나 추가: A열 `제출주소`, B열에 URL 붙여넣기
   → 이후 뷰어의 [의견 제출]은 파일 저장 없이 바로 전송되고,
     관리자는 뷰어의 [의견 검토]에서 [반영] 클릭으로 시트를 수정합니다.
   ============================================================ */

var ADMIN_PW = 'CHANGE-ME';   // ★ 사이트 관리자 비밀번호와 동일하게 변경하세요
var OP_SHEET = '의견';        // 의견이 쌓이는 탭 이름 (자동 생성)

function doPost(e) {
  var out = { ok: false, error: 'bad_request' };
  try {
    var req = JSON.parse(e.postData.contents);
    if (req.action === 'submit') out = submitOpinions(req);
    else if (req.action === 'apply' || req.action === 'close') {
      if (String(req.pw || '') !== ADMIN_PW) out = { ok: false, error: 'bad_pw' };
      else out = req.action === 'apply' ? applyOpinion(req) : closeOpinion(req);
    } else out = { ok: false, error: 'bad_action' };
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function dataSheet() { return ss().getSheets()[0]; }   // 첫 번째 탭 = 일정 데이터
function opSheet() {
  var s = ss().getSheetByName(OP_SHEET);
  if (!s) {
    s = ss().insertSheet(OP_SHEET);
    s.appendRow(['제출일시','부서명','학기','구분','대상일정','대상시작일',
                 '제안명칭','제안시작일','제안종료일','분류','의견','상태']);
  }
  return s;
}

/* 부서 의견 제출 : 의견 탭에 행 추가 */
function submitOpinions(req) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var s = opSheet();
    var now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
    var items = req.items || [];
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      s.appendRow([now, String(req.dept || ''), it.term || '', it.type || '',
                   it.targetName || '', it.targetIso || '', it.name || '',
                   it.iso || '', it.end || '', it.kind || '', it.note || '', '대기']);
    }
    return { ok: true, count: items.length };
  } finally { lock.releaseLock(); }
}

function normD(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
  var s = String(v || '').trim();
  var m = s.match(/^(\d{4})[.\-\/]\s*(\d{1,2})[.\-\/]\s*(\d{1,2})\.?$/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  return s;
}

/* 의견 행 확인 : 행 번호 + 제출일시·부서명 대조 (목록이 밀렸을 때 오반영 방지) */
function checkOpRow(req) {
  var s = opSheet();
  var row = +req.row;
  if (!(row >= 2 && row <= s.getLastRow())) return null;
  var v = s.getRange(row, 1, 1, 12).getValues()[0];
  if (String(v[0]) !== String(req.ts) && normD(v[0]) !== String(req.ts)) {
    // 표시형식 차이 허용: 문자열 비교 실패 시 부서·구분까지 맞으면 통과
    if (String(v[1]) !== String(req.dept)) return null;
  }
  if (String(v[11]).trim() !== '대기') return null;
  return { sheet: s, row: row, v: v };
}

/* [반영] : 제안 내용을 일정 데이터 탭에 적용하고 상태를 반영됨으로 */
function applyOpinion(req) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var op = checkOpRow(req);
    if (!op) return { ok: false, error: 'row_mismatch' };
    var v = op.v;
    var term = String(v[2]).trim(), type = String(v[3]).trim();
    var targetName = String(v[4]).trim(), targetIso = normD(v[5]);
    var name = String(v[6]).trim(), iso = normD(v[7]), end = normD(v[8]);
    var kind = String(v[9]).trim() || '학사';
    var d = dataSheet();

    if (type === '추가') {
      if (!name || !iso) return { ok: false, error: 'missing_fields' };
      d.appendRow(['일정', term, iso, end || iso, name, kind]);
    } else if (type === '수정' || type === '삭제') {
      var data = d.getDataRange().getValues();
      var idx = -1;
      for (var i = 0; i < data.length; i++) {
        var r = data[i];
        if (String(r[0]).trim() === '일정' && String(r[1]).trim() === term &&
            String(r[4]).trim() === targetName &&
            (!targetIso || normD(r[2]) === targetIso)) { idx = i + 1; break; }
      }
      if (idx < 0) return { ok: false, error: 'target_not_found' };
      if (type === '삭제') d.deleteRow(idx);
      else {
        if (name) d.getRange(idx, 5).setValue(name);
        if (iso) { d.getRange(idx, 3).setValue(iso); d.getRange(idx, 4).setValue(end || iso); }
        else if (end) d.getRange(idx, 4).setValue(end);
      }
    } else {
      return { ok: false, error: 'not_applicable' };  // 기타 의견은 [확인]으로 처리
    }
    op.sheet.getRange(op.row, 12).setValue('반영됨');
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* [반려]/[확인] : 일정 변경 없이 상태만 변경 */
function closeOpinion(req) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var op = checkOpRow(req);
    if (!op) return { ok: false, error: 'row_mismatch' };
    var st = req.status === '반려' ? '반려' : '확인';
    op.sheet.getRange(op.row, 12).setValue(st);
    return { ok: true };
  } finally { lock.releaseLock(); }
}
