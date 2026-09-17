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
     생성기의 [시트에 반영] 버튼도 이 스크립트를 통해 동작합니다.

   ★ 이 파일을 수정한 뒤에는: [배포] → [배포 관리] → 연필(수정)
     → 버전: 새 버전 → [배포] 를 해야 반영됩니다 (URL은 그대로 유지).
   ============================================================ */

var ADMIN_PW = 'CHANGE-ME';   // ★ 사이트 관리자 비밀번호와 동일하게 변경하세요
var OP_SHEET = '의견';        // 의견이 쌓이는 탭 이름 (자동 생성)

function doPost(e) {
  var out = { ok: false, error: 'bad_request' };
  try {
    var req = JSON.parse(e.postData.contents);
    if (req.action === 'submit') out = submitOpinions(req);
    else if (req.action === 'apply' || req.action === 'close' || req.action === 'replace') {
      if (String(req.pw || '') !== ADMIN_PW) out = { ok: false, error: 'bad_pw' };
      else out = req.action === 'apply' ? applyOpinion(req)
           : req.action === 'replace' ? replaceData(req)
           : closeOpinion(req);
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
      appendTextRow(s, [now, String(req.dept || ''), it.term || '', it.type || '',
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

/* 학기키 복원 : 시트가 '2027-1'을 날짜(2027년 1월)로 자동 변환해 저장하는 경우가 있어
   Date나 '2027. 1.' 같은 형태를 다시 '2027-1' 텍스트로 되돌린다 */
function normTerm(v) {
  if (v instanceof Date) return v.getFullYear() + '-' + (v.getMonth() + 1);
  var s = String(v || '').trim();
  var m = s.match(/^(\d{4})\D+([12])\D*$/);
  return m ? m[1] + '-' + m[2] : s;
}

/* 행을 텍스트 서식으로 고정해 추가 — 날짜·학기키가 자동 변환되지 않게 */
function appendTextRow(sheet, vals) {
  var row = sheet.getLastRow() + 1;
  var rg = sheet.getRange(row, 1, 1, vals.length);
  rg.setNumberFormat('@');
  rg.setValues([vals]);
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
    var term = normTerm(v[2]), type = String(v[3]).trim();
    var targetName = String(v[4]).trim(), targetIso = normD(v[5]);
    var name = String(v[6]).trim(), iso = normD(v[7]), end = normD(v[8]);
    var kind = String(v[9]).trim() || '학사';
    var d = dataSheet();

    if (type === '추가') {
      if (!name || !iso) return { ok: false, error: 'missing_fields' };
      appendTextRow(d, ['일정', term, iso, end || iso, name, kind]);
    } else if (type === '수정' || type === '삭제') {
      var data = d.getDataRange().getValues();
      var idx = -1;
      for (var i = 0; i < data.length; i++) {
        var r = data[i];
        if (String(r[0]).trim() === '일정' && normTerm(r[1]) === term &&
            String(r[4]).trim() === targetName &&
            (!targetIso || normD(r[2]) === targetIso)) { idx = i + 1; break; }
      }
      if (idx < 0) return { ok: false, error: 'target_not_found' };
      if (type === '삭제') d.deleteRow(idx);
      else {
        if (name) d.getRange(idx, 5).setValue(name);
        if (iso) { var rg = d.getRange(idx, 3, 1, 2); rg.setNumberFormat('@'); rg.setValues([[iso, end || iso]]); }
        else if (end) { var rg2 = d.getRange(idx, 4); rg2.setNumberFormat('@'); rg2.setValue(end); }
      }
    } else {
      return { ok: false, error: 'not_applicable' };  // 기타 의견은 [확인]으로 처리
    }
    op.sheet.getRange(op.row, 12).setValue('반영됨');
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* [시트에 반영] : 생성기의 현재 입력값(req.data)으로 일정 데이터 탭을 다시 씀
   — 메모·제출주소·비밀번호해시 행은 그대로 보존 */
function replaceData(req) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var d = req.data || {};
    if (!d.year || !d.terms || !Object.keys(d.terms).length)
      return { ok: false, error: 'bad_data' };
    var s = dataSheet();
    var keep = [];
    s.getDataRange().getValues().forEach(function (r) {
      var t = String(r[0] || '').trim();
      if (t === '메모' || t === '제출주소' || t === '비밀번호해시') keep.push(r.slice(0, 6));
    });
    var KIND_KO = { off: '휴업일', exam: '시험', makeup: '보강', event: '행사', admin: '학사' };
    var rows = [['종류', '학기', '시작일', '종료일', '명칭', '분류']];
    rows.push(['학년도', String(d.year), '', '', '', '']);
    rows.push(['표시학기', String(+d.sem === 2 ? 2 : 1), '', '', '', '']);
    rows.push(['주차', String(d.targetWeeks || 15), '', '', '', '']);
    var fb = d.feedback || {};
    rows.push(['의견수렴', '', fb.start || '', fb.end || '', '', '']);
    var pr = d.priority || {};
    if (pr.year) rows.push(['우선표시', String(pr.year), pr.start || '', pr.end || '', '', '']);
    rows.push(['연도잠금', d.lockYear === false ? '0' : '1', '', '', '', '']);
    Object.keys(d.terms).sort().forEach(function (k) {
      var t = d.terms[k] || {};
      if (/^\d{4}-[12]$/.test(k) && t.start && t.end) rows.push(['학기', k, t.start, t.end, '', '']);
    });
    Object.keys(d.user || {}).sort().forEach(function (k) {
      if (!/^\d{4}-[12]$/.test(k)) return;
      (d.user[k] || []).forEach(function (u) {
        if (!u || !u.iso || !u.name) return;
        rows.push(['일정', k, String(u.iso), String(u.end || u.iso), String(u.name), KIND_KO[u.kind] || '학사']);
      });
    });
    keep.forEach(function (r) { rows.push(r); });
    s.clearContents();
    var rg = s.getRange(1, 1, rows.length, 6);
    rg.setNumberFormat('@');   // 날짜가 자동 형식 변환되지 않게 텍스트로 고정
    rg.setValues(rows);
    return { ok: true, rows: rows.length };
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
