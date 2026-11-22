/**
 * 드론 관제 미션 · 전체 기록 서버
 * 연결된 스프레드시트: "드론 관제 미션 기록"
 * 사용법: 오른쪽 위 [배포 → 새 배포 → 유형: 웹 앱] / 실행 사용자: 나 / 액세스 권한: 모든 사용자 → 배포
 */
const SHEET_ID = '1llZEshS-5It_nGeqAid8HtIgwCksoG2VOfLDZRVat0w';
const HEADER = ['시각', '이름', '미션', '미션 이름', '결과', '별', '점수', '위반', '소요(초)', '착륙', '출발', '상황 조치', '실패 사유'];

function sheet_() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  if (sh.getLastRow() === 0) { sh.appendRow(HEADER); }
  return sh;
}
function safe_(v, n) {
  let s = String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, '').slice(0, n || 40);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}
function num_(v, max) { const n = Number(v); return isFinite(n) ? Math.max(-100000, Math.min(max || 100000, Math.round(n))) : 0; }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const name = safe_(d.name, 20);
    if (!name) return json_({ ok: false });
    sheet_().appendRow([
      new Date(), name, num_(d.mission, 15), safe_(d.title, 30), d.ok ? '성공' : '실패',
      num_(d.stars, 3), num_(d.score), num_(d.viol, 99), num_(d.secs, 600),
      num_(d.land, 99), num_(d.dep, 99), safe_(d.ev, 10), safe_(d.reason, 60)
    ]);
    return json_({ ok: true });
  } finally { lock.releaseLock(); }
}

function doGet() {
  const values = sheet_().getDataRange().getValues().slice(1).slice(-5000);
  const rows = values.map(r => ({
    time: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
    name: String(r[1]).replace(/^'/, ''), mission: Number(r[2]), title: String(r[3]),
    ok: r[4] === '성공', stars: Number(r[5]), score: Number(r[6]), viol: Number(r[7])
  }));
  return json_({ rows });
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/** 배포 전에 한 번 실행해 권한을 승인하고 연결을 확인하는 함수 */
function testConnection() { Logger.log('연결된 시트: ' + SpreadsheetApp.openById(SHEET_ID).getName() + ' / 행 수: ' + sheet_().getLastRow()); }
