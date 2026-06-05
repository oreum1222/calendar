/**
 * [오름] 마감 캘린더 — Apps Script 백엔드
 * ────────────────────────────────────────────────────────────
 * 같은 구글시트 안에 시트 3개를 만든다(이름·헤더 정확히):
 *
 *  ① 시트명 "일정"  (1행 헤더)
 *     제목 | 마감일 | 카테고리 | 담당 | 상태 | 반복 | 메모 | 링크
 *     - 마감일: 2026-06-07  /  담당: "전체" 또는 "은규T, 지연T"
 *     - 상태: todo / doing / done   반복: 없음 / 매주 / 매월
 *
 *  ② 시트명 "강사"  (1행 헤더)
 *     이름 | PIN
 *     가경T | 1212      ← PIN은 강사별 4자리 등 자유(원장만 알기)
 *     은규T | 1111
 *     ...(원철T까지 6명)
 *
 *  ③ 시트명 "조교"  (1행 헤더 — 강사와 같은 형식, 별도 탭)
 *     이름 | PIN
 *     이유섭 | 7001
 *     ...(조교 명단)
 *     ※ 일정의 담당 칸에 조교 이름을 직접 쓰거나 "조교전체"로 배정.
 *
 *  ④ 시트명 "체크"  (1행 헤더만, 내용은 자동 기록)
 *     제목 | 마감일 | 이름 | 완료 | 시각
 *
 * 배포: [배포] → [새 배포] → 유형 '웹 앱'
 *       실행 계정 = 나,  액세스 권한 = '모든 사용자' → 배포 → URL(...exec) 복사
 *       → index.html 의 CONFIG.apiUrl 에 붙여넣기
 * (시트/PIN 수정 후 코드 변경이 없으면 재배포 불필요)
 * ────────────────────────────────────────────────────────────
 */
const SS = SpreadsheetApp.getActiveSpreadsheet();
const TZ = Session.getScriptTimeZone() || 'Asia/Seoul';

function out(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function ymd(v){ return (v instanceof Date) ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v||'').trim(); }
function rowsOf(name){ const sh = SS.getSheetByName(name); return sh ? sh.getDataRange().getValues() : []; }

/* 일정 + 체크 + 강사명단 반환 */
function doGet(e){
  const ev = rowsOf('일정'), rows = [];
  for (let i = 1; i < ev.length; i++){
    const r = ev[i];
    if (!r[0]) continue;
    rows.push([ r[0], ymd(r[1]), r[2], r[3], r[4], r[5], r[6], r[7] ]);
  }
  const members = [];
  ['강사','조교'].forEach(function(name){
    const m = rowsOf(name);
    for (let i = 1; i < m.length; i++){ if (m[i][0]) members.push(String(m[i][0]).trim()); }
  });

  const ck = rowsOf('체크'), checks = {};
  for (let i = 1; i < ck.length; i++){
    const title = String(ck[i][0]||'').trim(), due = ymd(ck[i][1]), name = String(ck[i][2]||'').trim();
    if (!title || !name) continue;
    const v = ck[i][3];
    const done = (v === true || v === 1 || String(v).toUpperCase() === 'TRUE' || v === '완료' || v === 'O');
    (checks[title + '||' + due] = checks[title + '||' + due] || {})[name] = done;
  }
  return out({ rows, checks, members });
}

/* PIN 검증: '강사' + '조교' 시트의 이름↔PIN 대조 */
function verify(name, pin){
  return checkRoster('강사', name, pin) || checkRoster('조교', name, pin);
}
function checkRoster(sheet, name, pin){
  const mem = rowsOf(sheet);
  for (let i = 1; i < mem.length; i++){
    if (String(mem[i][0]).trim() === String(name).trim())
      return String(mem[i][1]).trim() === String(pin).trim();
  }
  return false;
}

/* 본인확인(auth) / 체크 기록(check) */
function doPost(e){
  let b = {};
  try { b = JSON.parse(e.postData.contents); } catch (err){ return out({ ok:false, error:'bad json' }); }

  if (b.action === 'auth') return out({ ok: verify(b.name, b.pin) });

  if (b.action === 'check'){
    if (!verify(b.name, b.pin)) return out({ ok:false, error:'PIN' });
    upsertCheck(b.title, b.due, b.name, !!b.done);
    return out({ ok:true });
  }
  return out({ ok:false, error:'unknown action' });
}

/* 체크 행 upsert (제목+마감일+이름 동일하면 갱신, 없으면 추가) */
function upsertCheck(title, due, name, done){
  const sh = SS.getSheetByName('체크');
  const data = sh.getDataRange().getValues();
  const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
  for (let i = 1; i < data.length; i++){
    if (String(data[i][0]).trim() === String(title).trim()
        && ymd(data[i][1]) === ymd(due)
        && String(data[i][2]).trim() === String(name).trim()){
      sh.getRange(i+1, 4).setValue(done ? '완료' : '');
      sh.getRange(i+1, 5).setValue(now);
      return;
    }
  }
  sh.appendRow([ title, due, name, done ? '완료' : '', now ]);
}
