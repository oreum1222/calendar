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
 *     이름 | PIN | 이메일
 *     가경T | 1212 | gakyeong@gmail.com   ← 이메일은 리마인드용(선택). 없으면 그 사람은 메일 안 감.
 *     은규T | 1111 | ...
 *     ...(원철T까지 6명)
 *
 *  ③ 시트명 "조교"  (1행 헤더 — 강사와 같은 형식, 별도 탭)
 *     이름 | PIN | 이메일
 *     이유섭 | 7001 | ...
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
 *
 * ★ 지메일 자동 리마인드(매일 14시·21시):
 *   1) 강사/조교 시트 C열에 이메일 입력
 *   2) 상단 함수 선택창에서 setupTriggers 골라 [실행] 1회 → 권한 승인
 *      → 이후 매일 14시·21시에 각자 '본인 마감', 원장(가경T)에겐 '전체 미이행 요약'까지 발송
 *   (테스트: sendReminders 를 직접 실행하면 지금 즉시 발송됨)
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

/* ════════════════════════════════════════════════════════════
   지메일 자동 리마인드
   - setupTriggers() 1회 실행 → 매일 14시·21시 sendReminders 자동
   - 각 담당자: 본인 미완 마감(지남~D+7) 메일
   - 원장(OWNER): 본인 마감 + 전체 미이행 요약
   ════════════════════════════════════════════════════════════ */
var OWNER = '가경T';        // 전체 미이행 요약을 받는 사람
var WIN_AHEAD = 7;          // 다가오는 마감 표시 일수
var WIN_PAST_REPEAT = 7;    // 반복 일정 지난 표시 일수
var WIN_PAST_ONCE = 30;     // 일반 일정 지난(미완) 표시 일수

function parseYmd(s){ var p = String(s).split('-'); return new Date(+p[0], (+p[1])-1, +p[2]); }
function midnight(d){ return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function ddays(a,b){ return Math.round((a-b)/86400000); }
function esc(s){ return String(s==null?'':s).replace(/[&<>"]/g,function(c){return({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]);}); }

function getRoster(){
  var emailOf={}, teachers=[], assistants=[];
  function read(sheet, arr){
    var m=rowsOf(sheet);
    for(var i=1;i<m.length;i++){
      var name=String(m[i][0]||'').trim(); if(!name) continue;
      arr.push(name);
      var email=String(m[i][2]||'').trim(); if(email) emailOf[name]=email;
    }
  }
  read('강사',teachers); read('조교',assistants);
  return {emailOf:emailOf, teachers:teachers, assistants:assistants};
}
function getEventsArr(){
  var ev=rowsOf('일정'), out=[];
  for(var i=1;i<ev.length;i++){ var r=ev[i]; if(!r[0]) continue;
    out.push({title:String(r[0]).trim(), due:ymd(r[1]), cat:String(r[2]||'').trim(),
      owner:String(r[3]||'전체').trim(), status:String(r[4]||'todo').trim().toLowerCase(),
      repeat:String(r[5]||'없음').trim(), memo:String(r[6]||'').trim()});
  }
  return out;
}
function getChecksMap(){
  var ck=rowsOf('체크'), checks={};
  for(var i=1;i<ck.length;i++){
    var title=String(ck[i][0]||'').trim(), due=ymd(ck[i][1]), name=String(ck[i][2]||'').trim();
    if(!title||!name) continue;
    var v=ck[i][3], done=(v===true||v===1||String(v).toUpperCase()==='TRUE'||v==='완료'||v==='O');
    (checks[title+'||'+due]=checks[title+'||'+due]||{})[name]=done;
  }
  return checks;
}
function isDone(checks,title,due,name){ var c=checks[title+'||'+due]; return !!(c&&c[name]); }
function assigneesOf(owner, roster){
  var s=String(owner||'전체').trim();
  if(s===''||s==='전체') return roster.teachers.slice();
  if(s==='조교전체') return roster.assistants.slice();
  return s.split(/[,·、\/]+/).map(function(x){return x.trim();}).filter(function(x){return x;});
}
function occurrences(dueStr, repeat, start, end){
  var res=[]; if(!dueStr) return res;
  var cur=parseYmd(dueStr);
  if(repeat==='매주'||repeat==='매월'){
    while(cur>start){ repeat==='매주'?cur.setDate(cur.getDate()-7):cur.setMonth(cur.getMonth()-1); }
    while(cur<start){ repeat==='매주'?cur.setDate(cur.getDate()+7):cur.setMonth(cur.getMonth()+1); }
    while(cur<=end){ res.push(new Date(cur)); repeat==='매주'?cur.setDate(cur.getDate()+7):cur.setMonth(cur.getMonth()+1); }
  } else if(cur>=start && cur<=end){ res.push(cur); }
  return res;
}
function ddInfo(due, today){
  var d=ddays(due,today);
  var lab = d<0 ? ('지남 '+(-d)+'일') : d===0 ? '오늘 마감' : ('D-'+d);
  var color = d<0 ? '#b3261e' : d===0 ? '#b3261e' : d<=2 ? '#9a6b00' : '#1f6f54';
  return {d:d, lab:lab, color:color};
}
function fmtDue(due){ return Utilities.formatDate(due, TZ, 'M/d'); }

function rowHTML(it, today, extra){
  var x=ddInfo(it.due, today);
  return '<tr>'
    +'<td style="padding:9px 10px;border-bottom:1px solid #eee;white-space:nowrap;vertical-align:top">'
      +'<span style="display:inline-block;font-weight:700;font-size:12px;color:'+x.color+';background:'+x.color+'14;border:1px solid '+x.color+'40;border-radius:7px;padding:3px 8px">'+x.lab+'</span>'
      +'<div style="font-size:11px;color:#888;margin-top:3px">'+fmtDue(it.due)+'</div>'
    +'</td>'
    +'<td style="padding:9px 10px;border-bottom:1px solid #eee;vertical-align:top">'
      +'<div style="font-weight:600;font-size:14px;color:#222">'+esc(it.title)+'</div>'
      +'<div style="font-size:12px;color:#777;margin-top:2px">'+esc(it.cat)+(it.memo?' · '+esc(it.memo):'')+(extra?' · '+extra:'')+'</div>'
    +'</td></tr>';
}

function sendReminders(){
  var roster=getRoster(), events=getEventsArr(), checks=getChecksMap();
  var today=midnight(new Date());
  var end=new Date(today); end.setDate(end.getDate()+WIN_AHEAD);
  var startRep=new Date(today); startRep.setDate(startRep.getDate()-WIN_PAST_REPEAT);
  var startOnce=new Date(today); startOnce.setDate(startOnce.getDate()-WIN_PAST_ONCE);

  var items=[];
  events.forEach(function(ev){
    if(!ev.due || ev.status==='done') return;
    var start=(ev.repeat==='매주'||ev.repeat==='매월')?startRep:startOnce;
    var asg=assigneesOf(ev.owner, roster);
    occurrences(ev.due, ev.repeat, start, end).forEach(function(d){
      items.push({title:ev.title, due:d, dueStr:ymd(d), cat:ev.cat, memo:ev.memo, assignees:asg});
    });
  });

  var stamp=Utilities.formatDate(new Date(), TZ, 'M/d HH시');
  var names=roster.teachers.concat(roster.assistants), sent=0;
  names.forEach(function(name){
    var email=roster.emailOf[name]; if(!email) return;
    var mine=items.filter(function(it){ return it.assignees.indexOf(name)>=0 && !isDone(checks,it.title,it.dueStr,name); })
                  .sort(function(a,b){return a.due-b.due;});
    var digestRows = (name===OWNER) ? buildDigest(items, checks, today) : [];
    if(name!==OWNER && mine.length===0) return;
    if(name===OWNER && mine.length===0 && digestRows.length===0) return;

    var today0=mine.filter(function(it){return ddays(it.due,today)===0;}).length;
    var over=mine.filter(function(it){return ddays(it.due,today)<0;}).length;
    var cnt='이번 마감 '+mine.length+(today0?(' · 오늘 '+today0):'')+(over?(' · 지남 '+over):'');

    var html=''
      +'<div style="font-family:\'Apple SD Gothic Neo\',\'Malgun Gothic\',sans-serif;max-width:600px;margin:0 auto;color:#222">'
      +'<div style="background:#1f6f54;color:#fff;padding:16px 18px;border-radius:12px 12px 0 0">'
        +'<div style="font-size:12px;letter-spacing:2px;opacity:.85">[오름] 국어학원 · 마감 리마인드</div>'
        +'<div style="font-size:18px;font-weight:700;margin-top:3px">'+esc(name)+' · '+esc(cnt)+'</div>'
        +'<div style="font-size:11px;opacity:.8;margin-top:2px">'+stamp+' 기준</div>'
      +'</div>'
      +'<div style="border:1px solid #e6e6e6;border-top:none;border-radius:0 0 12px 12px;padding:14px 14px 18px">';

    html+='<div style="font-size:13px;font-weight:700;color:#1f6f54;margin:4px 2px 8px">내 마감</div>';
    if(mine.length){
      html+='<table style="width:100%;border-collapse:collapse">';
      mine.forEach(function(it){ html+=rowHTML(it,today,''); });
      html+='</table>';
    } else {
      html+='<div style="font-size:13px;color:#888;padding:6px 2px">처리할 본인 마감이 없습니다. 수고하셨어요.</div>';
    }

    if(name===OWNER && digestRows.length){
      html+='<div style="font-size:13px;font-weight:700;color:#b3261e;margin:18px 2px 8px">전체 미이행 요약 (원장용)</div>';
      html+='<table style="width:100%;border-collapse:collapse">';
      digestRows.forEach(function(r){ html+=rowHTML(r.it,today,'완료 '+r.done+'/'+r.total+' · 미완 '+esc(r.missing)); });
      html+='</table>';
    }

    html+='<div style="text-align:center;margin-top:16px"><a href="https://oreum1222.github.io/calendar/" style="display:inline-block;background:#1f6f54;color:#fff;text-decoration:none;font-weight:700;font-size:13px;padding:10px 18px;border-radius:9px">캘린더 열기</a></div>';
    html+='<div style="font-size:11px;color:#aaa;text-align:center;margin-top:12px">처리한 항목은 캘린더에서 본인 PIN으로 체크하면 다음 메일에서 빠집니다.</div>';
    html+='</div></div>';

    MailApp.sendEmail({to:email, subject:'[오름] 마감 리마인드 · '+cnt+' ('+stamp+')', htmlBody:html});
    sent++;
  });
  Logger.log('reminders sent: '+sent);
  return sent;
}

/* 원장용 전체 미이행: 담당 중 미완자가 있는 항목 */
function buildDigest(items, checks, today){
  var rows=[];
  items.forEach(function(it){
    var total=it.assignees.length, done=0, missing=[];
    it.assignees.forEach(function(n){ if(isDone(checks,it.title,it.dueStr,n)) done++; else missing.push(n); });
    if(missing.length>0) rows.push({it:it, total:total, done:done, missing:missing.join(', ')});
  });
  rows.sort(function(a,b){return a.it.due-b.it.due;});
  return rows;
}

/* 트리거 설치 — 매일 14시·21시 (1회만 실행) */
function setupTriggers(){
  ScriptApp.getProjectTriggers().forEach(function(t){ if(t.getHandlerFunction()==='sendReminders') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendReminders').timeBased().atHour(14).nearMinute(0).everyDays(1).create();
  ScriptApp.newTrigger('sendReminders').timeBased().atHour(21).nearMinute(0).everyDays(1).create();
  Logger.log('triggers set: 매일 14시·21시');
}
