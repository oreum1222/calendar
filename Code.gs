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
  var emailOf={}, phoneOf={}, teachers=[], assistants=[];
  function read(sheet, arr){
    var m=rowsOf(sheet);
    for(var i=1;i<m.length;i++){
      var name=String(m[i][0]||'').trim(); if(!name) continue;
      arr.push(name);
      var email=String(m[i][2]||'').trim(); if(email) emailOf[name]=email;
      var phone=String(m[i][3]||'').trim(); if(phone) phoneOf[name]=phone;   // D열: 전화번호
    }
  }
  read('강사',teachers); read('조교',assistants);
  return {emailOf:emailOf, phoneOf:phoneOf, teachers:teachers, assistants:assistants};
}

/* 코드로 추가/수정하는 일정 — index.html CONFIG.extraEvents / overrides 와 동일하게 유지(메일·문자도 이걸 인식) */
var EXTRA_EVENTS = [
  ["한티메가 교재 1차 마감","2026-06-15","교재","가경T, 조교전체","done","없음","중요"],
  ["CSM 가천대 시즌1 강의 촬영","2026-06-20","출제","가경T, 다미T, 원철T","done","없음","촬영 6/14~6/20"],
  ["CSM 가천대 시즌2 출제 마감","2026-06-20","출제","다미T, 원철T","done","없음",""],
  ["CSM 가천대 시즌3 출제 마감","2026-06-30","출제","다미T, 원철T","done","없음",""],
  ["산본고 여름방학","2026-07-10","학사","전체","doing","없음","7/10(금)~8/21(금)"],
  ["군중고 여름방학","2026-07-20","학사","전체","doing","없음","7/20(월)~8/5(수)"],
  ["용호고 여름방학","2026-07-21","학사","전체","doing","없음","7/21(화)~8/6(목)"],
  ["수리고 여름방학","2026-07-20","학사","전체","doing","없음","7/20(월)~8/11(화) · 개학 8/12(수)"],
  ["홍진고 여름방학","2026-07-20","학사","전체","doing","없음","7/20(월)~8/9(일) · 개학 8/10(월)"],
  ["군포고 여름방학","2026-07-16","학사","전체","doing","없음","방학식 7/16(목) · 개학 미정"],
  ["새이솔 문학 — 작품선정·발췌·보기·유형","2026-06-24","출제","가경T","todo","없음","수 08:00 마감"],
  ["새이솔 문학 — 초고 출제","2026-07-01","출제","가경T","todo","없음","수 08:00 마감"],
  ["산본고1(토·일반) 시험 분석글 블로그 업로드","2026-06-26","홍보","다미T","todo","없음","!업로드 6/26 22시 ~ 6/27 04시"],
  ["산본고1(목·금반) 시험 분석글 블로그 업로드","2026-06-26","홍보","은규T","todo","없음","!업로드 6/26 22시 ~ 6/27 04시"],
  ["산본고2 시험 분석글 블로그 업로드","2026-06-29","홍보","은규T","todo","없음","!업로드 6/29 22시 ~ 6/30 04시"],
  ["용호고1 시험 분석글 블로그 업로드","2026-06-29","홍보","다미T","todo","없음","!업로드 6/29 22시 ~ 6/30 04시"],
  ["수원외고2·용호고2 시험 분석글 블로그 업로드","2026-06-30","홍보","은규T","todo","없음","!업로드 6/30 22시 ~ 7/1 04시"],
  ["중등 시험 분석글 블로그 업로드","2026-06-30","홍보","해솔T","todo","없음","!업로드 6/30 22시 ~ 7/1 04시"],
  ["흥진고1·군포고2·흥진고2 시험 분석글 블로그 업로드","2026-07-01","홍보","은규T","todo","없음","!업로드 7/1 22시 ~ 7/2 04시"],
  ["수리고(일요일반) 시험 분석글 블로그 업로드","2026-07-01","홍보","지연T","todo","없음","!업로드 7/1 22시 ~ 7/2 04시"],
  ["중등 시험 분석글 블로그 업로드","2026-07-01","홍보","해솔T","todo","없음","!업로드 7/1 22시 ~ 7/2 04시"],
  ["중등 시험 분석글 블로그 업로드","2026-07-02","홍보","해솔T","todo","없음","!업로드 7/2 22시 ~ 7/3 04시"],
  ["군포고1 시험 분석글 블로그 업로드","2026-07-02","홍보","은규T","todo","없음","!업로드 7/2 22시 ~ 7/3 04시"],
  ["수리고 시험 분석글 블로그 업로드","2026-07-05","홍보","지연T","todo","없음","!업로드 7/5 22시 ~ 7/6 04시"],
  ["수리고(일요일반) 시험 분석글 블로그 업로드","2026-07-06","홍보","지연T","todo","없음","!업로드 7/6 22시 ~ 7/7 04시"],
  ["군포중앙고2 시험 분석글 블로그 업로드","2026-07-07","홍보","은규T","todo","없음","!업로드 7/7 22시 ~ 7/8 04시"],
  ["군포중앙고1 시험 분석글 블로그 업로드","2026-07-07","홍보","다미T","todo","없음","!업로드 7/7 22시 ~ 7/8 04시"],
  ["[족보] 상반기 받은 자료 전원 업로드","2026-07-07","행정","전체","todo","없음","!7/7까지 원드라이브 [족보] 폴더에 본인 상반기 자료 모두 업로드"],
  ["2차고사 경향 입력(상담용 시트)","2026-07-05","행정","가경T, 은규T, 지연T, 다미T, 해솔T","todo","없음","!7/5 23:59 마감 · 원철T 제외 전원"],
  ["기말고사 학부모 상담 주간","2026-07-12","행정","은규T, 지연T, 다미T, 해솔T","todo","없음","~7/12까지 상담 주간 · 가경T·원철T 제외 전원"]
];
var OVERRIDES = { "수능특강 영상 촬영": { owner:"은규T, 가경T, 다미T" } };
var HIDE_TITLES = ["모고 진도(주 2회)"];
var INFO_CATS = ["학사"];   // 정보성 구분 — 리마인더(문자·메일)에서 제외(캘린더 화면엔 표시)

function getEventsArr(){
  var ev=rowsOf('일정'), out=[], seen={};
  for(var i=1;i<ev.length;i++){ var r=ev[i]; if(!r[0]) continue;
    var title=String(r[0]).trim();
    if(HIDE_TITLES.indexOf(title)>=0) continue;
    var o={title:title, due:ymd(r[1]), cat:String(r[2]||'').trim(),
      owner:String(r[3]||'전체').trim(), status:String(r[4]||'todo').trim().toLowerCase(),
      repeat:String(r[5]||'없음').trim(), memo:String(r[6]||'').trim()};
    if(INFO_CATS.indexOf(o.cat)>=0) continue;   // 학사 등 정보성은 리마인더 제외
    var ov=OVERRIDES[title]; if(ov){ if(ov.owner!==undefined)o.owner=ov.owner; if(ov.status!==undefined)o.status=ov.status; }
    out.push(o); seen[title+'||'+o.due]=true;
  }
  EXTRA_EVENTS.forEach(function(r){
    var o={title:String(r[0]).trim(), due:String(r[1]).trim(), cat:String(r[2]||'').trim(),
      owner:String(r[3]||'전체').trim(), status:String(r[4]||'todo').trim().toLowerCase(),
      repeat:String(r[5]||'없음').trim(), memo:String(r[6]||'').trim()};
    if(INFO_CATS.indexOf(o.cat)>=0) return;   // 학사 등 정보성은 리마인더·알림 제외(코드 일정도 동일)
    if(!seen[o.title+'||'+o.due]) out.push(o);
  });
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
  var out=[];
  s.split(/[,·、\/]+/).map(function(x){return x.trim();}).filter(function(x){return x;}).forEach(function(tok){
    if(tok==='전체') out=out.concat(roster.teachers);
    else if(tok==='조교전체') out=out.concat(roster.assistants);
    else out.push(tok);
  });
  return out.filter(function(v,i){return out.indexOf(v)===i;});
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

    // 긴급 = 미체크 + 내일까지(지남·오늘·D-1) → 독촉 대상
    var urgent=mine.filter(function(it){return ddays(it.due,today)<=1;});
    var rest=mine.filter(function(it){return ddays(it.due,today)>1;});
    var cnt='마감 '+mine.length+'건'+(urgent.length?(' · 임박 '+urgent.length):'');

    var html=''
      +'<div style="font-family:\'Apple SD Gothic Neo\',\'Malgun Gothic\',sans-serif;max-width:600px;margin:0 auto;color:#222">'
      +'<div style="background:#1f6f54;color:#fff;padding:16px 18px;border-radius:12px 12px 0 0">'
        +'<div style="font-size:12px;letter-spacing:2px;opacity:.85">[오름] 국어학원 · 마감 리마인드</div>'
        +'<div style="font-size:18px;font-weight:700;margin-top:3px">'+esc(name)+' · '+esc(cnt)+'</div>'
        +'<div style="font-size:11px;opacity:.8;margin-top:2px">'+stamp+' 기준</div>'
      +'</div>'
      +'<div style="border:1px solid #e6e6e6;border-top:none;border-radius:0 0 12px 12px;padding:14px 14px 18px">';

    if(urgent.length){
      html+='<div style="background:#fbe6e3;border:1px solid #e8c9be;border-left:4px solid #b3261e;border-radius:10px;padding:11px 13px;margin:2px 2px 10px">'
        +'<div style="font-weight:700;color:#b3261e;font-size:14px">[긴급] 내일까지 마감인데 아직 미체크 '+urgent.length+'건</div>'
        +'<div style="font-size:12px;color:#7a3b2f;margin-top:2px">오늘 안에 처리하고 캘린더에서 체크해주세요.</div></div>';
      html+='<table style="width:100%;border-collapse:collapse">';
      urgent.forEach(function(it){ html+=rowHTML(it,today,''); });
      html+='</table>';
    }
    if(rest.length){
      html+='<div style="font-size:13px;font-weight:700;color:#1f6f54;margin:'+(urgent.length?'16':'4')+'px 2px 8px">다가오는 마감</div>';
      html+='<table style="width:100%;border-collapse:collapse">';
      rest.forEach(function(it){ html+=rowHTML(it,today,''); });
      html+='</table>';
    }
    if(!mine.length){
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

    var subj = urgent.length
      ? '[오름][긴급] 내일까지 미체크 '+urgent.length+'건 — '+name+' ('+stamp+')'
      : '[오름] 마감 리마인드 · '+cnt+' ('+stamp+')';
    MailApp.sendEmail({to:email, subject:subj, htmlBody:html});
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

/* ════════════════════════════════════════════════════════════
   문자(SMS) 자동발송 — Solapi (마감 미체크자 대상)
   · 비밀값: 프로젝트 설정 → 스크립트 속성 SOLAPI_KEY / SOLAPI_SECRET / SOLAPI_SENDER
   · 명단: 강사/조교 시트 D열(전화번호)
   · 켜기: setupSmsTrigger() 1회 실행(매일 SMS_HOUR시). 테스트: 속성 SMS_DRYRUN=1 후 sendSmsReminders()
   · 인증: HMAC-SHA256 (date+salt 를 Secret으로 서명)  ← 기존 Solapi 방식 그대로
   ════════════════════════════════════════════════════════════ */
var SMS_DDAY_MAX = 0;   // (occurrences 범위용) 0=오늘까지. 실제 발송 필터는 '마감 당일(D-DAY)'만 — 아래 sendSmsReminders 참고.
var SMS_HOUR     = 18;  // 매일 발송 시각(시) — 마감 당일 18시에 1회. setupSmsTrigger로 등록(변경 시 재실행 필요).
var CAL_URL      = 'https://oreum1222.github.io/calendar/';

function solapiAuth_(key, secret){
  var date = new Date().toISOString();
  var salt = Utilities.getUuid().replace(/-/g, '');
  var raw  = Utilities.computeHmacSha256Signature(date + salt, secret);
  var hex  = raw.map(function(b){ var v=(b<0?b+256:b).toString(16); return v.length===1?'0'+v:v; }).join('');
  return 'HMAC-SHA256 apiKey=' + key + ', date=' + date + ', salt=' + salt + ', signature=' + hex;
}
function digits_(s){ return String(s||'').replace(/[^0-9]/g,''); }

/* 인증 확인용: 잔액 조회(200이면 키 정상) */
function checkSolapi(){
  var p=PropertiesService.getScriptProperties();
  var res=UrlFetchApp.fetch('https://api.solapi.com/cash/v1/balance',
    { headers:{ Authorization: solapiAuth_(p.getProperty('SOLAPI_KEY'), p.getProperty('SOLAPI_SECRET')) }, muteHttpExceptions:true });
  Logger.log(res.getResponseCode()+' '+res.getContentText());
}

function sendSmsReminders(){
  var p=PropertiesService.getScriptProperties();
  var key=p.getProperty('SOLAPI_KEY'), secret=p.getProperty('SOLAPI_SECRET'), sender=p.getProperty('SOLAPI_SENDER');
  var dryRun=(p.getProperty('SMS_DRYRUN')==='1') || !(key&&secret&&sender);

  var roster=getRoster(), events=getEventsArr(), checks=getChecksMap();
  var today=midnight(new Date());
  var occEnd=new Date(today); occEnd.setDate(occEnd.getDate()+Math.max(0,SMS_DDAY_MAX));
  var startOnce=new Date(today); startOnce.setDate(startOnce.getDate()-WIN_PAST_ONCE);
  var startRep=new Date(today); startRep.setDate(startRep.getDate()-WIN_PAST_REPEAT);

  var items=[];
  events.forEach(function(ev){
    if(!ev.due || ev.status==='done') return;
    if(ev.repeat==='매주' || ev.repeat==='매월') return;   // 반복 일정은 문자 제외(메일로만 안내)
    var s=startOnce;
    var asg=assigneesOf(ev.owner, roster);
    occurrences(ev.due, ev.repeat, s, occEnd).forEach(function(d){
      items.push({title:ev.title, due:d, dueStr:ymd(d), assignees:asg});
    });
  });

  var names=roster.teachers.concat(roster.assistants), messages=[];
  names.forEach(function(name){
    var phone=roster.phoneOf[name]; if(!phone) return;
    var miss=items.filter(function(it){
      return it.assignees.indexOf(name)>=0 && ddays(it.due,today)===0 && !isDone(checks,it.title,it.dueStr,name);   // 마감 당일(D-DAY)만 18시 1회
    }).sort(function(a,b){return a.due-b.due;});
    if(!miss.length) return;
    var head=miss.slice(0,3).map(function(it){ return it.title; }).join(', ');
    var more=miss.length>3?(' 외 '+(miss.length-3)+'건'):'';
    var text='[오름] '+name+' 오늘 마감 '+miss.length+'건: '+head+more+'\n오늘 안에 처리 후 캘린더에서 체크 부탁드립니다 '+CAL_URL;
    messages.push({to:digits_(phone), name:name, text:text, count:miss.length});
  });

  if(!messages.length){ Logger.log('SMS: 발송 대상 없음'); return 0; }

  if(dryRun){
    messages.forEach(function(m){ logSms_(m.name, m.to, m.count, '테스트(dryRun)', '', ''); });
    Logger.log('SMS dryRun '+messages.length+'건 (실제 발송 안 함). 키 없거나 SMS_DRYRUN=1');
    return messages.length;
  }

  var payload={ messages: messages.map(function(m){ return {to:m.to, from:digits_(sender), text:m.text}; }) };
  var res=UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send-many/detail', {
    method:'post', contentType:'application/json',
    headers:{ Authorization: solapiAuth_(key,secret) },
    payload: JSON.stringify(payload), muteHttpExceptions:true
  });
  var body={}; try{ body=JSON.parse(res.getContentText()||'{}'); }catch(_){}
  var failed={}; (body.failedMessageList||[]).forEach(function(f){ failed[digits_(f.to)]=f.statusMessage||'실패'; });
  var ok=0; messages.forEach(function(m){ var er=failed[m.to]; logSms_(m.name, m.to, m.count, er?'실패':'발송', (body.groupInfo||{}).groupId||'', er||''); if(!er)ok++; });
  Logger.log('SMS sent '+ok+'/'+messages.length+' (HTTP '+res.getResponseCode()+')');
  return ok;
}

function logSms_(name, to, count, result, groupId, err){
  var sh=SS.getSheetByName('문자로그');
  if(!sh){ sh=SS.insertSheet('문자로그'); sh.appendRow(['시각','이름','번호','건수','결과','groupId','오류']); }
  var masked = to.length>=7 ? (to.slice(0,3)+'****'+to.slice(-4)) : to;
  sh.appendRow([Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd HH:mm'), name, masked, count, result, groupId, err]);
}

function setupSmsTrigger(){
  ScriptApp.getProjectTriggers().forEach(function(t){ if(t.getHandlerFunction()==='sendSmsReminders') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendSmsReminders').timeBased().atHour(SMS_HOUR).nearMinute(0).everyDays(1).create();
  Logger.log('SMS 트리거: 매일 '+SMS_HOUR+'시');
}

/* ════════════════════════════════════════════════════════════
   새 일정 등록 알림 — 시트·코드 어느 쪽으로 새 일정을 올리든 담당자에게 메일 1회
   · setupNewEventNotify() 1회 실행 → 기존 일정은 '이미 통보'로 시드(소급 발송 방지) + 30분마다 검사
   · 이후 새로 추가된 일정만 담당자(이메일 등록자)에게 자동 메일. 중복·재발송 없음.
   · 학사(안내)·완료·지난 일정은 메일 생략(통보 기록만).
   ════════════════════════════════════════════════════════════ */
function notifyNewEvents(){
  var p=PropertiesService.getScriptProperties();
  var seen={}; try{ seen=JSON.parse(p.getProperty('NOTIFIED_KEYS')||'{}'); }catch(e){ seen={}; }
  var roster=getRoster(), events=getEventsArr(), today=midnight(new Date());
  var byPerson={}, changed=false;
  events.forEach(function(ev){
    if(!ev.due) return;
    var key=ev.title+'||'+ev.due;
    if(seen[key]) return;
    seen[key]=1; changed=true;                              // 새 일정 → 통보 기록
    if(ev.status==='done') return;                          // 완료 일정은 메일 생략
    if(ddays(parseYmd(ev.due), today) < 0) return;          // 지난 일정은 메일 생략
    assigneesOf(ev.owner, roster).forEach(function(name){
      (byPerson[name]=byPerson[name]||[]).push(ev);
    });
  });
  if(changed) p.setProperty('NOTIFIED_KEYS', JSON.stringify(seen));

  var stamp=Utilities.formatDate(new Date(), TZ, 'M/d HH:mm'), sent=0;
  Object.keys(byPerson).forEach(function(name){
    var email=roster.emailOf[name]; if(!email) return;
    var evs=byPerson[name].sort(function(a,b){ return parseYmd(a.due)-parseYmd(b.due); });
    var rows=evs.map(function(ev){
      var d=parseYmd(ev.due), x=ddInfo(d, today);
      return '<tr>'
        +'<td style="padding:9px 10px;border-bottom:1px solid #eee;white-space:nowrap;vertical-align:top">'
          +'<span style="display:inline-block;font-weight:700;font-size:12px;color:'+x.color+';background:'+x.color+'14;border:1px solid '+x.color+'40;border-radius:7px;padding:3px 8px">'+x.lab+'</span>'
          +'<div style="font-size:11px;color:#888;margin-top:3px">'+fmtDue(d)+'</div>'
        +'</td>'
        +'<td style="padding:9px 10px;border-bottom:1px solid #eee;vertical-align:top">'
          +'<div style="font-weight:600;font-size:14px;color:#222">'+esc(ev.title)+'</div>'
          +'<div style="font-size:12px;color:#777;margin-top:2px">'+esc(ev.cat)+(ev.memo?' · '+esc(ev.memo.replace(/^!+\s*/,'')):'')+'</div>'
        +'</td></tr>';
    }).join('');
    var html=''
      +'<div style="font-family:\'Apple SD Gothic Neo\',\'Malgun Gothic\',sans-serif;max-width:600px;margin:0 auto;color:#222">'
      +'<div style="background:#1f6f54;color:#fff;padding:16px 18px;border-radius:12px 12px 0 0">'
        +'<div style="font-size:12px;letter-spacing:2px;opacity:.85">[오름] 국어학원 · 새 일정 등록</div>'
        +'<div style="font-size:18px;font-weight:700;margin-top:3px">'+esc(name)+' · 새 담당 일정 '+evs.length+'건</div>'
        +'<div style="font-size:11px;opacity:.8;margin-top:2px">'+stamp+' 등록 · 캘린더에 추가되었습니다</div>'
      +'</div>'
      +'<div style="border:1px solid #e6e6e6;border-top:none;border-radius:0 0 12px 12px;padding:14px 14px 18px">'
      +'<table style="width:100%;border-collapse:collapse">'+rows+'</table>'
      +'<div style="text-align:center;margin-top:16px"><a href="'+CAL_URL+'" style="display:inline-block;background:#1f6f54;color:#fff;text-decoration:none;font-weight:700;font-size:13px;padding:10px 18px;border-radius:9px">캘린더에서 확인</a></div>'
      +'</div></div>';
    MailApp.sendEmail({to:email, subject:'[오름] 새 일정 '+evs.length+'건 등록 — '+name+' ('+stamp+')', htmlBody:html});
    sent++;
  });
  Logger.log('새 일정 알림: '+sent+'명에게 발송');
  return sent;
}

function setupNewEventNotify(){
  var p=PropertiesService.getScriptProperties(), seen={};
  getEventsArr().forEach(function(ev){ if(ev.due) seen[ev.title+'||'+ev.due]=1; });   // 기존 일정 시드(소급 발송 방지)
  p.setProperty('NOTIFIED_KEYS', JSON.stringify(seen));
  ScriptApp.getProjectTriggers().forEach(function(t){ if(t.getHandlerFunction()==='notifyNewEvents') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('notifyNewEvents').timeBased().everyMinutes(30).create();
  Logger.log('새 일정 알림 트리거: 30분마다 · 기존 '+Object.keys(seen).length+'건 시드(소급 발송 안 함)');
}
