const $ = id => document.getElementById(id);
const DEFAULTS = {weeklyHours:41, annualVacation:30, carryVacation:0, regularBreak:30, eveningBreak:15, eveningFrom:"16:30"};
let settings = {...DEFAULTS, ...(JSON.parse(localStorage.getItem("az-settings") || "null") || {})};
let records = JSON.parse(localStorage.getItem("az-records-v2") || "null");
let importedFiles = JSON.parse(localStorage.getItem("az-imported-files-v1") || "[]");
let legacyRecords = null;
if (!records) {
  legacyRecords = JSON.parse(localStorage.getItem("az-records") || "[]");
  records = legacyRecords.map(r => ({date:r.date, blocks:(r.start!=null&&r.end!=null)?[{start:r.start,end:r.end,reason:r.reason||""}]:[], dayReason:r.start==null&&r.end==null?r.reason||"":"", source:r.source||""}));
  localStorage.setItem("az-records-v2", JSON.stringify(records));
}

const ZERO_TARGET_ABSENCE = ["feiertag","dienstbefreiung","freistellung v. dst", "erlaubte abwesenheit"];
const NORMAL_WORK_WORDS = ["grundbetrieb"];
const COMP_WORD = "ausgleich mehrarbeit";
const VACATION_WORD = "erholungsurlaub";

function save(){
  localStorage.setItem("az-records-v2", JSON.stringify(records));
  localStorage.setItem("az-settings", JSON.stringify(settings));
  localStorage.setItem("az-imported-files-v1", JSON.stringify(importedFiles));
}
function timeToMin(v){
  if(v==null || v==="") return null;
  if(v instanceof Date) return (v.getHours()*60+v.getMinutes())%1440;
  if(typeof v === "number") return Math.round((v%1)*1440)%1440;
  const s=String(v).trim();
  const m=s.match(/^(\d{1,2})\s*[:.]\s*(\d{2})/);
  if(m) return Number(m[1])*60+Number(m[2]);
  return null;
}
function fmtTime(min){
  if(min==null) return "—";
  const m=((min%1440)+1440)%1440;
  return `${String(Math.floor(m/60)).padStart(2,"0")}:${String(m%60).padStart(2,"0")}`;
}
function fmtMin(min){
  const sign=min<0?"-":"", a=Math.abs(Math.round(min));
  return `${sign}${String(Math.floor(a/60)).padStart(2,"0")}:${String(a%60).padStart(2,"0")} h`;
}
function parseDate(v){
  if(v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(),v.getMonth(),v.getDate());
  if(typeof v === "number") { const d=XLSX.SSF.parse_date_code(v); return new Date(d.y,d.m-1,d.d); }
  if(typeof v === "string"){
    const s=v.trim();
    let m=s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if(m) return new Date(+m[3],+m[2]-1,+m[1]);
    const d=new Date(s); if(!isNaN(d)) return new Date(d.getFullYear(),d.getMonth(),d.getDate());
  }
  return null;
}
function dateKey(d){return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`}
function dateObj(key){const [y,m,d]=key.split("-").map(Number); return new Date(y,m-1,d)}
function dayName(d){return d.toLocaleDateString("de-DE",{weekday:"short"}).replace(".","")}
function dailyTarget(d){const n=d.getDay(); if(n>=1&&n<=4)return 540; if(n===5)return 300; return 0;}
function normalizeReason(r){return String(r||"").trim().toLowerCase();}
function reasonType(reason){
  const r=normalizeReason(reason);
  if(r.includes(VACATION_WORD)) return "vacation";
  if(r.includes(COMP_WORD)) return "comp";
  if(ZERO_TARGET_ABSENCE.some(x=>r.includes(x))) return "no-target";
  return "work";
}
function isKnownBookingReason(reason){
  const r=normalizeReason(reason); if(!r) return false;
  return r.includes("grundbetrieb") || r.includes("spähtrupp") || r.includes("schübz") || r.includes("bahn") || r.includes("ovwa") || r.includes("ovd") || r.includes("uvd") || r.includes("wache") || r.includes("atb") || r.includes("lehrgang") || r.includes("bereitschaft") || r.includes("rufbereitschaft") || r.includes("einsatz") || r.includes("ausnahmetatbestand") || r.includes("offzfü") || r.includes("siehe anlage");
}
function duration(start,end){
  if(start==null||end==null)return 0;
  let d=end-start;
  if(d<0)d+=1440;
  // 00:00 -> 00:00 is a full 24h shift in this spreadsheet context.
  if(d===0 && start===0 && end===0)return 1440;
  return d;
}
function eveningBreakFor(block){
  if(block.start==null||block.end==null)return 0;
  const threshold=timeToMin(settings.eveningFrom);
  return (block.end===0 || block.end>threshold) ? settings.eveningBreak : 0;
}
function regularBreakFor(block){
  if(block.start==null||block.end==null)return 0;
  const endAbs=block.end===0?1440:block.end;
  // Only deduct the regular 30-minute break when a single continuous block
  // spans the normal midday working period. Separate blocks already contain
  // an explicit gap and must not receive a second artificial break.
  if(block.start<=420 && endAbs>=960) return settings.regularBreak;
  return 0;
}
function calculateDay(day){
  const d=dateObj(day.date);
  const blocks=(day.blocks||[]).filter(b=>b.start!=null&&b.end!=null);
  const reasons=[...(day.blocks||[]).map(b=>b.reason), day.dayReason||""].filter(Boolean);
  const types=reasons.map(reasonType);
  // Vollständig leerer Tag: ausdrücklich ±0:00, kein automatisches Tages-Soll.
  if(blocks.length===0 && reasons.length===0){
    return {...day,date:day.date,weekday:dayName(d),target:0,raw:0,worked:0,balance:0,regularBreak:0,eveningBreak:0,status:"leer",reasons:[]};
  }
  const target=dailyTarget(d);
  const hasVacation=types.includes("vacation");
  const hasNoTarget=types.includes("no-target");
  let raw=0, regularBreak=0, eveningBreak=0;
  for(const b of blocks){raw+=duration(b.start,b.end); regularBreak+=regularBreakFor(b); eveningBreak+=eveningBreakFor(b);}
  // If there are multiple blocks, each block's own evening break is counted.
  // A regular break is only charged inside a continuous block.
  let worked=Math.max(0,raw-regularBreak-eveningBreak);
  let effectiveTarget=target, balance=worked;
  let status="ok";
  // Sondergrund mit Zeiten: Mo-Fr immer 9:00 Soll, Wochenende 0:00 Soll.
  const hasSpecialTimedReason = blocks.length>0 && reasons.some(r=>{
    const rr=normalizeReason(r);
    return rr && !NORMAL_WORK_WORDS.some(w=>rr.includes(w)) && !rr.includes(VACATION_WORD) && !rr.includes("feiertag") && !rr.includes(COMP_WORD);
  });
  if(hasSpecialTimedReason) effectiveTarget = d.getDay()>=1 && d.getDay()<=5 ? 540 : 0;
  balance=worked-effectiveTarget;
  if(hasVacation && blocks.length===0){worked=target; effectiveTarget=target; balance=0; status="urlaub";}
  else if(types.includes("comp") && blocks.length===0){
    // A compensatory day is free of duty, but it consumes the normal daily
    // overtime credit: 0 worked against the day's planned target.
    worked=0; effectiveTarget=target; balance=-target; status="ausgleich";
  }
  else if(hasNoTarget && blocks.length===0){worked=0; effectiveTarget=0; balance=0; status="feiertag/abwesenheit";}
  else if(blocks.length===0 && reasons.length>0 && !hasVacation && !hasNoTarget){
    effectiveTarget=0; balance=0; status="prüfen";
  }
  const unknownReason=reasons.find(r=>r && !isKnownBookingReason(r) && ![VACATION_WORD,"feiertag"].some(x=>normalizeReason(r).includes(x)) && !normalizeReason(r).includes("ausgleich mehrarbeit"));
  if(unknownReason && blocks.length===0) status="prüfen";
  return {...day, date:day.date, weekday:dayName(d), target:effectiveTarget, raw, worked, balance, regularBreak, eveningBreak, status, reasons};
}
function calculateAll(){return records.map(calculateDay).sort((a,b)=>a.date.localeCompare(b.date));}
function monthKey(date){return date.slice(0,7)}
function monthLabel(m){const [y,mo]=m.split("-");return new Date(+y,+mo-1,1).toLocaleDateString("de-DE",{month:"long",year:"numeric"})}
function uniqueMonths(){return [...new Set(records.map(r=>monthKey(r.date)))].sort()}

function render(){
  const data=calculateAll();
  const worked=data.reduce((a,r)=>a+r.worked,0), target=data.reduce((a,r)=>a+r.target,0), balance=data.reduce((a,r)=>a+r.balance,0);
  const vacationUsed=data.filter(r=>r.status==="urlaub").length;
  const vacationLeft=settings.annualVacation+settings.carryVacation-vacationUsed;
  $("totalBalance").textContent=(balance>=0?"+":"")+fmtMin(balance);
  $("worked").textContent=fmtMin(worked); $("target").textContent=fmtMin(target);
  $("vacationUsed").textContent=`${vacationUsed} ${vacationUsed===1?"Tag":"Tage"}`;
  $("vacationLeft").textContent=`${vacationLeft} Tage`;
  const months=uniqueMonths();
  $("periodLabel").textContent=months.length?`${months.length} Monat${months.length===1?"":"e"} im Konto`:"Noch keine Daten";
  $("rowCount").textContent=`${data.length} Tage`; $("monthCount").textContent=`${months.length}`;
  $("fileStatus").textContent=data.length?"Gespeichert":"Bereit";
  const review=data.filter(r=>r.status==="prüfen").length;
  $("reviewHint").textContent=review?`⚠ ${review} Tag${review===1?"":"e"} prüfen`:`✓ Keine offenen Prüfhinweise`;

  const byMonth={}; data.forEach(r=>(byMonth[monthKey(r.date)]??=[]).push(r));
  $("accountSummary").innerHTML=`
    <div class="summary-row"><span>Gesamte Arbeitszeit</span><strong>${fmtMin(worked)}</strong></div>
    <div class="summary-row"><span>Gesamte Sollzeit</span><strong>${fmtMin(target)}</strong></div>
    <div class="summary-row"><span>Arbeitszeitkonto</span><strong class="${balance>=0?"pos":"neg"}">${balance>=0?"+":""}${fmtMin(balance)}</strong></div>
    <div class="summary-row"><span>Urlaubsanspruch inkl. Übertrag</span><strong>${settings.annualVacation+settings.carryVacation} Tage</strong></div>
    <div class="summary-row"><span>Resturlaub</span><strong>${vacationLeft} Tage</strong></div>`;

  const renderDay=r=>{
    const d=dateObj(r.date);
    const blockText=r.blocks?.length ? r.blocks.map(b=>`${fmtTime(b.start)} – ${fmtTime(b.end)}${b.reason?` · ${b.reason}`:""}`).join("<br>") : (r.dayReason||"keine Buchung");
    const pause=(r.regularBreak+r.eveningBreak)?` · Pausen ${r.regularBreak+r.eveningBreak} min`:"";
    const badge=r.status==="urlaub"?" · Urlaub":r.status==="ausgleich"?" · Ausgleich Mehrarbeit":r.status==="feiertag/abwesenheit"?" · kein Soll":r.status==="prüfen"?" · ⚠ prüfen":"";
    return `<div class="day"><div class="date"><b>${d.toLocaleDateString("de-DE",{day:"2-digit",month:"2-digit"})}</b><small>${r.weekday}</small></div><div><div class="times">${blockText}</div><div class="reason">Soll ${fmtMin(r.target)}${pause}${badge}</div></div><div class="balance ${r.balance>0?"pos":r.balance<0?"neg":"zero"}">${r.balance>0?"+":""}${fmtMin(r.balance)}</div></div>`;
  };
  $("daysPreview").innerHTML=data.length?data.slice().reverse().slice(0,8).map(renderDay).join(""):`<div class="muted" style="padding:18px 0">Noch keine Daten. Importiere eine Excel-Datei.</div>`;

  const selected=window.selectedMonth && byMonth[window.selectedMonth] ? window.selectedMonth : (months[months.length-1]||null);
  window.selectedMonth=selected;
  $("monthTabs").innerHTML=months.length?months.map(m=>`<button class="month-tab ${m===selected?"active":""}" data-month="${m}">${monthLabel(m)}</button>`).join(""):`<div class="months-empty">Noch keine Monatsdaten.</div>`;
  document.querySelectorAll(".month-tab").forEach(b=>b.onclick=()=>{window.selectedMonth=b.dataset.month;renderMonth();});
  renderMonth();
  renderImportedFiles();
}
function escapeHtml(v){return String(v??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;","":"&quot;"}[c]));}
function renderImportedFiles(){
  const el=$("importedFiles"); if(!el)return;
  if(!importedFiles.length){el.innerHTML=`<div class="months-empty">Noch keine importierten Dateien gespeichert.</div>`;return;}
  el.innerHTML=importedFiles.slice().reverse().map(f=>`<div class="file-row"><div><strong>${escapeHtml(f.name)}</strong><small>${f.days} Tage · ${f.blocks} Zeitblöcke${f.importedAt?` · ${new Date(f.importedAt).toLocaleDateString("de-DE")}`:""}</small></div><button class="delete-file" data-source="${escapeHtml(f.source)}">Löschen</button></div>`).join("");
  el.querySelectorAll(".delete-file").forEach(b=>b.onclick=()=>deleteImportedFile(b.dataset.source));
}
function deleteImportedFile(source){
  const f=importedFiles.find(x=>x.source===source); if(!f)return;
  if(!confirm(`Importierte Datei „${f.name}“ und die darin gespeicherten Tage entfernen?`))return;
  records=records.filter(r=>r.source!==source); importedFiles=importedFiles.filter(x=>x.source!==source); save();
  if(window.selectedMonth && !uniqueMonths().includes(window.selectedMonth))window.selectedMonth=null;
  render(); $("importMessage").textContent=`${f.name} wurde entfernt.`;
}
function renderMonth(){
  const data=calculateAll(), m=window.selectedMonth;
  if(!m){$("monthDetail").innerHTML=`<div class="months-empty">Importiere zuerst eine Excel-Datei.</div>`;return;}
  const rows=data.filter(r=>monthKey(r.date)===m), w=rows.reduce((a,r)=>a+r.worked,0), t=rows.reduce((a,r)=>a+r.target,0), b=w-t, v=rows.filter(r=>r.status==="urlaub").length;
  const renderDay=r=>{
    const d=dateObj(r.date);
    const blockText=r.blocks?.length?r.blocks.map(x=>`${fmtTime(x.start)} – ${fmtTime(x.end)}${x.reason?` · ${x.reason}`:""}`).join("<br>"):(r.dayReason||"keine Buchung");
    const badge=r.status==="urlaub"?" · Urlaub":r.status==="ausgleich"?" · Ausgleich":r.status==="feiertag/abwesenheit"?" · kein Soll":r.status==="prüfen"?" · ⚠ prüfen":"";
    return `<div class="day"><div class="date"><b>${d.toLocaleDateString("de-DE",{day:"2-digit",month:"2-digit"})}</b><small>${r.weekday}</small></div><div><div class="times">${blockText}</div><div class="reason">Soll ${fmtMin(r.target)}${(r.regularBreak+r.eveningBreak)?` · Pausen ${r.regularBreak+r.eveningBreak} min`:""}${badge}</div></div><div class="balance ${r.balance>0?"pos":r.balance<0?"neg":"zero"}">${r.balance>0?"+":""}${fmtMin(r.balance)}</div></div>`;
  };
  $("monthDetail").innerHTML=`<div class="month-head"><div class="month-title">${monthLabel(m)}</div><div class="pill">${rows.length} Tage</div></div><div class="month-kpis"><div class="month-kpi"><span>Geleistet</span><strong>${fmtMin(w)}</strong></div><div class="month-kpi"><span>Soll</span><strong>${fmtMin(t)}</strong></div><div class="month-kpi"><span>Saldo</span><strong class="${b>=0?"pos":"neg"}">${b>=0?"+":""}${fmtMin(b)}</strong></div></div><div class="summary-row" style="margin-top:10px"><span>Urlaub</span><strong>${v} Tage</strong></div><div class="month-days">${rows.slice().reverse().map(renderDay).join("")}</div>`;
}
function showView(view){
  document.querySelectorAll(".view").forEach(v=>v.classList.remove("active"));
  $("view-"+view).classList.add("active");
  document.querySelectorAll(".tab").forEach(t=>t.classList.toggle("active",t.dataset.view===view));
  if(view==="months") renderMonth();
  window.scrollTo({top:0,behavior:"smooth"});
}

function parseSheet(ws,fileName){
  const raw=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null});
  const byDate=new Map(); let currentDate=null;
  for(let i=0;i<raw.length;i++){
    const row=raw[i]||[];
    const parsed=parseDate(row[0]);
    if(parsed && parsed.getFullYear()>=2000){currentDate=dateKey(parsed); if(!byDate.has(currentDate))byDate.set(currentDate,{date:currentDate,blocks:[],dayReason:"",source:fileName});}
    if(!currentDate)continue;
    const start=timeToMin(row[1]), end=timeToMin(row[2]), reason=String(row[3]??"").trim();
    const hasTime=start!=null||end!=null;
    // Only treat a row as a booking when it has at least one time. A reason-only
    // row is a day-level booking only if it is a recognized special status.
    if(hasTime){
      byDate.get(currentDate).blocks.push({start,end,reason});
    } else if(reason){
      const rt=reasonType(reason);
      if(rt!=="work" || isKnownBookingReason(reason)) {
        // Known reason-only entries such as Urlaub, Feiertag or Ausgleich belong to the day.
        // Other text like a location note (e.g. "Hammelburg") is ignored.
        if(rt!=="work") byDate.get(currentDate).dayReason=reason;
      }
    }
  }
  return [...byDate.values()].sort((a,b)=>a.date.localeCompare(b.date));
}
async function importFile(file){
  try{
    const buf=await file.arrayBuffer(), wb=XLSX.read(buf,{type:"array",cellDates:true});
    const ws=wb.Sheets["AZ"]||wb.Sheets[wb.SheetNames.find(n=>/^az$/i.test(n))];
    if(!ws)throw new Error("Kein Tabellenblatt „AZ“ gefunden.");
    const incoming=parseSheet(ws,file.name);
    if(!incoming.length)throw new Error("Keine Tagesdaten erkannt.");
    const source=`${file.name}::${incoming[0].date}::${incoming[incoming.length-1].date}`;
    incoming.forEach(r=>r.source=source);
    const incomingKeys=new Set(incoming.map(r=>r.date));
    const keep=records.filter(r=>!incomingKeys.has(r.date));
    records=[...keep,...incoming].sort((a,b)=>a.date.localeCompare(b.date));
    importedFiles=importedFiles.filter(f=>f.source!==source);
    const blockCount=incoming.reduce((n,r)=>n+r.blocks.length,0);
    importedFiles.push({source,name:file.name,days:incoming.length,blocks:blockCount,importedAt:new Date().toISOString()});
    save(); render();
    $("importMessage").textContent=`${file.name}: ${incoming.length} Tage und ${blockCount} Zeitblöcke erkannt. Tage aus dieser Datei wurden sauber neu eingelesen; doppelte alte Interpretationen werden ersetzt.`;
  }catch(e){$("importMessage").textContent="Importfehler: "+e.message}
}
function csvEscape(v){return `"${String(v??"").replaceAll('"','""')}"`}
function exportCSV(){
  const data=calculateAll(), head=["Datum","Wochentag","Zeiträume","Buchungsgründe","Rohzeit_min","Arbeitszeit_min","Sollzeit_min","Saldo_min","Normale_Pause_min","Abendpause_min","Status"];
  const lines=[head.map(csvEscape).join(";"),...data.map(r=>[
    r.date,r.weekday,r.blocks.map(b=>`${fmtTime(b.start)}-${fmtTime(b.end)}`).join(" | "),r.reasons.join(" | "),r.raw,r.worked,r.target,r.balance,r.regularBreak,r.eveningBreak,r.status
  ].map(csvEscape).join(";"))];
  const blob=new Blob(["\ufeff"+lines.join("\n")],{type:"text/csv;charset=utf-8"}),url=URL.createObjectURL(blob),a=document.createElement("a");
  a.href=url;a.download="arbeitszeitkonto-v3.csv";a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
}
$("fileInput").addEventListener("change",e=>{if(e.target.files[0])importFile(e.target.files[0]);e.target.value=""});
$("exportBtn").onclick=exportCSV;
document.querySelectorAll(".tab").forEach(t=>t.onclick=()=>showView(t.dataset.view));
function loadSettingsUI(){$("weeklyHours").value=settings.weeklyHours;$("annualVacation").value=settings.annualVacation;$("carryVacation").value=settings.carryVacation;$("regularBreak").value=settings.regularBreak;$("eveningBreak").value=settings.eveningBreak;$("eveningFrom").value=settings.eveningFrom;}
$("saveSettings").onclick=()=>{settings={weeklyHours:+$("weeklyHours").value,annualVacation:+$("annualVacation").value,carryVacation:+$("carryVacation").value,regularBreak:+$("regularBreak").value,eveningBreak:+$("eveningBreak").value,eveningFrom:$("eveningFrom").value||"16:30"};save();render()};
$("clearBtn").onclick=()=>{if(confirm("Alle lokal gespeicherten Arbeitszeitdaten und Einstellungen löschen?")){records=[];importedFiles=[];localStorage.removeItem("az-records-v2");localStorage.removeItem("az-records");localStorage.removeItem("az-imported-files-v1");render()}};
if("serviceWorker"in navigator)navigator.serviceWorker.register("sw.js").catch(()=>{});
loadSettingsUI();
render();
