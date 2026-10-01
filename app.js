const $=id=>document.getElementById(id);
const DEFAULTS={weeklyHours:41,annualVacation:30,carryVacation:0,regularBreak:30,eveningBreak:15,eveningFrom:"16:30"};
let settings=JSON.parse(localStorage.getItem("az-settings")||"null")||DEFAULTS;
let records=JSON.parse(localStorage.getItem("az-records")||"[]");

function save(){localStorage.setItem("az-records",JSON.stringify(records));localStorage.setItem("az-settings",JSON.stringify(settings))}
function timeToMin(v){
  if(v==null||v==="")return null;
  if(v instanceof Date)return v.getHours()*60+v.getMinutes();
  if(typeof v==="number")return Math.round((v%1)*1440)%1440;
  const m=String(v).trim().match(/(\d{1,2}):(\d{2})/);return m?Number(m[1])*60+Number(m[2]):null;
}
function duration(start,end){
  if(start==null||end==null)return 0;
  if(start===0&&end===0)return 1440;
  let d=end-start;if(d<0)d+=1440;return d;
}
function fmtMin(min){
  const sign=min<0?"-":"",a=Math.abs(Math.round(min));
  return `${sign}${String(Math.floor(a/60)).padStart(2,"0")}:${String(a%60).padStart(2,"0")} h`;
}
function parseDate(v){
  if(v instanceof Date)return v;
  if(typeof v==="number"){const d=XLSX.SSF.parse_date_code(v);return new Date(d.y,d.m-1,d.d)}
  const d=new Date(v);return isNaN(d)?null:d;
}
function dateKey(d){return d.toISOString().slice(0,10)}
function isWorkday(d){return d.getDay()>=1&&d.getDay()<=5}
function dailyTarget(d){
  const n=d.getDay();
  if(n>=1&&n<=4)return 540; // 09:00
  if(n===5)return 300;      // 05:00
  return 0;
}
function reasonType(reason){
  const r=(reason||"").toLowerCase();
  if(r.includes("feiertag"))return "holiday";
  if(r.includes("erholungsurlaub"))return "vacation";
  if(r.includes("ausgleich mehrarbeit"))return "comp";
  return "work";
}
function calculateRow(x){
  const d=new Date(x.date+"T12:00:00"),type=reasonType(x.reason),target=dailyTarget(d);
  let worked=0, regularBreak=0, eveningBreak=0;
  if(type==="vacation"){worked=target}
  else if(type==="holiday"||type==="comp"){worked=0}
  else if(x.start!=null&&x.end!=null){
    const raw=duration(x.start,x.end);
    worked=raw;
    // For normal Mon–Thu duty covering the regular schedule, the 30-min break is unpaid.
    if(d.getDay()>=1&&d.getDay()<=4 && x.start<=420 && (x.end>960||x.end===0)) regularBreak=settings.regularBreak;
    const eveningFrom=timeToMin(settings.eveningFrom);
    if(x.end===0 || (x.end!=null && x.end>eveningFrom)) eveningBreak=settings.eveningBreak;
    worked=Math.max(0,raw-regularBreak-eveningBreak);
  }
  return {...x,type,target,worked,balance:worked-target,regularBreak,eveningBreak};
}
function calculateAll(){return records.map(calculateRow).sort((a,b)=>a.date.localeCompare(b.date))}
function monthKey(date){return date.slice(0,7)}
function monthLabel(m){const [y,mo]=m.split("-");return new Date(Number(y),Number(mo)-1,1).toLocaleDateString("de-DE",{month:"long",year:"numeric"})}
function uniqueMonths(){return [...new Set(records.map(r=>monthKey(r.date)))].sort()}
function render(){
  const data=calculateAll();
  const worked=data.reduce((a,r)=>a+r.worked,0),target=data.reduce((a,r)=>a+r.target,0),balance=data.reduce((a,r)=>a+r.balance,0);
  const vacationUsed=data.filter(r=>r.type==="vacation").length;
  const vacationLeft=settings.annualVacation+settings.carryVacation-vacationUsed;
  $("totalBalance").textContent=(balance>=0?"+":"")+fmtMin(balance);
  $("worked").textContent=fmtMin(worked);$("target").textContent=fmtMin(target);
  $("vacationUsed").textContent=`${vacationUsed} ${vacationUsed===1?"Tag":"Tage"}`;
  $("vacationLeft").textContent=`${vacationLeft} Tage`;
  const months=uniqueMonths();
  $("periodLabel").textContent=months.length?`${months.length} Monat${months.length===1?"":"e"} im Konto`:"Noch keine Daten";
  $("rowCount").textContent=`${data.length} Tage`; $("monthCount").textContent=`${months.length}`;
  $("fileStatus").textContent=data.length?"Gespeichert":"Bereit";

  const byMonth={};data.forEach(r=>(byMonth[monthKey(r.date)]??=[]).push(r));
  $("accountSummary").innerHTML=`
    <div class="summary-row"><span>Gesamte Arbeitszeit</span><strong>${fmtMin(worked)}</strong></div>
    <div class="summary-row"><span>Gesamte Sollzeit</span><strong>${fmtMin(target)}</strong></div>
    <div class="summary-row"><span>Arbeitszeitkonto</span><strong class="${balance>=0?"pos":"neg"}">${balance>=0?"+":""}${fmtMin(balance)}</strong></div>
    <div class="summary-row"><span>Urlaubsanspruch inkl. Übertrag</span><strong>${settings.annualVacation+settings.carryVacation} Tage</strong></div>
    <div class="summary-row"><span>Resturlaub</span><strong>${vacationLeft} Tage</strong></div>`;
  $("months").innerHTML=months.length?months.map(m=>{
    const a=byMonth[m],w=a.reduce((x,r)=>x+r.worked,0),t=a.reduce((x,r)=>x+r.target,0),b=w-t,v=a.filter(r=>r.type==="vacation").length;
    return `<div class="months-row"><span>${monthLabel(m)}</span><span>${v} Urlaub</span><strong class="${b>=0?"pos":"neg"}">${b>=0?"+":""}${fmtMin(b)}</strong></div>`
  }).join(""):`<div class="muted">Noch keine Monatsdaten.</div>`;
  $("days").innerHTML=data.length?data.slice().reverse().map(r=>{
    const d=new Date(r.date+"T12:00:00"),start=r.start==null?"—":`${String(Math.floor(r.start/60)).padStart(2,"0")}:${String(r.start%60).padStart(2,"0")}`,end=r.end==null?"—":`${String(Math.floor(r.end/60)).padStart(2,"0")}:${String(r.end%60).padStart(2,"0")}`;
    return `<div class="day"><div class="date">${d.toLocaleDateString("de-DE",{day:"2-digit",month:"2-digit"})}</div><div><div class="times">${start} – ${end}</div><div class="reason">${r.reason||"keine Buchung"}${r.regularBreak||r.eveningBreak?` · Pausen ${r.regularBreak+r.eveningBreak} min`:""}</div></div><div class="balance ${r.balance>0?"pos":r.balance<0?"neg":"zero"}">${r.balance>0?"+":""}${fmtMin(r.balance)}</div></div>`
  }).join(""):`<div class="muted" style="padding:18px 0">Noch keine Daten. Importiere eine Excel-Datei.</div>`;
}
async function importFile(file){
  try{
    const buf=await file.arrayBuffer(),wb=XLSX.read(buf,{type:"array",cellDates:true});
    const ws=wb.Sheets["AZ"]||wb.Sheets[wb.SheetNames.find(n=>/^az$/i.test(n))];
    if(!ws)throw new Error("Kein Tabellenblatt „AZ“ gefunden.");
    const raw=XLSX.utils.sheet_to_json(ws,{header:1,raw:true});
    const incoming=[];
    for(let i=0;i<raw.length;i++){
      const row=raw[i]||[],d=parseDate(row[0]);
      if(!d||d.getFullYear()<2000)continue;
      const start=timeToMin(row[1]),end=timeToMin(row[2]),reason=String(row[3]??"").trim();
      if(start==null&&end==null&&!reason)continue;
      incoming.push({date:dateKey(d),start,end,reason,source:file.name});
    }
    if(!incoming.length)throw new Error("Keine Tagesbuchungen erkannt.");
    const before=records.length;
    const map=new Map(records.map(r=>[r.date,r]));
    for(const r of incoming)map.set(r.date,{...(map.get(r.date)||{}),...r});
    records=[...map.values()].sort((a,b)=>a.date.localeCompare(b.date));
    save();render();
    const added=Math.max(0,records.length-before),updated=incoming.length-added;
    $("importMessage").textContent=`${file.name} importiert: ${incoming.length} Buchungen erkannt. ${added} neue Tage, ${updated} bereits vorhandene Tage aktualisiert. Das Arbeitszeitkonto bleibt erhalten.`;
  }catch(e){$("importMessage").textContent="Importfehler: "+e.message}
}
function csvEscape(v){return `"${String(v??"").replaceAll('"','""')}"`}
function exportCSV(){
  const data=calculateAll(),head=["Datum","Beginn","Ende","Buchungsgrund","Arbeitszeit_min","Sollzeit_min","Saldo_min","Normale_Pause_min","Abendpause_min"];
  const lines=[head.map(csvEscape).join(";"),...data.map(r=>[r.date,r.start==null?"":r.start,r.end==null?"":r.end,r.reason,r.worked,r.target,r.balance,r.regularBreak,r.eveningBreak].map(csvEscape).join(";"))];
  const blob=new Blob(["\ufeff"+lines.join("\n")],{type:"text/csv;charset=utf-8"}),url=URL.createObjectURL(blob),a=document.createElement("a");
  a.href=url;a.download="arbeitszeitkonto.csv";a.click();URL.revokeObjectURL(url);
}
$("fileInput").addEventListener("change",e=>{if(e.target.files[0])importFile(e.target.files[0]);e.target.value=""});
$("exportBtn").onclick=exportCSV;
$("settingsBtn").onclick=()=>{$("settings").classList.toggle("hidden");$("weeklyHours").value=settings.weeklyHours;$("annualVacation").value=settings.annualVacation;$("carryVacation").value=settings.carryVacation;$("regularBreak").value=settings.regularBreak;$("eveningBreak").value=settings.eveningBreak;$("eveningFrom").value=settings.eveningFrom};
$("closeSettings").onclick=()=>$("settings").classList.add("hidden");
$("saveSettings").onclick=()=>{settings={weeklyHours:+$("weeklyHours").value,annualVacation:+$("annualVacation").value,carryVacation:+$("carryVacation").value,regularBreak:+$("regularBreak").value,eveningBreak:+$("eveningBreak").value,eveningFrom:$("eveningFrom").value||"16:30"};save();$("settings").classList.add("hidden");render()};
$("clearBtn").onclick=()=>{if(confirm("Alle lokal gespeicherten Arbeitszeitdaten und Einstellungen löschen?")){records=[];localStorage.removeItem("az-records");render()}};
if("serviceWorker"in navigator)navigator.serviceWorker.register("sw.js").catch(()=>{});
render();
