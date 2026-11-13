"use strict";
const blocks = [[14,18],[20,24],[26,33],[35,40],[42,48],[50,55],[59,77]];
const rows = blocks.flatMap(([a,b]) => Array.from({length:b-a+1}, (_,i) => a+i));
const actions = {
  amount: ["Enter Fixed Amount","Select a detail row. Enter the invoice amount for this period before retention.","Apply amount"],
  check: ["Check Before Billing","Review the active SOV application for billing errors and credits.","Check rows"],
  next: ["Close & Create Next Application","This copies the active application, carries forward billed amounts, locks the prior application and refreshes the log.","Close and create"],
  log: ["Refresh Billing Log","Rebuild the report from this workbook’s SOV applications.","Refresh log"],
  pdf: ["Export Active Tab to PDF","Download the active application or Billing Log. Other visible tabs are briefly hidden during export.","Download PDF"]
};
function issue(values, formulas, types) {
  if (types.includes("Error")) return "Excel error";
  if (types.some(t => !["Empty","Integer","Double"].includes(t))) return "Non-numeric billing entry";
  if (![2,5].every(i => typeof formulas[i]==="string" && formulas[i].startsWith("="))) return "Missing billing formula";
  if (values[0]<0 || values[0]>1) return "Completion outside 0% to 100%";
  if (values[5]<-0.00499999) return "Negative invoice amount (review credit)";
  return "";
}
async function snapshot(ctx,s) {
  ctx.workbook.application.calculate("Full");
  const data=s.getRange("A1:K80").load("values,formulas,valueTypes,text");
  s.load("name,id"); s.protection.load("protected");
  await ctx.sync();
  return data;
}
function matches(data) {
  return data.values[11][8]==="Invoice This Period" && data.values[79][1]==="Project Totals";
}
function issues(data) {
  return rows.map(r => [r,issue(...[data.values,data.formulas,data.valueTypes].map(grid => grid[r-1].slice(3,9)))])
    .filter(([,msg]) => msg);
}
function dateSerial(text) {
  const ms=Date.parse(text+"T00:00:00Z");
  if (!Number.isFinite(ms) || new Date(ms).toISOString().slice(0,10)!==text) throw Error("Choose a valid billing date.");
  return ms/86400000+25569;
}
async function catalog(ctx) {
  const sheets=ctx.workbook.worksheets.load("items/name");
  await ctx.sync();
  const info=sheets.items.map(s=>({s,invoice:s.getRange("I12").load("values"),total:s.getRange("B80").load("values"),number:s.getRange("I2").load("values"),project:s.getRange("C7").load("text,numberFormat")}));
  for (const {s} of info) s.protection.load("protected");
  await ctx.sync();
  return info.filter(x=>x.invoice.values[0][0]==="Invoice This Period" && x.total.values[0][0]==="Project Totals")
    .sort((a,b)=>a.number.values[0][0]-b.number.values[0][0]);
}
function prefix(name) { return "'"+name.replace(/'/g,"''")+"'!"; }
async function refresh(ctx) {
  const apps=await catalog(ctx);
  if (!apps.length) throw Error("No SOV application sheets found.");
  const sheets=ctx.workbook.worksheets;
  let log=sheets.getItemOrNullObject("Billing Log");
  await ctx.sync();
  if (log.isNullObject) log=sheets.add("Billing Log");
  else {
    const top=log.getRange("A1:A7").load("values");
    await ctx.sync();
    if (top.values[0][0]!=="PROJECT BILLING LOG") throw Error("Billing Log has a different layout.");
    if (top.values[4][0]==="Application #") log.getRange("2:3").insert("Down");
  }
  const used=log.getUsedRangeOrNullObject(true).load("rowIndex,rowCount");
  await ctx.sync();
  const start=prefix(apps[0].s.name);
  log.getRange("A1:A5").values=[["PROJECT BILLING LOG"],["Project Number"],["Project Name"],["Project Start"],["Project Finish"]];
  log.getRange("B2:B3").formulas=[['=IF('+start+'C7="","",'+start+'C7)'],['='+start+'C6&""']];
  log.getRange("B2").numberFormat=apps[0].project.numberFormat;
  log.getRange("B4:B5").numberFormat=[["mmm d, yyyy"],["mmm d, yyyy"]];
  log.getRange("C4").formulas=[['=IF(COUNT(B4:B5)<>2,"Enter both project dates.",IF(OR(B4<=0,B5<B4),"Check project dates.",""))']];
  const headers=log.getRange("A7:J7");
  headers.values=[["Application #","Period start","Period end","Contract incl. changes","Invoice this period","Invoiced to date","Schedule used","Invoiced %","Gap (pp)","Status"]];
  headers.format.fill.color="#DC143C"; headers.format.font.color="white"; headers.format.font.bold=true;
  const end=7+apps.length;
  if (!used.isNullObject && used.rowIndex+used.rowCount>end) log.getRange("A"+(end+1)+":J"+(used.rowIndex+used.rowCount)).clear("Contents");
  log.getRange("A8:J"+end).formulas=apps.map(({s},i)=>{
    const p=prefix(s.name), r=i+8;
    return ["I2","I4","I5","E80","I80"].map(a=>"="+p+a).concat(["="+p+"G80+"+p+"I80",
      '=IF(OR(COUNT($B$4:$B$5)<>2,$B$4<=0,$B$5<$B$4,C'+r+'<=0),"",MAX(0,MIN(1,(C'+r+'-$B$4+1)/($B$5-$B$4+1))))',
      '=IF(D'+r+'=0,"",F'+r+'/D'+r+')','=IF(COUNT(G'+r+':H'+r+')=2,(H'+r+'-G'+r+')*100,"")',s.protection.protected?"Completed / locked":"Open"]);
  });
  for (const [a,b,fmt] of [["B","C","mmm d, yyyy"],["D","F","$#,##0.00;[Red]($#,##0.00)"],["G","H","0.0%"],["I","I","+0.0;[Red]-0.0;0.0"]]) log.getRange(a+"8:"+b+end).numberFormat=apps.map(()=>Array(b.charCodeAt(0)-a.charCodeAt(0)+1).fill(fmt));
  log.getRange("A1:J"+end).format.autofitColumns();
  log.getRange("A1").format.font.bold=true;
  await ctx.sync(); return log;
}
async function rollover(ctx,s,data,start,end) {
  if (s.protection.protected) throw Error("Select the open billing application.");
  const n=data.values[1][8];
  if (!Number.isSafeInteger(n) || n<0) throw Error("Check the application number in I2.");
  if (end<start) throw Error("End date must be on or after start date.");
  const errors=issues(data).filter(([,msg])=>!msg.startsWith("Negative"));
  if (errors.length) throw Error("Row "+errors[0].join(" — "));
  const name="SOV App"+(n+1);
  const existing=ctx.workbook.worksheets.getItemOrNullObject(name);
  await ctx.sync();
  if (!existing.isNullObject) throw Error(name+" already exists.");
  const dest=s.copy("After",s);
  await ctx.sync();
  try {
    dest.name=name;
    dest.tabColor="";
    dest.getRange("I2").values=[[n+1]];
    dest.getRange("I4:I5").values=[[start],[end]];
    for (const [a,b] of blocks) dest.getRange("G"+a+":G"+b).values=data.values.slice(a-1,b).map(v=>[v[5]]);
    const result=await snapshot(ctx,dest);
    for (const r of rows) {
      const invoice=result.values[r-1][8];
      if (typeof invoice!=="number" || Math.abs(invoice)>0.000001) throw Error("Opening invoice did not reset on row "+r+".");
    }
    s.getUsedRange().format.protection.locked=true;
    s.protection.protect();
    await ctx.sync();
  } catch (err) {
    s.protection.load("protected"); await ctx.sync();
    if (!s.protection.protected) { dest.delete(); await ctx.sync(); }
    throw err;
  }
  dest.activate(); await ctx.sync();
  try { s.tabColor="#5A6E64"; await refresh(ctx); }
  catch(err) { return name+" created and prior application locked. Run Refresh Billing Log. "+err.message; }
  return name+" created. "+s.name+" is locked and the billing log is current.";
}
const el=id=>document.getElementById(id);
const action=new URLSearchParams(location.search).get("action");
async function execute(ctx) {
  if (action==="log") { const log=await refresh(ctx); log.activate(); await ctx.sync(); return "Billing Log refreshed."; }
  if (action==="pdf") return exportActivePdf(ctx);
  const s=ctx.workbook.worksheets.getActiveWorksheet(), data=await snapshot(ctx,s);
  if (!matches(data)) throw Error("Select a sheet with the original SOV application layout.");
  if (action==="check") {
    const found=issues(data);
    return found.length?found.length+" row(s) to review.\n"+found.map(([r,msg])=>"Row "+r+" — "+msg).join("\n"):"No issues found on the billing detail rows.";
  }
  if (action==="next") return rollover(ctx,s,data,dateSerial(el("start").value),dateSerial(el("end").value));
  if (s.protection.protected) throw Error("Select the open billing application.");
  const cell=ctx.workbook.getActiveCell().load("rowIndex"); await ctx.sync();
  const r=cell.rowIndex+1;
  if (!rows.includes(r)) throw Error("Select a milestone or change-order detail row.");
  const value=data.values[r-1], amount=el("amount").valueAsNumber;
  if (typeof value[4]!=="number" || value[4]===0) throw Error("Check the contract value in E"+r+".");
  const prior=value[6]===""?0:value[6];
  if (typeof prior!=="number") throw Error("Check the prior billing in G"+r+".");
  const pct=(prior+amount)/value[4];
  if (!Number.isFinite(pct) || pct<0 || pct>1) throw Error("This amount puts completion outside 0% to 100%.");
  s.getRange("D"+r).values=[[pct]]; ctx.workbook.application.calculate("Full"); await ctx.sync();
  return "Applied the fixed amount to row "+r+" on "+s.name+".";
}
if (Office.onReady) Office.onReady(({host})=>{
  if (host!==Office.HostType.Excel || !Object.hasOwn(actions,action)) return;
  const [title,help,button]=actions[action];
  el("title").textContent=title; el("help").textContent=help; el("run").textContent=button;
  const fields=action==="amount"?[["amount","Amount this period","number"]]:action==="next"?[["start","New billing start","date"],["end","New billing end","date"]]:[];
  for (const [id,label,type] of fields) {
    const wrapper=document.createElement("label"), input=document.createElement("input");
    input.id=id; input.type=type; input.required=true; if(type==="number") input.step="0.01";
    wrapper.append(label,document.createElement("br"),input); el("fields").append(wrapper);
  }
  el("form").hidden=false;
  el("form").onsubmit=async event=>{
    event.preventDefault(); if (action==="pdf") { const old=el("download").querySelector("a"); if(old) { URL.revokeObjectURL(old.href); old.remove(); } } el("run").disabled=true; el("status").textContent="Working…";
    try { el("status").textContent=await Excel.run(execute); }
    catch(err) { el("status").textContent=err.message; }
    finally { el("run").disabled=false; }
  };
});


