"use strict";
function pdfBounds(isLog, values, firstIndex) {
  if (isLog) {
    for (let i = values.length - 1; i >= 0; i--) {
      const value = values[i][0];
      if (typeof value === "number" && Number.isFinite(value)) return "A1:J" + (firstIndex+i+1);
    }
    throw Error("Column A has no numeric application number to print.");
  }
  const totals = [];
  values.forEach(([value], i) => {
    if (String(value).trim().toLowerCase() === "project totals") totals.push(firstIndex+i+1);
  });
  if (totals.length !== 1) throw Error("Expected one Project Totals row in column B; found " + totals.length + ".");
  return "B1:K" + totals[0];
}
function officeCall(register) {
  return new Promise((resolve,reject) => register(result => {
    if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
    else reject(Error(result.error.message));
  }));
}
async function workbookPdf() {
  const file = await officeCall(cb => Office.context.document.getFileAsync(Office.FileType.Pdf,cb));
  try {
    const slices = [];
    for (let i=0; i<file.sliceCount; i++) {
      const part = await officeCall(cb => file.getSliceAsync(i,cb));
      slices.push(new Uint8Array(part.data));
    }
    const pdf = new Blob(slices,{type:"application/pdf"});
    if (await pdf.slice(0,5).text() !== "%PDF-") throw Error("Excel returned an invalid PDF.");
    return pdf;
  } finally { await officeCall(cb => file.closeAsync(cb)); }
}
async function exportActivePdf(ctx) {
  const book=ctx.workbook, sheets=book.worksheets;
  const sheet=sheets.getActiveWorksheet().load("name,id");
  const heading=sheet.getRange("A1").load("text");
  const invoice=sheet.getRange("I12").load("text");
  await ctx.sync();
  const isLog=sheet.name==="Billing Log" && heading.text[0][0].trim().toUpperCase()==="PROJECT BILLING LOG";
  if (!isLog && invoice.text[0][0].trim()!=="Invoice This Period") throw Error("Select an SOV application or the Billing Log.");
  const column=sheet.getRange(isLog?"A:A":"B:B").getUsedRangeOrNullObject(true).load("rowIndex,values");
  const project=sheet.getRange(isLog?"B2":"C7").load("text");
  sheets.load("items/id,items/name,items/visibility");
  await ctx.sync();
  if (column.isNullObject) throw Error("The report column is empty.");
  const address=pdfBounds(isLog,column.values,column.rowIndex);
  const area=sheet.getRange(address);
  area.select();
  await ctx.sync();
  const fallback="Range "+address+" is selected. In Excel choose File → Print → selected cells to save this sheet as PDF.";
  try {
    const current=sheet.pageLayout.getPrintAreaOrNullObject().load("address");
    await ctx.sync();
    const existing=current.isNullObject ? "" : current.address.split("!").pop().replace(/\$/g,"");
    if (existing!==address) { sheet.pageLayout.setPrintArea(address); await ctx.sync(); }
  } catch(err) { throw Error("Excel could not set the print area. "+fallback+" "+err.message); }
  const others=sheets.items.filter(s=>s.id!==sheet.id && s.visibility==="Visible")
    .map(s=>({id:s.id,name:s.name}));
  let pdf, failure;
  try {
    for (const other of others) sheets.getItem(other.id).visibility="Hidden";
    await ctx.sync();
    sheets.load("items/id,items/visibility");
    await ctx.sync();
    const visible=sheets.items.filter(s=>s.visibility==="Visible");
    if (visible.length!==1 || visible[0].id!==sheet.id) throw Error("Excel could not isolate the active tab.");
    pdf=await workbookPdf();
  } catch(err) { failure=err; }
  finally {
    const failed=[];
    for (const other of others) {
      try {
        const s=sheets.getItemOrNullObject(other.id).load("visibility");
        await ctx.sync();
        if (!s.isNullObject && s.visibility==="Hidden") { s.visibility="Visible"; await ctx.sync(); }
      } catch { failed.push(other.name); }
    }
    if (failed.length) throw Error("Check worksheet visibility for: "+failed.join(", ")+".");
  }
  if (failure) throw Error(failure.message+" "+fallback);
  const filename=[project.text[0][0],sheet.name].filter(Boolean).join(" - ").replace(/[\\/:*?"<>|]/g,"-")+".pdf";
  const link=document.createElement("a");
  link.href=URL.createObjectURL(pdf); link.download=filename; link.textContent="Download "+filename;
  const result=document.getElementById("download");
  const previous=result.querySelector("a");
  if (previous) URL.revokeObjectURL(previous.href);
  result.replaceChildren(link);
  link.click();
  return sheet.name+" exported ("+address+"). If your browser did not start the download, use the link below.";
}

