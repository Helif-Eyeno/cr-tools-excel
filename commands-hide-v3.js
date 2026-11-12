/* global Office, Excel */
"use strict";

Office.onReady();

function notify(message) {
  const url = new URL("result.html", window.location.href);
  url.searchParams.set("message", message);
  Office.context.ui.displayDialogAsync(url.href, { height: 30, width: 35 }, result => {
    if (result.status === Office.AsyncResultStatus.Failed) console.error(result.error);
  });
}

function crSheets(items) {
  return items.map(sheet => ({ sheet, match: /^CR#(\d+)$/.exec(sheet.name) }))
    .filter(item => item.match)
    .map(item => ({ sheet: item.sheet, number: Number(item.match[1]) }))
    .filter(item => Number.isSafeInteger(item.number));
}

async function openLog() {
  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getItemOrNullObject("Log");
    await context.sync();
    if (sheet.isNullObject) throw new Error("The Log worksheet was not found.");
    sheet.activate();
    await context.sync();
  });
}

async function lastCR() {
  await Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    const found = crSheets(sheets.items).sort((a, b) => b.number - a.number);
    if (!found.length) throw new Error("No CR tabs found.");
    found[0].sheet.activate();
    await context.sync();
  });
}

async function createCR() {
  await Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    const template = sheets.getItemOrNullObject("TemplateCR");
    const project = sheets.getItem("Project_Info").getRange("C5");
    project.load("text");
    await context.sync();
    if (template.isNullObject) throw new Error("TemplateCR was not found.");
    const projectNumber = String(project.text[0][0]).trim();
    if (!projectNumber || projectNumber.startsWith("#")) throw new Error("Project_Info!C5 needs a valid project number.");
    const next = String(Math.max(0, ...crSheets(sheets.items).map(item => item.number)) + 1).padStart(2, "0");
    const name = `CR#${next}`;
    const sheet = template.copy(Excel.WorksheetPositionType.end);
    sheet.name = name;
    sheet.visibility = Excel.SheetVisibility.visible;
    sheet.activate();
    const date = new Date();
    sheet.getRange("AI66").values = [[(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) - Date.UTC(1899, 11, 30)) / 86400000]];
    sheet.getRange("AI68").hyperlink = { documentReference: `'${name}'!A1`, textToDisplay: `${projectNumber}-${next}-R00`, screenTip: name };
    await context.sync();
  });
}

async function preparePDF() {
  return Excel.run(async context => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    const number = sheet.getRange("AI68"), title = sheet.getRange("G73");
    number.load("text"); title.load("text");
    const breaks = sheet.horizontalPageBreaks;
    breaks.load("items/rowIndex");
    await context.sync();
    if (!breaks.items.some(pageBreak => pageBreak.rowIndex === 65)) {
      breaks.add("B66");
    }
    const layout = sheet.pageLayout;
    layout.setPrintArea("B1:AM131");
    layout.orientation = Excel.PageOrientation.portrait;
    layout.paperSize = Excel.PaperType.letter;
    layout.centerHorizontally = true; layout.centerVertically = false;
    // Percentage scaling preserves manual page breaks; fit-to-page does not.
    layout.zoom = { scale: 65 };
    layout.topMargin = 54; layout.bottomMargin = 54; layout.leftMargin = 54; layout.rightMargin = 54;
    layout.printGridlines = false;
    await context.sync();
    return `Print settings prepared for B1:AM131. Use Excel's active-sheet PDF export and verify the output. Suggested name: ${number.text[0][0]} ${title.text[0][0]} - ${sheet.name}.pdf`;
  });
}


async function testDirectPDF() {
  await preparePDF();
  const original = await Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    const active = sheets.getActiveWorksheet();
    sheets.load("items/id,items/name,items/visibility"); active.load("id");
    await context.sync();
    const states = sheets.items.filter(s => s.id !== active.id && s.visibility === Excel.SheetVisibility.visible).map(s => ({id: s.id, name: s.name}));
    return states;
  });
  let pdf;
  try {
  await Excel.run(async context => {
    for (const state of original) context.workbook.worksheets.getItem(state.name).visibility = Excel.SheetVisibility.hidden;
    await context.sync();
  });
  await Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    const active = sheets.getActiveWorksheet();
    sheets.load("items/name,items/id,items/visibility"); active.load("id");
    await context.sync();
    const unexpected = sheets.items.filter(s => s.id !== active.id && s.visibility === Excel.SheetVisibility.visible).map(s => s.name);
    if (unexpected.length) throw new Error("Visibility check failed; still visible: " + unexpected.join(", "));
  });
  pdf = await new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(Office.FileType.Pdf, result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) return reject(new Error(result.error.message));
      const file = result.value;
      const slices = [];
      const close = () => file.closeAsync(() => {});
      const next = index => {
        if (index >= file.sliceCount) { close(); resolve(new Blob(slices, {type: "application/pdf"})); return; }
        file.getSliceAsync(index, part => {
          if (part.status !== Office.AsyncResultStatus.Succeeded) { close(); reject(new Error(part.error.message)); return; }
          slices.push(new Uint8Array(part.value.data));
          next(index + 1);
        });
      };
      next(0);
    });
  });
  } finally {
    await Excel.run(async context => {
      const sheets = context.workbook.worksheets;
      for (const state of original) sheets.getItem(state.name).visibility = Excel.SheetVisibility.visible;
      await context.sync();
    });
  }
  const signature = new Uint8Array(await pdf.slice(0, 5).arrayBuffer());
  if (String.fromCharCode(...signature) !== "%PDF-") throw new Error("Excel did not return a valid PDF. No file was downloaded.");
  const url = URL.createObjectURL(pdf);
  const link = document.createElement("a");
  link.href = url; link.download = "CR-hide-v3-TEST.pdf";
  document.body.appendChild(link); link.click(); link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}

function command(operation, showSuccess) {
  return async event => {
    try {
      const result = await operation();
      if (showSuccess) notify(result);
    } catch (error) {
      console.error("CR Tools command failed", error);
      notify(`Could not complete the action: ${error.message || error}`);
    } finally {
      event.completed();
    }
  };
}

Office.actions.associate("command_create", command(createCR, false));
Office.actions.associate("command_log", command(openLog, false));
Office.actions.associate("command_last", command(lastCR, false));
Office.actions.associate("command_pdf", command(testDirectPDF, false));

