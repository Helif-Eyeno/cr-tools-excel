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
    layout.zoom = { horizontalFitToPages: 1 };
    layout.topMargin = 54; layout.bottomMargin = 54; layout.leftMargin = 54; layout.rightMargin = 54;
    layout.printGridlines = false;
    await context.sync();
    return `Print settings prepared for B1:AM131. Use Excel's active-sheet PDF export and verify the output. Suggested name: ${number.text[0][0]} ${title.text[0][0]} - ${sheet.name}.pdf`;
  });
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
Office.actions.associate("command_pdf", command(preparePDF, true));

