/**
 * 9.14.25 - testing synch up with github repo
 * 9.14.25 - added funtion to ensure one and only one START_HERE is tagged after running SORT and REPORT.
 * 8.25.25 - added logic for both URL as well as Article Title (cols B and C) dupe checking
 *
 * UPDATED 6.4.25  - changed dupe search for only article title - not URL dupes, also added abort if google doc not created
 * UPDATED 5.24.25 - adjusted dupe checking logic so that only latter duplicate records are labelled and the original is labelled 'Original'
 * UPDATED 5.24.25 - adjusted createDocFromSidebar to not change Col B or C
 * UPDATED 5.23.25 - fixed things so that script does not touch Col C (Article Title) and that Col C title is also used in Col L for Google Doc title.
 * UPDATED 5.22.25 - tried to fix dupe checking function
 * UPDATED 5.17.25 - added sort and report function and fixed the dupe count for the dupe function
 * UPDATED BY HAP 5.14.25 - with expanded dupe checking (both URL and Title Dupe checking) and Sort and Report function
 *
 * Google Apps Script for CRAP URLS 2.0 Sheet
 * - Adds sidebar tool to paste article content
 * - Creates Google Doc in a specific Drive folder
 * - Uses Article_Title (col C) as Google Doc title
 * - Writes doc title (col L), doc URL (col M), and sets status (col G)
 * - Archives rows where Column F = "Yes" into "Archive of CRAP URLS 2.0" tab
 * - Skips header and template rows
 * - Flags duplicates based on both URL (col B) and Article_Title (col C)
 * - Labels: Original, Dupe-Title, Dupe-URL, Dupe-Both
 * - Marks the earliest as "Original" by title; highlights all dupes in light red
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Article Tools")
    .addItem("Paste Clipboard to Google Doc", "showSidebar")
    .addItem("Archive Rows", "archiveRows")
    .addItem("Check for Duplicates", "checkAllForDuplicates")
    .addItem("Sort and Report", "sortAndReport")
    .addToUi();
}

function showSidebar() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CRAP URLS 2.0");
  const cell = sheet.getActiveCell();
  const row = cell.getRow();
  const articleUrl = sheet.getRange(row, 2).getValue(); // Column B

  const template = HtmlService.createTemplateFromFile("Sidebar");
  template.prefillUrl = articleUrl;
  template.activeRow = row;

  const html = template.evaluate()
    .setTitle("Paste Article Content")
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function createDocFromSidebar(articleUrl, articleContent, activeRow) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CRAP URLS 2.0");
  const rowToUpdate = activeRow;
  const articleTitle = sheet.getRange(rowToUpdate, 3).getValue(); // Column C - Article_Title

  try {
    const folder = DriveApp.getFolderById("1YgR6_EspJzoVTQjmIc2oE4EC-75iDkg8");
    const doc = DocumentApp.create(articleTitle); // Try to create the Google Doc
    DriveApp.getFileById(doc.getId()).moveTo(folder);
    doc.getBody().setText(articleContent);
    const docUrl = `https://docs.google.com/document/d/${doc.getId()}/edit`;

    // Do NOT overwrite Column B or C
    sheet.getRange(rowToUpdate, 7).setValue("Ready for Make"); // Column G - Status
    sheet.getRange(rowToUpdate, 12).setValue(articleTitle);    // Column L - Doc Title
    sheet.getRange(rowToUpdate, 13).setValue(docUrl);          // Column M - Doc URL
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Error creating Google Doc: ${error.message}`);
    return; // Stop further execution
  }
}

function archiveRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("CRAP URLS 2.0");
  const archiveSheet = ss.getSheetByName("Archive of CRAP URLS 2.0");

  if (!archiveSheet) {
    SpreadsheetApp.getUi().alert('Error: "Archive of CRAP URLS 2.0" sheet not found.');
    return;
  }

  const dataRange = sourceSheet.getDataRange();
  const data = dataRange.getValues();
  const rowsToArchive = [];

  for (let i = data.length - 1; i > 1; i--) {
    const row = data[i];
    if (row[5] === "Yes") {
      rowsToArchive.unshift(row);
      sourceSheet.deleteRow(i + 1);
    }
  }

  if (rowsToArchive.length > 0) {
    archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToArchive.length, rowsToArchive[0].length)
                .setValues(rowsToArchive);
  }

  SpreadsheetApp.getUi().alert(`Archived ${rowsToArchive.length} row(s) successfully.`);
}

/** Helpers for normalization **/
function normalizeTitle_(str) {
  return (str || "")
    .toString().toLowerCase().trim()
    .replace(/[.,\/#!$%\^&\*;:{}=\-_`~()"\[\]'’“”]/g, "")
    .replace(/\s+/g, " ");
}

function normalizeUrl_(url) {
  const s = (url || "").toString().trim().toLowerCase();
  if (!s) return "";
  // strip protocol and www
  let u = s.replace(/^https?:\/\//, "").replace(/^www\./, "");
  // drop query/hash and trailing slash
  u = u.split("?")[0].split("#")[0].replace(/\/+$/, "");
  return u;
}

/**
 * Labels:
 *  - Original: earliest record for a given normalized Title (by Col A timestamp)
 *  - Dupe-Title: later records sharing same normalized Title only
 *  - Dupe-URL: records sharing same normalized URL only (and not already a title dupe)
 *  - Dupe-Both: later records where both Title and URL are dupes
 * Any Dupe row is highlighted light red.
 */
function checkAllForDuplicates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CRAP URLS 2.0");
  const archiveSheet = ss.getSheetByName("Archive of CRAP URLS 2.0");

  if (!sheet || !archiveSheet) {
    SpreadsheetApp.getUi().alert('Error: One or both sheets not found.');
    return;
  }

  const lastRow = sheet.getLastRow();
  const archiveLastRow = archiveSheet.getLastRow();

  const data = lastRow > 2 ? sheet.getRange(3, 1, lastRow - 2, 13).getValues() : [];
  const archiveData = archiveLastRow > 2 ? archiveSheet.getRange(3, 1, archiveLastRow - 2, 13).getValues() : [];

  const allData = data.map((row, i) => ({ row, sheet, sheetName: "CRAP URLS 2.0", rowNum: i + 3 }))
    .concat(archiveData.map((row, i) => ({ row, sheet: archiveSheet, sheetName: "Archive of CRAP URLS 2.0", rowNum: i + 3 })));

  const titleGroups = {};   // normTitle -> entries[]
  const urlFirstSeen = {};  // normUrl   -> earliest entry
  const red = "#f4cccc";

  allData.forEach(e => {
    const dateAdded = new Date(e.row[0]);     // Col A: timestamp/date
    const rawUrl   = e.row[1];                // Col B
    const rawTitle = e.row[2];                // Col C

    const normTitle = normalizeTitle_(rawTitle);
    const normUrl   = normalizeUrl_(rawUrl);

    e.dateAdded = dateAdded;
    e.rawTitle = rawTitle;
    e.rawUrl = rawUrl;
    e.normTitle = normTitle;
    e.normUrl = normUrl;

    if (normTitle) {
      if (!titleGroups[normTitle]) titleGroups[normTitle] = [];
      titleGroups[normTitle].push(e);
    }

    if (normUrl) {
      const curr = urlFirstSeen[normUrl];
      if (!curr || (dateAdded && curr.dateAdded && dateAdded < curr.dateAdded)) {
        urlFirstSeen[normUrl] = e;
      }
    }
  });

  let dupeTitleCount = 0, dupeUrlCount = 0, dupeBothCount = 0;
  const touchedKey = new Set(); // sheetName#rowNum to avoid reprocessing

  // Step 1: Title groups — mark earliest as Original; later as Dupe-Title or Dupe-Both if URL dupes too
  Object.keys(titleGroups).forEach(normTitle => {
    const group = titleGroups[normTitle];
    if (group.length <= 1) return;

    group.sort((a, b) => a.dateAdded - b.dateAdded);
    const original = group[0];

    // Mark Original for earliest title
    original.sheet.getRange(original.rowNum, 7).setValue("Original");
    touchedKey.add(`${original.sheetName}#${original.rowNum}`);

    // Subsequent entries
    for (let i = 1; i < group.length; i++) {
      const e = group[i];
      const key = `${e.sheetName}#${e.rowNum}`;
      const firstSeenUrl = e.normUrl ? urlFirstSeen[e.normUrl] : null;
      const urlIsDupe = !!(firstSeenUrl &&
        (firstSeenUrl.sheetName !== e.sheetName || firstSeenUrl.rowNum !== e.rowNum) &&
        firstSeenUrl.dateAdded <= e.dateAdded);

      const label = urlIsDupe ? "Dupe-Both" : "Dupe-Title";
      e.sheet.getRange(e.rowNum, 7).setValue(label);
      e.sheet.getRange(e.rowNum, 1, 1, 13).setBackground(red);
      touchedKey.add(key);

      if (label === "Dupe-Both") dupeBothCount++;
      else dupeTitleCount++;
    }
  });

  // Step 2: URL-only dupes for rows not already labeled above
  allData.forEach(e => {
    const key = `${e.sheetName}#${e.rowNum}`;
    if (touchedKey.has(key)) return;
    if (!e.normUrl) return;

    const firstSeenUrl = urlFirstSeen[e.normUrl];
    const urlIsDupe = !!(firstSeenUrl &&
      (firstSeenUrl.sheetName !== e.sheetName || firstSeenUrl.rowNum !== e.rowNum) &&
      firstSeenUrl.dateAdded <= e.dateAdded);

    if (urlIsDupe) {
      e.sheet.getRange(e.rowNum, 7).setValue("Dupe-URL");
      e.sheet.getRange(e.rowNum, 1, 1, 13).setBackground(red);
      dupeUrlCount++;
      touchedKey.add(key);
    }
  });

  // Step 3: Build concise summary sheet
  const summaryRows = [];
  allData.forEach(e => {
    const status = e.sheet.getRange(e.rowNum, 7).getValue();
    if (["Dupe-Title", "Dupe-URL", "Dupe-Both"].includes(status)) {
      const firstUrlRef = e.normUrl && urlFirstSeen[e.normUrl]
        ? `${urlFirstSeen[e.normUrl].sheetName} R${urlFirstSeen[e.normUrl].rowNum}`
        : "";
      summaryRows.push([
        e.rawTitle || "",
        e.rawUrl || "",
        e.sheetName,
        `R${e.rowNum}`,
        status,
        firstUrlRef
      ]);
    }
  });

  const summaryName = "Dupe Summary";
  let sSheet = ss.getSheetByName(summaryName);
  if (sSheet) ss.deleteSheet(sSheet);
  sSheet = ss.insertSheet(summaryName);
  sSheet.getRange(1, 1, 1, 6).setValues([[
    "Raw Title", "URL", "Sheet", "Row", "Label", "First-URL Row (earliest)"
  ]]);
  if (summaryRows.length) {
    sSheet.getRange(2, 1, summaryRows.length, 6).setValues(summaryRows);
    sSheet.autoResizeColumns(1, 6);
  }

  SpreadsheetApp.getUi().alert(
    `Duplicate check completed.\n` +
    `Title-only: ${dupeTitleCount}\n` +
    `URL-only: ${dupeUrlCount}\n` +
    `Both: ${dupeBothCount}`
  );
}

function sortAndReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CRAP URLS 2.0');

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error: "CRAP URLS 2.0" sheet not found.');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 2) {
    SpreadsheetApp.getUi().alert('No data to sort.');
    return;
  }

  const range = sheet.getRange(3, 1, lastRow - 2, sheet.getLastColumn());
  range.sort({ column: 7, ascending: true });

  let jsonShelledCount = 0;
  let manualCount = 0;
  let readyToScrapeCount = 0;
  let readyForMakeCount = 0;
  let totalCount = 0;

  const statusRange = sheet.getRange(3, 7, lastRow - 2).getValues();
  statusRange.forEach(row => {
    const status = row[0];
    totalCount++;
    if (status === 'JSON_Shelled') jsonShelledCount++;
    if (status === 'Manual') manualCount++;
    if (status === 'Ready to Scrape') readyToScrapeCount++;
    if (status === 'Ready for Make') readyForMakeCount++;
  });

  const message = `Sorting complete.\n\nRecord Counts:\nJSON_Shelled: ${jsonShelledCount}\nManual: ${manualCount}\nReady to Scrape: ${readyToScrapeCount}\nReady for Make: ${readyForMakeCount}\nTotal Records: ${totalCount}`;

  SpreadsheetApp.getUi().alert(message);

  // After user clicks OK, handle Start_Here logic in Column N
  const columnNRange = sheet.getRange(3, 14, lastRow - 2); // Column N from row 3 to last row
  const columnNValues = columnNRange.getValues();

  // First, clear any existing "Start_Here" values by setting them to "Done"
  for (let i = 0; i < columnNValues.length; i++) {
    if (columnNValues[i][0] === "Start_Here") {
      sheet.getRange(i + 3, 14).setValue("Done");
    }
  }

  // Find the first row with "Ready to Scrape" in Column G
  const statusValues = sheet.getRange(3, 7, lastRow - 2).getValues();
  let firstReadyToScrapeRow = -1;

  for (let i = 0; i < statusValues.length; i++) {
    if (statusValues[i][0] === "Ready to Scrape") {
      firstReadyToScrapeRow = i + 3; // Convert to sheet row number
      break;
    }
  }

  // If we found a "Ready to Scrape" row, set "Start_Here" in the row immediately above it
  if (firstReadyToScrapeRow > 3) { // Make sure we're not trying to go above row 3
    sheet.getRange(firstReadyToScrapeRow - 1, 14).setValue("Start_Here");
  }
}