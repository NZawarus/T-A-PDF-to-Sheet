function convertPDFToText() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("How To Use");
  
  const pdfFileId = sheet.getRange("B5").getValue();
  Logger.log(pdfFileId);
  const pdfFile = DriveApp.getFileById(pdfFileId);
  const pdfBlob = pdfFile.getBlob();
  
  const resource = {
    title: pdfFile.getName().replace(/\.pdf$/i, ""),
    mimeType: MimeType.GOOGLE_DOCS
  };

  const options = {
    ocr: true,
    ocrLanguage: "en"
  };

  // Create a temporary Google Doc using Drive API OCR
  const tempDoc = Drive.Files.create(resource, pdfBlob, options);
  const tempDocId = tempDoc.id;

  // Open the temporary Doc and extract all the text
  const doc = DocumentApp.openById(tempDocId);
  const extractedText = doc.getBody().getText();
  
  // Create a brand new .txt file in your Google Drive with the extracted text
  const textFileName = resource.title + ".txt";
  const textFile = DriveApp.createFile(textFileName, extractedText, MimeType.PLAIN_TEXT);
  const txtFileId = textFile.getId();
  Logger.log("Text File ID:" + txtFileId);
  
  // Clean up and delete the temporary Google Doc to keep your Drive organized
  Drive.Files.remove(tempDocId);
  
  // Log the URL of your new plain text file
  Logger.log("Success! Text file created at: " + textFile.getUrl());
  return txtFileId;
}

function parseTimeAndAttendance(txtFileId) {
  // Use the created plain text file as rawData
  var txt = DriveApp.getFileById(txtFileId);
  var rawData = txt.getBlob().getDataAsString();

  // Using the active spreadsheet linked to this script, get the Timetable sheet, 
  // creating one if it does not yet exist, clearing the contents if it does
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Timetable");
  
  if (!sheet) {
    sheet = ss.insertSheet("Timetable");
  } else {
    sheet.clearContents();
  }

  // Define headers for the output
  var output = [["DATE", "TIME", "ACTIVITY 1", "ACTIVITY 2", "SOURCE"]];

  const headers = [
    "Date",
    "Day",
    "Assigned Shift Code",
    "Assigned Shift Start",
    "Assigned Shift Stop",
    "Time Outside Shift Code",
    "Time Outside Shift Start",
    "Time Outside Shift Stop",
    "Time Not Worked Code",
    "Time Not Worked Start",
    "Time Not Worked Stop",
    "Optional Type",
    "Optional Code",
    "Optional Start",
    "Optional Stop",
    "Pcnt"
  ];

  const START_MARKER = "Assigned Shift";
  const END_MARKER = "T&A Approval Status:";

  const NUM_COLUMNS = 16;

  // Split while preserving empty lines.
  const lines = rawData.split(/\r?\n/);

  const allRows = [headers];

  let i = 0;

  while (i < lines.length) {
    // Find the next "Assigned Shift" section.
    if (lines[i].trim() !== START_MARKER) {
      i++;
      continue;
    }

    // Find the corresponding T&A Approval Status.
    let sectionEnd = i + 1;

    while (
      sectionEnd < lines.length &&
      lines[sectionEnd].trim() !== END_MARKER
    ) {
      sectionEnd++;
    }

    if (sectionEnd >= lines.length) {
      break;
    }

    // Find the table header.
    let headerIndex = -1;

    for (let j = i; j < sectionEnd; j++) {
      if (lines[j].trim() === 'Date') {
        headerIndex = j;
        break;
      }
    }

    if (headerIndex === -1) {
      i = sectionEnd + 1;
      continue;
    }

    // The table data starts after the 16 header lines.
    let dataStart = headerIndex + NUM_COLUMNS;

    /*
     * Each row in the extracted document consists of exactly
     * 16 lines, including blank lines.
     */
    while (dataStart + NUM_COLUMNS <= sectionEnd) {
      
      let row = lines
        .slice(dataStart, dataStart + NUM_COLUMNS)
        .map(value => value.trim());

      // Ignore completely empty groups.
      if (row.some(value => value !== '')) {

        // Make sure the first cell looks like a date.
        if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(row[0])) {
          allRows.push(row);
        }
      }

      dataStart += NUM_COLUMNS;
    }

    // Continue looking for the next pay period.
    i = sectionEnd + 1;
  }

  // Write everything to the spreadsheet.

  sheet.clearContents();

  if (allRows.length > 0) {
    sheet
      .getRange(1, 1, allRows.length, NUM_COLUMNS)
      .setValues(allRows);
  }

  Logger.log(`Imported ${allRows.length - 1} data rows.`);
}

function splitMainSheetToTarget() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Source sheet
  const sourceSheet =
    ss.getSheetByName("Timetable") || ss.getActiveSheet();

  // Create/reset target sheet
  let targetSheet = ss.getSheetByName("ActivitySheet");

  if (!targetSheet) {
    targetSheet = ss.insertSheet("ActivitySheet");
  } else {
    targetSheet.clearContents();
  }

  // Output headers
  const outputHeaders = [
    "Date",
    "Time",
    "Activity1",
    "Activity2",
    "Source"
  ];

  targetSheet
    .getRange(1, 1, 1, outputHeaders.length)
    .setValues([outputHeaders]);

  // ============================================================
  // BUILD CODE LOOKUP TABLES
  // ============================================================

  /*
   * Normalizes a code for lookup.
   *
   * Some codes arrive from the parsed Timetable as numbers,
   * causing leading zeroes to disappear.
   *
   * Example:
   *   20  -> "20"
   *   020 -> "020"
   *
   * We will try both forms when looking up a code.
   */
  function normalizeCode(value) {
    if (value === null || value === undefined || value === "") {
      return "";
    }

    return String(value).trim();
  }

  /*
   * Build a simple Code -> Description lookup.
   */
  function buildCodeLookup(sheetName) {
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      Logger.log("Lookup sheet not found: " + sheetName);
      return {};
    }

    const values = sheet.getDataRange().getValues();
    const lookup = {};

    for (let i = 1; i < values.length; i++) {
      const code = normalizeCode(values[i][0]);
      const description = values[i][1];

      if (code !== "") {
        lookup[code] = description;
      }
    }

    return lookup;
  }

  /*
   * Optional is different because it has:
   *
   * Type | Code | Description
   *
   * Therefore the lookup key is Type + Code.
   */
  function buildOptionalLookup() {
    const sheet = ss.getSheetByName("Optional");

    if (!sheet) {
      Logger.log("Optional lookup sheet not found.");
      return {};
    }

    const values = sheet.getDataRange().getValues();
    const lookup = {};

    for (let i = 1; i < values.length; i++) {
      const type = normalizeCode(values[i][0]);
      const code = normalizeCode(values[i][1]);
      const description = values[i][2];

      if (type !== "" && code !== "") {
        lookup[type + "|" + code] = description;
      }
    }

    return lookup;
  }

  const assignedShiftLookup =
    buildCodeLookup("AssignedShift");

  const outsideShiftLookup =
    buildCodeLookup("OutsideShift");

  const timeNotWorkedLookup =
    buildCodeLookup("TimeNotWorked");

  const optionalLookup =
    buildOptionalLookup();


  // ============================================================
  // CODE LOOKUP FUNCTION
  // ============================================================

  /*
   * Look up a code while accounting for lost leading zeroes.
   *
   * For example:
   *
   *   "20" will try:
   *      "20"
   *      "020"
   *
   *   "46" will try:
   *      "46"
   *      "046"
   *
   *   "106" will try:
   *      "106"
   *      "0106"
   *
   * The exact code is always attempted first.
   */
  function lookupCode(lookup, code) {
    const normalized = normalizeCode(code);

    if (normalized === "") {
      return "";
    }

    // Exact match
    if (lookup.hasOwnProperty(normalized)) {
      return lookup[normalized];
    }

    // Try three-digit version
    if (/^\d+$/.test(normalized)) {
      const padded = normalized.padStart(3, "0");

      if (lookup.hasOwnProperty(padded)) {
        return lookup[padded];
      }
    }

    // No description found
    Logger.log(
      "No description found for code: " + normalized
    );

    // Return original code rather than losing information
    return normalized;
  }


  // ============================================================
  // GET SOURCE DATA
  // ============================================================

  const data = sourceSheet.getDataRange().getValues();

  if (data.length <= 1) {
    return;
  }

  const headers = data[0];
  const rows = data.slice(1);

  // Find a column by its exact header
  const getIndex = (headerName) => {
    return headers.findIndex(
      h =>
        String(h).trim().toLowerCase() ===
        headerName.toLowerCase()
    );
  };


  // ============================================================
  // COLUMN INDICES
  // ============================================================

  const dateIdx = getIndex("Date");

  const assignedShiftCodeIdx =
    getIndex("Assigned Shift Code");

  const assignedShiftStartIdx =
    getIndex("Assigned Shift Start");

  const assignedShiftStopIdx =
    getIndex("Assigned Shift Stop");

  const outsideShiftCodeIdx =
    getIndex("Time Outside Shift Code");

  const outsideShiftStartIdx =
    getIndex("Time Outside Shift Start");

  const outsideShiftStopIdx =
    getIndex("Time Outside Shift Stop");

  const timeNotWorkedCodeIdx =
    getIndex("Time Not Worked Code");

  const timeNotWorkedStartIdx =
    getIndex("Time Not Worked Start");

  const timeNotWorkedStopIdx =
    getIndex("Time Not Worked Stop");

  const optionalTypeIdx =
    getIndex("Optional Type");

  const optionalCodeIdx =
    getIndex("Optional Code");

  const optionalStartIdx =
    getIndex("Optional Start");

  const optionalStopIdx =
    getIndex("Optional Stop");


  // ============================================================
  // ACTIVITY DEFINITIONS
  // ============================================================

  const activities = [
    {
      name: "Assigned Shift",

      codeIdx: assignedShiftCodeIdx,
      startIdx: assignedShiftStartIdx,
      stopIdx: assignedShiftStopIdx,

      lookup: assignedShiftLookup
    },

    {
      name: "Time Outside Shift",

      codeIdx: outsideShiftCodeIdx,
      startIdx: outsideShiftStartIdx,
      stopIdx: outsideShiftStopIdx,

      lookup: outsideShiftLookup
    },

    {
      name: "Time Not Worked",

      codeIdx: timeNotWorkedCodeIdx,
      startIdx: timeNotWorkedStartIdx,
      stopIdx: timeNotWorkedStopIdx,

      lookup: timeNotWorkedLookup
    }
  ];


  const outputRows = [];


  // ============================================================
  // PROCESS EACH TIMETABLE ROW
  // ============================================================

  rows.forEach(row => {

    const date =
      dateIdx !== -1
        ? row[dateIdx]
        : "";


    // ----------------------------------------------------------
    // STANDARD ACTIVITIES
    // ----------------------------------------------------------

    activities.forEach(activity => {

      const code =
        activity.codeIdx !== -1
          ? row[activity.codeIdx]
          : "";

      const start =
        activity.startIdx !== -1
          ? row[activity.startIdx]
          : "";

      const stop =
        activity.stopIdx !== -1
          ? row[activity.stopIdx]
          : "";


      const hasCode =
        code !== "" &&
        code !== null &&
        code !== undefined;

      const hasStart =
        start !== "" &&
        start !== null &&
        start !== undefined;

      const hasStop =
        stop !== "" &&
        stop !== null &&
        stop !== undefined;


      // Convert encoded code to plaintext description
      const description =
        hasCode
          ? lookupCode(activity.lookup, code)
          : "";


      // --------------------------------------------------------
      // START
      // --------------------------------------------------------

      if (hasStart) {

        outputRows.push([
          date,
          start,
          "Start - " + activity.name,
          description,
          "CASTLE"
        ]);

      }


      // --------------------------------------------------------
      // STOP
      // --------------------------------------------------------

      if (hasStop) {

        outputRows.push([
          date,
          stop,
          "Stop - " + activity.name,
          description,
          "CASTLE"
        ]);

      }


      // --------------------------------------------------------
      // CODE ONLY
      // --------------------------------------------------------

      if (hasCode && !hasStart && !hasStop) {

        outputRows.push([
          date,
          "",
          description,
          "",
          "CASTLE"
        ]);

      }

    });


    // ==========================================================
    // OPTIONAL ACTIVITY
    // ==========================================================

    const optionalType =
      optionalTypeIdx !== -1
        ? row[optionalTypeIdx]
        : "";

    const optionalCode =
      optionalCodeIdx !== -1
        ? row[optionalCodeIdx]
        : "";

    const optionalStart =
      optionalStartIdx !== -1
        ? row[optionalStartIdx]
        : "";

    const optionalStop =
      optionalStopIdx !== -1
        ? row[optionalStopIdx]
        : "";


    const hasOptionalType =
      optionalType !== "" &&
      optionalType !== null &&
      optionalType !== undefined;

    const hasOptionalCode =
      optionalCode !== "" &&
      optionalCode !== null &&
      optionalCode !== undefined;

    const hasOptionalStart =
      optionalStart !== "" &&
      optionalStart !== null &&
      optionalStart !== undefined;

    const hasOptionalStop =
      optionalStop !== "" &&
      optionalStop !== null &&
      optionalStop !== undefined;


    const hasOptional =
      hasOptionalType || hasOptionalCode;


    /*
     * Optional lookup requires BOTH type and code:
     *
     * Type = 4
     * Code = 31
     *
     * becomes:
     *
     * Enroute TDY
     */
    let optionalDescription = "";

    if (hasOptionalCode) {

      const optionalKey =
        normalizeCode(optionalType) +
        "|" +
        normalizeCode(optionalCode);

      if (optionalLookup.hasOwnProperty(optionalKey)) {
        optionalDescription =
          optionalLookup[optionalKey];
      } else {
        Logger.log(
          "No Optional description found for Type=" +
          optionalType +
          ", Code=" +
          optionalCode
        );

        optionalDescription =
          normalizeCode(optionalCode);
      }
    }


    // ----------------------------------------------------------
    // OPTIONAL START
    // ----------------------------------------------------------

    if (hasOptionalStart) {

      outputRows.push([
        date,
        optionalStart,
        "Start - Optional",
        optionalDescription,
        "CASTLE"
      ]);

    }


    // ----------------------------------------------------------
    // OPTIONAL STOP
    // ----------------------------------------------------------

    if (hasOptionalStop) {

      outputRows.push([
        date,
        optionalStop,
        "Stop - Optional",
        optionalDescription,
        "CASTLE"
      ]);

    }


    // ----------------------------------------------------------
    // OPTIONAL CODE ONLY
    // ----------------------------------------------------------

    if (
      hasOptional &&
      !hasOptionalStart &&
      !hasOptionalStop
    ) {

      outputRows.push([
        date,
        "",
        optionalDescription,
        "",
        "CASTLE"
      ]);

    }

  });


  // ============================================================
  // WRITE OUTPUT
  // ============================================================

  if (outputRows.length > 0) {

    targetSheet
      .getRange(
        2,
        1,
        outputRows.length,
        outputHeaders.length
      )
      .setValues(outputRows);


    // Format date
    targetSheet
      .getRange(
        2,
        1,
        outputRows.length,
        1
      )
      .setNumberFormat("MM/dd/yyyy");


    // Format time
    targetSheet
      .getRange(
        2,
        2,
        outputRows.length,
        1
      )
      .setNumberFormat("h:mm AM/PM");

  }


  // Resize columns
  targetSheet.autoResizeColumns(
    1,
    outputHeaders.length
  );


  Logger.log(
    "Created " +
    outputRows.length +
    " activity rows."
  );
}

function convertToSheet() {
  txtFileId = convertPDFToText();
  parseTimeAndAttendance(txtFileId);
  splitMainSheetToTarget();
}

