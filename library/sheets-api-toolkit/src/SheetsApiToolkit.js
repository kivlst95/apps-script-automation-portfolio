/**
 * Get spreadsheet id part from the file URL
 * 
 * @param {string} url - URL to check and ID extract    
 * Return {string} extracted ID from the inputed URL
 */

function getSpreadsheetIdFromUrl(url) { return String(url).match(/[-\w]{25,}/)[0] || "" }



/**
 * Get sheet id part from the file URL
 * 
 * @param {string} url - URL to check and ID extract    
 * Return {string} extracted ID from the inputed URL
 */

function getSheetIdFromUrl(url) { return String(url).match(/[#&]gid=([0-9]+)/)[1] || "" }


/**
 * Custom function for the Google Sheet to calculate the number of columns in r1c1 notation
 *
 * @param {string} ranges - string with r1c1 notation
 * @return {int} number of columns beetween srart and end column 
 */

function columnsFromRange(a1Notation) {
  // Locale function for the reduce callback
  function getColumnNumber_(acc, value, index, array) {
    var charCode = array.join("").charCodeAt(index) - 64
    acc += charCode * Math.pow(26, array.length - (index + 1))
    return acc
  }

  if (a1Notation.map) {            // Test whether input is an array to use with an ArrayFormula

    //Recursive use of the same function os the array 
    return a1Notation.map(columnsFromRange); // Recurse over array if so.
  }

  else {
    if (a1Notation) {

      // Split Start and End of the range
      a1notation = a1Notation.split(":");

      // Regex to extract columns (letters), rows (digits)
      try {
        var columnStart = a1notation?.[0].match(/[A-Z]+/)[0];
        var columnEnd = a1notation?.[1]?.match(/[A-Z]+/)?.[0] || "";
      }

      catch { throw new Error(`Check r1c1 notation correctness`) }

      // Identify columnStart and End using reduce
      columnStart = columnStart.toUpperCase().split("").reduce(getColumnNumber_, 0) - 1
      columnEnd = columnEnd.toUpperCase().split("").reduce(getColumnNumber_, 0) - columnStart || ""

      return columnEnd
    }
    return ""
  }
}


/**
 * Function for the spreadsheet data getting using GridRange filter. 
 * Transpose A1 notation range to GridRange // Borrowed from => https://gist.github.com/tanaikech/95c7cd650837f33a564babcaf013cae0
 * 
 * 
 * @param {int} fullSpreadsheetURL - full Google Spreadsheet with spreadsheet ID and sheet ID mentioning 
 * @param {int} range - A1 notation range (e.g A1:B1)
 * 
 * Return gridRange {object} => https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/other#GridRange
 * 
 */

function a1Notation2GridRange(fullSpreadsheetURL, a1Notation) {
  // Locale function for the reduce callback
  function getColumnNumber_(acc, value, index, array) {
    var charCode = array.join("").charCodeAt(index) - 64
    acc += charCode * Math.pow(26, array.length - (index + 1))
    return acc
  }

  // Split Start and End of the range
  a1notation = a1Notation.split(":");

  // Regex to extract columns (letters), rows (digits)
  try {
    var rowStart = a1notation?.[0]?.match(/\d+/g)[0] * 1 - 1
    var rowEnd = a1notation?.[1]?.match(/\d+/g)?.[0] * 1 || ""
    var columnStart = a1notation?.[0].match(/[A-Z]+/)[0];
    var columnEnd = a1notation?.[1]?.match(/[A-Z]+/)?.[0] || "";
  }

  catch { throw new Error(`Check r1c1 notation correctness`) }

  // Identify columnStart and End using reduce
  columnStart = columnStart.toUpperCase().split("").reduce(getColumnNumber_, 0) - 1
  columnEnd = columnEnd.toUpperCase().split("").reduce(getColumnNumber_, 0)

  // Initialize sheetID for the gridRange object
  var pageId = getSheetIdFromUrl(fullSpreadsheetURL) || "";
  var gridRange =
  {
    sheetId: pageId,
    startRowIndex: rowStart,
    endRowIndex: rowEnd,
    startColumnIndex: columnStart,
    endColumnIndex: columnEnd
  }

  // Check if the end parameters are not specifies and delete them if yes
  if (!gridRange["endRowIndex"]) {
    delete gridRange["endRowIndex"]
  }
  if (!gridRange["endColumnIndex"]) {
    delete gridRange["endColumnIndex"]
  }
  return gridRange;
}

function setValues(fullSpreadsheetURL, a1Notation, values, inputOption = "USER_ENTERED", dateRender = "SERIAL_NUMBER", valueRender = "UNFORMATTED_VALUE") {

  const spreadsheetID = getSpreadsheetIdFromUrl(fullSpreadsheetURL)

  // Edge case when full URL isn`t correct
  if (!spreadsheetID) {
    throw new Error(`Please, check the correctness of the full URLs mentioned. URL must contain SpreadsheetID and SheetID`)
  }


  //  Initializing payload object for the further post request
  var payload = {
    "valueInputOption": inputOption,
    "responseDateTimeRenderOption": dateRender,
    "responseValueRenderOption": valueRender,
    "data": [
      {
        "dataFilter": {
          "gridRange": a1Notation2GridRange(fullSpreadsheetURL, a1Notation),

        },
        "values": values,
        "majorDimension": "ROWS"
      }
    ]
  }


  Sheets.Spreadsheets.Values.batchUpdateByDataFilter(payload, spreadsheetID)
}



function getValues(fullSpreadsheetURL, a1Notation, valueRender = "UNFORMATTED_VALUE", dateRender = "SERIAL_NUMBER") {

  const spreadsheetID = getSpreadsheetIdFromUrl(fullSpreadsheetURL)

  // Edge case when full URL isn`t correct
  if (!spreadsheetID) {
    throw new Error(`Please, check the correctness of the full URLs mentioned. URL must contain SpreadsheetID and SheetID`)
  }


  //  Initializing payload object for the further post request
  var payload = {
    "dataFilters": [
      {
        "gridRange": a1Notation2GridRange(fullSpreadsheetURL, a1Notation)
      }
    ],
    "valueRenderOption": valueRender,
    "dateTimeRenderOption": dateRender
  }



  return Sheets.Spreadsheets.Values.batchGetByDataFilter(payload, spreadsheetID)
}


function clearValuesSheets(fullSpreadsheetURL, a1Notation) {

  var spreadSheetId = getSpreadsheetIdFromUrl(fullSpreadsheetURL)

  // Edge case when full URL isn`t correct
  if (!spreadSheetId) {
    throw new Error(`Please, check the correctness of the full URLs mentioned. URL must contain SpreadsheetID and SheetID`)
  }

  //  Initializing payload object for the further get request
  var payload = {
    "dataFilters": [
      {
        "gridRange": a1Notation2GridRange(fullSpreadsheetURL, a1Notation)
      }
    ]
  }

  // Use built-in Sheets service to clear data
  var response = Sheets.Spreadsheets.Values.batchClearByDataFilter(payload, spreadSheetId)

  Log.info(response)
}



function ensureEnoughRowsSheetsApi(spreadsheetURL, rangeA1, data) {
  const neededRows = data.length;
  const neededCols = Math.max(...data.map(r => r.length));

  const spreadsheetId = getSpreadsheetIdFromUrl(spreadsheetURL);
  let sheetId = getSheetIdFromUrl(spreadsheetURL);

  // Get spreadsheet metadata
  const meta = Sheets.Spreadsheets.get(spreadsheetId, {
    fields: "sheets.properties(sheetId,title,gridProperties(rowCount,columnCount))"
  });

  // If sheetId not provided, use first sheet
  if (!sheetId) {
    sheetId = meta.sheets[0].properties.sheetId;
  }

  const sheet = meta.sheets.find(s => s.properties.sheetId === Number(sheetId));
  if (!sheet) throw new Error(`Sheet with ID ${sheetId} not found`);

  const sheetName = sheet.properties.title;
  const fullRange = `${sheetName}!${rangeA1.split("!").pop()}`;

  const { rowCount, columnCount } = sheet.properties.gridProperties;

  // Determine start row from range (e.g., "A2:H" → 2)
  const startRow = parseInt(rangeA1.match(/\d+/)?.[0] || "1", 10);
  const availableRows = rowCount - (startRow - 1);
  const missingRows = neededRows - availableRows;

  const requests = [];

  // Expand rows if needed
  if (missingRows > 0) {
    requests.push({
      updateSheetProperties: {
        properties: {
          sheetId,
          gridProperties: { rowCount: rowCount + missingRows }
        },
        fields: "gridProperties.rowCount"
      }
    });
    Log.warn(`↕️ Added ${missingRows} rows (now ${rowCount + missingRows})`);
  }

  // Expand columns if needed
  if (columnCount < neededCols) {
    requests.push({
      updateSheetProperties: {
        properties: {
          sheetId,
          gridProperties: { columnCount: neededCols }
        },
        fields: "gridProperties.columnCount"
      }
    });
    Log.info(`↔️ Expanded columns to ${neededCols}`);
  }

  if (requests.length) {
    Sheets.Spreadsheets.batchUpdate({ requests }, spreadsheetId);
  }

  return { spreadsheetId, sheetId, sheetName, fullRange };
}


function deleteBasicFilterSheetsApi(spreadsheetFullURL) {
  const spreadsheetId = getSpreadsheetIdFromUrl(spreadsheetFullURL);
  const sheetId = getSheetIdFromUrl(spreadsheetFullURL);

  const meta = Sheets.Spreadsheets.get(spreadsheetId, {
    fields: "sheets.properties(sheetId,title),sheets.basicFilter"
  });

  const targetSheet =
    meta.sheets.find(s => s.properties.sheetId === Number(sheetId)) ||
    meta.sheets[0];

  if (!targetSheet.basicFilter) {
    Log.info(`ℹ️ No basic filter found on sheet "${targetSheet.properties.title}"`);
    return;
  }

  const requests = [
    {
      clearBasicFilter: {
        sheetId: targetSheet.properties.sheetId
      }
    }
  ];

  Sheets.Spreadsheets.batchUpdate({ requests }, spreadsheetId);
  Log.info(`🧹 Removed basic filter from sheet "${targetSheet.properties.title}"`);
}

function addBasicFilterSheetsApi(spreadsheetFullURL) {
  const spreadsheetId = getSpreadsheetIdFromUrl(spreadsheetFullURL);
  const sheetId = getSheetIdFromUrl(spreadsheetFullURL);

  // Get sheet metadata (for title, row/column counts)
  const meta = Sheets.Spreadsheets.get(spreadsheetId, {
    fields: "sheets.properties(sheetId,title,gridProperties(rowCount,columnCount))"
  });

  const targetSheet =
    meta.sheets.find(s => s.properties.sheetId === Number(sheetId)) ||
    meta.sheets[0]; // fallback to first sheet if no gid in URL

  const { title, gridProperties } = targetSheet.properties;
  const rowCount = gridProperties.rowCount;
  const columnCount = gridProperties.columnCount;

  // Add filter for the whole used grid (usually first row as headers)
  const requests = [
    {
      setBasicFilter: {
        filter: {
          range: {
            sheetId: targetSheet.properties.sheetId,
            startRowIndex: 0, // typically header row at top
            startColumnIndex: 0,
            endRowIndex: rowCount,
            endColumnIndex: columnCount,
          },
        },
      },
    },
  ];

  Sheets.Spreadsheets.batchUpdate({ requests }, spreadsheetId);
  Log.info(`✅ Added basic filter to sheet "${title}" (${spreadsheetId})`);
}