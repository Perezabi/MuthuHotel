function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = JSON.parse(e.postData.contents);

    saveSheet(ss, "Bills_History", data.bills);
    saveSheet(ss, "Godown_Inventory", data.inventory);
    saveSheet(ss, "Menu_Master", data.menu);
    saveSheet(ss, "Advance_Bookings", data.bookings);
    saveSheet(ss, "Daily_Expenses", data.expenses);

    return response({ status: "success" });
  } catch (error) {
    return response({
      status: "error",
      message: error.toString(),
    });
  }
}

function saveSheet(ss, sheetName, records) {
  if (!records) return;

  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.clearContents();

  if (!records.length) return;

  const headers = [
    ...new Set(records.flatMap((record) => Object.keys(record))),
  ];
  const rows = records.map((record) =>
    headers.map((header) => {
      const value = record[header];
      return typeof value === "object" ? JSON.stringify(value) : (value ?? "");
    }),
  );

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function response(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(
    ContentService.MimeType.JSON,
  );
}

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const action = e && e.parameter ? e.parameter.action : "";

  if (action !== "GET_ALL_DATA") {
    return response({
      status: "success",
      message: "MKS sync endpoint is online",
    });
  }

  return response({
    status: "success",
    menu: readSheet(ss, "Menu_Master"),
    inventory: readSheet(ss, "Godown_Inventory"),
    bills: readSheet(ss, "Bills_History"),
    bookings: readSheet(ss, "Advance_Bookings"),
    expenses: readSheet(ss, "Daily_Expenses"),
  });
}

function readSheet(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet
    .getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
    .getValues();
  const headers = values.shift();
  return values.map((row) =>
    headers.reduce((record, header, index) => {
      record[header] = row[index];
      return record;
    }, {}),
  );
}
