/**
 * fit.Hive Admin Portal - Backend v8.1
 * Features: Auto-Expiry, Multi-Sheet Support, Row Locking, Photo Upload to Drive, Trainer Field Restrictions
 * @OnlyCurrentDoc false
 */
const SPREADSHEET_ID = "1iIBGAmV9ixsxs1K-O827pMiAncULk4PAxIftTQ96vMw";
const DRIVE_FOLDER_ID = "1nk5j54WwSflM01Sc-4fbfiJtk1WTSM3h";

const TRAINER_HIDDEN_HEADER_NAMES = ["Phone", "Phone Number", "Email"];
const TRAINER_HIDDEN_KEYS = ["Phone", "PhoneNumber", "Email"];

function normalizeRole(role) {
  const r = String(role || "").trim().toLowerCase();
  return r === "trainer" ? "trainer" : "admin";
}

function getRoleFromRequestObject(obj) {
  return normalizeRole(obj && obj.userRole);
}

function toKey(headerName) {
  return String(headerName || "").replace(/[^a-zA-Z0-9]/g, "");
}

function jsonOut(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  const role = getRoleFromRequestObject(params);
  const sheetName = params.sheet || "Regular Tracker";

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return jsonOut([]);

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  if (!values || values.length === 0) return jsonOut([]);

  const headers = values[0];

  // Auto-expiry engine
  const expiryIdx = headers.indexOf("Expiry Date");
  const statusIdx = headers.indexOf("Payment Status");
  if (expiryIdx !== -1 && statusIdx !== -1) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let updatesMade = false;
    for (let i = 1; i < values.length; i++) {
      const expiryVal = values[i][expiryIdx];
      const currentStatus = values[i][statusIdx];

      if (expiryVal instanceof Date) {
        const checkDate = new Date(expiryVal);
        checkDate.setHours(0, 0, 0, 0);

        if (checkDate < today && currentStatus !== "Expired") {
          sheet.getRange(i + 1, statusIdx + 1).setValue("Expired");
          updatesMade = true;
        } else if (checkDate >= today && currentStatus === "Expired") {
          sheet.getRange(i + 1, statusIdx + 1).setValue("Paid");
          updatesMade = true;
        }
      }
    }

    if (updatesMade) SpreadsheetApp.flush();
  }

  const displayData = sheet.getDataRange().getValues();
  const rawHeaders = displayData.shift();

  const json = displayData.map((row, index) => {
    const obj = { rowNumber: index + 2 };
    rawHeaders.forEach((header, i) => {
      const key = toKey(header);
      obj[key] = row[i];
    });

    if (role === "trainer") {
      TRAINER_HIDDEN_KEYS.forEach(k => delete obj[k]);
    }

    return obj;
  });

  return jsonOut(json);
}

function parsePostParams(e) {
  if (e && e.postData && e.postData.contents) {
    return JSON.parse(e.postData.contents || "{}");
  }
  if (e && e.parameter) return e.parameter;
  return {};
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const params = parsePostParams(e);
    const action = params.action;
    const role = getRoleFromRequestObject(params);

    if (!action) throw new Error("Missing action");

    // Photo upload
    if (action === "uploadPhoto") {
      if (!params.fileName) throw new Error("Missing fileName");
      if (!params.mimeType) throw new Error("Missing mimeType");
      if (!params.data) throw new Error("Missing data");

      const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      const timestamp = new Date().getTime();
      const uniqueFileName = "client_photo_" + timestamp + "_" + params.fileName;

      const decodedBytes = Utilities.base64Decode(params.data);
      const blob = Utilities.newBlob(decodedBytes, params.mimeType, uniqueFileName);
      const file = folder.createFile(blob);
      const fileId = file.getId();

      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const url = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w400";

      return jsonOut({
        success: true,
        url: url,
        fileId: fileId,
        fileName: uniqueFileName
      });
    }

    if (action !== "add" && action !== "edit") {
      throw new Error("Unsupported action");
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = params.sheetName;
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error("Sheet not found: " + sheetName);

    const rowData = typeof params.data === "string" ? JSON.parse(params.data) : params.data;
    if (!Array.isArray(rowData)) throw new Error("Invalid data payload");

    if (action === "add") {
      sheet.appendRow(rowData);
      return jsonOut({ success: true, message: "Data saved successfully" });
    }

    // Edit flow
    const rowNumber = parseInt(params.rowNumber, 10);
    if (!rowNumber || rowNumber < 2) throw new Error("Invalid rowNumber for edit");

    // Server-side protection: trainer cannot overwrite hidden fields
    if (role === "trainer") {
      const lastCol = sheet.getLastColumn();
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      const existingRow = sheet.getRange(rowNumber, 1, 1, lastCol).getValues()[0];

      headers.forEach((h, idx) => {
        if (TRAINER_HIDDEN_HEADER_NAMES.indexOf(String(h)) !== -1) {
          rowData[idx] = existingRow[idx];
        }
      });
    }

    sheet.getRange(rowNumber, 1, 1, rowData.length).setValues([rowData]);

    return jsonOut({
      success: true,
      message: "Data saved successfully"
    });
  } catch (error) {
    return jsonOut({
      success: false,
      error: String(error)
    });
  } finally {
    lock.releaseLock();
  }
}


///-------------------------------------------------------------


{
  "timeZone": "Asia/Kolkata",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/script.external_request"
  ],
  "webapp": {
    "executeAs": "USER_DEPLOYING",
    "access": "ANYONE_ANONYMOUS"
  }
}
