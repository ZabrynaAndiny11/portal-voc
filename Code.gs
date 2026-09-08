var SPREADSHEET_ID =
  "1n4sc2PIiD7UuMCd7LAM_zdn3EUSMvJGJXXWsfNYUn8Y";

var SOURCE_SPREADSHEET_URL =
  "https://docs.google.com/spreadsheets/d/1K82wN91piAbwRwMNF1vXnLz0znQbfP_P9t75pcUzTbU/edit?gid=943874303#gid=943874303";

var SOURCE_SHEET_GID = 943874303;

var SHEET_NAME_USERS = "Data_User";
var SHEET_NAME_VOC = "Data_VOC";


function doPost(e) {
  var lock = LockService.getScriptLock();
  var lockAcquired = false;

  try {
    lockAcquired = lock.tryLock(10000);

    if (!lockAcquired) {
      return jsonResponse({
        result: "error",
        message: "Server sedang sibuk."
      });
    }

    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({
        result: "error",
        message: "Tidak ada data yang dikirim."
      });
    }

    var data = JSON.parse(e.postData.contents);
    var type = String(data.type || "voc").toLowerCase();

    if (type === "register") {
      var userSpreadsheet =
        SpreadsheetApp.openById(SPREADSHEET_ID);

      return saveUser(userSpreadsheet, data);
    }

    if (type === "voc") {
      return saveVocToBothSpreadsheets(data);
    }

    return jsonResponse({
      result: "error",
      message: "Tipe data tidak dikenali."
    });

  } catch (error) {
    return jsonResponse({
      result: "error",
      message: error.toString()
    });

  } finally {
    if (lockAcquired) {
      lock.releaseLock();
    }
  }
}


function saveUser(spreadsheet, data) {
  var name = normalizeName(data.name);
  var username = normalizeUsername(data.username);
  var password = String(data.password || "").trim();

  if (!name || !username || !password) {
    return jsonResponse({
      result: "error",
      message: "Nama, username, dan password wajib diisi."
    });
  }

  var sheet = getOrCreateSheet(
    spreadsheet,
    SHEET_NAME_USERS,
    [
      "Timestamp",
      "Nama Lengkap",
      "Username",
      "Password",
      "Role"
    ]
  );

  var rows = sheet.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++) {
    var existingName = normalizeName(rows[i][1]);
    var existingUsername = normalizeUsername(rows[i][2]);

    if (existingName === name) {
      return jsonResponse({
        result: "error",
        message: "Nama sudah terdaftar."
      });
    }

    if (existingUsername === username) {
      return jsonResponse({
        result: "error",
        message: "Username sudah terdaftar."
      });
    }
  }

  sheet.appendRow([
    data.timestamp || new Date(),
    name,
    username,
    password,
    data.role || "agent"
  ]);

  return jsonResponse({
    result: "success",
    message: "Registrasi berhasil disimpan."
  });
}


function saveVocToBothSpreadsheets(data) {
  var idNumber = String(data.idNumber || "").trim();
  var namaAm = String(data.namaAm || "").trim();
  var namaCustomer = String(data.namaCustomer || "").trim();
  var voc = String(data.voc || "").trim();
  var keterangan = String(data.keterangan || "").trim();

  if (!idNumber || !voc) {
    return jsonResponse({
      result: "error",
      message: "ID Number dan VOC wajib diisi."
    });
  }

  var attachmentUrl = "";

  try {
    attachmentUrl = saveAttachment(
      data.attachment,
      idNumber
    );
  } catch (error) {
    return jsonResponse({
      result: "error",
      message: "Gagal menyimpan lampiran: " + error.toString()
    });
  }

  var headers = [
    "Timestamp",
    "Agent",
    "Role",
    "ID Number",
    "Nama AM",
    "Nama Customer",
    "VOC",
    "Keterangan",
    "Lampiran URL"
  ];

  var rowData = [
    data.timestamp || new Date(),
    data.agent || "",
    data.role || "agent",
    idNumber,
    namaAm,
    namaCustomer,
    voc,
    keterangan,
    attachmentUrl
  ];

  try {
    var vocSpreadsheet =
      SpreadsheetApp.openById(SPREADSHEET_ID);

    var vocSheet = getOrCreateSheet(
      vocSpreadsheet,
      SHEET_NAME_VOC,
      headers
    );

    vocSheet.appendRow(rowData);

  } catch (error) {
    return jsonResponse({
      result: "error",
      message:
        "Gagal menyimpan ke VOC INPUT: " +
        error.toString()
    });
  }

  try {
    var sourceSpreadsheet =
      SpreadsheetApp.openByUrl(SOURCE_SPREADSHEET_URL);

    var sourceVocSheet = getOrCreateSheet(
      sourceSpreadsheet,
      SHEET_NAME_VOC,
      headers
    );

    sourceVocSheet.appendRow(rowData);

  } catch (error) {
    return jsonResponse({
      result: "partial_success",
      message:
        "Data tersimpan di VOC INPUT, tetapi gagal ke sheet sumber: " +
        error.toString(),
      attachmentUrl: attachmentUrl
    });
  }

  return jsonResponse({
    result: "success",
    message: "Data VOC berhasil disimpan ke kedua spreadsheet.",
    attachmentUrl: attachmentUrl
  });
}


function doGet(e) {
  var action = "";

  if (e && e.parameter && e.parameter.action) {
    action = String(e.parameter.action).toLowerCase();
  }

  if (action === "source") {
    return getSourceData();
  }

  if (action === "login") {
    return checkLogin(
      e.parameter.username,
      e.parameter.password
    );
  }

  if (action === "debug") {
    return debugSourceSpreadsheet();
  }

  return jsonResponse({
    result: "success",
    message: "VOC Portal Apps Script aktif."
  });
}


function checkLogin(username, password) {
  try {
    var uname = normalizeUsername(username);
    var pass = String(password || "").trim();

    if (!uname || !pass) {
      return jsonResponse({
        result: "error",
        message: "Username dan password wajib diisi."
      });
    }

    var spreadsheet =
      SpreadsheetApp.openById(SPREADSHEET_ID);

    var sheet = spreadsheet.getSheetByName(SHEET_NAME_USERS);

    if (!sheet) {
      return jsonResponse({
        result: "error",
        message: "Username belum terdaftar atau Password salah."
      });
    }

    var rows = sheet.getDataRange().getValues();

    for (var i = 1; i < rows.length; i++) {
      var rowUsername = normalizeUsername(rows[i][2]);
      var rowPassword = String(rows[i][3] || "").trim();

      if (rowUsername === uname && rowPassword === pass) {
        return jsonResponse({
          result: "success",
          name: rows[i][1],
          username: rowUsername,
          role: rows[i][4] || "agent"
        });
      }
    }

    return jsonResponse({
      result: "error",
      message: "Username belum terdaftar atau Password salah."
    });

  } catch (error) {
    return jsonResponse({
      result: "error",
      message: "Gagal memeriksa akun: " + error.toString()
    });
  }
}


function getSourceData() {
  try {
    var spreadsheet =
      SpreadsheetApp.openByUrl(SOURCE_SPREADSHEET_URL);

    var sheets = spreadsheet.getSheets();
    var sourceSheet = null;

    for (var i = 0; i < sheets.length; i++) {
      if (
        Number(sheets[i].getSheetId()) ===
        Number(SOURCE_SHEET_GID)
      ) {
        sourceSheet = sheets[i];
        break;
      }
    }

    if (!sourceSheet) {
      return jsonResponse({
        result: "error",
        message:
          "Sheet dengan GID " +
          SOURCE_SHEET_GID +
          " tidak ditemukan."
      });
    }

    var lastRow = sourceSheet.getLastRow();

    if (lastRow < 1) {
      return jsonResponse([]);
    }

    // Kolom C = ID Number
    // Kolom D = Nama AM
    // Kolom E = Nama
    var values = sourceSheet
      .getRange(1, 3, lastRow, 3)
      .getDisplayValues();

    var result = [];

    for (var rowIndex = 0; rowIndex < values.length; rowIndex++) {
      var idNumber = String(values[rowIndex][0] || "").trim();
      var namaAm = String(values[rowIndex][1] || "").trim();
      var nama = String(values[rowIndex][2] || "").trim();

      var normalizedId = idNumber
        .toLowerCase()
        .replace(/\s+/g, "");

      if (
        !idNumber ||
        normalizedId === "id" ||
        normalizedId === "idnumber"
      ) {
        continue;
      }

      result.push({
        idNumber: idNumber,
        namaAm: namaAm || "-",
        nama: nama || "-"
      });
    }

    return jsonResponse(result);

  } catch (error) {
    return jsonResponse({
      result: "error",
      message:
        "Gagal mengambil data sumber: " +
        error.toString()
    });
  }
}


function saveAttachment(attachment, idNumber) {
  if (!attachment || !attachment.base64) {
    return "";
  }

  var bytes = Utilities.base64Decode(
    attachment.base64
  );

  var blob = Utilities.newBlob(
    bytes,
    attachment.mimeType || "application/octet-stream",
    attachment.name || ("lampiran-" + idNumber)
  );

  var file = DriveApp.createFile(blob);
  file.setName(idNumber + " - " + file.getName());

  return file.getUrl();
}


function debugSourceSpreadsheet() {
  try {
    var spreadsheet =
      SpreadsheetApp.openByUrl(SOURCE_SPREADSHEET_URL);

    var sheets = spreadsheet.getSheets();
    var result = [];

    for (var i = 0; i < sheets.length; i++) {
      result.push({
        name: sheets[i].getName(),
        gid: sheets[i].getSheetId(),
        rows: sheets[i].getLastRow(),
        columns: sheets[i].getLastColumn()
      });
    }

    return jsonResponse({
      result: "success",
      spreadsheetName: spreadsheet.getName(),
      sheets: result
    });

  } catch (error) {
    return jsonResponse({
      result: "error",
      message: error.toString()
    });
  }
}


function normalizeName(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}


function normalizeUsername(value) {
  return String(value || "")
    .trim()
    .toLowerCase();
}


function getOrCreateSheet(spreadsheet, sheetName, headers) {
  var sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }

  return sheet;
}


function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
