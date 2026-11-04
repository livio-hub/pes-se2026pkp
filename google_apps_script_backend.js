// =====================================================================
// KODE GOOGLE APPS SCRIPT: BACKEND DATABASE REAL-TIME PES SE2026
// Tempelkan kode ini di Google Spreadsheet: Ekstensi -> Apps Script
// =====================================================================

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "get_all";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (action === "get_all" || action === "get_ub" || action === "get_sls") {
    // 1. Data Sampel 38 UB
    var sheetUb = getOrCreateSheet(ss, "Data_UB", ["IDSBR", "Nama Perusahaan", "Status", "PPL", "Pengawas", "Terakhir Diupdate"]);
    var dataUb = sheetUb.getDataRange().getValues();
    var ubList = [];
    for (var i = 1; i < dataUb.length; i++) {
      if (dataUb[i][0]) {
        ubList.push({
          idsbr: String(dataUb[i][0]),
          usaha: dataUb[i][1],
          status: dataUb[i][2],
          ppl: dataUb[i][3],
          pml: dataUb[i][4],
          updated_at: dataUb[i][5]
        });
      }
    }
    
    // 2. Data Petugas SLS (22 PPL)
    var sheetSls = getOrCreateSheet(ss, "Data_SLS", ["Kode PPL", "Nama", "Target", "Realisasi", "Approved", "Submit", "Draft", "Reject", "Open", "Terakhir Diupdate"]);
    var dataSls = sheetSls.getDataRange().getValues();
    var slsList = [];
    for (var j = 1; j < dataSls.length; j++) {
      if (dataSls[j][0]) {
        slsList.push({
          kodePpl: String(dataSls[j][0]),
          nama: dataSls[j][1],
          target: Number(dataSls[j][2]) || 0,
          realisasi: Number(dataSls[j][3]) || 0,
          completed: Number(dataSls[j][4]) || 0,
          submitted: Number(dataSls[j][5]) || 0,
          draft: Number(dataSls[j][6]) || 0,
          rejected: Number(dataSls[j][7]) || 0,
          open: Number(dataSls[j][8]) || 0,
          updated_at: dataSls[j][9]
        });
      }
    }

    // 3. Data Progres Petugas UB (5 PPL)
    var sheetUbPetugas = getOrCreateSheet(ss, "Data_UB_Petugas", ["Nama", "Target", "Realisasi", "Dicacah", "Submit", "Diolah", "Draft", "Open", "Pengawas", "Terakhir Diupdate"]);
    var dataUbP = sheetUbPetugas.getDataRange().getValues();
    var ubPetugasList = [];
    for (var k = 1; k < dataUbP.length; k++) {
      if (dataUbP[k][0]) {
        ubPetugasList.push({
          nama: dataUbP[k][0],
          target: Number(dataUbP[k][1]) || 0,
          realisasi: Number(dataUbP[k][2]) || 0,
          dicacah: Number(dataUbP[k][3]) || 0,
          submit: Number(dataUbP[k][4]) || 0,
          diolah: Number(dataUbP[k][5]) || 0,
          draft: Number(dataUbP[k][6]) || 0,
          open: Number(dataUbP[k][7]) || 0,
          pml: dataUbP[k][8],
          updated_at: dataUbP[k][9]
        });
      }
    }

    var result = {
      status: "success",
      timestamp: new Date().toISOString(),
      ub_data: ubList,
      sls_data: slsList,
      ub_petugas_data: ubPetugasList
    };

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: "ready" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var raw = e.postData.contents;
    var payload = JSON.parse(raw);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var action = payload.action;
    var now = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss WIB");

    // 1. UPDATE STATUS UB SATUAN DARI WEB
    if (action === "update_ub") {
      var sheetUb = getOrCreateSheet(ss, "Data_UB", ["IDSBR", "Nama Perusahaan", "Status", "PPL", "Pengawas", "Terakhir Diupdate"]);
      var data = sheetUb.getDataRange().getValues();
      var foundRow = -1;
      var targetIdsbr = String(payload.idsbr);

      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === targetIdsbr) {
          foundRow = i + 1; // 1-indexed baris spreadsheet
          break;
        }
      }

      if (foundRow !== -1) {
        sheetUb.getRange(foundRow, 3).setValue(payload.status);
        sheetUb.getRange(foundRow, 6).setValue(now);
      } else {
        sheetUb.appendRow([
          targetIdsbr,
          payload.usaha || "-",
          payload.status || "Belum Berhasil Didata",
          payload.ppl || "-",
          payload.pml || "-",
          now
        ]);
      }

      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Status UB berhasil disimpan!" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 2. SINKRONISASI DATA SLS (22 PPL & FASIH)
    if (action === "sync_sls") {
      var sheetSls = getOrCreateSheet(ss, "Data_SLS", ["Kode PPL", "Nama", "Target", "Realisasi", "Approved", "Submit", "Draft", "Reject", "Open", "Terakhir Diupdate"]);
      var items = payload.items || [];
      sheetSls.clearContents();
      sheetSls.appendRow(["Kode PPL", "Nama", "Target", "Realisasi", "Approved", "Submit", "Draft", "Reject", "Open", "Terakhir Diupdate"]);
      sheetSls.getRange(1, 1, 1, 10).setFontWeight("bold").setBackground("#e2e8f0");
      
      var rowsSls = [];
      for (var s = 0; s < items.length; s++) {
        var it = items[s];
        var fd = it.fasihDetail || {};
        rowsSls.push([
          it.kodePpl || ("PPL-" + (s + 1)),
          it.nama || "-",
          it.target || 0,
          it.realisasi || 0,
          fd.completed || 0,
          fd.submitted || 0,
          fd.draft || 0,
          fd.rejected || 0,
          fd.open || 0,
          now
        ]);
      }
      if (rowsSls.length > 0) {
        sheetSls.getRange(2, 1, rowsSls.length, 10).setValues(rowsSls);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "success", count: rowsSls.length, message: "Data SLS berhasil disimpan di Google Sheets!" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 3. SINKRONISASI DATA PETUGAS UB (5 PPL)
    if (action === "sync_ub_petugas") {
      var sheetUbP = getOrCreateSheet(ss, "Data_UB_Petugas", ["Nama", "Target", "Realisasi", "Dicacah", "Submit", "Diolah", "Draft", "Open", "Pengawas", "Terakhir Diupdate"]);
      var pList = payload.petugas || payload.items || [];
      sheetUbP.clearContents();
      sheetUbP.appendRow(["Nama", "Target", "Realisasi", "Dicacah", "Submit", "Diolah", "Draft", "Open", "Pengawas", "Terakhir Diupdate"]);
      sheetUbP.getRange(1, 1, 1, 10).setFontWeight("bold").setBackground("#fef3c7");
      
      var rowsUbP = [];
      for (var u = 0; u < pList.length; u++) {
        var p = pList[u];
        rowsUbP.push([
          p.nama || "-",
          p.target || 0,
          p.realisasi || 0,
          p.dicacah || 0,
          p.submit || 0,
          p.diolah || 0,
          p.draft || 0,
          p.open || 0,
          p.pml || "-",
          now
        ]);
      }
      if (rowsUbP.length > 0) {
        sheetUbP.getRange(2, 1, rowsUbP.length, 10).setValues(rowsUbP);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "success", count: rowsUbP.length, message: "Data Petugas UB berhasil disimpan di Google Sheets!" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 4. BULK SYNC DATA 38 SAMPEL UB
    if (action === "bulk_sync_ub") {
      var sheetUb = getOrCreateSheet(ss, "Data_UB", ["IDSBR", "Nama Perusahaan", "Status", "PPL", "Pengawas", "Terakhir Diupdate"]);
      var items = payload.items || [];

      sheetUb.clearContents();
      sheetUb.appendRow(["IDSBR", "Nama Perusahaan", "Status", "PPL", "Pengawas", "Terakhir Diupdate"]);
      sheetUb.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#e2e8f0");
      
      var rowsToInsert = [];
      for (var k = 0; k < items.length; k++) {
        var itm = items[k];
        rowsToInsert.push([
          String(itm.idsbr),
          itm.usaha || "-",
          itm.status || "Belum Berhasil Didata",
          itm.ppl || "-",
          itm.pml || "-",
          now
        ]);
      }
      if (rowsToInsert.length > 0) {
        sheetUb.getRange(2, 1, rowsToInsert.length, 6).setValues(rowsToInsert);
      }

      return ContentService.createTextOutput(JSON.stringify({ status: "success", count: rowsToInsert.length, message: "38 Sampel UB berhasil disimpan di Google Sheets!" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 5. SYNC SEMUA (SLS, UB, PETUGAS UB)
    if (action === "sync_all") {
      if (payload.sls_items) {
        var shS = getOrCreateSheet(ss, "Data_SLS", ["Kode PPL", "Nama", "Target", "Realisasi", "Approved", "Submit", "Draft", "Reject", "Open", "Terakhir Diupdate"]);
        shS.clearContents();
        shS.appendRow(["Kode PPL", "Nama", "Target", "Realisasi", "Approved", "Submit", "Draft", "Reject", "Open", "Terakhir Diupdate"]);
        var rS = [];
        for (var i1 = 0; i1 < payload.sls_items.length; i1++) {
          var x = payload.sls_items[i1];
          var f = x.fasihDetail || {};
          rS.push([x.kodePpl, x.nama, x.target, x.realisasi, f.completed||0, f.submitted||0, f.draft||0, f.rejected||0, f.open||0, now]);
        }
        if (rS.length > 0) shS.getRange(2, 1, rS.length, 10).setValues(rS);
      }
      if (payload.ub_items) {
        var shU = getOrCreateSheet(ss, "Data_UB", ["IDSBR", "Nama Perusahaan", "Status", "PPL", "Pengawas", "Terakhir Diupdate"]);
        shU.clearContents();
        shU.appendRow(["IDSBR", "Nama Perusahaan", "Status", "PPL", "Pengawas", "Terakhir Diupdate"]);
        var rU = [];
        for (var i2 = 0; i2 < payload.ub_items.length; i2++) {
          var y = payload.ub_items[i2];
          rU.push([String(y.idsbr), y.usaha, y.status, y.ppl, y.pml, now]);
        }
        if (rU.length > 0) shU.getRange(2, 1, rU.length, 6).setValues(rU);
      }
      if (payload.ub_petugas) {
        var shP = getOrCreateSheet(ss, "Data_UB_Petugas", ["Nama", "Target", "Realisasi", "Dicacah", "Submit", "Diolah", "Draft", "Open", "Pengawas", "Terakhir Diupdate"]);
        shP.clearContents();
        shP.appendRow(["Nama", "Target", "Realisasi", "Dicacah", "Submit", "Diolah", "Draft", "Open", "Pengawas", "Terakhir Diupdate"]);
        var rP = [];
        for (var i3 = 0; i3 < payload.ub_petugas.length; i3++) {
          var z = payload.ub_petugas[i3];
          rP.push([z.nama, z.target, z.realisasi, z.dicacah||0, z.submit||0, z.diolah||0, z.draft||0, z.open||0, z.pml||"-", now]);
        }
        if (rP.length > 0) shP.getRange(2, 1, rP.length, 10).setValues(rP);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Seluruh data berhasil disinkronkan ke Google Sheets!" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Aksi tidak dikenali" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// FUNGSI HELPER: BIKIN SHEET JIKA BELUM ADA
function getOrCreateSheet(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e2e8f0");
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}
