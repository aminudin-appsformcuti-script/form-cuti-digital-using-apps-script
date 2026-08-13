// ==========================================
// CONFIGURATION
// ==========================================
const SEATALK_WEBHOOK_URL = "WEBHOOK_SEATALK_MILIK _KAMU_TEMPEL_DISINI";

// ID Folder Khusus PDF Cuti
const FOLDER_PDF_ID      = "FOLDER_DRIVE_PDF_ID_MILIK_KAMU_TEMPEL_DISINI"; 

// ID Folder Khusus Foto Evidence
const FOLDER_EVIDENCE_ID = "FOLDER_DRIVE_EVIDENCE_ID_KAMU_TEMPEL_DISINI";

// ==========================================
// CORE WEB APP FUNCTIONS
// ==========================================
function doGet(e) {
  var role = (e.parameter.role || "karyawan").toLowerCase().trim();
  var reqId = e.parameter.id || e.parameter.reqId || ""; 
  
  if (reqId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("sheet_database_punya_kamu");
    if (sheet) {
      var data = sheet.getDataRange().getDisplayValues();
      for (var i = 1; i < data.length; i++) {
        if (data[i][0].toString().trim() === reqId.toString().trim()) {
          var currentStatus = data[i][16]; // Kolom Q (Index 16) - Status Pengajuan

          if (currentStatus.indexOf("REJECTED") !== -1) {
            return HtmlService.createHtmlOutput("<h3 style='font-family:Arial;color:#f44336;text-align:center;margin-top:50px;'>Pengajuan ini telah DITOLAK (" + currentStatus + ").</h3>");
          }
          if ((role === "tl" || role === "team lead") && currentStatus !== "PENDING_TL") {
            return HtmlService.createHtmlOutput("<h3 style='font-family:Arial;color:#ff5722;text-align:center;margin-top:50px;'>Pengajuan ini sudah diproses oleh Team Lead.</h3>");
          }
          if ((role === "sl" || role === "shift lead") && currentStatus !== "PENDING_SL") {
            return HtmlService.createHtmlOutput("<h3 style='font-family:Arial;color:#ff5722;text-align:center;margin-top:50px;'>Pengajuan ini sudah diproses oleh Shift Lead.</h3>");
          }
          if ((role === "tp" || role === "tp lead" || role === "tplead") && currentStatus === "APPROVED") {
            return HtmlService.createHtmlOutput("<h3 style='font-family:Arial;color:#4caf50;text-align:center;margin-top:50px;'>Pengajuan ini sudah selesai dan disetujui (APPROVED).</h3>");
          }
          break;
        }
      }
    }
  }
  
  var tmp = HtmlService.createTemplateFromFile('Index');
  tmp.role = role;
  tmp.reqId = reqId;
  
  return tmp.evaluate()
    .setTitle("Form Cuti SPX Express")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getPengajuanData(reqId) {
  if (!reqId) return null;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DataCuti");
  if (!sheet) return null;
  
  var data = sheet.getDataRange().getDisplayValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === reqId.toString().trim()) {
      return {
        reqId: data[i][0],
        nama: data[i][1],
        id_spx: data[i][2],
        unit_kerja: data[i][3],
        team: data[i][4],
        email: data[i][5],
        tgl_gabung: data[i][6],
        tgl_mulai: data[i][7],
        tgl_selesai: data[i][8],
        lama_cuti: data[i][9],
        alasan: data[i][10],
        backup: data[i][11],
        status: data[i][16],
        kategori_cuti: data[i][20] || "Cuti Tahunan", // Kolom U (Index 20)
        link_evidence: data[i][21] || ""             // Kolom V (Index 21)
      };
    }
  }
  return null;
}

// ==========================================
// SUBMIT WORKFLOW
// ==========================================
function submitWorkflow(dataForm) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DataCuti");
  if (!sheet) {
    sheet = ss.insertSheet("DataCuti");
    sheet.appendRow([
      "ID Request", "Nama", "ID SPX", "Unit Kerja", "Team", "Email Karyawan", 
      "Tgl Bergabung", "Tgl Mulai", "Tgl Selesai", "Lama Cuti", "Alasan", "Backup",
      "TTD TL", "TTD SL", "TTD TP", "TTD Karyawan", "Status Pengajuan", "PDF URL", "Fitur Kolom S", "Link Approval TTD", "Kategori Cuti", "Link Evidence Foto"
    ]);
  }
  
  var role = dataForm.role ? dataForm.role.toLowerCase().trim() : "karyawan";
  var reqId = dataForm.reqId || "";
  var action = dataForm.action || "approve";
  var signatureData = dataForm.ttd3 || dataForm.ttd || dataForm.signature || dataForm.ttdKaryawan || "";

  // 1. PENGAJUAN AWAL OLEH KARYAWAN
  if (role === "karyawan" || reqId === "") {
    var newId = "REQ-" + new Date().getTime();
    var baseUrl = ScriptApp.getService().getUrl();
    var initialApprovalUrl = baseUrl + "?role=tl&id=" + newId;

    // Simpan foto evidence ke folder FOLDER_EVIDENCE_ID
    var linkEvidence = "-";
    if (dataForm.base64Evidence && dataForm.base64Evidence.length > 50) {
      try {
        var targetFolder;
        if (typeof FOLDER_EVIDENCE_ID !== 'undefined' && FOLDER_EVIDENCE_ID !== "") {
          targetFolder = DriveApp.getFolderById(FOLDER_EVIDENCE_ID);
        } else {
          var folderIter = DriveApp.getFoldersByName("Evidence_Cuti_SPX");
          targetFolder = folderIter.hasNext() ? folderIter.next() : DriveApp.createFolder("Evidence_Cuti_SPX");
        }

        var splitData = dataForm.base64Evidence.split(",");
        var contentType = splitData[0].match(/:(.*?);/)[1];
        var bytes = Utilities.base64Decode(splitData[1]);
        var blob = Utilities.newBlob(bytes, contentType, "EVIDENCE_" + dataForm.nama.replace(/[^a-zA-Z0-9]/g, "_") + "_" + newId + ".jpg");

        var fileEvidence = targetFolder.createFile(blob);
        fileEvidence.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        linkEvidence = fileEvidence.getUrl();
      } catch(e) {
        Logger.log("Gagal menyimpan foto evidence: " + e.toString());
      }
    }

    var newRowData = [
      "'" + newId,
      dataForm.nama,
      dataForm.id_spx,
      dataForm.unit_kerja,
      dataForm.team,
      dataForm.email_karyawan,
      dataForm.tgl_gabung,
      dataForm.tgl_mulai,
      dataForm.tgl_selesai,
      dataForm.lama_cuti,
      dataForm.alasan,
      dataForm.backup,
      "",              // Kolom M (Index 12) - TTD TL
      "",              // Kolom N (Index 13) - TTD SL
      "",              // Kolom O (Index 14) - TTD TP LEAD
      signatureData,   // Kolom P (Index 15) - TTD KARYAWAN
      "PENDING_TL",    // Kolom Q (Index 16)
      "",              // Kolom R (Index 17)
      "",              // Kolom S (Index 18)
      initialApprovalUrl,                       // Kolom T (Index 19)
      dataForm.kategori_cuti || "Cuti Tahunan", // Kolom U (Index 20)
      linkEvidence                              // Kolom V (Index 21)
    ];

    sheet.insertRowBefore(2);
    sheet.getRange(2, 1, 1, newRowData.length).setValues([newRowData]);
    SpreadsheetApp.flush();

    // Kirim Notifikasi SeaTalk ke Team Lead
    kirimNotifikasiSeaTalk(initialApprovalUrl, dataForm, "TL", newId);

    return { success: true, next: "TL", reqId: newId };
  }
  
  // 2. PROSES APPROVAL / REJECT
  var data = sheet.getDataRange().getDisplayValues();
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === reqId.toString().trim()) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) throw new Error("Data pengajuan ID (" + reqId + ") tidak ditemukan di database Spreadsheet.");

  function updateLinkApproval(nextRole) {
    var rawUrl = ScriptApp.getService().getUrl() + "?role=" + nextRole.toLowerCase().replace(" ", "") + "&id=" + reqId;
    sheet.getRange(rowIndex, 20).setValue(rawUrl);
    return rawUrl;
  }

  // A. TEAM LEAD (TL)
  if (role === "tl" || role === "team lead") {
    if (action === "reject") {
      sheet.getRange(rowIndex, 13).setValue("REJECTED");
      sheet.getRange(rowIndex, 17).setValue("REJECTED_BY_TL");
      sheet.getRange(rowIndex, 20).setValue("-");
      return { success: true, action: "rejected", by: "TL" };
    } else {
      sheet.getRange(rowIndex, 13).setValue(signatureData);
      sheet.getRange(rowIndex, 17).setValue("PENDING_SL");
      var nextUrl = updateLinkApproval("SL");
      
      kirimNotifikasiSeaTalk(nextUrl, dataForm, "SL", reqId);
      return { success: true, next: "SL" };
    }
    
  // B. SHIFT LEAD (SL)
  } else if (role === "sl" || role === "shift lead") {
    if (action === "reject") {
      sheet.getRange(rowIndex, 14).setValue("REJECTED");
      sheet.getRange(rowIndex, 17).setValue("REJECTED_BY_SL");
      sheet.getRange(rowIndex, 20).setValue("-");
      return { success: true, action: "rejected", by: "SL" };
    } else {
      sheet.getRange(rowIndex, 14).setValue(signatureData);
      sheet.getRange(rowIndex, 17).setValue("PENDING_TPLEAD");
      var nextUrl = updateLinkApproval("TP LEAD");
      
      kirimNotifikasiSeaTalk(nextUrl, dataForm, "TP LEAD", reqId);
      return { success: true, next: "TP LEAD" };
    }
    
  // C. TP LEAD
  } else if (role === "tp" || role === "tp lead" || role === "tplead") {
    if (action === "reject") {
      sheet.getRange(rowIndex, 15).setValue("REJECTED");
      sheet.getRange(rowIndex, 17).setValue("REJECTED_BY_TPLEAD");
      sheet.getRange(rowIndex, 20).setValue("-");
      return { success: true, action: "rejected", by: "TP LEAD" };
    } else {
      sheet.getRange(rowIndex, 15).setValue(signatureData);
      sheet.getRange(rowIndex, 17).setValue("APPROVED");
      
      SpreadsheetApp.flush();
      
      var rowDataUpdated = sheet.getRange(rowIndex, 1, 1, 22).getDisplayValues()[0];
      var pdfResult = buatDokumenPDFCuti(rowDataUpdated, dataForm);
      
      sheet.getRange(rowIndex, 18).setValue(pdfResult.fileUrl);
      sheet.getRange(rowIndex, 20).setValue("-");
      
      return { success: true, next: "APPROVED", pdfUrl: pdfResult.fileUrl, folderUrl: pdfResult.folderUrl };
    }
  }
  throw new Error("Akses Peran (Role) tidak valid.");
}

// ==========================================
// SEATALK NOTIFICATION INTEGRATION
// ==========================================
function kirimNotifikasiSeaTalk(approvalUrl, dataForm, targetRole, reqId) {
  if (!SEATALK_WEBHOOK_URL) return;

  var nama = dataForm.nama || "Karyawan";
  var team = dataForm.team || "-";
  var lama = dataForm.lama_cuti || "0";
  var alasan = dataForm.alasan || "-";
  var katCuti = dataForm.kategori_cuti || "Cuti Tahunan";

  var payload = {
    "tag": "text",
    "text": {
      "content": "📢 PEMBERITAHUAN PERMOHONAN FORM CUTI SPX EXPRESS\n\n" +
                 "• *Nama:* " + nama + "\n" +
                 "• *Team:* " + team + "\n" +
                 "• *Kategori:* " + katCuti + "\n" +
                 "• *Durasi:* " + lama + " Hari\n" +
                 "• *Alasan:* " + alasan + "\n\n" +
                 "Mohon untuk *" + targetRole + "* meninjau & memberikan TTD persetujuan melalui link di bawah ini:\n" + approvalUrl
    }
  };

  try {
    UrlFetchApp.fetch(SEATALK_WEBHOOK_URL, {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload)
    });
  } catch(e) {
    Logger.log("Gagal kirim Webhook SeaTalk: " + e.toString());
  }
}

// ==========================================
// GENERATE PDF FUNCTION
// ==========================================
function buatDokumenPDFCuti(rowData, dataForm) {
  if (!rowData || !Array.isArray(rowData)) {
    return { fileUrl: "", folderUrl: "", error: "Parameter rowData kosong." };
  }

  var targetFolder;
  try {
    if (typeof FOLDER_PDF_ID !== 'undefined' && FOLDER_PDF_ID !== "") {
      targetFolder = DriveApp.getFolderById(FOLDER_PDF_ID);
    } else { throw new Error(); }
  } catch(e) {
    var folderIter = DriveApp.getFoldersByName("PDF_Cuti_SPX");
    targetFolder = folderIter.hasNext() ? folderIter.next() : DriveApp.createFolder("PDF_Cuti_SPX");
  }

  try { targetFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}

  // Format Helper Tanggal (Menangani format teks YYYY-MM-DD maupun Date Object)
  var formatTanggal = function(val) {
    if (!val || val === "-") return "-";
    try {
      var d = new Date(val);
      if (isNaN(d.getTime())) return val; // Jika sudah berformat string dd/MM/yyyy
      return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
    } catch(e) {
      return val;
    }
  };

  var idReq     = rowData[0] || "REQ-" + new Date().getTime();
  var namaKary  = rowData[1] || "-";
  var idSpx     = rowData[2] || "-";
  var unitKerja = rowData[3] || "unit kerja kamu";
  var teamKary  = rowData[4] || "-";
  var tglGabung = formatTanggal(rowData[6]);
  var tglMulai  = formatTanggal(rowData[7]);
  var tglSelesi = formatTanggal(rowData[8]);
  var lamaCuti  = rowData[9] || "0";
  var alasan    = rowData[10] || "-";
  var backup    = rowData[11] || "-";
  var katCuti   = rowData[20] || "Cuti Tahunan";

  var ttdTL   = rowData[12] || "";
  var ttdSL   = rowData[13] || "";
  var ttdTP   = rowData[14] || "";
  var ttdKary = rowData[15] || "";

  var renderImg = function(val) {
    if (!val || val === "REJECTED") return "&nbsp;";
    var valStr = val.toString().trim();
    if (valStr.length > 50) {
      var base64Src = valStr.indexOf("data:image") === -1 ? "data:image/png;base64," + valStr : valStr;
      return '<img src="' + base64Src + '" style="height:55px; width:auto; max-width:110px;" />';
    }
    return "&nbsp;";
  };

  var htmlContent = `
  <html>
    <head>
      <style>
        body { font-family: 'Helvetica', 'Arial', sans-serif; color: #222; padding: 20px; line-height: 1.2; }
        .header { text-align: center; margin-bottom: 25px; }
        .logo-spx { font-size: 28px; font-weight: bold; font-style: italic; }
        .logo-orange { color: #ff5722; }
        .logo-dark { color: #2c3e50; font-style: normal; }
        .title { font-size: 13px; font-weight: bold; text-decoration: underline; margin-top: 5px; }
        
        .content-table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 12px; }
        .content-table td { padding: 4px 0; vertical-align: bottom; }
        .label-col { width: 140px; }
        .titik-col { width: 15px; }
        .value-col { border-bottom: 1px solid #000; padding-left: 5px; }
        .no-border td { border-bottom: none !important; padding: 15px 0 5px 0; }
        
        .matrix-table { width: 100%; border-collapse: collapse; margin-top: 30px; text-align: center; font-size: 11px; }
        .matrix-table td { width: 25%; vertical-align: bottom; text-align: center; }
        .sign-box { height: 60px; line-height: 60px; text-align: center; vertical-align: bottom; }
        .line-space { border-bottom: 1px dotted #444; margin: 5px auto 4px auto; width: 85%; }
        .job-title { font-weight: bold; color: #333; font-size: 10px; }
      </style>
    </head>
    <body>
      <div class="header">
        <div class="logo-spx"><span class="logo-orange">SPX</span> <span class="logo-dark">Express</span></div>
        <div class="title">FORMULIR PENGAJUAN CUTI KARYAWAN</div>
      </div>
      
      <table class="content-table">
        <tr><td class="label-col">Nama</td><td class="titik-col">:</td><td class="value-col">${namaKary}</td></tr>
        <tr><td class="label-col">Unit Kerja</td><td class="titik-col">:</td><td class="value-col">${unitKerja}</td></tr>
        <tr><td class="label-col">ID SPX</td><td class="titik-col">:</td><td class="value-col">${idSpx}</td></tr>
        <tr><td class="label-col">Team</td><td class="titik-col">:</td><td class="value-col">${teamKary}</td></tr>
        <tr><td class="label-col">Tanggal Bergabung</td><td class="titik-col">:</td><td class="value-col">${tglGabung}</td></tr>
        <tr class="no-border"><td colspan="3">Dengan ini mengajukan cuti,</td></tr>
        <tr><td class="label-col">Kategori Cuti</td><td class="titik-col">:</td><td class="value-col">${katCuti}</td></tr>
        <tr><td class="label-col">Selama</td><td class="titik-col">:</td><td class="value-col">${lamaCuti} Hari</td></tr>
        <tr><td class="label-col">Tanggal</td><td class="titik-col">:</td><td class="value-col">${tglMulai} s/d ${tglSelesi}</td></tr>
        <tr><td class="label-col">Alasan cuti</td><td class="titik-col">:</td><td class="value-col">${alasan}</td></tr>
        <tr><td class="label-col">Back-Up</td><td class="titik-col">:</td><td class="value-col">${backup}</td></tr>
      </table>
      
      <table class="matrix-table">
        <tr>
          <td>Menyetujui,</td>
          <td>Menyetujui,</td>
          <td>Menyetujui,</td>
          <td>Tangerang, ${tglMulai}</td>
        </tr>
        <tr>
          <td class="sign-box">${renderImg(ttdTL)}</td>
          <td class="sign-box">${renderImg(ttdSL)}</td>
          <td class="sign-box">${renderImg(ttdTP)}</td>
          <td class="sign-box">${renderImg(ttdKary)}</td>
        </tr>
        <tr>
          <td><div class="line-space"></div></td>
          <td><div class="line-space"></div></td>
          <td><div class="line-space"></div></td>
          <td><div class="line-space"></div></td>
        </tr>
        <tr>
          <td class="job-title">TP TEAM LEAD</td>
          <td class="job-title">SHIFT LEAD</td>
          <td class="job-title">TP LEAD</td>
          <td class="job-title">Yang Mengajukan</td>
        </tr>
      </table>
    </body>
  </html>
  `;

  var htmlBlob = Utilities.newBlob(htmlContent, "text/html", "Form_Cuti_Temp.html");
  var pdfBlob = htmlBlob.getAs("application/pdf");
  pdfBlob.setName("Form_Cuti_" + namaKary + "_" + idReq + ".pdf");
  
  var filePdf = targetFolder.createFile(pdfBlob);
  filePdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return { 
    fileUrl: filePdf.getUrl(),
    folderUrl: targetFolder.getUrl() 
  };
}
