// ==========================================
// REMINDER SEATALK PENDING & CUTI MENDEKATI TANGGAL (H-1)
// ==========================================

function kirimReminderPending() {
  var webhookUrl = "WEBHOOK_PUNYA_KAMU_TEMPEL_DISINI";
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("nama_sheet_punya_kamu");
  if (!sheet) {
    Logger.log("❌ ERROR: Sheet 'nama_sheet_punya_kamu' tidak ditemukan!");
    return;
  }

  var data = sheet.getDataRange().getDisplayValues();
  var listPendingData = [];

  // Hitung tanggal besok (Format YYYY-MM-DD / DD/MM/YYYY)
  var today = new Date();
  var tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);
  
  var tomorrowFormattedStr = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), "yyyy-MM-dd");
  var tomorrowFormattedIndo = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), "dd/MM/yyyy");

  // 1. Kumpulkan semua data PENDING
  for (var i = 1; i < data.length; i++) {
    var nama        = data[i][1];  // Kolom B: Nama
    var team        = data[i][4];  // Kolom E: Team
    var rawTglMulai = data[i][7];  // Kolom H: Tgl Mulai Cuti
    var tglSelesai  = data[i][8];  // Kolom I: Tgl Selesai Cuti
    var lamaCuti    = data[i][9];  // Kolom J: Lama Cuti
    var status      = data[i][16] ? data[i][16].toString().trim().toUpperCase() : ""; // Kolom Q
    var linkAppr    = data[i][19]; // Kolom T: Link Approval TTD

    if (status.indexOf("PENDING") !== -1) {
      // Cek apakah tanggal cuti adalah BESOK (H-1)
      var isUrgentTomorrow = false;
      if (rawTglMulai) {
        try {
          var tglMulaiObj = new Date(rawTglMulai);
          var tglMulaiStr = Utilities.formatDate(tglMulaiObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
          var tglMulaiIndo = Utilities.formatDate(tglMulaiObj, Session.getScriptTimeZone(), "dd/MM/yyyy");

          if (tglMulaiStr === tomorrowFormattedStr || tglMulaiIndo === tomorrowFormattedIndo || rawTglMulai.trim() === tomorrowFormattedIndo) {
            isUrgentTomorrow = true;
          }
        } catch(e) {
          if (rawTglMulai.trim() === tomorrowFormattedIndo) {
            isUrgentTomorrow = true;
          }
        }
      }

      listPendingData.push({
        nama: nama || "-",
        team: team || "-",
        tglMulai: rawTglMulai || "-",
        tglSelesai: tglSelesai || "-",
        lamaCuti: lamaCuti || "0",
        status: status,
        linkAppr: linkAppr || "-",
        isUrgent: isUrgentTomorrow
      });
    }
  }

  // Urutkan data: Yang URGENT (Besok Cuti) ditaruh paling ATAS!
  listPendingData.sort(function(a, b) {
    return (b.isUrgent === true ? 1 : 0) - (a.isUrgent === true ? 1 : 0);
  });

  var totalPending = listPendingData.length;
  Logger.log("📊 Total data status PENDING ditemukan: " + totalPending);

  if (totalPending === 0) {
    Logger.log("⚠️ Tidak ada data PENDING untuk dikirim.");
    return;
  }

  // 2. Maksimal 8 data per pesan agar aman & muat banyak
  var BATCH_SIZE = 8; 
  var totalBatch = Math.ceil(totalPending / BATCH_SIZE);

  for (var b = 0; b < totalBatch; b++) {
    var currentBatch = listPendingData.slice(b * BATCH_SIZE, (b + 1) * BATCH_SIZE);
    
    var pesan = "📌 *FOLLOW UP CUTI BELUM APPROVE (" + (b + 1) + "/" + totalBatch + ")*\n" +
                "Total: *" + totalPending + " pengajuan pending*\n" +
                "───────────────────────────\n\n";

    currentBatch.forEach(function(item, index) {
      var nomor = (b * BATCH_SIZE) + index + 1;
      
      // Jika Cuti BESOK, beri label/badge khusus 🚨 URGENT
      var badgeUrgent = item.isUrgent ? "🚨 *[URGENT: CUTI BESOK!]*\n   " : "";

      pesan += nomor + ". " + badgeUrgent + "*" + item.nama + "* (" + item.team + ")\n" +
               "   🗓️ Mulai: " + item.tglMulai + "\n" +
               "   🏁 Selesai: " + item.tglSelesai + " (" + item.lamaCuti + " Hari)\n" +
               "   ⏳ Status: *" + item.status + "*\n" +
               "   🔗 Link: " + item.linkAppr + "\n\n";
    });

    if (b === totalBatch - 1) {
      pesan += "───────────────────────────\nMohon dibantu untuk TTD-nya Bapak/Ibu 🙏";
    }

    // Kirim ke SeaTalk
    var payload = {
      "tag": "text",
      "text": { "content": pesan }
    };

    try {
      var response = UrlFetchApp.fetch(webhookUrl, {
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify(payload),
        "muteHttpExceptions": true
      });
      Logger.log("🚀 Batch " + (b + 1) + " terkirim: " + response.getContentText());
    } catch(e) {
      Logger.log("❌ Gagal kirim Batch " + (b + 1) + ": " + e.toString());
    }

    Utilities.sleep(1000); // Jeda 1 detik
  }
}
