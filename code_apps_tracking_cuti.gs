function doGetTracking(e) {
  return HtmlService.createTemplateFromFile('Index apps tracking cuti')
      .evaluate()
      .setTitle('Tracking Cuti Karyawan SPX')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function trackCutiBySpxId(spxId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('nama_sheet_database_kamu'); 
    
    if (!sheet) {
      return { status: 'error', message: 'Tab sheet "DataCuti" tidak ditemukan. Cek nama tab Anda.' };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { status: 'empty', message: 'Data di Sheet masih kosong.' };
    }

    // Ambil baris pertama (Header) untuk mencari indeks kolom
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().trim().toLowerCase());
    
    // Fungsi pembantu fleksibel untuk mencari kolom berdasarkan beberapa kata kunci
    const findIdx = (keywords) => {
      return headers.findIndex(h => keywords.some(kw => h.includes(kw)));
    };

    const idxSpxId = findIdx(['id spx', 'ops id', 'spx id', 'id ops']);
    const idxReqId = findIdx(['id request', 'req id', 'request id', 'no request']);
    const idxNama = findIdx(['nama']);
    const idxTglMulai = findIdx(['tgl mulai', 'tanggal mulai', 'start date']);
    const idxTglSelesai = findIdx(['tgl selesai', 'tanggal selesai', 'end date']);
    const idxAlasan = findIdx(['alasan', 'keterangan']);
    
    // Pencarian TTD yang fleksibel (mencari kombinasi kata TL, SL, TP)
    const idxTtdTL = findIdx(['tl', 'team lead']);
    const idxTtdSL = findIdx(['sl', 'shift lead']);
    const idxTtdTP = findIdx(['tp', 'tp lead']);
    const idxStatus = findIdx(['status']);

    if (idxSpxId === -1 || idxReqId === -1) {
      return { 
        status: 'error', 
        message: 'Header "ID SPX" atau "ID Request" tidak ditemukan di Sheet.' 
      };
    }

    // Ambil seluruh data sesuai baris & kolom yang ada
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const results = [];
    const searchId = spxId.toString().trim().toLowerCase();

    for (let i = data.length - 1; i >= 0; i--) {
      const row = data[i];
      const rowSpxId = row[idxSpxId] ? row[idxSpxId].toString().trim().toLowerCase() : '';

      if (rowSpxId === searchId) {
        const getVal = (idx) => (idx !== -1 && row[idx] !== undefined && row[idx] !== '') ? row[idx].toString().trim() : 'Belum';

        const ttdTL = getVal(idxTtdTL);
        const ttdSL = getVal(idxTtdSL);
        const ttdTP = getVal(idxTtdTP);
        const statusVal = idxStatus !== -1 && row[idxStatus] ? row[idxStatus].toString().trim() : 'PENDING';
        
        // Logika kelengkapan TTD
        const isApproved = (val) => {
          const v = val.toLowerCase();
          return v !== 'belum' && v !== '-' && v !== '' && !v.includes('pending');
        };

        const isComplete = isApproved(ttdTL) && isApproved(ttdSL) && isApproved(ttdTP);

        results.push({
          reqId: row[idxReqId] || '-',
          nama: idxNama !== -1 ? row[idxNama] : '-',
          tglMulai: idxTglMulai !== -1 ? formatDate(row[idxTglMulai]) : '-',
          tglSelesai: idxTglSelesai !== -1 ? formatDate(row[idxTglSelesai]) : '-',
          alasan: idxAlasan !== -1 ? row[idxAlasan] : '-',
          ttdTL: ttdTL,
          ttdSL: ttdSL,
          ttdTP: ttdTP,
          status: statusVal,
          canDownloadPdf: isComplete
        });
      }
    }

    if (results.length === 0) {
      return { status: 'empty', message: 'ID SPX / Ops ID tidak ditemukan.' };
    }

    return { status: 'success', data: results };

  } catch (error) {
    return { status: 'error', message: 'Backend Error: ' + error.toString() };
  }
}

function formatDate(dateVal) {
  if (!dateVal) return '-';
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'dd-MM-yyyy');
  }
  return dateVal.toString();
}

function generateCutiPdf(reqId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('DataCuti');
    
    if (!sheet) return { status: 'error', message: 'Tab sheet "DataCuti" tidak ditemukan.' };

    const lastRow = sheet.getLastRow();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().trim().toLowerCase());
    
    const idxReqId = headers.findIndex(h => h.includes('request') || h.includes('req'));
    const idxPdf = 17; // Kolom R (Index 17)

    if (idxReqId === -1) return { status: 'error', message: 'Kolom ID Request tidak ditemukan.' };

    const reqIds = sheet.getRange(1, idxReqId + 1, lastRow, 1).getValues();

    let targetRow = -1;
    for (let i = 1; i < reqIds.length; i++) {
      if (reqIds[i][0] == reqId) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) return { status: 'error', message: 'ID Request tidak ditemukan.' };

    const pdfData = sheet.getRange(targetRow, idxPdf + 1).getValue();

    if (!pdfData) return { status: 'error', message: 'Isi Kolom R masih kosong.' };

    return {
      status: 'success',
      pdfUrl: pdfData.toString().trim()
    };

  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}
