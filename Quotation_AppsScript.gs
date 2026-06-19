// ============================================================
// 1M1H Quotation Writer — Google Apps Script
// Deploy as Web App: Execute as Me, Access: Anyone
// ============================================================

const SPREADSHEET_ID = '1hFwG8hSnGDQV9AOUv8hXTFq0HUDBlyOYtxDEHfPLx2M';
const SECRET_TOKEN   = 'REPLACE_WITH_YOUR_SECRET_TOKEN';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // Simple auth check
    if (data.token !== SECRET_TOKEN) {
      return respond({ status: 'unauthorized' });
    }

    const ss            = SpreadsheetApp.openById(SPREADSHEET_ID);
    const templateSheet = ss.getSheetByName('TEMPLATE');
    const quotSheet     = ss.getSheetByName('QUOTATION');

    if (!templateSheet || !quotSheet) {
      return respond({ status: 'error', message: 'TEMPLATE or QUOTATION tab not found' });
    }

    // --- Find next empty row in QUOTATION ---
    const lastRow  = quotSheet.getLastRow();
    const startRow = lastRow > 0 ? lastRow + 2 : 1;

    // --- Copy entire template to QUOTATION (preserves merge, format, formulas) ---
    const tmplRows = templateSheet.getLastRow();
    const tmplCols = templateSheet.getLastColumn();
    const tmplRange = templateSheet.getRange(1, 1, tmplRows, tmplCols);
    const destCell  = quotSheet.getRange(startRow, 1);
    tmplRange.copyTo(destCell, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

    // --- Scan column A of copied block to find label rows ---
    const colA = quotSheet.getRange(startRow, 1, tmplRows, 1).getValues();

    // rate rows: [rate, unit] → columns B and C
    const rateTargets = {
      'AIR FREIGHT CHARGE':  [data.sellingAirFreight, data.billableCW],
      'FUEL SURCHARGE':      [data.fuelRate,          data.fuelUnit],
      'SECURITY SURCHARGE':  [data.secRate,           data.secUnit],
      'SCREENING SURCHARGE': [data.screenRate,        data.screenUnit],
    };

    // single-value rows: value → column B only
    const singleTargets = {
      'POD :':          data.pod          || '',
      'AIR CARRIER :':  data.airlineCode  || '',
      'TRANSIT TIME :': data.transitTime  || '',
    };

    for (let i = 0; i < colA.length; i++) {
      const label  = String(colA[i][0] || '').trim().toUpperCase();
      const absRow = startRow + i;

      if (rateTargets[label]) {
        const [rate, unit] = rateTargets[label];
        quotSheet.getRange(absRow, 2).setValue(rate);   // B = RATE
        quotSheet.getRange(absRow, 3).setValue(unit);   // C = UNIT
      }

      if (singleTargets.hasOwnProperty(label)) {
        quotSheet.getRange(absRow, 2).setValue(singleTargets[label]); // B = value
      }
    }

    return respond({ status: 'ok', startRow, message: 'Quotation written successfully' });

  } catch (err) {
    return respond({ status: 'error', message: err.message });
  }
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
