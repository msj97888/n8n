// ============================================================
// 1M1H Logistics — Quotation Auto-Writer
// Google Apps Script — Gmail trigger version
//
// HOW TO SET UP:
// 1. Paste this entire file into Apps Script (script.google.com)
// 2. Fill in the 3 constants below
// 3. Run setupTrigger() ONCE manually to install the 5-min trigger
// 4. Done — it will poll Gmail every 5 min automatically
// ============================================================

// ── CONFIG ─────────────────────────────────────────────────
const SPREADSHEET_ID       = '1hFwG8hSnGDQV9AOUv8hXTFq0HUDBlyOYtxDEHfPLx2M';
const TRIGGER_SENDER_EMAIL = 'YOUR_GMAIL_ADDRESS@gmail.com'; // the address you forward rates to (your own gmail)
const GEMINI_API_KEY       = 'REPLACE_WITH_YOUR_GEMINI_API_KEY';

// Markup config — adjust these to change selling rates
const MARKUP = {
  airFreightPerKg:    0.50,  // USD/kg added on top of airline rate
  fuelRatePerKg:      0.10,
  secRatePerKg:       0.05,
  screenRatePerKg:    0.10,
  minChargeableWeight: 45    // minimum billable CW in kg
};
// ── END CONFIG ─────────────────────────────────────────────


// ─────────────────────────────────────────────────────────────
// MAIN: called every 5 minutes by time-based trigger
// ─────────────────────────────────────────────────────────────
function checkGmailAndWriteQuotation() {
  // Search for unread emails forwarded to this inbox that haven't been labelled yet
  const query   = 'is:unread label:RATES-RECEIVED';
  const threads = GmailApp.search(query, 0, 10);

  if (threads.length === 0) return;

  const ss            = SpreadsheetApp.openById(SPREADSHEET_ID);
  const templateSheet = ss.getSheetByName('TEMPLATE');
  const quotSheet     = ss.getSheetByName('QUOTATION');

  if (!templateSheet || !quotSheet) {
    Logger.log('ERROR: TEMPLATE or QUOTATION tab not found');
    return;
  }

  for (const thread of threads) {
    const message = thread.getMessages()[thread.getMessageCount() - 1]; // latest message
    const subject = message.getSubject();
    const body    = message.getPlainBody();

    Logger.log('Processing email: ' + subject);

    try {
      // 1. Extract rates via Gemini
      const extracted = callGemini(subject, body);
      if (!extracted) {
        Logger.log('Gemini returned nothing for: ' + subject);
        continue;
      }

      // 2. Compute selling rates with markup
      const data = applyMarkup(extracted);

      // 3. Get next quote number
      data.quoteNumber = getNextQuoteNumber(quotSheet);

      // 4. Copy template → write quotation
      writeQuotation(ss, templateSheet, quotSheet, data);

      // 5. Mark thread as read and apply label so it won't be processed again
      thread.markRead();
      applyProcessedLabel(thread);

      Logger.log('Quotation written. Quote#: ' + data.quoteNumber + ' | POD: ' + data.pod);

    } catch (err) {
      Logger.log('Error processing thread "' + subject + '": ' + err.message);
    }
  }
}


// ─────────────────────────────────────────────────────────────
// GEMINI: extract rates from email
// ─────────────────────────────────────────────────────────────
function callGemini(subject, body) {
  const prompt = `You are a freight rate extractor for an air freight company.
Read the email below and extract the airline's quoted rates.
Return ONLY valid JSON — no explanation, no markdown.

Email Subject: ${subject}
Email Body: ${body}

Extract these fields:
- pod: destination airport/city (e.g. "LAX", "Los Angeles", "JFK") — get from subject if possible
- airlineCode: airline name or IATA code mentioned
- transitTime: transit time if mentioned (e.g. "3-4 days"), else ""
- airFreightRatePerKg: the base air freight rate per kg as a number
- fuelRatePerKg: fuel surcharge per kg as a number (0 if not mentioned)
- secRatePerKg: security surcharge per kg as a number (0 if not mentioned)
- screenRatePerKg: screening surcharge per kg as a number (0 if not mentioned)
- chargeableWeightKg: chargeable weight in kg if mentioned, else 0
- currency: currency code (default "USD")`;

  const url     = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`;
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: {
        type: 'OBJECT',
        properties: {
          pod:                 { type: 'STRING' },
          airlineCode:         { type: 'STRING' },
          transitTime:         { type: 'STRING' },
          airFreightRatePerKg: { type: 'NUMBER' },
          fuelRatePerKg:       { type: 'NUMBER' },
          secRatePerKg:        { type: 'NUMBER' },
          screenRatePerKg:     { type: 'NUMBER' },
          chargeableWeightKg:  { type: 'NUMBER' },
          currency:            { type: 'STRING' }
        },
        required: ['pod', 'airlineCode', 'airFreightRatePerKg']
      }
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method:      'post',
    contentType: 'application/json',
    payload:     JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const json = JSON.parse(response.getContentText());
  const text = json?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) return null;

  return JSON.parse(text);
}


// ─────────────────────────────────────────────────────────────
// Apply markup to raw extracted rates
// ─────────────────────────────────────────────────────────────
function applyMarkup(extracted) {
  const cwRaw    = extracted.chargeableWeightKg || 0;
  const billableCW = Math.max(MARKUP.minChargeableWeight, cwRaw);

  return {
    pod:               extracted.pod          || '',
    airlineCode:       extracted.airlineCode  || '',
    transitTime:       extracted.transitTime  || '',
    currency:          extracted.currency     || 'USD',
    billableCW:        billableCW,
    sellingAirFreight: round2((extracted.airFreightRatePerKg || 0) + MARKUP.airFreightPerKg),
    fuelRate:          round2(Math.max(extracted.fuelRatePerKg   || 0, MARKUP.fuelRatePerKg)),
    secRate:           round2(Math.max(extracted.secRatePerKg    || 0, MARKUP.secRatePerKg)),
    screenRate:        round2(Math.max(extracted.screenRatePerKg || 0, MARKUP.screenRatePerKg)),
  };
}

function round2(n) { return Math.round(n * 100) / 100; }


// ─────────────────────────────────────────────────────────────
// Get next sequential quote number from QUOTATION sheet
// ─────────────────────────────────────────────────────────────
function getNextQuoteNumber(quotSheet) {
  const today   = new Date();
  const yy      = String(today.getFullYear()).slice(-2);
  const mm      = String(today.getMonth() + 1).padStart(2, '0');
  const prefix  = `QT${yy}${mm}-`;

  // Find all existing quote numbers in column A that match this month's prefix
  const data    = quotSheet.getDataRange().getValues();
  let maxSeq    = 0;

  for (const row of data) {
    const cell = String(row[0] || '');
    if (cell.startsWith(prefix)) {
      const seq = parseInt(cell.replace(prefix, ''), 10);
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  }

  return prefix + String(maxSeq + 1).padStart(3, '0');
}


// ─────────────────────────────────────────────────────────────
// Copy TEMPLATE → QUOTATION and fill in values
// ─────────────────────────────────────────────────────────────
function writeQuotation(ss, templateSheet, quotSheet, data) {
  // Find next empty slot (2 blank rows gap between quotations)
  const lastRow  = quotSheet.getLastRow();
  const startRow = lastRow > 0 ? lastRow + 2 : 1;

  // Copy entire TEMPLATE block (preserves merges, bold, formulas, currency format)
  const tmplRows  = templateSheet.getLastRow();
  const tmplCols  = templateSheet.getLastColumn();
  const tmplRange = templateSheet.getRange(1, 1, tmplRows, tmplCols);
  const destCell  = quotSheet.getRange(startRow, 1);
  tmplRange.copyTo(destCell, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  // Scan column A of the copied block to find label rows
  const colA = quotSheet.getRange(startRow, 1, tmplRows, 1).getValues();

  // Labels that get RATE (col B) + UNIT (col C)
  const rateRows = {
    'AIR FREIGHT CHARGE':  [data.sellingAirFreight, data.billableCW],
    'FUEL SURCHARGE':      [data.fuelRate,          'PER KG'],
    'SECURITY SURCHARGE':  [data.secRate,            'PER KG'],
    'SCREENING SURCHARGE': [data.screenRate,         'PER KG'],
  };

  // Labels that get a single value (col B)
  const singleRows = {
    'QUOTE NUMBER':   data.quoteNumber  || '',
    'QUOTE NO':       data.quoteNumber  || '',
    'QUOTE NO.':      data.quoteNumber  || '',
    'POD :':          data.pod          || '',
    'POD:':           data.pod          || '',
    'AIR CARRIER :':  data.airlineCode  || '',
    'AIR CARRIER:':   data.airlineCode  || '',
    'TRANSIT TIME :': data.transitTime  || '',
    'TRANSIT TIME:':  data.transitTime  || '',
  };

  for (let i = 0; i < colA.length; i++) {
    const label  = String(colA[i][0] || '').trim().toUpperCase();
    const absRow = startRow + i;

    if (rateRows[label]) {
      const [rate, unit] = rateRows[label];
      quotSheet.getRange(absRow, 2).setValue(rate);
      quotSheet.getRange(absRow, 3).setValue(unit);
    }

    if (singleRows.hasOwnProperty(label)) {
      quotSheet.getRange(absRow, 2).setValue(singleRows[label]);
    }
  }
}


// ─────────────────────────────────────────────────────────────
// Gmail label helpers
// ─────────────────────────────────────────────────────────────
function applyProcessedLabel(thread) {
  let label = GmailApp.getUserLabelByName('RATES-PROCESSED');
  if (!label) label = GmailApp.createLabel('RATES-PROCESSED');
  thread.addLabel(label);

  // Remove the RATES-RECEIVED label so it won't be picked up again
  const ratesLabel = GmailApp.getUserLabelByName('RATES-RECEIVED');
  if (ratesLabel) thread.removeLabel(ratesLabel);
}


// ─────────────────────────────────────────────────────────────
// ONE-TIME SETUP: run this manually once to install the trigger
// ─────────────────────────────────────────────────────────────
function setupTrigger() {
  // Delete any existing triggers for this function to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'checkGmailAndWriteQuotation') {
      ScriptApp.deleteTrigger(t);
    }
  }

  // Create a new trigger: every 5 minutes
  ScriptApp.newTrigger('checkGmailAndWriteQuotation')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('Trigger installed — checkGmailAndWriteQuotation runs every 5 minutes.');
}
