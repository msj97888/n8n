# 1M1H RATE Workflow — Node-by-Node Build Guide (Outlook + Gemini)

**Rebuilt version.** This corrects a real wiring bug from the previous version: "Read Quote Log" and "Read Airline Contacts" were chained inline into the main flow. Since both return every row in their sheet as separate items, chaining them meant the node right after each one was receiving a sheet row instead of your actual shipment data — `Compute Quote Number + Airlines` and `Build Draft Fields` would have been reading garbage. The fix: both lookup sheets now run as independent branches straight off the Schedule Trigger, and the rest of the pipeline reads them by reference (`$('Read Quote Log').all()`, `$('Read Airline Contacts').all()`) rather than by direct connection. Node names below match your actual canvas (`Get many messages`, `Update a message`, etc.) rather than the names from the first draft.

All the math (quote number sequencing, the most-recent-airline lookup, subject/body construction, the Gemini-response parser) was unit-tested in plain Node.js against your real sheet data before going into these Code nodes, including the sequence-gap edge case, a brand-new POD, and a failed extraction. That logic is unchanged from before — only the wiring and node-name references changed.

---

## 0. Prerequisites

**Credentials:**
- Microsoft Outlook (OAuth2)
- Google Sheets (OAuth2)
- Gemini API key (Google AI Studio, paid Tier 2)

**Outlook setup:**
- Category `1M1H RATE` — what the workflow polls for.
- Category `1M1H RATE - PROCESSED` (or similar) — applied once an inquiry is handled, so the next poll doesn't pick it up again.

**Sheets:**
1. **Quote Log** — your existing sheet (`DATE | EMAIL SUBJECT | QUOTE NUMBER | AGENT/CUSTOMER | AIRLINE | DEST | CW (KGS) | SENT BY | STATUS | REMARK`). Doubles as both the log and the airline-routing history source.
2. **Airline Contacts** — `AIRLINE CODE` / `EMAIL ADDRESS` columns assumed. If yours differ, update the two bracket references flagged in Node 9.

**Workflow setting to check once:** Settings → Execution Order should be **v1 (recommended)**, not the v0 legacy mode. v1 completes one branch fully before starting the next, ordered top-to-bottom on the canvas — this is what makes the corrected wiring below actually safe. v0 interleaves branches node-by-node instead, which would break the ordering this guide depends on.

---

## Node 1 — Schedule Trigger

**Type:** Schedule Trigger
Interval: every 5 minutes (adjust to taste).

This node now has **three** outgoing connections instead of one — see the wiring diagram at the end.

---

## Node 2 — "Get many messages" (Microsoft Outlook)

**Type:** Microsoft Outlook
**Resource:** Message · **Operation:** Get Many · **Return All:** true
**Folder:** Inbox
**Filters → Filter Query:**
```
categories/any(c:c eq '1M1H RATE')
```
This is the verified, working OData syntax for category filtering — `startswith()`/`contains()` are not supported on this field, only `eq`.

Confirm the output includes `subject`, `body.content` (or at minimum `bodyPreview`), `id`, and `from`. If `body` isn't present, add a Select/Fields option requesting `subject,body,bodyPreview,from,receivedDateTime,id,categories`.

---

## Node 3 — "Gemini - Extract Shipment Details" (HTTP Request)

**Method:** POST
**URL:** `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent`
**Headers:** `x-goog-api-key` = your Gemini API key
**Body Content Type:** JSON, field switched to **expression mode** (`fx` toggle):

```javascript
{
  "contents": [
    {
      "role": "user",
      "parts": [
        {
          "text": "Extract air freight shipment details from this rate inquiry email body. If a field is genuinely not present, return an empty string for text fields or null for numbers. Email body:\n\n" + ((($json.body && $json.body.content) ? $json.body.content : ($json.bodyPreview || '')).replace(/<[^>]+>/g, ' '))
        }
      ]
    }
  ],
  "generationConfig": {
    "responseMimeType": "application/json",
    "responseSchema": {
      "type": "object",
      "properties": {
        "pol": { "type": "string", "description": "Port of loading / origin airport code, e.g. ORD" },
        "pod": { "type": "string", "description": "Port of discharge / destination airport code. Empty string if not found." },
        "shipper": { "type": "string", "description": "Shipper name, or the literal word UNKNOWN if the email says unknown shipper" },
        "cargoType": { "type": "string", "description": "Cargo type as written, e.g. GENERAL CARGO" },
        "pieces": { "type": "string", "description": "Piece/crate count description exactly as written, e.g. '1 CRATE'" },
        "dimensions": { "type": "string", "description": "Dimensions exactly as written, e.g. '65X53X65 INCHES'" },
        "grossWeightKg": { "type": "number", "description": "Total gross weight in kg, numeric only" },
        "chargeableWeightKg": { "type": "number", "description": "Total chargeable weight in kg, numeric only" }
      },
      "required": ["pol", "pod", "shipper", "cargoType", "pieces", "dimensions", "grossWeightKg", "chargeableWeightKg"]
    }
  }
}
```

This version prefers the full `body.content` field but automatically falls back to `bodyPreview` if `body` isn't present on the item — so it won't hard-fail either way. Worth still checking the Outlook node's field selection so you're getting the untruncated `body`, since `bodyPreview` caps at 255 characters and your inquiry format (8 bullet lines) can run past that.

---

## Node 4 — "Parse Extraction" (Code)

**Mode:** Run Once for Each Item

```javascript
const geminiResponse = $input.item.json;
const rawText = geminiResponse?.candidates?.[0]?.content?.parts?.[0]?.text;

let parsed = {};
let needsManualReview = false;

if (!rawText) {
  needsManualReview = true;
} else {
  const cleaned = rawText.replace(/^```json\s*/i, '').replace(/```\s*$/i, '').trim();
  try {
    parsed = JSON.parse(cleaned);
  } catch (e) {
    needsManualReview = true;
  }
}

if (!parsed.pod || String(parsed.pod).trim() === '') {
  needsManualReview = true;
}

// References the upstream Outlook node by its actual name on your canvas.
const sourceEmail = $('Get many messages').item.json;

return {
  json: {
    messageId: sourceEmail.id,
    originalSubject: sourceEmail.subject,
    pol: parsed.pol || '',
    pod: parsed.pod || '',
    shipper: parsed.shipper || 'UNKNOWN',
    cargoType: parsed.cargoType || '',
    pieces: parsed.pieces || '',
    dimensions: parsed.dimensions || '',
    grossWeightKg: parsed.grossWeightKg ?? '',
    chargeableWeightKg: parsed.chargeableWeightKg ?? '',
    needsManualReview,
  },
};
```

---

## Node 5 — "Read Quote Log" (Google Sheets)

**Operation:** Get Row(s) — no filters, return all rows.

**Wiring — this is the fix:** connect this node's input directly from **Schedule Trigger**, as its own branch. Do **not** chain it after Parse Extraction. Place it above the `Get many messages` branch on the canvas so it finishes first (see Execution Order note above).

Its output isn't consumed by a direct connection downstream — `Compute Quote Number + Airlines` reads it by name via `$('Read Quote Log').all()`.

---

## Node 6 — "Compute Quote Number + Airlines" (Code)

**Mode:** Run Once for Each Item
**Wiring:** input comes directly from **Parse Extraction** now (not from Read Quote Log).

```javascript
function pad2(n) { return String(n).padStart(2, '0'); }
function normalizeDate(raw) {
  if (!raw) return '';
  return new Date(raw).toISOString().slice(0, 10);
}

const todayISO = $now.toFormat('yyyy-LL-dd');
const todayYYMMDD = $now.toFormat('yyLLdd');

const logRows = $('Read Quote Log').all().map(i => i.json);

// ---- 1) Quote number: highest existing sequence for today, + 1 ----
let maxSeq = 0;
for (const row of logRows) {
  if (normalizeDate(row['DATE']) !== todayISO) continue;
  const match = /^AECQ(\d{6})(\d{2})$/.exec(row['QUOTE NUMBER'] || '');
  if (match && match[1] === todayYYMMDD) {
    const seq = parseInt(match[2], 10);
    if (seq > maxSeq) maxSeq = seq;
  }
}
const quoteNumber = `AECQ${todayYYMMDD}${pad2(maxSeq + 1)}`;

// ---- 2) Most recent airline(s) for this POD ----
const current = $input.item.json;
let airlineCodes = [];
let fallbackReason = null;

if (current.needsManualReview) {
  airlineCodes = [null];
  fallbackReason = 'Could not extract POD automatically - flagged for manual review';
} else {
  const matches = logRows.filter(
    r => (r['DEST'] || '').trim().toUpperCase() === current.pod.trim().toUpperCase()
  );

  if (matches.length === 0) {
    airlineCodes = [null];
    fallbackReason = `No airline history on file for POD "${current.pod}"`;
  } else {
    matches.sort((a, b) => new Date(normalizeDate(b['DATE'])) - new Date(normalizeDate(a['DATE'])));
    const mostRecentDate = normalizeDate(matches[0]['DATE']);
    const mostRecentRows = matches.filter(r => normalizeDate(r['DATE']) === mostRecentDate);

    const codes = new Set();
    for (const row of mostRecentRows) {
      (row['AIRLINE'] || '').split('/').forEach(code => {
        const trimmed = code.trim();
        if (trimmed) codes.add(trimmed);
      });
    }
    airlineCodes = Array.from(codes);
  }
}

return {
  json: {
    ...current,
    quoteNumber,
    airlineCodes,
    fallbackReason,
  },
};
```

---

## Node 7 — "Split Out Airlines" (Split Out)

**Field To Split Out:** `airlineCodes`
**Destination Field Name:** `airlineCode`

Unchanged.

---

## Node 8 — "Read Airline Contacts" (Google Sheets)

**Operation:** Get Row(s) — no filters, return all rows.

**Wiring — the other half of the fix:** connect this node's input directly from **Schedule Trigger** as well, as its own independent branch (alongside Read Quote Log and the main Get-many-messages branch). Do **not** chain it after Split Out Airlines. Place it above the `Get many messages` branch on the canvas too.

`Build Draft Fields` reads it by name via `$('Read Airline Contacts').all()`, not by direct connection.

---

## Node 9 — "Build Draft Fields" (Code)

**Mode:** Run Once for Each Item
**Wiring:** input comes directly from **Split Out Airlines** now (not from Read Airline Contacts).

**If your Airline Contacts sheet headers differ from `AIRLINE CODE` / `EMAIL ADDRESS`, update those two references below.**

```javascript
const current = $input.item.json;
const contacts = $('Read Airline Contacts').all().map(i => i.json);

let contactEmail = '';
if (current.airlineCode) {
  const match = contacts.find(
    c => (c['AIRLINE CODE'] || '').trim().toUpperCase() === current.airlineCode.trim().toUpperCase()
  );
  if (match) contactEmail = match['EMAIL ADDRESS'] || '';
}

let remark;
if (current.fallbackReason) {
  remark = current.fallbackReason;
} else if (!contactEmail) {
  remark = `No contact email on file for airline code "${current.airlineCode}"`;
} else {
  remark = '';
}

function buildSubject() {
  const cargoSegment = current.cargoType && current.cargoType.toUpperCase() === 'GENERAL CARGO' ? ' / GENERAL CARGO' : '';
  const airlineTag = current.airlineCode || 'UNASSIGNED';
  return `[${airlineTag}] ${current.pol}-${current.pod || '?'} / RATE QUOTE${cargoSegment} / ${current.quoteNumber}`;
}

function buildBody() {
  const shipperLine = current.shipper && current.shipper.toUpperCase() !== 'UNKNOWN'
    ? current.shipper.toUpperCase() + ' SHIPPER'
    : 'UNKNOWN SHIPPER';
  return [
    'Hello team,',
    '',
    'Please check the below shipment and provide us with rate and flight frequency.',
    '',
    `* POL : ${current.pol}`,
    `* POD : ${current.pod}`,
    `* ${shipperLine}`,
    `* ${current.cargoType}`,
    `* TOTAL ${current.pieces}`,
    `* ${current.dimensions}`,
    `* TOTAL G.W ${current.grossWeightKg} KGS`,
    `* TOTAL C.W ${current.chargeableWeightKg} KGS`,
    '',
    'Thank you.',
  ].join('\n');
}

return {
  json: {
    ...current,
    to: contactEmail,
    subject: buildSubject(),
    body: buildBody(),
    remark,
  },
};
```

---

## Node 10 — "Create Draft" (Microsoft Outlook)

**Resource:** Draft · **Operation:** Create
- To Recipients: `{{$json.to}}`
- Subject: `{{$json.subject}}`
- Message Body: `{{$json.body}}` · Body Content Type: Text

Unchanged.

---

## Node 11 — "Log Quote Row" (Google Sheets)

**Operation:** Append Row

| Column | Value |
|---|---|
| DATE | `{{$now.toFormat('yyyy-LL-dd')}}` |
| EMAIL SUBJECT | `{{$json.originalSubject}}` |
| QUOTE NUMBER | `{{$json.quoteNumber}}` |
| AGENT/CUSTOMER | `{{$json.shipper}}` |
| AIRLINE | `{{$json.airlineCode || ''}}` |
| DEST | `{{$json.pod}}` |
| CW (KGS) | `{{$json.chargeableWeightKg}}` |
| SENT BY | *(blank)* |
| STATUS | `DRAFT` |
| REMARK | `{{$json.remark}}` |

Unchanged.

---

## Node 12 — "Update a message" (Microsoft Outlook)

**Resource:** Message · **Operation:** Update
- Message ID: `{{$json.messageId}}`
- Update Fields → Category Names or IDs: `1M1H RATE - PROCESSED`

Fires once per airline-draft item; harmless if it sets the same category twice for a two-airline inquiry.

---

## Wiring it together (corrected)

```
Schedule Trigger
  ├─→ Read Quote Log                         (independent branch, runs first)
  ├─→ Read Airline Contacts                  (independent branch, runs second)
  └─→ Get many messages
        → Gemini - Extract Shipment Details
        → Parse Extraction
        → Compute Quote Number + Airlines     ← fed directly by Parse Extraction
        → Split Out Airlines
        → Build Draft Fields                  ← fed directly by Split Out Airlines
        → Create Draft
        → Log Quote Row
        → Update a message
```

Put "Read Quote Log" and "Read Airline Contacts" higher up on the canvas than the "Get many messages" branch — with Execution Order set to v1, that vertical position is what guarantees both lookups finish before the main branch needs them.

---

## What changed vs. what's still untested

**Fixed this round:** the branch wiring for Read Quote Log and Read Airline Contacts (was the actual cause of bad data reaching Nodes 6 and 9), and the node-name reference in Node 4 (now matches `Get many messages` instead of a name that didn't exist on your canvas).

**Still worth one live test:** the body/bodyPreview fallback in Node 3 — run it once and check whether `$json.body.content` or the `bodyPreview` fallback is actually what's coming through, and whether it's getting cut off; the OData category filter against your tenant; and the empty-recipient behavior on Create Draft.
