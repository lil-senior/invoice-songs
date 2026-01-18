/**
 * Extracts Qobuz purchases from Gmail and stores them in a Google Sheet.
 **/

function extractQobuzPurchases() {
  const query = [
    '(from:account@account.qobuz.com OR from:account.qobuz.com OR from:qobuz.com)',
    'subject:"Summary of your order"',
  ].join(' ');
  const threads = GmailApp.search(query);
  Logger.log(`Qobuz: Found ${threads.length} email threads.`);

  const sheetName = 'Qobuz Purchases';
  const headers = [
    "Date", "Track", "Artist", "Currency",
    "Subtotal", "Tax", "Total",
    "Subtotal EUR", "Tax EUR", "Total EUR",
    "Transaction ID"
  ];

  const sheet = getOrCreateSheet(sheetName, headers);
  const existingTxIds = getExistingTxIds(sheet, 11); // Column K = index 11

  const rowsToAppend = [];

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      // Qobuz provides a plain-text part that is usually easiest to parse.
      const rawPlain = msg.getPlainBody() || "";
      const rawHtml  = msg.getBody() || "";

      const plain = normalizeQobuzText(rawPlain);
      const html  = normalizeQobuzText(rawHtml);

      const { dateStr, purchaseDate } = extractQobuzDate(plain, msg);
      const transactionID = extractQobuzTransactionID(plain, html);

      if (!transactionID || existingTxIds.includes(String(transactionID).trim())) {
        Logger.log(`Qobuz: Skipping ${transactionID ? "duplicate" : "invalid"} TX ID.`);
        return;
      }

      const { track, artist } = extractQobuzTrackAndArtist(plain, html, transactionID);
      const { subtotal, tax, total, currency } = extractQobuzFinancials(plain, html);

      rowsToAppend.push([
        purchaseDate, track, artist, currency,
        subtotal, tax, total,
        convertToEUR(subtotal, currency, dateStr),
        convertToEUR(tax, currency, dateStr),
        convertToEUR(total, currency, dateStr),
        transactionID
      ]);
    });
  });

  if (rowsToAppend.length) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, headers.length).setValues(rowsToAppend);
    Logger.log(`Qobuz: Appended ${rowsToAppend.length} new records.`);
  } else {
    Logger.log("Qobuz: No new purchases to append.");
  }
}

/**
 * Normalize quoted-printable artifacts and common Qobuz encodings.
 * - Removes soft line breaks "=\r\n"
 * - Converts "=E2=82=AC" to "€"
 * - Collapses excessive whitespace
 **/

function normalizeQobuzText(s) {
  if (!s) return "";
  let out = String(s);

  // Remove quoted-printable soft wraps
  out = out.replace(/=\r?\n/g, "");

  // Decode the most common currency encoding seen in Qobuz emails
  out = out.replace(/=E2=82=AC/gi, "€");

  // Reduce noisy zero-width / odd spacing artifacts (safe, conservative)
  out = out.replace(/\u200C/g, ""); 
  out = out.replace(/[ \t]+\n/g, "\n");
  out = out.replace(/\n{3,}/g, "\n\n");

  return out;
}

/**
 * Extracts the order date. Example:
 *   "Order n°54714713 confirmed on January  5, 2026."
 * Falls back to message date.
 **/
function extractQobuzDate(plain, msg) {
  let purchaseDate = null;

  const m = plain.match(/confirmed\s+on\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})/i);
  if (m && m[1]) {
    purchaseDate = new Date(m[1].replace(/\s{2,}/g, " ").trim());
    if (isNaN(purchaseDate.getTime())) purchaseDate = null;
  }

  if (!purchaseDate) purchaseDate = msg.getDate();

  const dateStr = Utilities.formatDate(purchaseDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return { dateStr, purchaseDate };
}

/**
 * Extracts the Qobuz order number as Transaction ID.
 * Matches:
 *   "Order n°54714713 confirmed ..."
 * Also tolerates "Order n° 54714713" etc.
 */
function extractQobuzTransactionID(plain, html) {
  const sources = [plain || "", html || ""].join("\n");

  // n° may appear as "n°" or "n=C2=B0" depending on encoding/part
  const m =
    sources.match(/Order\s+n(?:°|=C2=B0)\s*([0-9]{5,})/i) ||
    sources.match(/Order\s+n[°\s]*([0-9]{5,})/i);

  return (m && m[1] ? String(m[1]).trim() : "");
}

/**
 * Extracts track & artist from the typical Qobuz plain-text layout.
 * Example (plain):
 *   Made in Coracao
 *   -
 *   Selected ('09 24bit Digital...
 *   Sadao Watanabe
 **/
function extractQobuzTrackAndArtist(plain, html, txId) {
  const text = plain || "";
  const lines = text.split(/\r?\n/).map(l => l.trim());

  let track = "";
  let artist = "";

  // Find a line that is just "-"
  const dashIdx = lines.findIndex(l => l === "-");
  if (dashIdx > 0) {
    // Track is the closest non-empty line above "-"
    for (let i = dashIdx - 1; i >= 0; i--) {
      if (lines[i]) { track = lines[i]; break; }
    }

    // Artist: look forward for "Selected" then next non-empty line
    let start = dashIdx + 1;
    for (let i = dashIdx + 1; i < Math.min(lines.length, dashIdx + 20); i++) {
      if (/^Selected\b/i.test(lines[i])) { start = i + 1; break; }
    }
    for (let i = start; i < Math.min(lines.length, start + 10); i++) {
      if (lines[i] && !/^Taxes\b/i.test(lines[i]) && !/^Total\b/i.test(lines[i])) {
        artist = lines[i];
        break;
      }
    }
  }

  // Fallback: try a weaker pattern across the full text if heuristic fails
  if (!track || !artist) {
    // Try: "<track>\n-\n... \n<artist>\n\nTaxes"
    const m = text.match(/^\s*([^\n\r]+)\s*\r?\n\s*-\s*\r?\n[\s\S]{0,400}?\r?\n\s*([^\n\r]+)\s*\r?\n[\s\S]{0,200}?\bTaxes\b/im);
    if (m) {
      track = track || (m[1] ? m[1].trim() : "");
      artist = artist || (m[2] ? m[2].trim() : "");
    }
  }

  if (track || artist) {
    Logger.log(`Qobuz: Extracted track for TX ${txId}: ${track} by ${artist}`);
  } else {
    Logger.log(`Qobuz: Could not extract track/artist for TX ${txId}`);
  }

  return { track, artist };
}

/**
 * Extracts subtotal, tax, total, and currency.
 * Your sample shows:
 *   Taxes €0.4
 *   Total 2.39€
 *
 * Notes:
 * - Qobuz receipts sometimes omit an explicit subtotal. If not found, we infer:
 *   subtotal = max(total - tax, 0)
 */
function extractQobuzFinancials(plain, html) {
  const sources = [plain || "", html || ""].join("\n");

  // Normalize any lingering "=E2=82=AC"
  const s = normalizeQobuzText(sources);

  let tax = 0;
  let total = 0;
  let subtotal = 0;
  let currency = "EUR";

  // Taxes: "Taxes €0.4" or "Taxes\n€0.4"
  const taxMatch =
    s.match(/\bTaxes\b[^\d€$£]*([€$£])?\s*([\d.,]+)/i) ||
    s.match(/\bTaxes\b[\s\r\n]*([€$£])?\s*([\d.,]+)/i);

  if (taxMatch) {
    const sym = (taxMatch[1] || "").trim();
    tax = parseMoneyNumber_(taxMatch[2]);
    currency = getCurrencyFromSymbol(sym) || currency;
  }

  // Total: "Total\n2.39€" or "Total 2.39€" or "Total €2.39"
  const totalMatch =
    s.match(/\bTotal\b[^\d€$£]*([€$£])?\s*([\d.,]+)\s*([€$£])?/i) ||
    s.match(/\bTotal\b[\s\r\n]*([\d.,]+)\s*([€$£])/i);

  if (totalMatch) {
    // Handle both regex shapes
    let sym = "";
    let num = "";
    if (totalMatch.length >= 4) {
      sym = (totalMatch[1] || totalMatch[3] || "").trim();
      num = totalMatch[2];
    } else {
      num = totalMatch[1];
      sym = (totalMatch[2] || "").trim();
    }
    total = parseMoneyNumber_(num);
    if (!currency || currency === "EUR") currency = getCurrencyFromSymbol(sym) || currency;
  }

  // Subtotal: not always present; try anyway
  const subMatch =
    s.match(/\bSubtotal\b[^\d€$£]*([€$£])?\s*([\d.,]+)/i) ||
    s.match(/\bSub-?total\b[^\d€$£]*([€$£])?\s*([\d.,]+)/i);

  if (subMatch) {
    const sym = (subMatch[1] || "").trim();
    subtotal = parseMoneyNumber_(subMatch[2]);
    currency = getCurrencyFromSymbol(sym) || currency;
  } else {
    // Infer subtotal if possible
    if (total > 0) subtotal = Math.max(total - tax, 0);
  }

  return { subtotal, tax, total, currency };
}

/**
 * Parses a number like "2.39", "0.4", "2,39" or "1 234,56" into float.
 **/
function parseMoneyNumber_(x) {
  if (x == null) return 0;
  let s = String(x).trim();

  // Remove spaces used as thousand separators
  s = s.replace(/\s+/g, "");

  // If comma is used as decimal separator, convert it.
  // (Simple approach, consistent with your Beatport parser.)
  s = s.replace(",", ".");

  const v = parseFloat(s);
  return isNaN(v) ? 0 : v;
}
