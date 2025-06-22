/**
 * Extracts Beatport purchases from Gmail and stores them in a Google Sheet.
 * Expects emails from reply@beatport-em.com with subject "order receipt".
 */
function extractBeatportPurchases() {
  const query = 'from:reply@beatport-em.com subject:"order receipt"';
  const threads = GmailApp.search(query);
  Logger.log(`Beatport: Found ${threads.length} email threads.`);

  const sheetName = 'Beatport Purchases';
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
      const html = msg.getBody();
      const plain = msg.getPlainBody();

      const { dateStr, purchaseDate } = extractBeatportDate(html, plain, msg);
      const transactionID = extractBeatportTransactionID(html, plain);

      if (!transactionID || existingTxIds.includes(transactionID)) {
        Logger.log(`Beatport: Skipping ${transactionID ? "duplicate" : "invalid"} TX ID.`);
        return;
      }

      const { track, artist } = extractTrackAndArtistFromHTML(html, transactionID);
      const { subtotal, tax, total, currency } = extractFinancialsFromPlain(plain);

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
    Logger.log(`Beatport: Appended ${rowsToAppend.length} new records.`);
  } else {
    Logger.log("Beatport: No new purchases to append.");
  }
}

/**
 * Ensures a sheet exists and has the correct header row.
 */
function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);

  if (sheet.getLastRow() === 0 || sheet.getRange("A1").getValue() !== headers[0]) {
    sheet.clear();
    sheet.appendRow(headers);
  }
  return sheet;
}

/**
 * Gets existing transaction IDs from the sheet.
 */
function getExistingTxIds(sheet, colIndex) {
  const range = sheet.getLastRow() > 1 ? sheet.getRange(2, colIndex, sheet.getLastRow() - 1, 1) : null;
  return range
    ? range.getValues().flat().map(v => String(v).trim()).filter(v => v)
    : [];
}

/**
 * Extracts the purchase date with fallbacks from HTML or plain text.
 */
function extractBeatportDate(html, plain, msg) {
  let purchaseDate = null;
  let dateStr = "UNKNOWN_DATE";

  const htmlMatch = html.match(/Purchase Date:<\/p>\s*<p[^>]*>(.*?)<\/p>/i);
  const plainMatch = plain.match(/Purchase Date:\s*([^\r\n]+)/i);

  try {
    purchaseDate = htmlMatch ? new Date(htmlMatch[1].trim()) : null;
  } catch (e) {
    Logger.log("Beatport: HTML date parse failed: " + e.message);
  }

  if (!purchaseDate && plainMatch) {
    try {
      purchaseDate = new Date(plainMatch[1].trim());
    } catch (e) {
      Logger.log("Beatport: Plain text date parse failed: " + e.message);
    }
  }

  if (!purchaseDate) {
    purchaseDate = msg.getDate();
  }

  dateStr = Utilities.formatDate(purchaseDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return { dateStr, purchaseDate };
}

/**
 * Extracts the transaction ID (invoice number).
 */
function extractBeatportTransactionID(html, plain) {
  const htmlMatch = html.match(/Invoice Number:<\/p>\s*<p[^>]*>#?(\d+)<\/p>/i);
  const plainMatch = plain.match(/Invoice Number:\s*#?(\d+)/i);
  return (
    (htmlMatch && htmlMatch[1]) ||
    (plainMatch && plainMatch[1]) ||
    ""
  ).toString().trim();
}

/**
 * Extracts track title and artist from the Beatport HTML layout.
 */
function extractTrackAndArtistFromHTML(html, txId = "") {
  const trackRegex = /<p[^>]*>Tracks<\/p>.*?<table.*?>(?:(?!<\/table>).)*?<tr>.*?<td[^>]*><p[^>]*>.*?<\/p><\/td>\s*<td[^>]*><p[^>]*>(.*?)<\/p><\/td>\s*<td[^>]*><p[^>]*>(.*?)<\/p><\/td>/is;
  const match = html.match(trackRegex);

  if (match) {
    Logger.log(`Beatport: Extracted track for TX ${txId}: ${match[1]} by ${match[2]}`);
    return { track: match[1].trim(), artist: match[2].trim() };
  }

  Logger.log(`Beatport: Could not extract track/artist for TX ${txId}`);
  return { track: "", artist: "" };
}

/**
 * Extracts subtotal, tax, total, and currency symbol.
 */
function extractFinancialsFromPlain(plain) {
  let subtotal = 0, tax = 0, total = 0;
  let currency = "EUR"; // default

  const subtotalMatch = plain.match(/Subtotal\s*([€$£])?([\d.,]+)/i);
  const taxMatch = plain.match(/Tax\s*([€$£])?([\d.,]+)/i);
  const totalMatch = plain.match(/Total\s*([€$£])?([\d.,]+)/i);

  if (subtotalMatch) {
    subtotal = parseFloat(subtotalMatch[2].replace(',', '.'));
    currency = getCurrencyFromSymbol(subtotalMatch[1]) || currency;
  }

  if (taxMatch) {
    tax = parseFloat(taxMatch[2].replace(',', '.'));
  }

  if (totalMatch) {
    total = parseFloat(totalMatch[2].replace(',', '.'));
    if (!subtotalMatch) currency = getCurrencyFromSymbol(totalMatch[1]) || currency;
  }

  return { subtotal, tax, total, currency };
}

/**
 * Converts a currency symbol to a currency code.
 */
function getCurrencyFromSymbol(symbol) {
  if (!symbol) return null;
  const s = symbol.trim();
  return s === "€" ? "EUR" : s === "$" ? "USD" : s === "£" ? "GBP" : null;
}
