/**
 * Extracts Bandcamp purchases from Gmail and stores them in a Google Sheet.
 * Expects emails from noreply@bandcamp.com with subject "Thank you!".
 */
function extractBandcampPurchases() {
  const query = 'from:noreply@bandcamp.com subject:"Thank you!"';
  const threads = GmailApp.search(query);
  Logger.log(`Found ${threads.length} email threads with the query: "${query}"`);

  const sheetName = 'Bandcamp Purchases';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  // Add header if not present
  const headers = [
    "Date", "Track", "Artist", "Currency",
    "Subtotal", "VAT", "Total",
    "Subtotal EUR", "VAT EUR", "Total EUR",
    "Transaction ID"
  ];
  if (sheet.getLastRow() === 0 || sheet.getRange("A1").getValue() !== headers[0]) {
    sheet.clear(); // Avoid duplicate headers
    sheet.appendRow(headers);
  }

  const existingTxIds = sheet.getLastRow() > 1
    ? sheet.getRange("K2:K" + sheet.getLastRow()).getValues()
        .flat()
        .map(v => String(v).trim())  // Force to string & trim
        .filter(v => v)              // Remove falsy values
    : [];
  
  const rowsToAppend = [];

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const html = msg.getBody();
      const plain = msg.getPlainBody();

      const { dateStr, purchaseDate } = extractPurchaseDate(html, plain, msg);
      const track = (html.match(/<(b|strong)>(.*?)<\/(b|strong)>, by/i) || [])[2] || "";
      const artist = (html.match(/, by (.*?)<\/div>/) || [])[1] || "";

      const subtotal = extractAmount(plain, /Subtotal:\s+([\d.]+)/);
      const vat = extractAmount(plain, /VAT.*?:\s+([\d.]+)/);
      const totalMatch = /Total:\s+([€$£]?)([\d.]+)\s*([A-Z]{3})?/.exec(plain) || /Total:\s+([€$£]?)([\d.]+)\s*([A-Z]{3})?/.exec(html);
      const total = totalMatch ? parseFloat(totalMatch[2]) : 0;

      let currency = "EUR";
      if (totalMatch) {
        if (totalMatch[3]) currency = totalMatch[3].trim();
        else if (totalMatch[1]) {
          const symbol = totalMatch[1].trim();
          if (symbol === "$") currency = "USD";
          else if (symbol === "£") currency = "GBP";
          else if (symbol === "€") currency = "EUR";
        }
      }

      const txId = extractTransactionID(plain);
      if (!txId || existingTxIds.includes(txId)) {
        Logger.log(`Skipping ${txId ? 'duplicate' : 'invalid'} transaction.`);
        return;
      }

      rowsToAppend.push([
        purchaseDate, track, artist, currency,
        subtotal, vat, total,
        convertToEUR(subtotal, currency, dateStr),
        convertToEUR(vat, currency, dateStr),
        convertToEUR(total, currency, dateStr),
        txId
      ]);
    });
  });

  if (rowsToAppend.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    Logger.log(`Appended ${rowsToAppend.length} new purchase records.`);
  } else {
    Logger.log("No new purchase records to append.");
  }
}

/**
 * Extracts the purchase date from HTML or plain text fallback.
 */
function extractPurchaseDate(htmlBody, plainBody, msg) {
  let dateStr = "UNKNOWN_DATE";
  let purchaseDate;

  const htmlMatch = htmlBody.match(/<span class="label"[^>]*>Purchased:<\/span>\s*<span class="value"[^>]*>(.*?)<\/span>/i);
  const plainMatch = plainBody.match(/Purchased:\s*([^\r\n]+)/);

  try {
    purchaseDate = htmlMatch ? new Date(htmlMatch[1].trim()) : null;
  } catch (e) {
    Logger.log("HTML date parse error: " + e.message);
  }

  if (!purchaseDate && plainMatch) {
    try {
      purchaseDate = new Date(plainMatch[1].trim());
    } catch (e) {
      Logger.log("Plain text date parse error: " + e.message);
    }
  }

  if (!purchaseDate) {
    purchaseDate = msg.getDate();
  }

  dateStr = Utilities.formatDate(purchaseDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return { dateStr, purchaseDate };
}

/**
 * Extracts a transaction ID from plain text.
 */
function extractTransactionID(text) {
  return (
    (text.match(/Payment\s+(\d+)/) || [])[1] ||
    (text.match(/Bandcamp transaction ID:\s*(\d+)/i) || [])[1] ||
    (text.match(/payment_id=(\d+)/) || [])[1] ||
    ""
  ).toString().trim();
}

/**
 * Extracts a float value using a regex match.
 */
function extractAmount(text, regex) {
  const match = text.match(regex);
  return match ? parseFloat(match[1]) : 0;
}
