/**
 * Decodes Quoted-Printable email content (Apple invoice emails use this).
 */

function decodeQuotedPrintable(input) {
  return input
    .replace(/=\r?\n/g, '')  // Remove soft line breaks (like '=')
    .replace(/=([A-Fa-f0-9]{2})/g, (match, hex) => String.fromCharCode(parseInt(hex, 16)));
}


/**
 * Extracts Apple purchase information from Gmail invoice emails
 * and stores it in a Google Sheet.
 */
function extractApplePurchases() {
  const query = 'from:no_reply@email.apple.com "Factuur Apple"';
  const threads = GmailApp.search(query);
  Logger.log(`Apple: Found ${threads.length} threads.`);

  const sheetName = 'Apple Purchases';
  const headers = [
    "Date", "Track", "Artist", "Currency",
    "Subtotal", "VAT", "Total",
    "Subtotal EUR", "VAT EUR", "Total EUR",
    "Transaction ID"
  ];

  const sheet = getOrCreateSheet(sheetName, headers);
  const existingOrderIds = getExistingTxIds(sheet, 11); // Column K = index 11

  const rowsToAppend = [];

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const rawContent = msg.getRawContent();
      const html = decodeQuotedPrintable(rawContent);

      const { dateStr, purchaseDate } = extractAppleDate(html, msg);
      const orderId = extractAppleOrderId(html);

      if (!orderId || existingOrderIds.includes(orderId)) {
        Logger.log(`Apple: Skipping ${orderId ? 'duplicate' : 'missing'} Order ID.`);
        return;
      }

      const { item, artist } = extractAppleItemDetails(html);
      const { subtotal, vat, total, currency } = extractAppleFinancials(html);

      rowsToAppend.push([
        dateStr, item, artist, currency,
        subtotal, vat, total,
        convertToEUR(subtotal, currency, dateStr),
        convertToEUR(vat, currency, dateStr),
        convertToEUR(total, currency, dateStr),
        orderId
      ]);
    });
  });

  if (rowsToAppend.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, headers.length).setValues(rowsToAppend);
    Logger.log(`Apple: Appended ${rowsToAppend.length} new records.`);
  } else {
    Logger.log("Apple: No new records to append.");
  }
}

/**
 * Extracts purchase date from Apple email HTML or uses email timestamp.
 */
function extractAppleDate(html, msg) {
  let purchaseDate;
  let dateStr = "UNKNOWN_DATE";

  const dateMatch = html.match(/FACTUURDATUM<\/span>\s*<br\s*\/?>\s*([^<]+)/i);
  try {
    if (dateMatch && dateMatch[1]) {
      const rawDate = dateMatch[1].replace(/[\u00A0\u2022]/g, '').trim();
      const parts = rawDate.split('-');
      if (parts.length === 3) {
        purchaseDate = new Date(`${parts[2]}-${parts[1]}-${parts[0]}`);
      } else {
        throw new Error("Unexpected date format: " + rawDate);
      }
    }
  } catch (e) {
    Logger.log("Apple: Failed parsing date. Falling back. " + e.message);
  }

  if (!purchaseDate) {
    purchaseDate = msg.getDate();
  }

  dateStr = Utilities.formatDate(purchaseDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return { dateStr, purchaseDate };
}

/**
 * Extracts BESTEL-ID (order ID) from the Apple invoice HTML.
 */
function extractAppleOrderId(html) {
  const match = html.match(/BESTEL-ID<\/span><br><span[^>]*>[^>]*>([^<]+)<\/a>/i);
  return (match && match[1]) ? match[1].trim() : "";
}

/**
 * Extracts item name and artist from the Apple email HTML.
 */
function extractAppleItemDetails(html) {
  const match = html.match(
    /<img[^>]+src="https:\/\/is1-ssl\.mzstatic\.com\/image\/thumb\/Music[^"]+"[^>]*>\s*<\/td>\s*<td[^>]*>\s*<span[^>]+class="title"[^>]*>([^<]+)<\/span><br>\s*<span[^>]+class="artist"[^>]*>([^<]+)<\/span>/i
  );
  return {
    item: match ? match[1].trim() : "",
    artist: match ? match[2].trim() : ""
  };
}

/**
 * Extracts subtotal, VAT, total, and currency from Apple email HTML.
 */
function extractAppleFinancials(html) {
  let subtotal = 0, vat = 0, total = 0;
  const currency = "EUR";

  const subtotalMatch = html.match(/Subtotaal[\s\S]*?â¬Â\s*([\d.,]+)/i);
  if (subtotalMatch) subtotal = parseFloat(subtotalMatch[1].replace(',', '.'));

  const vatMatch = html.match(/Inclusief btw[\s\S]*?â¬Â\s*([\d.,]+)/i);
  if (vatMatch) vat = parseFloat(vatMatch[1].replace(',', '.'));

  const totalMatch = html.match(
    /font-weight:600[^>]*>\s*â¬Â\s*([\d.,]+)<\/span><\/td><\/tr>(?:\s*<tr>.*?Inclusief btw.*?<td[^>]*>[^<]*<\/td>\s*<td[^>]*><span[^>]*>\s*â¬Â\s*([\d.,]+)<\/span>)?/i
  );
  if (totalMatch) total = parseFloat(totalMatch[1].replace(',', '.'));

  return { subtotal, vat, total, currency };
}

/**
 * Reusable helpers (already defined in other files if shared)
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

function getExistingTxIds(sheet, colIndex) {
  const range = sheet.getLastRow() > 1 ? sheet.getRange(2, colIndex, sheet.getLastRow() - 1, 1) : null;
  return range
    ? range.getValues().flat().map(v => String(v).trim()).filter(v => v)
    : [];
}
