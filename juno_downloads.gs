/**
 * Extracts Juno Download purchases from Gmail and stores them in a Google Sheet.
 * Expects emails from sales@junodownload.com with subject "Juno Download/Order ref".
 */
function extractJunoDownloadPurchases() {
  const query = 'from:sales@junodownload.com subject:"Juno Download/Order ref"';
  const threads = GmailApp.search(query);
  Logger.log(`Juno Download: Found ${threads.length} email threads.`);

  const sheetName = 'Juno Download Purchases';
  const headers = [
    "Date", "Track", "Artist", "Label", "Cat No", "Format", "Currency",
    "Price", "Price EUR", "Order ID"
  ];

  const sheet = getOrCreateSheet(sheetName, headers);
  const existingOrderIds = getExistingTxIds(sheet, 10); // Column J = index 10 (Order ID)

  const rowsToAppend = [];

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const html = msg.getBody();
      const plain = msg.getPlainBody();

      const { dateStr, purchaseDate } = extractJunoDate(html, plain, msg);
      const orderID = extractJunoOrderID(html, plain, msg);

      if (!orderID || existingOrderIds.includes(orderID)) {
        Logger.log(`Juno Download: Skipping ${orderID ? "duplicate" : "invalid"} Order ID.`);
        return;
      }

      const tracks = extractJunoTracks(html, orderID);
      
      tracks.forEach(trackInfo => {
        rowsToAppend.push([
          purchaseDate, 
          trackInfo.track, 
          trackInfo.artist, 
          trackInfo.label, 
          trackInfo.catNo, 
          trackInfo.format, 
          trackInfo.currency,
          trackInfo.price, 
          convertToEUR(trackInfo.price, trackInfo.currency, dateStr),
          orderID
        ]);
      });
    });
  });

  if (rowsToAppend.length) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, headers.length).setValues(rowsToAppend);
    Logger.log(`Juno Download: Appended ${rowsToAppend.length} new records.`);
  } else {
    Logger.log("Juno Download: No new purchases to append.");
  }
}

/**
 * Extracts the purchase date from Juno Download emails.
 */
function extractJunoDate(html, plain, msg) {
  let purchaseDate = null;
  let dateStr = "UNKNOWN_DATE";

  // Try to extract from HTML - look for date in order receipt section
  const htmlMatch = html.match(/<span[^>]*>Date<\/span><br>\s*([^<]+)/i);
  
  try {
    if (htmlMatch) {
      purchaseDate = new Date(htmlMatch[1].trim());
    }
  } catch (e) {
    Logger.log("Juno Download: HTML date parse failed: " + e.message);
  }

  if (!purchaseDate) {
    purchaseDate = msg.getDate();
  }

  dateStr = Utilities.formatDate(purchaseDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return { dateStr, purchaseDate };
}

/**
 * Extracts the order ID from Juno Download emails.
 */
function extractJunoOrderID(html, plain, msg) {
  // Try to extract from subject first
  const subject = msg.getSubject();
  const subjectMatch = subject.match(/Order ref\s+([A-Z0-9]+)/i);
  if (subjectMatch) {
    return subjectMatch[1].trim();
  }

  // Fallback to HTML
  const htmlMatch = html.match(/<span[^>]*>Order number<\/span><br>\s*([A-Z0-9]+)/i);
  if (htmlMatch) {
    return htmlMatch[1].trim();
  }

  return "";
}

/**
 * Extracts track information from the Juno Download HTML table.
 */
function extractJunoTracks(html, orderId = "") {
  const tracks = [];
  
  // Find all table rows that contain track data
  // Table structure: Artist - Title | Label (Cat No) | Format | Price
  const rowRegex = /<tr><td[^>]*>(.*?)<\/td>\s*<td[^>]*>(.*?)<\/td>\s*<td[^>]*>(.*?)<\/td>\s*<td[^>]*>(.*?)<\/td><\/tr>/gis;
  
  let rowMatch;
  while ((rowMatch = rowRegex.exec(html)) !== null) {
    const col1 = rowMatch[1]; // Artist - Title
    const col2 = rowMatch[2]; // Label (Cat No)
    const col3 = rowMatch[3]; // Format
    const col4 = rowMatch[4]; // Price
    
    // Clean up the text content
    const artistTitle = col1.replace(/<[^>]*>/g, '').trim();
    const labelCatNo = col2.replace(/<[^>]*>/g, '').trim();
    const format = col3.replace(/<[^>]*>/g, '').trim();
    const priceText = col4.replace(/<[^>]*>/g, '').replace(/&#128;/g, '€').replace(/&pound;/g, '£').replace(/&#36;/g, '$').trim();
    
    // Skip header rows, total rows, and empty rows
    if (artistTitle.includes('Artist - Title') || 
        artistTitle.includes('Subtotal') || 
        artistTitle.includes('Grand total') ||
        artistTitle === '' ||
        !priceText.match(/[€$£]?[\d.,]+/)) {
      continue;
    }
    
    // Parse artist and title
    const { artist, track } = parseJunoArtistTitle(artistTitle);
    
    // Parse label and catalog number
    const { label, catNo } = parseJunoLabelCatNo(labelCatNo);
    
    // Parse price and currency
    const { price, currency } = parseJunoPrice(priceText);
    
    if (track && artist && price > 0) {
      tracks.push({
        artist,
        track,
        label,
        catNo,
        format,
        price,
        currency
      });
      
      Logger.log(`Juno Download: Found track: ${artist} - ${track} (${currency}${price})`);
    }
  }
  
  Logger.log(`Juno Download: Extracted ${tracks.length} tracks for order ${orderId}`);
  return tracks;
}

/**
 * Parses artist and title from Juno Download format.
 */
function parseJunoArtistTitle(artistTitleText) {
  // Handle various formats like "Artist - Title", "Artist1 / Artist2 - Title", etc.
  const parts = artistTitleText.split(' - ');
  if (parts.length >= 2) {
    return {
      artist: parts[0].trim(),
      track: parts.slice(1).join(' - ').trim()
    };
  }
  
  return {
    artist: artistTitleText,
    track: ""
  };
}

/**
 * Parses label and catalog number from Juno Download format.
 */
function parseJunoLabelCatNo(labelCatNoText) {
  // Format is typically "Label (Cat No)"
  const match = labelCatNoText.match(/^(.*?)\s*\(([^)]+)\)$/);
  if (match) {
    return {
      label: match[1].trim(),
      catNo: match[2].trim()
    };
  }
  
  return {
    label: labelCatNoText,
    catNo: ""
  };
}

/**
 * Parses price and currency from Juno Download format.
 */
function parseJunoPrice(priceText) {
  // Extract currency symbol and amount
  const match = priceText.match(/([€$£])?([\d.,]+)/);
  if (match) {
    const symbol = match[1] || "€"; // Default to EUR for Juno
    const amount = parseFloat(match[2].replace(',', '.'));
    const currency = getCurrencyFromSymbol(symbol) || "EUR";
    
    return {
      price: amount,
      currency: currency
    };
  }
  
  return {
    price: 0,
    currency: "EUR"
  };
}

// You'll need to add this to your existing helper functions if not already present
function getCurrencyFromSymbol(symbol) {
  if (!symbol) return null;
  const s = symbol.trim();
  return s === "€" ? "EUR" : s === "$" ? "USD" : s === "£" ? "GBP" : null;
}
