/**
 * Converts an amount from a foreign currency to EUR using historical exchange rates.
 * Uses the Frankfurter.dev API.
 *
 * @param {number} amount - The amount to convert.
 * @param {string} currency - The original currency code (e.g., "USD", "GBP").
 * @param {string} dateStr - The date in "yyyy-MM-dd" format for historical rate.
 * @returns {number|string} - The converted amount rounded to 2 decimals, or error string.
 */
function convertToEUR(amount, currency, dateStr) {
  if (currency === "EUR") return amount;
  if (!amount || amount === 0) return 0;

  const url = `https://api.frankfurter.app/${dateStr}?base=${currency}&symbols=EUR`;

  try {
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());

    const rate = data?.rates?.EUR;
    if (rate === undefined) {
      Logger.log(`[convertToEUR] No rate found for ${currency}->EUR on ${dateStr}.`);
      return "NO_RATE_FOUND";
    }

    const converted = Math.round((amount * rate + Number.EPSILON) * 100) / 100;
    Logger.log(`[convertToEUR] Converted ${amount} ${currency} to ${converted} EUR (Rate: ${rate} on ${dateStr})`);
    return converted;
  } catch (e) {
    Logger.log(`[convertToEUR] API error for ${currency} on ${dateStr}: ${e.message} (URL: ${url})`);
    return "API_ERROR";
  }
}
