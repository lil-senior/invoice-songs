# 🎧 Automated Music Purchase Extractor (Bandcamp / Beatport / Apple / Juno Download)

As I couldn't be asked to do this myself, I used Google Apps Script together with Google Sheets to automate it.

Shoutout to Gemini, ChatGPT and Claude for writing the code.

What this setup does:
- 📥 Reads the **HTML content** of incoming **Gmail** receipts
- 🧠 Extracts relevant info like **track title, artist, subtotal, VAT, total, transaction ID**
- 💱 Automatically converts all amounts into a base currency (default: **EUR**, using [frankfurter.dev](https://frankfurter.dev/))
- 📊 Appends the results into a neat tab in Google Sheets
- 🤖 With triggering in App Script, it looks for new invoices in my mailbox every Monday 10 AM 

No more manual copy-pasting from your Bandcamp, Beatport, Apple or Juno Download receipts. 🙌

---

## 💡 Use Case

If you're:
- A DJ, artist, or label manager
- Someone who buys tracks regularly from multiple platforms
- Living in a place like the Netherlands and have to report **VAT** or business expenses...

...this tool saves hours of repetitive work.

---

## 🛠 Platforms Supported

- ✅ Bandcamp
- ✅ Beatport
- ✅ Apple (iTunes, Music purchases — tested on Dutch invoices)
- ✅ Juno Download

---

## 📬 Email Requirements

- You **must use Gmail** (free or Google Workspace).
- If your receipts go elsewhere (e.g. iCloud, Outlook), set up **automatic forwarding** to a Gmail inbox.
- Scripts run entirely inside Google Sheets (using `Extensions > Apps Script`).

---

## 🔧 Setup Instructions

1. **Create a new Google Sheet** (sheets.google.com)

2. Click `Extensions > Apps Script` 

3. Paste in the provided code files:
   - `bandcamp.gs`
   - `beatport.gs`
   - `apple.gs`
   - `juno_downloads.gs`
   - `utils.gs`
   - `call_all.gs`

4. It will through a 'security warning', just click yes and you'll be good to go!

5. Check the apple/bandcamp/beatport html constants such that they will work for your invoice email. I think Bandcamp & Beatport will be fine, but Apple sends mine in Dutch so I can almost guarantee that it won't work for a non-dutch invoice email.
   An easy way to do that is by copy/pasting the whole apple.gs script into chatgpt or gemini together with original email (it is called **Show original** in Gmail) and ask to make it work with your type of email. 

7. Set a trigger via `Triggers > Add Trigger` to run `run_all_scripts` weekly or monthly.
   This way, each week your inbox gets scanned on new invoices and automatically adds them to your spreadsheet. 

---

## 📦 What the Sheet Includes

Each tab (Bandcamp / Beatport / Apple / Juno Download) will include:

| Date       | Track | Artist | Currency | Subtotal | VAT | Total | Subtotal EUR | VAT EUR | Total EUR | Transaction ID |
|------------|-------|--------|----------|----------|-----|-------|---------------|----------|-----------|----------------|

All amounts are converted to **EUR** using daily exchange rates from [Frankfurter.dev](https://frankfurter.dev).
And yes, if you run the script retrospectively, it will use the exchange rate of the day when you bought your song.

---

## 📉 What This Script Does *Not* Do

- ❌ Create quarterly/annual summaries — use a pivot table, formula, or ChatGPT for that

- I mean, a simple formula using '=SOM('Apple Purchases'!F2:F)' in google sheets automatically calculates everything. You can even separate it based on year or month using the date from the generated sheets.
  
- ❌ Support other currencies beyond what's supported by [Frankfurter.dev](https://frankfurter.dev)
- ❌ Export to accounting software (although CSV download works fine)

---

## 🔒 Privacy Notes

- This script **only accesses emails** from known senders (`noreply@bandcamp.com`, `reply@beatport-em.com`, `no_reply@email.apple.com`)
- It does **not** store or share your data externally
- API calls are made to [Frankfurter.dev](https://frankfurter.dev), an open-source exchange rate provider

---

## 🙋 FAQ

**Do I need to be technical to use this?**  
Not really. If you know how to copy-paste code into Apps Script, you're good.

**Can I use this with other platforms?**  
Sure, but you'll need to write a new extractor for those emails. The structure for Apple/Bandcamp/Beatport is a good starting point.

**Can I change the base currency?**  
Yes! Modify the `convertToEUR()` function in `utils.gs` to use a different target like `"USD"` or `"GBP"`.

---

## 🧠 Author's Note

> This was a lazy automation attempt turned into something mildly useful.  
> Feel free to fork, tweak, and build on it.

PRs welcome if you want to add more platforms (Juno, Boomkat, Traxsource, etc.) or analytics features.

---

## 📎 License

GNU GPLv3. Do whatever you want with it — just don't resell it as a SaaS please.
