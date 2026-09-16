/**
 * রিল আর্কাইভ — Google Apps Script ব্যাকএন্ড
 * এই কোডটা Google Sheet-এর Extensions > Apps Script-এ পেস্ট করবে।
 *
 * সেটআপ (README.md-তে বিস্তারিত আছে):
 * 1) নিচের SHEET_ID বদলে তোমার Google Sheet-এর ID বসাও।
 * 2) শিটের প্রথম ট্যাবের নাম রাখো "Movies" এবং প্রথম সারিতে (header) এই কলামগুলো বসাও:
 *    Timestamp | Name | Year | Country | Language | Category | WatchLink | DownloadLink | Poster
 * 3) Project Settings > Script Properties-এ একটা প্রপার্টি বানাও:
 *    key = ADMIN_PASSWORD,  value = তোমার পছন্দের পাসওয়ার্ড
 * 4) Deploy > New deployment > Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 *    সেভ করে যে URL পাবে সেটাই index.html-এর API_URL।
 */

const SHEET_ID = "PASTE_YOUR_GOOGLE_SHEET_ID_HERE";
const SHEET_NAME = "Movies";

function getSheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// GET — সব এন্ট্রি লিস্ট হিসেবে পাঠায়
function doGet(e) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  const rows = values
    .filter(row => row.join("").trim() !== "")
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
  return jsonOut_(rows);
}

// POST — অ্যাডমিন পাসওয়ার্ড যাচাই (action: "verify") অথবা নতুন এন্ট্রি যোগ (action: "add")
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut_({ success: false, error: "Invalid request" });
  }

  const storedPassword = PropertiesService.getScriptProperties().getProperty("ADMIN_PASSWORD");

  if (body.action === "verify") {
    const ok = storedPassword && body.password === storedPassword;
    return jsonOut_({ success: !!ok });
  }

  if (body.action === "add") {
    if (!storedPassword || body.password !== storedPassword) {
      return jsonOut_({ success: false, error: "Unauthorized — পাসওয়ার্ড মেলেনি।" });
    }
    if (!body.name || !body.year || !body.country || !body.language || !body.watchLink || !body.downloadLink) {
      return jsonOut_({ success: false, error: "সব প্রয়োজনীয় তথ্য দাও।" });
    }
    const sheet = getSheet_();
    sheet.appendRow([
      new Date(),
      body.name,
      body.year,
      body.country,
      body.language,
      body.category || "Movie",
      body.watchLink,
      body.downloadLink,
      body.poster || ""
    ]);
    return jsonOut_({ success: true });
  }

  return jsonOut_({ success: false, error: "Unknown action" });
}
