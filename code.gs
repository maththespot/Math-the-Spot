const ORDER_SHEET_NAME = "Orders";
const ORDER_HEADERS = [
  "Submitted at",
  "Name",
  "Email",
  "Phone",
  "Product",
  "Quantity",
  "Needed by",
  "Shipping method",
  "Delivery address",
  "Mode of payment"
];

function setupOrdersSpreadsheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error("Open this script from the Google Sheet that should receive orders.");
  }

  let sheet = spreadsheet.getSheetByName(ORDER_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(ORDER_SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(ORDER_HEADERS);
  }

  PropertiesService.getScriptProperties().setProperty(
    "ORDERS_SPREADSHEET_ID",
    spreadsheet.getId()
  );
}

function doPost(e) {
  const form = e && e.parameter;
  if (!form) {
    return responsePage("We could not read your order. Please try again.");
  }

  if (form._honey) {
    return responsePage(
      "Thank you, your order request was received, we will send an invoice and payment instructions."
    );
  }

  const order = {
    name: readField(form.name, 120),
    email: readField(form.email, 254),
    phone: readField(form.phone, 40),
    product: readField(form.product, 120),
    quantity: readField(form.quantity, 3),
    neededBy: readField(form.needed_by, 10),
    fulfilment: readField(form.fulfilment, 80),
    address: readField(form.address, 500),
    payment: readField(form.mode_of_payment, 40)
  };
  const neededByDate = new Date(order.neededBy + "T00:00:00Z");
  const validNeededBy =
    /^\d{4}-\d{2}-\d{2}$/.test(order.neededBy) &&
    !Number.isNaN(neededByDate.getTime()) &&
    neededByDate.toISOString().slice(0, 10) === order.neededBy;

  const valid =
    order.name &&
    /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(order.email) &&
    order.phone &&
    order.product === "Calculus Bookmark" &&
    /^\d+$/.test(order.quantity) &&
    Number(order.quantity) >= 1 &&
    Number(order.quantity) <= 100 &&
    validNeededBy &&
    ["Delivery 1", "Delivery 2", "Pick-up"].includes(order.fulfilment) &&
    ["gcash", "maya"].includes(order.payment);

  if (!valid) {
    return responsePage("Some order details are missing or invalid. Please go back and check the form.");
  }

  const spreadsheetId = PropertiesService.getScriptProperties().getProperty(
    "ORDERS_SPREADSHEET_ID"
  );
  if (!spreadsheetId) {
    return responsePage("Orders are not set up yet. Please contact the store and try again later.");
  }
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    let sheet = spreadsheet.getSheetByName(ORDER_SHEET_NAME);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(ORDER_SHEET_NAME);
    }

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(ORDER_HEADERS);
    }

    sheet.appendRow([
      new Date(),
      safeCell(order.name),
      safeCell(order.email),
      safeCell(order.phone),
      safeCell(order.product),
      Number(order.quantity),
      safeCell(order.neededBy),
      safeCell(order.fulfilment),
      safeCell(order.address),
      safeCell(order.payment)
    ]);
  } finally {
    lock.releaseLock();
  }

  return responsePage(
    "Thank you, your order request was received, we will send an invoice and payment instructions."
  );
}

function readField(value, maxLength) {
  return String(value || "").trim().slice(0, maxLength);
}

function safeCell(value) {
  return /^[=+\-@]/.test(value) ? "'" + value : value;
}

function responsePage(message) {
  return HtmlService.createHtmlOutput(
    '<!doctype html><html lang="en"><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<title>Order status - Math the Spot</title>' +
    '<style>' +
    '*{box-sizing:border-box}' +
    'body{min-width:320px;min-height:100vh;margin:0;padding:24px;display:grid;place-items:center;' +
    'background:linear-gradient(rgba(18,14,33,.62),rgba(18,14,33,.76)),linear-gradient(135deg,#171126,#30234c);' +
    'font:16px "Lexend","Segoe UI",Arial,sans-serif;color:#20233a}' +
    '.status-card{width:min(100%,520px);padding:clamp(28px,6vw,48px);border:1px solid rgba(255,255,255,.65);' +
    'border-radius:20px;background:rgba(255,255,255,.97);box-shadow:0 24px 70px rgba(8,5,20,.35);text-align:center}' +
    'h1{margin:0 0 16px;background:linear-gradient(120deg,#212d60,#438857);-webkit-background-clip:text;' +
    'background-clip:text;color:transparent;font-size:clamp(2rem,7vw,3rem);line-height:1.15}' +
    'p{margin:0;color:#4b5062;font-size:clamp(1rem,3vw,1.2rem);line-height:1.6}' +
    '</style></head><body><main class="status-card"><h1>Math the Spot</h1><p>' +
    escapeHtml(message) +
    '</p></main></body></html>'
  );
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}