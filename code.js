function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('سیستم جمع‌آوری سفارش‌ها - نسخه حرفه‌ای')
    .setWidth(600)
    .setHeight(500);
}

function getNextOrder() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('report');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const orderNumber = data[i][2]; // ستون C
    const log = data[i][10]; // ستون K
    
    if (orderNumber && !log) {
      const serialsStr = data[i][5] || ''; // ستون F
      const addressesStr = data[i][9] || ''; // ستون J
      const serials = serialsStr ? serialsStr.split('-') : [];
      const addresses = addressesStr ? addressesStr.split('-') : [];
      
      return {
        row: i + 1,
        orderNumber: orderNumber,
        serials: serials,
        addresses: addresses
      };
    }
  }
  return null;
}

function updateLog(row, reason) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('report');
  const date = new Date().toLocaleString('fa-IR');
  let logMessage = '';
  if (reason && reason !== 'جمع‌آوری شده') {
    logMessage = `رد شده: ${reason} در ${date}`;
  } else {
    logMessage = `جمع‌آوری شده در ${date}`;
  }
  sheet.getRange(row, 11).setValue(logMessage);
  return `سفارش ${sheet.getRange(row, 3).getValue()}: ${logMessage}`;
}

function getRecentLogs() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('report');
  const data = sheet.getDataRange().getValues();
  const logs = [];
  
  for (let i = data.length - 1; i >= 1 && logs.length < 5; i--) {
    const log = data[i][10]; // ستون K
    if (log) {
      logs.push({
        orderNumber: data[i][2],
        log: log
      });
    }
  }
  return logs;
}
fetch("https://script.google.com/macros/s/AKfycbx.../exec")
  .then(response => response.json())
  .then(data => {
    console.log("Data from Google Sheets:", data);
    // مثلاً نمایش در صفحه:
    document.body.innerHTML += `<pre>${JSON.stringify(data, null, 2)}</pre>`;
  })
  .catch(error => console.error("Error fetching data:", error));
function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("report");
  var data = sheet.getDataRange().getValues();
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

