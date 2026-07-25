function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚡ AutoFill Dashboard')
    .addItem('Open UI Control Panel', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Active Page Controller')
      .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getPageData(pageNo) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DataSheet");
  if (!sheet) return { error: "DataSheet nahi mili!" };
  
  var data = sheet.getDataRange().getValues();
  var pageIndex = parseInt(pageNo, 10);
  
  if (isNaN(pageIndex) || pageIndex < 1 || pageIndex >= data.length) {
    return { error: "Page data available nahi hai!" };
  }
  
  var row = data[pageIndex];
  return {
    pageNo: row[0],
    customer: row[1],
    product: row[2],
    qty: row[3],
    price: row[4],
    baseAmt: row[5],
    discountPct: row[6],
    discountAmt: row[7],
    taxableAmt: row[8],
    gstPct: row[9],
    gstAmt: row[10],
    netPayable: row[11],
    description: row[12]
  };
}
