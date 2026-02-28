function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Order Picker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Load items from the Excel-imported sheet
function getItems() {
  const ss = SpreadsheetApp.openById('YOUR_SHEET_ID');
  const sheet = ss.getSheetByName('Sheet1'); // rename if needed
  const values = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues(); 
  // Column B = Description
  return values.flat().filter(String);
}

// Submit selected items into Google Form
function submitOrder(selectedItems) {
  const form = FormApp.openById('YOUR_FORM_ID');
  const item = form.getItems(FormApp.ItemType.TEXT)[0].asTextItem();

  const formResponse = form.createResponse();
  formResponse.withItemResponse(
    item.createResponse(selectedItems.join(', '))
  );
  formResponse.submit();

  return true;
}
