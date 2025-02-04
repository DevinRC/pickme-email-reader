var spreadsheet = SpreadsheetApp.openById("");
var sheet = spreadsheet.getSheetByName("")

function main() {
  var query = "{\"PickMe delivery email receipt for\"} after:2024/10/01 before:2024/11/01  ";
  var threads = GmailApp.search(query)
  threads.reverse();
  Logger.log("Threads in Search: " + threads.length);

  for (var i = 0; i < threads.length; i++) {
    var content = threads[i].getMessages()[0].getBody()
    var $ = Cheerio.load(content);

    var orderId = $('#main_table > tbody > tr:nth-child(1) > td > table > tbody > tr:nth-child(1) > td:nth-child(2) > div').text().split("-")[1].trim();
    var restaurantName = $('#service-banner > tbody > tr:nth-child(3) > td > div').text().trim();
    var date = threads[i].getMessages()[0].getDate();
    var total = $('#trip-details-outer > tbody > tr:nth-child(1) > td > table > tbody > tr > td:nth-child(2) > div').text().trim().replace(/[^0-9.]/g, '')

    var parent = $('#trip-details-outer > tbody > tr:nth-child(3) > td > table > tbody > tr');

    var items = []
    var noOfItems = 0;
    parent.each(function () {
      var itemQty = $(this).find('.itemQty').text();
      var itemName = $(this).find('.itemQty').parent().text().trim().split(/\s+/).slice(1).join(' ');
      var itemPrice = $(this).find('.itemQty').parent().parent().next().find('td:nth-child(2) > div').text().trim().replace(/[^0-9.]/g, '')
      if (itemQty && itemName) {
        items.push({
          itemName: itemName,
          quantity: itemQty,
          price: itemPrice
        })
        noOfItems += parseInt(itemQty)
      }
    })

    var adjustments = {};
    var excludedAdjustments = ["Sub Total", "Delivery Fee"];
    parent.each(function () {
      if ($(this).find('.itemQty').length === 0) {
        var label = $(this).find('td').first().text().trim();
        var value = $(this).find('td').last().text().trim().replace(/[^0-9.+-]/g, '');

        if (excludedAdjustments.includes(label)) {
          return
        }

        if (label && value) {
          adjustments[label] = parseFloat(value);
        }
      }
    });

    var paymentMethod = $('#trip-details-outer > tbody > tr:nth-child(4) > td > table > tbody > tr:nth-child(2) > td:nth-child(2) > div').text().trim();

    // Write to Sheet
    var startRow = getNextUsableRow(sheet);
    var numRows = Math.max(noOfItems, Object.keys(adjustments).length);

    sheet.getRange(startRow, 1, numRows, 1).merge().setValue(orderId);
    sheet.getRange(startRow, 2, numRows, 1).merge().setValue(restaurantName);
    sheet.getRange(startRow, 3, numRows, 1).merge().setValue(date);

    for (var j = 0, row = startRow; j < items.length; j++) {
      var item = items[j];
      var quantity = parseInt(item.quantity);
      var individualPrice = parseFloat(item.price) / quantity;

      // Add each item as an individual row according to their quantity
      for (var k = 0; k < quantity; k++) {
        sheet.getRange(row, 4).insertCheckboxes();
        sheet.getRange(row, 5).setValue(item.itemName);
        sheet.getRange(row, 6).setValue(individualPrice);
        sheet.getRange(row, 7).setFormula(`=IF(D${row}, F${row}, 0)`);
        row++;
      }
    }

    var adjustmentKeys = Object.keys(adjustments);
    var orderStartRow = startRow;
    var orderEndRow = row - 1;
    for (var l = 0; l < adjustmentKeys.length; l++) {
      var label = adjustmentKeys[l];
      var value = adjustments[label];
      sheet.getRange(startRow + l, 8).setValue(label);
      sheet.getRange(startRow + l, 9).setValue(value);
      sheet.getRange(startRow + l, 10).setFormula(`=I${startRow + l} * (SUM(G${orderStartRow}:G${orderEndRow}) / SUM(F${orderStartRow}:F${orderEndRow}))`);
    }

    sheet.getRange(startRow, 11, numRows, 1).merge().setValue(parseFloat(total));
    sheet.getRange(startRow, 12, numRows, 1).merge().setFormula(`=SUM(G${orderStartRow}:G${orderEndRow}, J${orderStartRow}:J${orderEndRow})`);
    sheet.getRange(startRow, 13, numRows, 1).merge().setValue(paymentMethod);

    sheet.getRange(orderStartRow, 1, orderEndRow - orderStartRow + 1, 13).setBorder(true, true, true, true, null, null);
    sheet.getRange(orderStartRow, 4, orderEndRow - orderStartRow + 1, 4).setBorder(null, true, null, true, null, null);
    sheet.getRange(orderStartRow, 8, orderEndRow - orderStartRow + 1, 3).setBorder(null, true, null, true, null, null);

    console.log(`[${i + 1}/${threads.length}] Completed write for: ${restaurantName}, ID: ${orderId} - ${noOfItems} items`)
  }
}

function getNextUsableRow(sheet) {
  var range = sheet.getRange(sheet.getLastRow(), 1);
  var nextUsableRow;

  if (range.isPartOfMerge()) {
    nextUsableRow = range.getMergedRanges()[0].getLastRow() + 1
  } else {
    nextUsableRow = range.getRow() + 1
  }

  // console.log("Next Usable Row:", nextUsableRow);
  return nextUsableRow;
}

function initNewSheet() {
  sheet.clear();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).removeCheckboxes();

  var headerRow = [
    "PickMe ID",
    "Restaurant",
    "Date",
    "Claim",
    "Item",
    "Price",
    "Effective Price to Claim",
    "Adjustment",
    "Value",
    "Effective Adjustment Value",
    "Total",
    "Effective Total to Claim",
    "Payment Method"
  ];
  sheet.appendRow(headerRow);

  var headerRange = sheet.getRange(1, 1, 1, headerRow.length);

  // Text Formatting
  headerRange.setFontWeight("bold");
  headerRange.setHorizontalAlignment("center");
  headerRange.setVerticalAlignment("middle");
  headerRange.setWrap(true);

  // Date Formats
  var dateColumns = [3]; // Replace with your specific date columns
  dateColumns.forEach(function (column) {
    var range = sheet.getRange(2, column, sheet.getMaxRows() - 1);
    range.setNumberFormat("dd MMM yyyy hh:mm AM/PM");
  });

  // Number Formats
  var currencyColumns = [6, 7, 9, 10, 11];
  currencyColumns.forEach(function (column) {
    var range = sheet.getRange(2, column, sheet.getMaxRows() - 1);
    range.setNumberFormat("LKR #,##0.00");
  });

  // Cell Alignments
  var centerAlignColumns = [1, 2, 3, 11, 12, 13]
  centerAlignColumns.forEach(function (column) {
    var range = sheet.getRange(2, column, sheet.getMaxRows() - 1);
    range.setHorizontalAlignment("center");
    range.setVerticalAlignment("middle");
  })

  // Borders
  sheet.getRange(1, 4, 1, 4).setBorder(null, true, null, true, null, null); // Claim -- Effective Price to Claim
  sheet.getRange(1, 8, 1, 3).setBorder(null, true, null, true, null, null); // Adjustment -- Effective Adjustment Value
  sheet.getRange(1, 13, 1, 1).setBorder(null, null, null, true, null, null); // Effective Total to Claim

  // Row Freezing
  sheet.setFrozenRows(1);

  // Conditional Formatting Rules
  var rules = []; // Reset existing rules

  // Rule 1: Highlight Cash
  var paymentMethodConditionalFormatRange = sheet.getRange("K2:M")
  var paymentMethodConditionalFormatRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$M2=\"Cash\"")
    .setBackground("#b7e1cd") // Light green
    .setRanges([paymentMethodConditionalFormatRange])
    .build()

  // Rule 2: Highlight duplicate IDs
  var duplicateIDRange = sheet.getRange("A2:A");
  var duplicateIDRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=COUNTIF(A:A, A2) > 1")
    .setBackground("#ff9999") // Light red
    .setRanges([duplicateIDRange])
    .build();

  rules.push(paymentMethodConditionalFormatRule);
  rules.push(duplicateIDRule);
  sheet.setConditionalFormatRules(rules);
}
