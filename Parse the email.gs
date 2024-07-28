function parseSupplierEmails() {
  var searchQuery = 'subject:"Order Approval Request"';
  var threads = GmailApp.search(searchQuery);

  if (threads.length === 0) {
    Logger.log("No emails found with the given subject.");
    return;
  }

  Logger.log("Found " + threads.length + " threads.");

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RFQ SEND');
  var data = sheet.getDataRange().getValues();
  var emailColumn = 11; // Email column index (0-based)
  var timestampColumn = 1; // Timestamp column index (0-based)
  var orderIdColumn = 2; // Order ID column index (0-based)

  var latestEmails = {};

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var body = message.getPlainBody();
      var sender = extractEmail(message.getFrom());
      var orderId = extractOrderId(body);

      Logger.log("Processing reply from: " + sender + " with Order ID: " + orderId);

      // Extract details from the email body
      var unitPrice = extractDetail(body, 'Unit price/cost');
      var availability = extractDetail(body, 'Is it available');
      var totalPrice = extractDetail(body, 'Total Price/Cost');
      var estimatedDelivery = extractDetail(body, 'Estimated Delivery Date');

      if (unitPrice && availability && totalPrice && estimatedDelivery) {
        // Track the latest email for each supplier and order ID
        var key = sender + "|" + orderId;
        var timestamp = message.getDate().getTime();

        if (!latestEmails[key] || timestamp > latestEmails[key].timestamp) {
          latestEmails[key] = {
            timestamp: timestamp,
            unitPrice: unitPrice,
            availability: availability,
            totalPrice: totalPrice,
            estimatedDelivery: estimatedDelivery
          };
        }
      }
    }
  }

  // Update the sheet with the latest data
  for (var k = 1; k < data.length; k++) {
    var row = data[k];
    var rowEmail = extractEmail(row[emailColumn]);
    var rowOrderId = row[orderIdColumn];
    var timestamp = parseDate(row[timestampColumn]);

    var key = rowEmail + "|" + rowOrderId;
    if (latestEmails[key] && (!timestamp || timestamp.getTime() <= latestEmails[key].timestamp)) {
      updateOrderSheet(k + 1, latestEmails[key].unitPrice, latestEmails[key].availability, latestEmails[key].totalPrice, latestEmails[key].estimatedDelivery);
    }
  }
}

function extractEmail(fullEmail) {
  var match = fullEmail.match(/<(.*)>/);
  return match ? match[1] : fullEmail;
}

function extractOrderId(body) {
  var regex = /Order ID\s*[:\-]\s*(\S+)/i;
  var match = body.match(regex);
  return match ? match[1].trim() : null;
}

function parseDate(dateValue) {
  var parsedDate = new Date(dateValue);
  return isNaN(parsedDate.getTime()) ? null : parsedDate;
}

function updateOrderSheet(rowToUpdate, unitPrice, availability, totalPrice, deliveryDate) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RFQ SEND');
  
  // Update the row with extracted data
  sheet.getRange(rowToUpdate, 14).setValue(unitPrice);
  sheet.getRange(rowToUpdate, 13).setValue(availability);
  sheet.getRange(rowToUpdate, 15).setValue(totalPrice);
  sheet.getRange(rowToUpdate, 16).setValue(deliveryDate);
}

function extractDetail(body, keyword) {
  var regex = new RegExp(keyword + '\\s*[:\\-]\\s*(.*)', 'i');
  var match = body.match(regex);
  var value = match ? match[1].trim() : null;
  Logger.log('Extracted value for ' + keyword + ': ' + value);
  return value;
}
