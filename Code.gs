function sendOrderEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var orderSheet = ss.getSheetByName('Purchase Request Form'); 
  var supplierSheet = ss.getSheetByName('Supplier'); // 
  
  var orders = orderSheet.getDataRange().getValues();
  var suppliers = supplierSheet.getDataRange().getValues();
  
  for (var i = 1; i < orders.length; i++) { 
    var order = orders[i];
    var orderId = order[0]; 
    var supplierName = order[18];
    var status = order[17]; 
    
    if (status !== 'In Progress') continue;
    
    Logger.log('Processing order ID: ' + orderId + ' for supplier: ' + supplierName);
    
    for (var j = 1; j < suppliers.length; j++) {
      var supplier = suppliers[j];
      Logger.log('Checking supplier: ' + supplier[1] + ', email: ' + supplier[3]);
      
      if (supplier[1] === supplierName) { 
        var email = supplier[3]; 
        
        if (email) { 
          Logger.log('Sending email to: ' + email);
          sendEmail(orderId, email, order);
        } else {
          Logger.log('Error: No email address found for supplier ' + supplierName);
        }
        break;
      }
    }
  }
}

function sendEmail(orderId, email, order) {
  var webAppUrl = "https://script.google.com/macros/s/AKfycbxLBAOOyPxSJOe2ith7Dej-5-PhBdaUOCkz5N1wSzPepY4exmZuL26ReV7ChyR53AA/exec"; // Replace with your actual web app URL
  var agreeUrl = webAppUrl + "?orderId=" + orderId + "&action=agree";
  var rejectUrl = webAppUrl + "?orderId=" + orderId + "&action=reject";
  
  var subject = "Order Approval Request";

  // Include order details in the email body with requested fields
  var body = "Dear Supplier,<br><br>" +
             "We are pleased to inform you that we have chosen you as the supplier for our recent purchase request.<br><br>" +
             "Order Details:<br>" +
             "Order ID: " + order[0] + "<br>" +
             "Item Name: " + order[6] + "<br>" +
             "Quantity: " + order[7] + "<br>" +
             "Item Description: " + order[8] + "<br>" +
             "Preferred Delivery Date: " + order[12] + "<br><br>" +
             "Please provide the following details:<br>" +
             "- Unit price/cost<br>" +
             "- Is it available<br>" +
             "- Total Price/Cost<br>" +
             "- Estimated Delivery Date<br><br>" +
             "Thank you,<br>" +
             "YipiYipiYah";

  try {
    Logger.log('Email content: ' + body);
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: body
    });
  } catch (e) {
    Logger.log('Failed to send email: ' + e.message);
  }
}


function onEdit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  

  if (sheet.getName() === 'Purchase Request Form') {
    var range = e.range;
    var row = range.getRow();
    var column = range.getColumn();
    
  
    if (column === 18 && row > 1) {
      var status = sheet.getRange(row, column).getValue();
      var supplierName = sheet.getRange(row, 19).getValue(); 
      var orderId = sheet.getRange(row, 1).getValue(); 

      if (status === 'Approved') {
        Logger.log('Order ID: ' + orderId + ' is approved. Sending Google Form to supplier: ' + supplierName);
        sendGoogleFormEmailToSupplier(supplierName, orderId);
      } else if (status === 'Rejected') {
        Logger.log('Order ID: ' + orderId + ' is rejected. Sending rejection email to supplier: ' + supplierName);
        sendRejectionEmailToSupplier(supplierName, orderId);
      }
    }
  }
}

function sendGoogleFormEmailToSupplier(supplierName, orderId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var supplierSheet = ss.getSheetByName('Supplier');
  
  if (!supplierSheet) {
    Logger.log('Error: Could not find the Supplier sheet.');
    return;
  }

  var suppliers = supplierSheet.getDataRange().getValues();

  for (var i = 1; i < suppliers.length; i++) { 
    var supplier = suppliers[i];
    if (supplier[1] === supplierName) { 
      var email = supplier[3]; 
      if (email) {
        Logger.log('Sending Google Form email to: ' + email);
        sendGoogleFormEmail(email, orderId);
      } else {
        Logger.log('Error: No email address found for supplier ' + supplierName);
      }
      break;
    }
  }
}

function sendGoogleFormEmail(supplierEmail, orderId) {
  var formUrl = "https://forms.gle/wF7jLwdhGtvqFcZe7"; 
  var subject = "Please Complete the Attached Form";
  var body = "Dear Supplier,\n\n" +
             "Your order has been approved. Please complete the following form for further processing:\n\n" +
             "Order ID: " + orderId + "\n" +
             "Form URL: " + formUrl + "\n\n" +
             "Thank you!";

  try {
    MailApp.sendEmail({
      to: supplierEmail,
      subject: subject,
      body: body
    });
    Logger.log('Form email sent to: ' + supplierEmail);
  } catch (e) {
    Logger.log('Failed to send form email: ' + e.message);
  }
}


function sendRejectionEmailToSupplier(supplierName, orderId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var supplierSheet = ss.getSheetByName('Supplier');
  
  if (!supplierSheet) {
    Logger.log('Error: Could not find the Supplier sheet.');
    return;
  }

  var suppliers = supplierSheet.getDataRange().getValues();

  for (var i = 1; i < suppliers.length; i++) { 
    var supplier = suppliers[i];
    if (supplier[1] === supplierName) { 
      var email = supplier[3]; 
      if (email) {
        Logger.log('Sending rejection email to: ' + email);
        sendRejectionEmail(email, orderId);
      } else {
        Logger.log('Error: No email address found for supplier ' + supplierName);
      }
      break;
    }
  }
}

function sendRejectionEmail(supplierEmail, orderId) {
  var subject = "Order Rejection Notification";
  var body = "Dear Supplier,\n\n" +
             "We regret to inform you that your order (Order ID: " + orderId + ") has been rejected.\n\n" +
             "Thank you for your understanding.\n\n" +
             "Best regards,\n" +
             "YipiYipiYah";

  try {
    MailApp.sendEmail({
      to: supplierEmail,
      subject: subject,
      body: body
    });
    Logger.log('Rejection email sent to: ' + supplierEmail);
  } catch (e) {
    Logger.log('Failed to send rejection email: ' + e.message);
  }
}

function notifyAdmin(orderId, status) {
  var adminEmail = "hackathonyah@gmail.com";
  var subject = "Order " + status;
  var body = "Order ID: " + orderId + " has been " + status;

  MailApp.sendEmail(adminEmail, subject, body);
}
function onFormSubmit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var transactionSheet = ss.getSheetByName('Transaction Receipt');
  var orderSheet = ss.getSheetByName('Purchase Request Form');
  var invoiceSheet = ss.getSheetByName('Invoice');
  

  var lastRow = transactionSheet.getLastRow();
  var transactionData = transactionSheet.getRange(lastRow, 1, 1, transactionSheet.getLastColumn()).getValues()[0];
  
  var orderId = transactionData[1]; 
  var fileUrl = transactionData[5]; 
  var additionalComment = transactionData[6]; 


  var fileId = extractFileIdFromUrl(fileUrl);

  Logger.log('File ID from form submission: ' + fileId);
  

  var orders = orderSheet.getDataRange().getValues();
  var supplierEmail = '';
  var supplierName = '';
  
  for (var i = 1; i < orders.length; i++) {
    var order = orders[i];
    if (order[0] == orderId) { 
      supplierEmail = order[19]; 
      supplierName = order[18]; 
      break;
    }
  }
  
 
  var currentStatus = transactionSheet.getRange(lastRow, 8).getValue(); 

  if (supplierEmail && fileId && currentStatus !== 'Complete') {
    Logger.log('Sending transaction receipt to supplier: ' + supplierName + ', email: ' + supplierEmail);
    sendTransactionReceiptEmail(supplierEmail, supplierName, orderId, transactionData, fileId, additionalComment);
    

    var invoiceData = invoiceSheet.getDataRange().getValues();
    for (var j = 1; j < invoiceData.length; j++) { 
      if (invoiceData[j][0] == orderId) { 
        invoiceSheet.getRange(j + 1, 14).setValue('Pending'); 
        break;
      }
    }
    
    transactionSheet.getRange(lastRow, 8).setValue('Complete'); 
    
    // Notify admin
    notifyAdmin(orderId, 'Complete');
  } else {
    Logger.log('Error: No matching supplier or file ID found for order ID ' + orderId + ' or the status is already complete.');
  }
}


function extractFileIdFromUrl(url) {
  var fileId = '';
  try {
    // Extract file ID from the URL
    var matches = url.match(/[-\w]{25,}/);
    if (matches) {
      fileId = matches[0];
    }
  } catch (e) {
    Logger.log('Error extracting file ID: ' + e.message);
  }
  return fileId;
}


function sendTransactionReceiptEmail(email, supplierName, orderId, transactionData, fileId, additionalComment) {
  var subject = "Transaction Receipt for Order ID: " + orderId;
  

  var body = "Dear " + supplierName + ",<br><br>" +
             "Thank you for your submission. Here are the details of your transaction:<br><br>" +
             "Order ID: " + transactionData[1] + "<br>" +
             "Transaction Date: " + transactionData[0] + "<br>" +
             "Amount: " + transactionData[4] + "<br>" +
             "Status: Pending<br><br>" +
             "Additional Comment: " + additionalComment + "<br><br>" +
             "If you have any questions, please do not hesitate to contact us.<br><br>" +
             "Best regards,<br>" +
             "YipiYipiYah";
             
  
  var file;
  try {
    file = DriveApp.getFileById(fileId);
    Logger.log('File fetched successfully: ' + file.getName());
  } catch (e) {
    Logger.log('Error fetching file: ' + e.message);
    return;
  }

  try {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: body,
      attachments: [file.getAs(MimeType.PDF)]
    });
    Logger.log('Transaction receipt email sent to: ' + email);
  } catch (e) {
    Logger.log('Failed to send transaction receipt email: ' + e.message);
  }
}

function notifyAdmin(orderId, status) {
  var adminEmail = "hackathonyah@gmail.com";
  var subject = "Order " + status;
  var body = "Your transaction receipt successfully sent to the supplier. Order ID: " + orderId + " has been " + status;

  try {
    MailApp.sendEmail(adminEmail, subject, body);
    Logger.log('Admin notified about order ID: ' + orderId);
  } catch (e) {
    Logger.log('Failed to notify admin: ' + e.message);
  }
}

function authorizeDrive() {
  DriveApp.getRootFolder();
}

