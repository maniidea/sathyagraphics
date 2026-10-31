var SHEET_NAME_ORDERS = "Orders";
var SHEET_NAME_PRODUCTS = "Products";
var SHEET_NAME_CUSTOMERS = "Customers";

function doPost(e) {
  try {
    var contents = e.postData.contents;
    var data = JSON.parse(contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Handle Saving Customer Details from Store Order/Enquiry
    if (data.action === "saveCustomer") {
      var sheet = ss.getSheetByName(SHEET_NAME_CUSTOMERS) || ss.insertSheet(SHEET_NAME_CUSTOMERS);
      
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(["Customer ID", "Name", "Mobile Number", "Email", "Registered Date"]);
      }
      
      var mobile = data.mobileNumber ? data.mobileNumber.toString().trim() : "";
      var rows = sheet.getDataRange().getValues();
      var exists = false;
      
      // Check if customer already exists by mobile number (Column C, index 2)
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][2] && rows[i][2].toString() === mobile) {
          exists = true;
          break;
        }
      }
      
      if (!exists && mobile !== "") {
        var customerId = 'CUST-' + Math.floor(100000 + Math.random() * 900000);
        sheet.appendRow([
          customerId,
          data.customerName || "Customer",
          mobile,
          data.email || "",
          new Date()
        ]);
      }
      
      return ContentService.createTextOutput(JSON.stringify({
        status: "success"
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Handle Add Product from Admin 2 Dashboard
    if (data.action === "addProduct") {
      var sheet = ss.getSheetByName(SHEET_NAME_PRODUCTS) || ss.insertSheet(SHEET_NAME_PRODUCTS);
      
      // If sheet is empty, write headers matching your exact sheet columns
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(["Product ID", "Category", "Name", "Price", "Stock", "Description", "Image Name", "Image URL"]);
      }
      
      var fileUrl = "";
      var fileName = "";
      
      // Handle Image Upload to Google Drive if provided
      if (data.fileData && data.fileName) {
        try {
          var decodedFile = Utilities.base64Decode(data.fileData);
          var blob = Utilities.newBlob(decodedFile, data.mimeType, data.fileName);
          var masterFolderId = "13tdIy2qqMExxIjuw6AZ1m_zOxZWgIYqO"; 
          var masterFolder = DriveApp.getFolderById(masterFolderId);
          var file = masterFolder.createFile(blob);
          
          try {
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          } catch (err) {}
          
          fileUrl = file.getUrl();
          fileName = data.fileName;
        } catch (driveErr) {
          var rootFolder = DriveApp.getRootFolder();
          var decodedFile = Utilities.base64Decode(data.fileData);
          var blob = Utilities.newBlob(decodedFile, data.mimeType, data.fileName);
          var file = rootFolder.createFile(blob);
          fileUrl = file.getUrl();
          fileName = data.fileName;
        }
      }
      
      var productId = 'PROD-' + Math.floor(100000 + Math.random() * 900000);
      
      // Columns: Product ID (A), Category (B), Name (C), Price (D), Stock (E), Description (F), Image Name (G), Image URL (H)
      sheet.appendRow([
        productId,
        data.category,
        data.productName,
        data.price,
        data.stock || "In Stock",
        data.details || "",
        fileName,
        fileUrl
      ]);
      
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        productId: productId,
        imageUrl: fileUrl
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var sheet = ss.getSheetByName(SHEET_NAME_ORDERS) || ss.getActiveSheet();

    // Handle Order Updates from Admin Dashboard[cite: 12]
    if (data.action === "updateOrder") {
      var orderId = data.orderId;
      var rows = sheet.getDataRange().getValues();
      var targetRow = -1;

      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0] && rows[i][0].toString() === orderId.toString()) {
          targetRow = i + 1;
          break;
        }
      }

      if (targetRow !== -1) {
        sheet.getRange(targetRow, 10).setValue(data.orderStatus || "Order Received");
        sheet.getRange(targetRow, 11).setValue(data.paymentStatus || "Pending");
        sheet.getRange(targetRow, 12).setValue(data.amount || "");

        return ContentService.createTextOutput(JSON.stringify({
          status: "success",
          message: "Order updated successfully"
        })).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({
          status: "error",
          message: "Order ID not found"
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // Handle New Order & File Upload to Drive[cite: 12]
    var decodedFile = Utilities.base64Decode(data.fileData);
    var blob = Utilities.newBlob(decodedFile, data.mimeType, data.fileName);
    
    var masterFolderId = "13tdIy2qqMExxIjuw6AZ1m_zOxZWgIYqO"; 
    var targetFolder;
    var targetFolderName = "Customer Uploads";

    try {
      var masterFolder = DriveApp.getFolderById(masterFolderId);
      var service = data.service ? data.service.toLowerCase() : "";
      
      if (service.includes("xerox") || service.includes("photocopy")) {
        targetFolderName = "Xerox Orders";
      } else if (service.includes("printing") || service.includes("book") || service.includes("binding") || service.includes("lamination")) {
        targetFolderName = "Printing Orders";
      } else if (service.includes("logo") || service.includes("design") || service.includes("photo")) {
        targetFolderName = "Logo Designs";
      }
      
      var subFolders = masterFolder.getFoldersByName(targetFolderName);
      targetFolder = subFolders.hasNext() ? subFolders.next() : masterFolder;
    } catch (driveErr) {
      targetFolder = DriveApp.getRootFolder();
      targetFolderName = "Root (Fallback)";
    }
    
    var file = targetFolder.createFile(blob);
    
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {}
    
    var fileUrl = file.getUrl();
    var fileName = data.fileName; 
    
    sheet.appendRow([
      data.orderId,
      new Date(),
      data.customerName,
      data.mobileNumber,
      data.service,
      data.quantity,
      data.requirements || "",
      fileName,  
      fileUrl,   
      "Order Received", 
      "Pending",        
      ""                
    ]);

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      fileUrl: fileUrl,
      folderSorted: targetFolderName
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var action = e && e.parameter ? e.parameter.action : "";
  var callback = e && e.parameter ? e.parameter.callback : "";

  // Handle Fetching Orders[cite: 12]
  if (action === "getOrders") {
    var sheet = ss.getSheetByName(SHEET_NAME_ORDERS) || ss.getActiveSheet();
    var rows = sheet.getDataRange().getValues();
    var orders = [];

    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0]) { 
        orders.push({
          orderId: rows[i][0],
          date: rows[i][1],
          customerName: rows[i][2],
          mobileNumber: rows[i][3],
          service: rows[i][4],
          quantity: rows[i][5],
          requirements: rows[i][6],
          fileName: rows[i][7] || "View File",
          fileUrl: rows[i][8] || "#",
          orderStatus: rows[i][9] || "Order Received",
          paymentStatus: rows[i][10] || "Pending",
          amount: rows[i][11] || ""
        });
      }
    }

    var dataObj = { status: "success", orders: orders.reverse() };
    if (callback) {
      return ContentService.createTextOutput(callback + "(" + JSON.stringify(dataObj) + ");").setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(JSON.stringify(dataObj)).setMimeType(ContentService.MimeType.JSON);
  }

  // Handle Fetching Store Products for shop.html and admin2.html
  if (action === "getProducts") {
    var sheet = ss.getSheetByName(SHEET_NAME_PRODUCTS);
    var products = [];
    if (sheet && sheet.getLastRow() > 1) {
      var rows = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0]) {
          products.push({
            productId: rows[i][0],
            category: rows[i][1],
            productName: rows[i][2],
            price: rows[i][3],
            stock: rows[i][4],
            details: rows[i][5],
            imageUrl: rows[i][7] // Column H stores Image URL from Drive
          });
        }
      }
    }
    var dataObj = { status: "success", products: products.reverse() };
    if (callback) {
      return ContentService.createTextOutput(callback + "(" + JSON.stringify(dataObj) + ");").setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(JSON.stringify(dataObj)).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: "active" })).setMimeType(ContentService.MimeType.JSON);
}