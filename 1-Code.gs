/**
 * ============================================================
 * Code.gs - MABONENG ART eCOMMERCE PLATFORM
 * ============================================================
 * 
 * Full Stack Google App Script Application
 * Adapted from Mabon Suites Hotel to eCommerce
 * 
 * Features:
 * - Public storefront with product catalog API
 * - Shopping cart checkout with PayPal + EFT
 * - Order management admin dashboard
 * - Inventory management with inline editing
 * - Financial ledger with VAT calculation
 * - PDF order confirmation generation
 * - Automated email system
 * - Soft-delete with restore
 * - System audit logging
 * ============================================================
 */

var CONFIG = {
  SPREADSHEET_ID: '1NnsMcz3r6yGfzkXhI3SyF5RwMfKTR2USrxt89Qn5gfI',
  DRIVE_FOLDER_ID: '1TfQhap8ZU0vbdyfLDcBh9vJFixo03JTZ',
  INVENTORY_SHEET: 'Inventory',
  ORDERS_SHEET: 'Orders',
  TRANSACTIONS_SHEET: 'Transactions',
  COMMENTS_SHEET: 'Comments',
  DELETED_SHEET: 'Deleted',
  LOGS_SHEET: 'SystemLogs',
  APP_NAME: 'Maboneng Art',
  COMPANY_EMAIL: 'orders@maboneng.art',
  COMPANY_PHONE: '+27 11 XXX XXXX',
  COMPANY_ADDRESS: 'Johannesburg, South Africa',
  VAT_RATE: 0.15,
  FREE_SHIPPING_THRESHOLD: 50,
  BANK_NAME: 'First National Bank',
  BANK_ACCOUNT: '62XXXXXXXXX',
  BANK_BRANCH: '250655'
};

/* =========================
   ENTRY POINTS (ROUTING)
========================= */

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'form';

  // Page routing
  if (action === 'form' || action === 'shop')
    return serve('Form', 'Shop | ' + CONFIG.APP_NAME);

  if (action === 'admin' || action === 'orders')
    return serve('Submissions', 'Admin Dashboard | ' + CONFIG.APP_NAME);

  if (action === 'inventory')
    return serve('Inventory', 'Inventory Manager | ' + CONFIG.APP_NAME);

  if (action === 'new-inventory')
    return serve('newInventory', 'Add Product | ' + CONFIG.APP_NAME);

  if (action === 'transactions' || action === 'billing')
    return serve('Transactions', 'Transactions | ' + CONFIG.APP_NAME);

  if (action === 'deleted')
    return serve('Deleted', 'Cancelled Orders | ' + CONFIG.APP_NAME);

  // API routing (JSON responses)
  if (action === 'get-inventory')
    return jsonResponse(getInventoryForStorefront());

  if (action === 'get-orders')
    return jsonResponse(getAllOrders());

  if (action === 'get-orders-with-comments')
    return jsonResponse(getAllOrdersWithComments());

  if (action === 'get-order')
    return jsonResponse(getOrderByReference(e.parameter.reference));

  if (action === 'get-comments')
    return jsonResponse({ comments: getCommentsForOrder(e.parameter.reference) });

  if (action === 'get-deleted')
    return jsonResponse(getDeletedOrders());

  if (action === 'get-deleted-count')
    return jsonResponse(getDeletedOrdersCount());

  if (action === 'get-stats')
    return jsonResponse(getStats());

  if (action === 'get-transactions')
    return jsonResponse(getTransactionsByReference(e.parameter.reference));

  if (action === 'get-all-transactions')
    return jsonResponse(getAllTransactions());

  if (action === 'get-inventory-item')
    return jsonResponse(getInventoryItem(e.parameter.id));

  if (action === 'get-system-logs')
    return jsonResponse(getSystemLogs());

  return jsonResponse({ error: 'Unknown action: ' + action });
}

function doPost(e) {
  var data = e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
  var action = data.action;

  // Order operations
  if (action === 'submit-order')     return jsonResponse(submitOrder(data));
  if (action === 'update-field')     return jsonResponse(updateOrderField(data));
  if (action === 'send-email')       return jsonResponse(sendCustomEmail(data));
  if (action === 'send-confirmation') return jsonResponse(sendOrderConfirmation(data));

  // Delete / Restore
  if (action === 'delete')           return jsonResponse(moveToDeleted(data));
  if (action === 'restore')          return jsonResponse(restoreFromDeleted(data));
  if (action === 'permanent-delete') return jsonResponse(permanentDelete(data));

  // Transactions
  if (action === 'add-transaction')    return jsonResponse(addTransaction(data));
  if (action === 'delete-transaction') return jsonResponse(deleteTransaction(data));
  if (action === 'send-statement')     return jsonResponse(sendStatementEmail(data));

  // Comments
  if (action === 'add-comment') return jsonResponse(addCommentInternal(data));

  // Inventory
  if (action === 'add-inventory')    return jsonResponse(addInventoryItem(data));
  if (action === 'update-inventory') return jsonResponse(updateInventoryItem(data));

  return jsonResponse({ error: 'Unknown action: ' + action });
}

function serve(file, title) {
  return HtmlService.createHtmlOutputFromFile(file)
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Gets the spreadsheet using container-bound approach first,
 * falls back to openById. This fixes the "Unexpected error
 * while getting the method or property openById" error.
 */
function getSpreadsheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch(e) { /* not container-bound, use openById */ }
  return getSpreadsheet();
}

function jsonResponse(data) {
  var result = {};
  if (data) {
    for (var key in data) {
      if (data.hasOwnProperty(key)) result[key] = data[key];
    }
  }
  result.success = !data || !data.error;
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/* =========================
   ORDER SUBMISSION (CHECKOUT)
========================= */

function submitOrder(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    if (!data.email) return { error: 'Email address is required' };
    if (!data.name)  return { error: 'Name is required' };
    if (!data.items || data.items.length === 0) return { error: 'Cart is empty' };

    var ss = getSpreadsheet();
    var orderSheet = getOrCreateSheet(ss, CONFIG.ORDERS_SHEET, [
      'Reference','Timestamp','Status','Customer_Name','Customer_Surname','Email','Phone',
      'Country','Shipping_Address','City','Postal_Code','Items_JSON','Item_Count',
      'Subtotal','Shipping_Fee','Discount_Amount','Total_Savings','Total_Amount',
      'Payment_Method','Payment_Status','PayPal_Transaction','EFT_Reference',
      'Confirmation_Sent','Tracking_Number','Notes','Last_Updated'
    ]);

    var surname = sanitize(data.surname || '');
    var reference = generateReference(surname);
    var timestamp = new Date().toISOString();

    // Calculate totals
    var items = data.items;
    var itemCount = items.reduce(function(s, i) { return s + (i.quantity || 1); }, 0);
    var subtotal = items.reduce(function(s, i) { return s + (i.price * (i.quantity || 1)); }, 0);
    var totalSavings = items.reduce(function(s, i) { return s + ((i.original - i.price) * (i.quantity || 1)); }, 0);
    var shippingFee = subtotal >= CONFIG.FREE_SHIPPING_THRESHOLD ? 0 : 9.99;
    var discountAmount = parseFloat(data.discountAmount) || 0;
    var totalAmount = subtotal + shippingFee - discountAmount;
    var paymentMethod = data.paymentMethod || 'pending';

    // Append order row (26 columns)
    orderSheet.appendRow([
      reference,                              // 0  Reference
      timestamp,                              // 1  Timestamp
      'new',                                  // 2  Status
      sanitize(data.name),                    // 3  Customer_Name
      surname,                                // 4  Customer_Surname
      sanitize(data.email),                   // 5  Email
      sanitize(data.phone || ''),             // 6  Phone
      sanitize(data.country || 'South Africa'), // 7  Country
      sanitize(data.address || ''),           // 8  Shipping_Address
      sanitize(data.city || ''),              // 9  City
      sanitize(data.postalCode || ''),        // 10 Postal_Code
      JSON.stringify(items),                  // 11 Items_JSON
      itemCount,                              // 12 Item_Count
      subtotal,                               // 13 Subtotal
      shippingFee,                            // 14 Shipping_Fee
      discountAmount,                         // 15 Discount_Amount
      totalSavings,                           // 16 Total_Savings
      totalAmount,                            // 17 Total_Amount
      paymentMethod,                          // 18 Payment_Method
      paymentMethod === 'paypal' ? 'paid' : 'pending', // 19 Payment_Status
      data.paypalTransactionId || '',         // 20 PayPal_Transaction
      '',                                     // 21 EFT_Reference
      false,                                  // 22 Confirmation_Sent
      '',                                     // 23 Tracking_Number
      '',                                     // 24 Notes
      timestamp                               // 25 Last_Updated
    ]);

    // Deduct stock from inventory
    deductStock(ss, items);

    // Create initial transaction
    createInitialTransaction(ss, reference, totalAmount, itemCount);

    // Add system comment
    addCommentInternal({ reference: reference, author: 'SYSTEM', text: 'Order submitted via ' + paymentMethod });

    // Log system event
    logSystemEvent('ORDER_CREATED', 'SYSTEM', 'Order ' + reference + ' created. Total: $' + totalAmount.toFixed(2));

    // Send acknowledgement email
    sendOrderAcknowledgement(data, reference, {
      items: items, itemCount: itemCount, subtotal: subtotal,
      shippingFee: shippingFee, totalSavings: totalSavings,
      totalAmount: totalAmount, paymentMethod: paymentMethod
    });

    return { success: true, reference: reference };

  } catch(e) {
    Logger.log('submitOrder error: ' + e.message);
    return { error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// Called from Form.html via google.script.run
function processCheckout(formData) {
  var result = submitOrder(formData);
  if (result.error) throw new Error(result.error);
  return { success: true, reference: result.reference };
}

/* =========================
   INVENTORY MANAGEMENT
========================= */

function getInventoryForStorefront() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.INVENTORY_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return { products: [] };

    var rows = sheet.getDataRange().getValues();
    var products = [];

    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0] || r[23] === 'inactive' || r[23] === 'archived') continue;
      if (r[15] <= 0) continue; // Skip out-of-stock

      products.push({
        id: String(r[0]),
        sku: String(r[1] || ''),
        name: String(r[2] || ''),
        desc: String(r[3] || ''),
        category: String(r[4] || 'Uncategorized'),
        size: String(r[6] || 'Medium'),
        colour: String(r[9] || ''),
        artist: String(r[11] || ''),
        original: parseFloat(r[12]) || 0,
        price: parseFloat(r[13]) || 0,
        stock: parseInt(r[15]) || 0,
        image: String(r[17] || ''),
        image2: String(r[18] || ''),
        image3: String(r[19] || ''),
        video: String(r[20] || ''),
        dateListed: r[21] ? (r[21] instanceof Date ? r[21].toISOString() : String(r[21])) : '',
        featured: r[24] === true || r[24] === 'TRUE',
        tags: String(r[25] || '')
      });
    }

    return { products: products };
  } catch(e) {
    Logger.log('getInventoryForStorefront error: ' + e.message);
    return { products: [], error: e.message };
  }
}

function getInventoryItem(id) {
  if (!id) return { error: 'Product ID required' };
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.INVENTORY_SHEET);
  if (!sheet) return { error: 'Inventory sheet not found' };

  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      return { product: rowToInventoryObject(rows[i]) };
    }
  }
  return { error: 'Product not found' };
}

function addInventoryItem(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var ss = getSpreadsheet();
    var sheet = getOrCreateSheet(ss, CONFIG.INVENTORY_SHEET, []);
    var productId = 'PROD-' + generateUUID().substring(0, 8).toUpperCase();
    var timestamp = new Date().toISOString();

    sheet.appendRow([
      productId,                                  // 0  Product_ID
      sanitize(data.sku || ''),                   // 1  SKU
      sanitize(data.name || ''),                  // 2  Name
      sanitize(data.description || ''),           // 3  Description
      sanitize(data.category || 'Uncategorized'), // 4  Category
      sanitize(data.subCategory || ''),           // 5  Sub_Category
      sanitize(data.size || 'Medium'),            // 6  Size
      sanitize(data.dimensions || ''),            // 7  Dimensions
      parseFloat(data.weight) || 0,               // 8  Weight_KG
      sanitize(data.colour || ''),                // 9  Colour
      sanitize(data.material || ''),              // 10 Material
      sanitize(data.artist || ''),                // 11 Artist
      parseFloat(data.originalPrice) || 0,        // 12 Original_Price
      parseFloat(data.currentPrice) || 0,         // 13 Current_Price
      parseFloat(data.costPrice) || 0,            // 14 Cost_Price
      parseInt(data.stock) || 0,                  // 15 Stock
      parseInt(data.minStock) || 5,               // 16 Min_Stock
      sanitize(data.image1 || ''),                // 17 Image_1
      sanitize(data.image2 || ''),                // 18 Image_2
      sanitize(data.image3 || ''),                // 19 Image_3
      sanitize(data.video || ''),                 // 20 Video_URL
      timestamp,                                  // 21 Date_Listed
      timestamp,                                  // 22 Last_Updated
      'active',                                   // 23 Status
      data.featured || false,                     // 24 Featured
      sanitize(data.tags || ''),                  // 25 Tags
      sanitize(data.shippingClass || 'standard'), // 26 Shipping_Class
      sanitize(data.countryOrigin || ''),         // 27 Country_Origin
      0,                                          // 28 Total_Sold
      sanitize(data.notes || '')                  // 29 Notes
    ]);

    logSystemEvent('INVENTORY_ADDED', data.addedBy || 'Admin', 'Product added: ' + data.name + ' (ID: ' + productId + ')');

    return { success: true, productId: productId };
  } catch(e) {
    Logger.log('addInventoryItem error: ' + e.message);
    return { error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function updateInventoryItem(data) {
  if (!data.productId || !data.field) return { error: 'Product ID and field required' };

  var fieldMap = {
    name: 3, description: 4, category: 5, size: 7, colour: 10,
    originalPrice: 13, currentPrice: 14, costPrice: 15,
    stock: 16, image1: 18, image2: 19, image3: 20, video: 21,
    status: 24, featured: 25, notes: 30
  };

  var col = fieldMap[data.field];
  if (!col) return { error: 'Invalid field: ' + data.field };

  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.INVENTORY_SHEET);
    var rows = sheet.getDataRange().getValues();

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.productId)) {
        sheet.getRange(i + 1, col).setValue(data.value);
        sheet.getRange(i + 1, 23).setValue(new Date().toISOString()); // Last_Updated
        logSystemEvent('INVENTORY_UPDATED', data.updatedBy || 'Admin',
          'Product ' + data.productId + ': ' + data.field + ' changed to ' + data.value);
        return { success: true };
      }
    }
    return { error: 'Product not found' };
  } catch(e) {
    return { error: e.message };
  }
}

function deductStock(ss, items) {
  try {
    var sheet = ss.getSheetByName(CONFIG.INVENTORY_SHEET);
    if (!sheet) return;
    var rows = sheet.getDataRange().getValues();

    for (var j = 0; j < items.length; j++) {
      var item = items[j];
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(item.id)) {
          var currentStock = parseInt(rows[i][15]) || 0;
          var qty = parseInt(item.quantity) || 1;
          var newStock = Math.max(0, currentStock - qty);
          var totalSold = (parseInt(rows[i][28]) || 0) + qty;
          sheet.getRange(i + 1, 16).setValue(newStock);     // Stock column
          sheet.getRange(i + 1, 29).setValue(totalSold);    // Total_Sold column
          sheet.getRange(i + 1, 23).setValue(new Date().toISOString());
          break;
        }
      }
    }
  } catch(e) {
    Logger.log('deductStock error: ' + e.message);
  }
}

function rowToInventoryObject(r) {
  return {
    id: String(r[0] || ''), sku: String(r[1] || ''), name: String(r[2] || ''),
    description: String(r[3] || ''), category: String(r[4] || ''),
    subCategory: String(r[5] || ''), size: String(r[6] || ''),
    dimensions: String(r[7] || ''), weight: parseFloat(r[8]) || 0,
    colour: String(r[9] || ''), material: String(r[10] || ''),
    artist: String(r[11] || ''), originalPrice: parseFloat(r[12]) || 0,
    currentPrice: parseFloat(r[13]) || 0, costPrice: parseFloat(r[14]) || 0,
    stock: parseInt(r[15]) || 0, minStock: parseInt(r[16]) || 5,
    image1: String(r[17] || ''), image2: String(r[18] || ''),
    image3: String(r[19] || ''), video: String(r[20] || ''),
    dateListed: String(r[21] || ''), lastUpdated: String(r[22] || ''),
    status: String(r[23] || 'active'), featured: r[24] === true,
    tags: String(r[25] || ''), shippingClass: String(r[26] || 'standard'),
    countryOrigin: String(r[27] || ''), totalSold: parseInt(r[28]) || 0,
    notes: String(r[29] || '')
  };
}

/* =========================
   READ OPERATIONS (ORDERS)
========================= */

function getAllOrders() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.ORDERS_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return { orders: [] };

    var rows = sheet.getDataRange().getValues();
    var orders = [];

    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      orders.push(rowToOrder(r));
    }

    return { orders: orders };
  } catch(e) {
    Logger.log('getAllOrders error: ' + e.message);
    return { orders: [], error: e.message };
  }
}

function getAllOrdersWithComments() {
  try {
    var result = getAllOrders();
    var orders = result.orders || [];
    if (result.error) return result;

    var ss = getSpreadsheet();
    var commentsSheet = ss.getSheetByName(CONFIG.COMMENTS_SHEET);
    var txnSheet = ss.getSheetByName(CONFIG.TRANSACTIONS_SHEET);

    // Comment counts
    var commentCounts = {};
    if (commentsSheet && commentsSheet.getLastRow() > 1) {
      var cRows = commentsSheet.getDataRange().getValues();
      for (var c = 1; c < cRows.length; c++) {
        var ref = cRows[c][1];
        if (ref) commentCounts[ref] = (commentCounts[ref] || 0) + 1;
      }
    }

    // Balance from transactions
    var balanceMap = {};
    if (txnSheet && txnSheet.getLastRow() > 1) {
      var tRows = txnSheet.getDataRange().getValues();
      for (var t = 1; t < tRows.length; t++) {
        var tRef = tRows[t][1];
        var amt = parseFloat(tRows[t][4]) || 0;
        if (tRef) {
          if (!balanceMap[tRef]) balanceMap[tRef] = { charges: 0, payments: 0 };
          if (amt >= 0) balanceMap[tRef].charges += amt;
          else balanceMap[tRef].payments += Math.abs(amt);
        }
      }
    }

    for (var j = 0; j < orders.length; j++) {
      orders[j].commentCount = commentCounts[orders[j].reference] || 0;
      var txn = balanceMap[orders[j].reference] || { charges: 0, payments: 0 };
      orders[j].balance = (txn.charges === 0 && txn.payments === 0)
        ? orders[j].totalAmount
        : txn.charges - txn.payments;
    }

    return { orders: orders };
  } catch(e) {
    return { orders: [], error: e.message };
  }
}

function getOrderByReference(reference) {
  if (!reference) return null;
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.ORDERS_SHEET);
  if (!sheet) return null;
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(reference)) {
      var order = rowToOrder(rows[i]);
      order.comments = getCommentsForOrder(reference);
      return order;
    }
  }
  return null;
}

function rowToOrder(r) {
  var timestamp = r[1] instanceof Date ? r[1].toISOString() : String(r[1] || '');
  return {
    reference: String(r[0] || ''), timestamp: timestamp,
    status: String(r[2] || 'new').toLowerCase().trim(),
    customerName: String(r[3] || ''), customerSurname: String(r[4] || ''),
    email: String(r[5] || ''), phone: String(r[6] || ''),
    country: String(r[7] || ''), shippingAddress: String(r[8] || ''),
    city: String(r[9] || ''), postalCode: String(r[10] || ''),
    itemsJson: String(r[11] || '[]'), itemCount: parseInt(r[12]) || 0,
    subtotal: parseFloat(r[13]) || 0, shippingFee: parseFloat(r[14]) || 0,
    discountAmount: parseFloat(r[15]) || 0, totalSavings: parseFloat(r[16]) || 0,
    totalAmount: parseFloat(r[17]) || 0,
    paymentMethod: String(r[18] || 'pending'),
    paymentStatus: String(r[19] || 'pending').toLowerCase(),
    paypalTransaction: String(r[20] || ''),
    eftReference: String(r[21] || ''),
    confirmationSent: r[22] === true || r[22] === 'TRUE',
    trackingNumber: String(r[23] || ''),
    notes: String(r[24] || ''),
    lastUpdated: String(r[25] || '')
  };
}

function getStats() {
  var result = getAllOrders();
  var orders = result.orders || [];
  var stats = { total: orders.length, new: 0, processing: 0, shipped: 0, delivered: 0, cancelled: 0, totalRevenue: 0, paidRevenue: 0 };
  for (var i = 0; i < orders.length; i++) {
    var s = orders[i].status;
    if (stats.hasOwnProperty(s)) stats[s]++;
    stats.totalRevenue += orders[i].totalAmount;
    if (orders[i].paymentStatus === 'paid') stats.paidRevenue += orders[i].totalAmount;
  }
  return stats;
}

/* =========================
   UPDATE ORDER FIELDS
========================= */

function updateOrderField(data) {
  var reference = data.reference, field = data.field, value = data.value;
  if (!reference || !field) return { error: 'Missing reference or field' };

  var fieldMap = {
    status: 3, customerName: 4, email: 6, phone: 7,
    paymentStatus: 20, paymentMethod: 19, trackingNumber: 24,
    confirmationSent: 23, notes: 25
  };
  var col = fieldMap[field];
  if (!col) return { error: 'Invalid field: ' + field };

  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.ORDERS_SHEET);
    var rows = sheet.getDataRange().getValues();

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(reference).trim()) {
        sheet.getRange(i + 1, col).setValue(value);
        sheet.getRange(i + 1, 26).setValue(new Date().toISOString());

        if (field === 'status') addCommentInternal({ reference: reference, author: 'Admin', text: 'Status changed to: ' + value });
        if (field === 'paymentStatus') addCommentInternal({ reference: reference, author: 'Admin', text: 'Payment status: ' + value });
        if (field === 'trackingNumber') addCommentInternal({ reference: reference, author: 'Admin', text: 'Tracking number added: ' + value });

        logSystemEvent('ORDER_UPDATED', data.updatedBy || 'Admin', reference + ': ' + field + ' = ' + value);
        return { success: true };
      }
    }
    return { error: 'Order not found: ' + reference };
  } catch(e) {
    return { error: e.message };
  }
}

/* =========================
   COMMENTS
========================= */

function addCommentInternal(data) {
  if (!data.reference || !data.text) return { error: 'Invalid comment data' };
  var ss = getSpreadsheet();
  getOrCreateSheet(ss, CONFIG.COMMENTS_SHEET, ['Comment_ID','Reference','Author','Text','Timestamp']).appendRow([
    generateUUID(), data.reference, data.author || 'Admin', sanitize(data.text), new Date().toISOString()
  ]);
  return { success: true };
}

function addComment(reference, text, author) {
  var result = addCommentInternal({ reference: reference, author: author || 'Admin', text: text });
  if (result.error) throw new Error(result.error);
  return result;
}

function getCommentsForOrder(reference) {
  try {
    if (!reference) return { comments: [] };
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.COMMENTS_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return { comments: [] };
    var rows = sheet.getDataRange().getValues();
    var comments = [];
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === String(reference)) {
        comments.push({
          id: String(rows[i][0]),
          author: String(rows[i][2]),
          text: String(rows[i][3]),
          timestamp: rows[i][4] ? new Date(rows[i][4]).toISOString() : ''
        });
      }
    }
    comments.reverse();
    return { comments: comments };
  } catch(e) {
    Logger.log('getCommentsForOrder error: ' + e.message);
    return { comments: [], error: e.message };
  }
}

/* =========================
   DELETE / RESTORE
========================= */

function moveToDeleted(data) {
  if (!data.reference) return { error: 'Reference required' };
  if (!data.deleteReason) return { error: 'Reason required' };

  var ss = getSpreadsheet();
  var orderSheet = ss.getSheetByName(CONFIG.ORDERS_SHEET);
  var delSheet = getOrCreateSheet(ss, CONFIG.DELETED_SHEET, ['Reference','Deleted_At','Reason','Deleted_By','Snapshot']);

  var rows = orderSheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.reference)) {
      delSheet.appendRow([data.reference, new Date().toISOString(), data.deleteReason, data.deletedBy || 'Admin', JSON.stringify(rows[i])]);
      orderSheet.deleteRow(i + 1);
      addCommentInternal({ reference: data.reference, author: 'SYSTEM', text: 'Order cancelled: ' + data.deleteReason });
      logSystemEvent('ORDER_DELETED', data.deletedBy || 'Admin', 'Order ' + data.reference + ' cancelled: ' + data.deleteReason);
      return { success: true };
    }
  }
  return { error: 'Order not found' };
}

function restoreFromDeleted(data) {
  var ss = getSpreadsheet();
  var delSheet = ss.getSheetByName(CONFIG.DELETED_SHEET);
  var orderSheet = getOrCreateSheet(ss, CONFIG.ORDERS_SHEET, []);

  var rows = delSheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.reference)) {
      var snapshot = JSON.parse(rows[i][4]);
      snapshot[2] = 'processing';
      snapshot[25] = new Date().toISOString();
      orderSheet.appendRow(snapshot);
      delSheet.deleteRow(i + 1);
      addCommentInternal({ reference: data.reference, author: 'SYSTEM', text: 'Order restored' });
      logSystemEvent('ORDER_RESTORED', data.restoredBy || 'Admin', 'Order ' + data.reference + ' restored');
      return { success: true };
    }
  }
  return { error: 'Not found in deleted' };
}

function permanentDelete(data) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.DELETED_SHEET);
  if (!sheet) return { error: 'Deleted sheet not found' };
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.reference)) {
      sheet.deleteRow(i + 1);
      logSystemEvent('ORDER_PURGED', data.deletedBy || 'Admin', 'Order ' + data.reference + ' permanently deleted');
      return { success: true };
    }
  }
  return { error: 'Not found' };
}

function getDeletedOrders() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.DELETED_SHEET);
  if (!sheet) return { deleted: [] };
  var rows = sheet.getDataRange().getValues();
  var deleted = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    var snap = {};
    try { snap = JSON.parse(rows[i][4]); } catch(e) {}
    deleted.push({
      reference: rows[i][0], deletedAt: rows[i][1], deleteReason: rows[i][2],
      deletedBy: rows[i][3], customerName: snap[3] || '', email: snap[5] || '',
      totalAmount: snap[17] || 0, itemCount: snap[12] || 0, snapshot: rows[i][4]
    });
  }
  return { deleted: deleted };
}

function getDeletedOrdersCount() {
  try {
    var result = getDeletedOrders();
    return { deletedCount: (result.deleted || []).length };
  } catch(e) { return { deletedCount: 0 }; }
}

/* =========================
   TRANSACTIONS / BILLING
========================= */

function addTransaction(data) {
  if (!data.reference) return { error: 'Reference required' };
  if (!data.description) return { error: 'Description required' };
  if (data.amount === undefined) return { error: 'Amount required' };

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = getSpreadsheet();
    var sheet = getOrCreateSheet(ss, CONFIG.TRANSACTIONS_SHEET, []);
    var txnId = generateUUID();
    sheet.appendRow([
      txnId, data.reference, data.date || new Date().toISOString().split('T')[0],
      sanitize(data.description), parseFloat(data.amount),
      (data.type || 'CHARGE').toUpperCase(), data.createdBy || 'Admin', new Date().toISOString()
    ]);
    addCommentInternal({ reference: data.reference, author: 'BILLING', text: 'Transaction: ' + data.description + ' - $' + Math.abs(parseFloat(data.amount)).toFixed(2) });
    logSystemEvent('TRANSACTION_ADDED', data.createdBy || 'Admin', 'Txn ' + txnId + ' for ' + data.reference);
    return { success: true, transactionId: txnId };
  } catch(e) { return { error: e.message }; }
  finally { lock.releaseLock(); }
}

function createInitialTransaction(ss, reference, totalAmount, itemCount) {
  try {
    var sheet = getOrCreateSheet(ss, CONFIG.TRANSACTIONS_SHEET, []);
    sheet.appendRow([
      generateUUID(), reference, new Date().toISOString().split('T')[0],
      'Order: ' + itemCount + ' item(s)', totalAmount, 'CHARGE', 'SYSTEM', new Date().toISOString()
    ]);
  } catch(e) { Logger.log('createInitialTransaction error: ' + e.message); }
}

function deleteTransaction(data) {
  if (!data.transactionId) return { error: 'Transaction ID required' };
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.TRANSACTIONS_SHEET);
    if (!sheet) return { error: 'Sheet not found' };
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.transactionId)) {
        sheet.deleteRow(i + 1);
        logSystemEvent('TRANSACTION_DELETED', data.deletedBy || 'Admin', 'Txn ' + data.transactionId + ' deleted');
        return { success: true };
      }
    }
    return { error: 'Transaction not found' };
  } catch(e) { return { error: e.message }; }
}

function updateTransaction(data) {
  if (!data.transactionId) return { error: 'Transaction ID required' };
  if (!data.description) return { error: 'Description required' };
  if (data.amount === undefined) return { error: 'Amount required' };

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.TRANSACTIONS_SHEET);
    if (!sheet) return { error: 'Sheet not found' };
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.transactionId)) {
        var row = i + 1;
        sheet.getRange(row, 4).setValue(sanitize(data.description));
        sheet.getRange(row, 5).setValue(parseFloat(data.amount));
        sheet.getRange(row, 6).setValue((data.type || 'CHARGE').toUpperCase());
        logSystemEvent('TRANSACTION_UPDATED', data.updatedBy || 'Admin', 'Txn ' + data.transactionId + ' updated: ' + data.description + ' $' + Math.abs(parseFloat(data.amount)).toFixed(2));
        return { success: true };
      }
    }
    return { error: 'Transaction not found' };
  } catch(e) { return { error: e.message }; }
  finally { lock.releaseLock(); }
}

function getTransactionsByReference(reference) {
  try {
    if (!reference) return { transactions: [] };
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.TRANSACTIONS_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return { transactions: [] };
    var rows = sheet.getDataRange().getValues();
    var txns = [];
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === String(reference)) {
        txns.push({
          id: String(rows[i][0]),
          reference: String(rows[i][1]),
          date: rows[i][2] ? String(rows[i][2]) : '',
          description: String(rows[i][3]),
          amount: parseFloat(rows[i][4]) || 0,
          type: String(rows[i][5]),
          createdBy: String(rows[i][6]),
          createdAt: rows[i][7] ? new Date(rows[i][7]).toISOString() : ''
        });
      }
    }
    txns.sort(function(a, b) { return new Date(a.createdAt || a.date) - new Date(b.createdAt || b.date); });
    return { transactions: txns };
  } catch(e) {
    Logger.log('getTransactionsByReference error: ' + e.message);
    return { transactions: [], error: e.message };
  }
}

function getAllTransactions() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.TRANSACTIONS_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return { transactions: [] };
    var rows = sheet.getDataRange().getValues();
    var txns = [];
    for (var i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      txns.push({
        id: String(rows[i][0]),
        reference: String(rows[i][1]),
        date: rows[i][2] ? String(rows[i][2]) : '',
        description: String(rows[i][3]),
        amount: parseFloat(rows[i][4]) || 0,
        type: String(rows[i][5]),
        createdBy: String(rows[i][6]),
        createdAt: rows[i][7] ? new Date(rows[i][7]).toISOString() : ''
      });
    }
    return { transactions: txns };
  } catch(e) {
    Logger.log('getAllTransactions error: ' + e.message);
    return { transactions: [], error: e.message };
  }
}

/* =========================
   EMAIL SYSTEM
========================= */

function sendOrderAcknowledgement(data, reference, pricing) {
  try {
    var email = data.email;
    if (!email) return false;

    var name = data.name || 'Valued Customer';
    var paymentMethod = pricing.paymentMethod;
    var total = pricing.totalAmount;

    var eftSection = '';
    if (paymentMethod === 'eft' || paymentMethod === 'pending') {
      eftSection = '<div style="background:#fff3cd;border-left:4px solid #f59e0b;padding:16px;margin:20px 0;border-radius:4px;">' +
        '<h3 style="margin:0 0 12px;color:#92400e;">EFT Payment Details</h3>' +
        '<table style="width:100%;">' +
        '<tr><td style="padding:4px 0;color:#78350f;">Bank:</td><td style="font-weight:600;">' + CONFIG.BANK_NAME + '</td></tr>' +
        '<tr><td style="padding:4px 0;color:#78350f;">Account:</td><td style="font-weight:600;">' + CONFIG.BANK_ACCOUNT + '</td></tr>' +
        '<tr><td style="padding:4px 0;color:#78350f;">Branch:</td><td style="font-weight:600;">' + CONFIG.BANK_BRANCH + '</td></tr>' +
        '<tr><td style="padding:4px 0;color:#78350f;">Reference:</td><td style="font-weight:700;color:#dc2626;font-size:16px;">' + reference + '</td></tr>' +
        '</table>' +
        '<p style="margin:12px 0 0;color:#92400e;font-weight:600;">IMPORTANT: Use your order reference <strong>' + reference + '</strong> as payment reference.</p>' +
        '</div>';
    }

    var itemsHtml = '';
    for (var i = 0; i < pricing.items.length; i++) {
      var item = pricing.items[i];
      itemsHtml += '<tr><td style="padding:8px 0;border-bottom:1px solid #eee;">' + escapeHtml(item.name) + '</td><td style="padding:8px 0;border-bottom:1px solid #eee;text-align:center;">' + (item.quantity || 1) + '</td><td style="padding:8px 0;border-bottom:1px solid #eee;text-align:right;">$' + (item.price * (item.quantity || 1)).toFixed(2) + '</td></tr>';
    }

    var html = '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f5f5f5;font-family:Arial,sans-serif;">' +
      '<table width="100%" cellpadding="0" cellspacing="0" style="padding:24px;"><tr><td align="center">' +
      '<table width="100%" style="max-width:600px;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,0.1);">' +
      '<tr><td style="background:#18181b;padding:32px;text-align:center;">' +
      '<h1 style="margin:0;color:#c9a962;font-size:28px;">' + CONFIG.APP_NAME + '</h1>' +
      '<p style="margin:8px 0 0;color:#a3a3a3;font-size:14px;">Order Confirmation</p>' +
      '</td></tr>' +
      '<tr><td style="padding:32px;">' +
      '<p>Dear <strong>' + escapeHtml(name) + '</strong>,</p>' +
      '<p>Thank you for your order! Your reference number is:</p>' +
      '<div style="background:#f9fafb;border:2px solid #c9a962;border-radius:8px;padding:20px;text-align:center;margin:20px 0;">' +
      '<div style="font-size:12px;color:#888;">Order Reference</div>' +
      '<div style="font-size:28px;font-weight:700;color:#c9a962;letter-spacing:2px;">' + reference + '</div></div>' +
      eftSection +
      '<h3 style="border-bottom:2px solid #c9a962;padding-bottom:8px;">Order Items</h3>' +
      '<table style="width:100%;border-collapse:collapse;"><tr><th style="text-align:left;padding:8px 0;border-bottom:2px solid #eee;">Item</th><th style="text-align:center;padding:8px 0;border-bottom:2px solid #eee;">Qty</th><th style="text-align:right;padding:8px 0;border-bottom:2px solid #eee;">Price</th></tr>' + itemsHtml + '</table>' +
      '<div style="background:#f9fafb;padding:16px;border-radius:8px;margin-top:16px;">' +
      '<div style="display:flex;justify-content:space-between;margin-bottom:8px;"><span>Subtotal:</span><span>$' + pricing.subtotal.toFixed(2) + '</span></div>' +
      '<div style="display:flex;justify-content:space-between;margin-bottom:8px;color:#22c55e;"><span>You Saved:</span><span>$' + pricing.totalSavings.toFixed(2) + '</span></div>' +
      '<div style="display:flex;justify-content:space-between;margin-bottom:8px;"><span>Shipping:</span><span>' + (pricing.shippingFee === 0 ? 'FREE' : '$' + pricing.shippingFee.toFixed(2)) + '</span></div>' +
      '<div style="display:flex;justify-content:space-between;font-size:20px;font-weight:700;padding-top:12px;border-top:2px solid #eee;color:#c9a962;"><span>Total:</span><span>$' + total.toFixed(2) + '</span></div></div>' +
      '<p style="margin-top:24px;">We look forward to serving you!</p>' +
      '<p><strong style="color:#c9a962;">' + CONFIG.APP_NAME + ' Team</strong></p>' +
      '</td></tr>' +
      '<tr><td style="background:#18181b;padding:16px;text-align:center;font-size:12px;color:#888;">' +
      '<p>' + CONFIG.APP_NAME + ' | ' + CONFIG.COMPANY_ADDRESS + '</p></td></tr>' +
      '</table></td></tr></table></body></html>';

    var pdfBlob = null;
    try { pdfBlob = generateOrderPDF(data, reference, pricing); } catch(e) {}

    var emailOpts = { to: email, subject: 'Order Confirmation - ' + reference + ' | ' + CONFIG.APP_NAME, body: 'Your order ' + reference + ' has been confirmed. Total: $' + total.toFixed(2), htmlBody: html };
    if (pdfBlob) emailOpts.attachments = [pdfBlob];
    MailApp.sendEmail(emailOpts);

    return true;
  } catch(e) {
    Logger.log('sendOrderAcknowledgement error: ' + e.message);
    return false;
  }
}

function sendOrderConfirmation(data) {
  var order = getOrderByReference(data.reference);
  if (!order) return { error: 'Order not found' };
  var items = [];
  try { items = JSON.parse(order.itemsJson || '[]'); } catch(e) {}

  var pricing = {
    items: items, itemCount: order.itemCount, subtotal: order.subtotal,
    shippingFee: order.shippingFee, totalSavings: order.totalSavings,
    totalAmount: order.totalAmount, paymentMethod: order.paymentMethod
  };

  var sent = sendOrderAcknowledgement({
    name: order.customerName, email: data.email || order.email, surname: order.customerSurname
  }, order.reference, pricing);

  if (sent) {
    updateOrderField({ reference: data.reference, field: 'confirmationSent', value: true });
    return { success: true };
  }
  return { error: 'Failed to send' };
}

function sendCustomEmail(data) {
  if (!data.to || !data.subject || !data.body) return { error: 'Missing email parameters' };
  try {
    MailApp.sendEmail({ to: data.to, subject: data.subject + ' | ' + CONFIG.APP_NAME, body: data.body + '\n\n--\n' + CONFIG.APP_NAME });
    if (data.reference) addCommentInternal({ reference: data.reference, author: 'Admin', text: 'Email sent: ' + data.subject });
    logSystemEvent('EMAIL_SENT', 'Admin', 'To: ' + data.to + ' Subject: ' + data.subject);
    return { success: true };
  } catch(e) { return { error: e.message }; }
}

function sendStatementEmail(data) {
  if (!data.reference || !data.email) return { error: 'Reference and email required' };
  try {
    var order = getOrderByReference(data.reference);
    if (!order) return { error: 'Order not found' };
    var txnResult = getTransactionsByReference(data.reference);
    var txns = txnResult.transactions || [];

    var html = '<h2>Account Statement - ' + data.reference + '</h2><p>Customer: ' + escapeHtml(order.customerName) + '</p>';
    html += '<table border="1" cellpadding="8" style="border-collapse:collapse;width:100%;"><tr><th>Date</th><th>Description</th><th>Debit</th><th>Credit</th><th>Balance</th></tr>';
    var balance = 0;
    for (var i = 0; i < txns.length; i++) {
      balance += txns[i].amount;
      html += '<tr><td>' + txns[i].date + '</td><td>' + escapeHtml(txns[i].description) + '</td>';
      html += '<td style="color:red;">' + (txns[i].amount >= 0 ? '$' + txns[i].amount.toFixed(2) : '') + '</td>';
      html += '<td style="color:green;">' + (txns[i].amount < 0 ? '$' + Math.abs(txns[i].amount).toFixed(2) : '') + '</td>';
      html += '<td style="font-weight:bold;">$' + balance.toFixed(2) + '</td></tr>';
    }
    html += '</table><p style="font-size:18px;font-weight:bold;">Balance Due: $' + balance.toFixed(2) + '</p>';

    MailApp.sendEmail({ to: data.email, subject: 'Account Statement - ' + data.reference + ' | ' + CONFIG.APP_NAME, body: 'Please see your account statement attached.', htmlBody: html });
    addCommentInternal({ reference: data.reference, author: 'BILLING', text: 'Statement sent to ' + data.email });
    return { success: true };
  } catch(e) { return { error: e.message }; }
}

/* =========================
   DEBT COLLECTION
========================= */

function sendDebtReminder(data) {
  try {
    if (!data.reference || !data.email) return { error: 'Reference and email required' };
    var order = getOrderByReference(data.reference);
    var customerName = order ? (order.customerName + ' ' + (order.customerSurname || '')).trim() : data.reference;

    var txnResult = getTransactionsByReference(data.reference);
    var txns = txnResult.transactions || [];
    var balance = 0;
    for (var i = 0; i < txns.length; i++) balance += txns[i].amount;

    if (balance <= 0.01) return { error: 'This account has no outstanding balance ($' + balance.toFixed(2) + ')' };

    var subject = (data.subject || 'Outstanding Balance Reminder') + ' | ' + CONFIG.APP_NAME;
    var message = (data.message || 'Dear {name}, your account {reference} has an outstanding balance of {balance}.')
      .replace(/\{name\}/g, customerName)
      .replace(/\{balance\}/g, '$' + balance.toFixed(2))
      .replace(/\{reference\}/g, data.reference);

    // Build HTML statement table
    var html = '<div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;">';
    html += '<div style="background:#1a1a1a;color:#c9a962;padding:20px;text-align:center;border-radius:8px 8px 0 0;">';
    html += '<h1 style="margin:0;font-size:24px;">' + CONFIG.APP_NAME + '</h1></div>';
    html += '<div style="padding:24px;background:#ffffff;border:1px solid #e0e0e0;">';
    html += '<p style="white-space:pre-line;color:#333;font-size:15px;line-height:1.6;">' + escapeHtml(message) + '</p>';
    html += '<hr style="border:none;border-top:2px solid #c9a962;margin:20px 0;">';
    html += '<h3 style="color:#1a1a1a;margin-bottom:12px;">Account Statement: ' + escapeHtml(data.reference) + '</h3>';
    html += '<table style="width:100%;border-collapse:collapse;font-size:14px;">';
    html += '<tr style="background:#f5f5f5;"><th style="padding:10px;text-align:left;border:1px solid #ddd;">Date</th><th style="padding:10px;text-align:left;border:1px solid #ddd;">Description</th><th style="padding:10px;text-align:right;border:1px solid #ddd;">Debit</th><th style="padding:10px;text-align:right;border:1px solid #ddd;">Credit</th><th style="padding:10px;text-align:right;border:1px solid #ddd;">Balance</th></tr>';
    var runBal = 0;
    for (var j = 0; j < txns.length; j++) {
      runBal += txns[j].amount;
      html += '<tr>';
      html += '<td style="padding:8px;border:1px solid #ddd;">' + (txns[j].date || '') + '</td>';
      html += '<td style="padding:8px;border:1px solid #ddd;">' + escapeHtml(txns[j].description) + '</td>';
      html += '<td style="padding:8px;text-align:right;border:1px solid #ddd;color:#dc2626;">' + (txns[j].amount >= 0 ? '$' + txns[j].amount.toFixed(2) : '') + '</td>';
      html += '<td style="padding:8px;text-align:right;border:1px solid #ddd;color:#16a34a;">' + (txns[j].amount < 0 ? '$' + Math.abs(txns[j].amount).toFixed(2) : '') + '</td>';
      html += '<td style="padding:8px;text-align:right;border:1px solid #ddd;font-weight:bold;color:' + (runBal > 0 ? '#dc2626' : '#16a34a') + ';">$' + runBal.toFixed(2) + '</td>';
      html += '</tr>';
    }
    html += '</table>';
    html += '<div style="margin-top:20px;padding:16px;background:#fef2f2;border:2px solid #dc2626;border-radius:8px;text-align:center;">';
    html += '<p style="margin:0;font-size:20px;font-weight:bold;color:#dc2626;">Balance Due: $' + balance.toFixed(2) + '</p></div>';

    // Add bank details for EFT
    html += '<div style="margin-top:20px;padding:16px;background:#f5f5f5;border-radius:8px;">';
    html += '<h4 style="color:#c9a962;margin:0 0 8px;">Payment Details (EFT):</h4>';
    html += '<p style="margin:4px 0;color:#333;">Bank: ' + CONFIG.BANK_NAME + '</p>';
    html += '<p style="margin:4px 0;color:#333;">Account: ' + CONFIG.BANK_ACCOUNT + '</p>';
    html += '<p style="margin:4px 0;color:#333;">Branch: ' + CONFIG.BANK_BRANCH + '</p>';
    html += '<p style="margin:4px 0;font-weight:bold;color:#333;">Reference: ' + escapeHtml(data.reference) + '</p></div>';

    html += '</div>';
    html += '<div style="background:#1a1a1a;color:#999;padding:16px;text-align:center;font-size:12px;border-radius:0 0 8px 8px;">';
    html += '<p style="margin:0;">This is an automated reminder from ' + CONFIG.APP_NAME + '.</p></div></div>';

    MailApp.sendEmail({ to: data.email, subject: subject, body: message, htmlBody: html });
    addCommentInternal({ reference: data.reference, author: 'DEBT_COLLECTION', text: 'Debt reminder sent to ' + data.email + ' | Balance: $' + balance.toFixed(2) });
    logSystemEvent('DEBT_REMINDER', 'Admin', 'Ref: ' + data.reference + ' Email: ' + data.email + ' Balance: $' + balance.toFixed(2));
    return { success: true };
  } catch(e) {
    Logger.log('sendDebtReminder error: ' + e.message);
    return { error: e.message };
  }
}

function sendBulkDebtReminders(data) {
  try {
    if (!data.subject || !data.message) return { error: 'Subject and message required' };

    // Get all orders to find emails
    var ss = getSpreadsheet();
    var orderSheet = ss.getSheetByName(CONFIG.ORDERS_SHEET);
    if (!orderSheet || orderSheet.getLastRow() < 2) return { error: 'No orders found' };

    var orderRows = orderSheet.getDataRange().getValues();
    var orderHeaders = orderRows[0];
    var refCol = orderHeaders.indexOf('Reference');
    var emailCol = orderHeaders.indexOf('Email');
    var nameCol = orderHeaders.indexOf('Customer_Name');
    var surnameCol = orderHeaders.indexOf('Customer_Surname');

    // Build ref -> {email, name} map
    var orderMap = {};
    for (var i = 1; i < orderRows.length; i++) {
      var ref = String(orderRows[i][refCol] || '');
      if (ref) {
        orderMap[ref] = {
          email: String(orderRows[i][emailCol] || ''),
          name: (String(orderRows[i][nameCol] || '') + ' ' + String(orderRows[i][surnameCol] || '')).trim()
        };
      }
    }

    // Get all transactions and compute balances per reference
    var txnSheet = ss.getSheetByName(CONFIG.TRANSACTIONS_SHEET);
    if (!txnSheet || txnSheet.getLastRow() < 2) return { error: 'No transactions found' };
    var txnRows = txnSheet.getDataRange().getValues();
    var balances = {};
    for (var j = 1; j < txnRows.length; j++) {
      var tRef = String(txnRows[j][1] || '');
      if (!tRef) continue;
      if (!balances[tRef]) balances[tRef] = 0;
      balances[tRef] += parseFloat(txnRows[j][4]) || 0;
    }

    // Find owing accounts that have emails
    var sentCount = 0;
    var errors = [];
    for (var refKey in balances) {
      if (!balances.hasOwnProperty(refKey)) continue;
      if (balances[refKey] <= 0.01) continue; // not owing

      var orderInfo = orderMap[refKey];
      if (!orderInfo || !orderInfo.email) {
        errors.push(refKey + ': no email on file');
        continue;
      }

      try {
        var result = sendDebtReminder({
          reference: refKey,
          email: orderInfo.email,
          subject: data.subject,
          message: data.message
        });
        if (result.success) sentCount++;
        else errors.push(refKey + ': ' + (result.error || 'failed'));
      } catch(innerErr) {
        errors.push(refKey + ': ' + innerErr.message);
      }
    }

    logSystemEvent('BULK_DEBT_REMINDER', 'Admin', 'Sent: ' + sentCount + (errors.length > 0 ? ' | Errors: ' + errors.join('; ') : ''));
    return { success: true, sentCount: sentCount, errors: errors };
  } catch(e) {
    Logger.log('sendBulkDebtReminders error: ' + e.message);
    return { error: e.message };
  }
}

/* =========================
   PDF GENERATION
========================= */

function generateOrderPDF(data, reference, pricing) {
  var name = data.name || 'Customer';
  var tempDoc = DocumentApp.create(CONFIG.APP_NAME + '_Order_' + reference);
  var body = tempDoc.getBody();

  var h = body.appendParagraph(CONFIG.APP_NAME);
  h.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  h.setForegroundColor('#c9a962');
  h.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('ORDER CONFIRMATION').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('Reference: ' + reference).setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('Date: ' + new Date().toLocaleDateString()).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendHorizontalRule();

  body.appendParagraph('CUSTOMER').setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph('Name: ' + name);
  body.appendParagraph('Email: ' + (data.email || 'N/A'));
  body.appendHorizontalRule();

  body.appendParagraph('ORDER ITEMS').setHeading(DocumentApp.ParagraphHeading.HEADING3);
  var tableData = [['Item', 'Qty', 'Price']];
  for (var i = 0; i < pricing.items.length; i++) {
    var item = pricing.items[i];
    tableData.push([item.name, String(item.quantity || 1), '$' + (item.price * (item.quantity || 1)).toFixed(2)]);
  }
  body.appendTable(tableData);

  body.appendParagraph('');
  body.appendParagraph('Subtotal: $' + pricing.subtotal.toFixed(2));
  body.appendParagraph('Shipping: ' + (pricing.shippingFee === 0 ? 'FREE' : '$' + pricing.shippingFee.toFixed(2)));
  body.appendParagraph('TOTAL: $' + pricing.totalAmount.toFixed(2)).setBold(true);

  if (pricing.paymentMethod === 'eft' || pricing.paymentMethod === 'pending') {
    body.appendHorizontalRule();
    body.appendParagraph('EFT PAYMENT DETAILS').setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendParagraph('Bank: ' + CONFIG.BANK_NAME);
    body.appendParagraph('Account: ' + CONFIG.BANK_ACCOUNT);
    body.appendParagraph('Branch: ' + CONFIG.BANK_BRANCH);
    body.appendParagraph('Reference: ' + reference).setBold(true);
  }

  tempDoc.saveAndClose();
  var docFile = DriveApp.getFileById(tempDoc.getId());
  var pdfBlob = docFile.getAs('application/pdf').setName(CONFIG.APP_NAME + '_Order_' + reference + '.pdf');
  docFile.setTrashed(true);
  return pdfBlob;
}

/* =========================
   SYSTEM LOGGING
========================= */

function logSystemEvent(action, user, details) {
  try {
    var ss = getSpreadsheet();
    var sheet = getOrCreateSheet(ss, CONFIG.LOGS_SHEET, ['Log_ID','Timestamp','Action','User','Details']);
    sheet.appendRow([generateUUID(), new Date().toISOString(), action, user || 'SYSTEM', details || '']);
  } catch(e) { Logger.log('logSystemEvent error: ' + e.message); }
}

function getSystemLogs() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.LOGS_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return { logs: [] };
    var rows = sheet.getDataRange().getValues();
    var logs = [];
    for (var i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      logs.push({ id: rows[i][0], timestamp: rows[i][1], action: String(rows[i][2]), user: String(rows[i][3]), details: String(rows[i][4]) });
    }
    logs.reverse();
    return { logs: logs };
  } catch(e) { return { logs: [], error: e.message }; }
}

/* =========================
   HELPERS
========================= */

function generateReference(surname) {
  var now = new Date();
  var yy = String(now.getFullYear()).slice(-2);
  var mm = String(now.getMonth() + 1).padStart(2, '0');
  var dd = String(now.getDate()).padStart(2, '0');
  var surnameCode = (surname || 'XXXX').substring(0, 4).toUpperCase().replace(/[^A-Z]/g, 'X');
  while (surnameCode.length < 4) surnameCode += 'X';
  var random = String(Math.floor(Math.random() * 1000)).padStart(3, '0');
  return yy + mm + dd + surnameCode + random;
}

function generateUUID() { return Utilities.getUuid(); }

function sanitize(str) {
  if (typeof str !== 'string') return str;
  return str.replace(/[<>]/g, '').trim();
}

function escapeHtml(str) {
  if (!str) return '';
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length > 0) sheet.appendRow(headers);
  }
  return sheet;
}
