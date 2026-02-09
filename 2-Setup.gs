/**
 * ============================================================
 * Setup.gs - MABONENG ART eCOMMERCE PLATFORM SETUP
 * ============================================================
 *
 * RUN: runSetup()
 *
 * Creates / fixes these sheets:
 * 1. Inventory    (30 columns) - Product catalog
 * 2. Orders       (26 columns) - Customer orders
 * 3. Transactions (8 columns)  - Financial ledger
 * 4. Comments     (5 columns)  - Event log
 * 5. Deleted      (5 columns)  - Tombstone pattern
 * 6. SystemLogs   (5 columns)  - Audit trail
 *
 * Adapted from Mabon Suites Hotel to eCommerce
 * ============================================================
 */

function runSetup() {
  Logger.log('========================================');
  Logger.log('MABONENG ART eCOMMERCE PLATFORM SETUP');
  Logger.log('========================================');

  // *** REPLACE THESE WITH YOUR OWN IDs ***
  var SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
  var DRIVE_FOLDER_ID = 'YOUR_DRIVE_FOLDER_ID_HERE';

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log('Spreadsheet: ' + ss.getName());

  /* =====================================================
     INVENTORY (PRODUCT CATALOG) - 30 COLUMNS
     Source of truth for storefront products
  ===================================================== */

  var inventoryHeaders = [
    'Product_ID',          // 0  - Auto-generated unique ID
    'SKU',                 // 1  - Stock Keeping Unit
    'Name',                // 2  - Product name
    'Description',         // 3  - Product description
    'Category',            // 4  - Art category (Paintings, Sculptures, etc.)
    'Sub_Category',        // 5  - Sub-category
    'Size',                // 6  - Small | Medium | Large
    'Dimensions',          // 7  - e.g. "40x60cm"
    'Weight_KG',           // 8  - Weight in kg
    'Colour',              // 9  - Primary colour
    'Material',            // 10 - Canvas, Metal, Wood, etc.
    'Artist',              // 11 - Artist name
    'Original_Price',      // 12 - Original price (USD)
    'Current_Price',       // 13 - Sale price (USD)
    'Cost_Price',          // 14 - Cost to business (for profit calc)
    'Stock',               // 15 - Current stock quantity
    'Min_Stock',           // 16 - Minimum stock threshold (alerts)
    'Image_1',             // 17 - Primary image URL
    'Image_2',             // 18 - Secondary image URL
    'Image_3',             // 19 - Tertiary image URL
    'Video_URL',           // 20 - Video/GIF URL (.mp4 or .gif)
    'Date_Listed',         // 21 - When product was listed
    'Last_Updated',        // 22 - Last modification timestamp
    'Status',              // 23 - active | inactive | archived
    'Featured',            // 24 - true | false (for homepage)
    'Tags',                // 25 - Comma-separated tags
    'Shipping_Class',      // 26 - standard | express | fragile
    'Country_Origin',      // 27 - Country of origin
    'Total_Sold',          // 28 - Running total of units sold
    'Notes'                // 29 - Internal admin notes
  ];

  setupSheet(ss, 'Inventory', inventoryHeaders, '#c9a962');

  /* =====================================================
     ORDERS (CUSTOMER ORDERS) - 26 COLUMNS
     Replaces hotel Bookings
  ===================================================== */

  var ordersHeaders = [
    'Reference',           // 0  - Auto-generated order reference
    'Timestamp',           // 1  - Order submission time
    'Status',              // 2  - new | processing | shipped | delivered | cancelled | refunded
    'Customer_Name',       // 3  - Full name
    'Customer_Surname',    // 4  - Surname (for ref generation)
    'Email',               // 5  - Email address
    'Phone',               // 6  - Phone number
    'Country',             // 7  - Country
    'Shipping_Address',    // 8  - Full shipping address
    'City',                // 9  - City
    'Postal_Code',         // 10 - Postal/ZIP code
    'Items_JSON',          // 11 - JSON array of cart items [{id,name,qty,price,original}]
    'Item_Count',          // 12 - Total number of items
    'Subtotal',            // 13 - Sum of item prices
    'Shipping_Fee',        // 14 - Shipping cost
    'Discount_Amount',     // 15 - Any discount applied
    'Total_Savings',       // 16 - Total savings from original prices
    'Total_Amount',        // 17 - Final amount (subtotal + shipping - discount)
    'Payment_Method',      // 18 - paypal | eft | pending
    'Payment_Status',      // 19 - pending | paid | partial | refunded
    'PayPal_Transaction',  // 20 - PayPal transaction ID (if applicable)
    'EFT_Reference',       // 21 - EFT banking reference
    'Confirmation_Sent',   // 22 - true | false
    'Tracking_Number',     // 23 - Shipping tracking number
    'Notes',               // 24 - Admin notes
    'Last_Updated'         // 25 - Last update timestamp
  ];

  setupSheet(ss, 'Orders', ordersHeaders, '#18181b');

  /* =====================================================
     TRANSACTIONS (FINANCIAL LEDGER) - 8 COLUMNS
     Bank-style double-entry for all money movement
  ===================================================== */

  var transactionsHeaders = [
    'Transaction_ID',  // 0 - UUID
    'Reference',       // 1 - FK -> Orders.Reference
    'Date',            // 2 - Transaction date
    'Description',     // 3 - What it's for
    'Amount',          // 4 - Positive=charge, Negative=payment
    'Type',            // 5 - CHARGE | PAYMENT | REFUND | ADJUSTMENT
    'Created_By',      // 6 - Admin or SYSTEM
    'Created_At'       // 7 - Timestamp
  ];

  setupSheet(ss, 'Transactions', transactionsHeaders, '#059669');

  /* =====================================================
     COMMENTS (EVENT LOG) - 5 COLUMNS
  ===================================================== */

  var commentsHeaders = [
    'Comment_ID',    // UUID
    'Reference',     // FK -> Orders.Reference
    'Author',        // Admin | SYSTEM | BILLING
    'Text',
    'Timestamp'
  ];

  setupSheet(ss, 'Comments', commentsHeaders, '#a88a4a');

  /* =====================================================
     DELETED (TOMBSTONE PATTERN) - 5 COLUMNS
  ===================================================== */

  var deletedHeaders = [
    'Reference',     // 0
    'Deleted_At',    // 1
    'Reason',        // 2
    'Deleted_By',    // 3
    'Snapshot'       // 4 (JSON string of full order row)
  ];

  setupSheet(ss, 'Deleted', deletedHeaders, '#dc3545');

  /* =====================================================
     SYSTEM LOGS (AUDIT TRAIL) - 5 COLUMNS
  ===================================================== */

  var logsHeaders = [
    'Log_ID',        // UUID
    'Timestamp',     // When
    'Action',        // What happened
    'User',          // Who did it
    'Details'        // JSON or text details
  ];

  setupSheet(ss, 'SystemLogs', logsHeaders, '#6366f1');

  /* =====================================================
     VERIFY DRIVE FOLDER
  ===================================================== */

  try {
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    Logger.log('Drive folder OK: ' + folder.getName());
  } catch (e) {
    Logger.log('WARNING: Drive folder not accessible. Create one and update DRIVE_FOLDER_ID.');
  }

  Logger.log('');
  Logger.log('========================================');
  Logger.log('SETUP COMPLETE');
  Logger.log('========================================');
  Logger.log('');
  Logger.log('ARCHITECTURE:');
  Logger.log('  Inventory (catalog) -> Storefront API');
  Logger.log('  Orders (1) ----< Transactions (N) [BILLING]');
  Logger.log('        |');
  Logger.log('        +----< Comments (N) [EVENT LOG]');
  Logger.log('        |');
  Logger.log('        +----< Deleted (1, tombstone)');
  Logger.log('');
  Logger.log('  SystemLogs - Full audit trail');
  Logger.log('');

  return 'Setup complete. 6 sheets created and synced.';
}

/* =====================================================
   HELPERS
===================================================== */

function setupSheet(ss, name, headers, color) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    Logger.log('Created sheet: ' + name);
  }

  // Set headers
  sheet.getRange(1, 1, 1, sheet.getLastColumn() || headers.length).clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground(color)
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);

  Logger.log(name + ': ' + headers.length + ' columns verified');
  return sheet;
}

function testSetup() {
  var SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ['Inventory', 'Orders', 'Transactions', 'Comments', 'Deleted', 'SystemLogs'];
  for (var i = 0; i < sheets.length; i++) {
    var sh = ss.getSheetByName(sheets[i]);
    Logger.log(sheets[i] + ': ' + (sh ? 'OK (' + sh.getLastColumn() + ' cols)' : 'MISSING'));
  }
}
