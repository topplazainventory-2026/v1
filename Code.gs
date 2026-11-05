function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Stock Management System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/* =========================================================
   GET ITEM CATALOG
   ITEMS SHEET

   A = CODE
   B = NAME
   C = MAXIMUM STOCKS
   ========================================================= */

function getItemCatalog() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Items');

  if (!sheet) {
    throw new Error('Sheet "Items" was not found.');
  }

  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  var data =
    sheet
      .getRange(2, 1, lastRow - 1, 3)
      .getValues();

  var items = [];

  data.forEach(function(row) {

    var code =
      String(row[0] || '').trim();

    var name =
      String(row[1] || '').trim();

    var maxStock =
      Number(row[2]);

    if (
      !maxStock ||
      maxStock <= 0
    ) {
      maxStock = 80;
    }

    if (code || name) {

      items.push({

        code: code,

        name: name,

        maxStock: maxStock

      });

    }

  });

  return items;

}


/* =========================================================
   GET ITEM LIST
   ========================================================= */

function getItemList() {

  // Load the catalog and transaction totals once for this request.
  // Each item returned to the browser includes its current remaining stock.
  var items = getItemCatalog();
  var totals = getAllTransactionTotals();

  return items.map(function(item) {
    var key = String(item.name || '').trim().toLowerCase();
    var itemTotals = totals[key] || { totalIn: 0, totalOut: 0 };
    var remaining =
      (Number(itemTotals.totalIn) || 0) -
      (Number(itemTotals.totalOut) || 0);

    return {
      code: String(item.code || ''),
      name: String(item.name || ''),
      maxStock: Number(item.maxStock) || 80,
      stocksLeft: remaining
    };
  });

}


/* =========================================================
   FIND ITEM
   SEARCH BY CODE OR NAME
   ========================================================= */

function findItem(searchValue) {

  var search = String(searchValue || '').trim().replace(/\s+/g, ' ').toLowerCase();
  if (!search) {
    return { found: false, code: '', name: '', stocksLeft: 0, maxStock: 80 };
  }

  var items = getItemList();
  for (var i = 0; i < items.length; i++) {
    var code = String(items[i].code || '').trim().replace(/\s+/g, ' ').toLowerCase();
    var name = String(items[i].name || '').trim().replace(/\s+/g, ' ').toLowerCase();
    if (code === search || name === search) {
      return {
        found: true,
        code: items[i].code,
        name: items[i].name,
        stocksLeft: Number(items[i].stocksLeft) || 0,
        maxStock: Number(items[i].maxStock) || 80
      };
    }
  }

  return { found: false, code: '', name: '', stocksLeft: 0, maxStock: 80 };

}


/* =========================================================
   BUILD ITEM LOOKUP
   BOTH CODE AND NAME POINT TO SAME ITEM
   ========================================================= */

function getItemLookup() {

  var items =
    getItemCatalog();

  var lookup = {};

  items.forEach(function(item) {

    var codeKey =
      String(item.code || '')
        .trim()
        .toLowerCase();

    var nameKey =
      String(item.name || '')
        .trim()
        .toLowerCase();

    if (codeKey) {

      lookup[codeKey] = {

        code: item.code,

        name: item.name,

        maxStock: item.maxStock

      };

    }

    if (nameKey) {

      lookup[nameKey] = {

        code: item.code,

        name: item.name,

        maxStock: item.maxStock

      };

    }

  });

  return lookup;

}


/* =========================================================
   RESOLVE ITEM NAME
   CODE -> NAME
   NAME -> NAME
   ========================================================= */

function resolveItemName(item) {

  item =
    String(item || '').trim();

  if (!item) {
    return '';
  }

  var lookup =
    getItemLookup();

  var key =
    item.toLowerCase();

  if (lookup[key]) {

    return lookup[key].name;

  }

  return item;

}


/* =========================================================
   GET ALL TRANSACTION TOTALS
   =========================================================

   Sheet1:

   A = DATE
   B = TIME
   C = ITEM
   D = STOCKS
   E = STATUS

   This function reads EVERY transaction row.

   TOTAL IN  = all rows with STATUS = IN
   TOTAL OUT = all rows with STATUS = OUT

   It does NOT use only the last transaction.
   ========================================================= */

function getAllTransactionTotals() {

  var ss =
    SpreadsheetApp.getActiveSpreadsheet();

  var sheet =
    ss.getSheetByName('Entry');

  if (!sheet) {

    throw new Error(
      'Sheet "Entry" was not found.'
    );

  }

  var totals = {};

  var lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {

    return totals;

  }

  /*
   * READ ALL TRANSACTION ROWS
   *
   * C = ITEM
   * D = STOCKS
   * E = STATUS
   */
  var data =
    sheet
      .getRange(
        2,
        3,
        lastRow - 1,
        3
      )
      .getValues();

  var lookup =
    getItemLookup();


  data.forEach(function(row) {

    var transactionItem =
      String(row[0] || '').trim();

    var quantity =
      Number(row[1]) || 0;

    var status =
      String(row[2] || '')
        .trim()
        .toUpperCase();


    if (!transactionItem) {
      return;
    }

    if (quantity <= 0) {
      return;
    }

    if (
      status !== 'IN' &&
      status !== 'OUT'
    ) {
      return;
    }


    /*
     * IMPORTANT:
     *
     * If Sheet1 contains the CODE,
     * convert it to the catalog NAME.
     *
     * If Sheet1 contains the NAME,
     * keep the catalog NAME.
     *
     * This makes old and new transactions
     * count toward the SAME item.
     */
    var transactionKey =
      transactionItem.toLowerCase();

    var catalogItem =
      lookup[transactionKey];


    var canonicalName;

    if (catalogItem) {

      canonicalName =
        String(catalogItem.name || '')
          .trim();

    } else {

      canonicalName =
        transactionItem;

    }


    if (!canonicalName) {
      return;
    }


    var key =
      canonicalName.toLowerCase();


    if (!totals[key]) {

      totals[key] = {

        name: canonicalName,

        totalIn: 0,

        totalOut: 0

      };

    }


    /*
     * SUM EVERY IN
     */
    if (status === 'IN') {

      totals[key].totalIn +=
        quantity;

    }


    /*
     * SUM EVERY OUT
     */
    if (status === 'OUT') {

      totals[key].totalOut +=
        quantity;

    }

  });


  return totals;

}


/* =========================================================
   GET STOCKS LEFT
   =========================================================

   VERY IMPORTANT:

   STOCKS LEFT =
   TOTAL ALL IN - TOTAL ALL OUT

   This checks ALL transaction rows in Sheet1.

   Example:

   IN 50
   IN 20
   OUT 10
   IN 30
   OUT 5

   Total IN  = 100
   Total OUT = 15

   STOCKS LEFT = 85

   It does NOT depend on the last entry
   or page refresh.
   ========================================================= */

function getStocksLeft(item) {

  var itemName =
    resolveItemName(item);

  if (!itemName) {
    return 0;
  }


  var allTotals =
    getAllTransactionTotals();


  var key =
    itemName
      .trim()
      .toLowerCase();


  var itemTotals =
    allTotals[key];


  if (!itemTotals) {

    return 0;

  }


  var totalIn =
    Number(itemTotals.totalIn) || 0;


  var totalOut =
    Number(itemTotals.totalOut) || 0;


  /*
   * EXACT CALCULATION REQUESTED:
   *
   * TOTAL IN - TOTAL OUT
   */
  var stocksLeft =
    totalIn - totalOut;


  return stocksLeft;

}


/* =========================================================
   GET STOCK SUMMARY
   ========================================================= */

function getStockSummary() {

  var ss =
    SpreadsheetApp.getActiveSpreadsheet();

  var itemSheet =
    ss.getSheetByName('Items');

  if (!itemSheet) {

    throw new Error(
      'Sheet "Items" was not found.'
    );

  }


  var itemLastRow =
    itemSheet.getLastRow();


  if (itemLastRow < 2) {

    return {

      noStocks: 0,

      lowStocks: 0,

      onStocks: 0,

      totalItems: 0,

      totalStocks: 0,

      items: []

    };

  }


  /*
   * GET ALL ITEMS
   */
  var itemData =
    itemSheet
      .getRange(
        2,
        1,
        itemLastRow - 1,
        3
      )
      .getValues();


  /*
   * GET ALL TRANSACTION TOTALS
   *
   * This is the important part.
   * It reads ALL Sheet1 entries.
   */
  var transactionTotals =
    getAllTransactionTotals();


  var noStocks = 0;

  var lowStocks = 0;

  var onStocks = 0;

  var totalStocks = 0;

  var items = [];


  itemData.forEach(function(row) {

    var code =
      String(row[0] || '').trim();

    var name =
      String(row[1] || '').trim();

    var maxStock =
      Number(row[2]);


    if (!name) {
      return;
    }


    if (
      !maxStock ||
      maxStock <= 0
    ) {

      maxStock = 80;

    }


    /*
     * Find ALL transactions belonging
     * to this catalog item.
     */
    var key =
      name.toLowerCase();


    var itemTotals =
      transactionTotals[key];


    var totalIn = 0;

    var totalOut = 0;


    if (itemTotals) {

      totalIn =
        Number(
          itemTotals.totalIn
        ) || 0;

      totalOut =
        Number(
          itemTotals.totalOut
        ) || 0;

    }


    /*
     * EXACT STOCK CALCULATION
     *
     * ALL IN - ALL OUT
     */
    var stocks =
      totalIn - totalOut;


    /*
     * Do not show negative stocks
     * in the monitoring dashboard.
     *
     * The actual calculation above
     * is still ALL IN - ALL OUT.
     */
    var displayStocks =
      stocks < 0
        ? 0
        : stocks;


    /*
     * STOCK PERCENTAGE
     */
    var percentage = 0;


    if (maxStock > 0) {

      percentage =
        (
          displayStocks /
          maxStock
        ) * 100;

    }


    var status = '';


    /*
     * 0% = NO STOCKS
     */
    if (displayStocks <= 0) {

      status =
        'NO STOCKS';

      noStocks++;

    }

    /*
     * 1% TO 20% = LOW STOCKS
     */
    else if (percentage <= 20) {

      status =
        'LOW STOCKS';

      lowStocks++;

    }

    /*
     * 21% OR HIGHER = ON STOCKS
     */
    else {

      status =
        'ON STOCKS';

      onStocks++;

    }


    totalStocks +=
      displayStocks;


    items.push({

      code: code,

      name: name,

      stocks: displayStocks,

      maxStock: maxStock,

      percentage: percentage,

      status: status,

      /*
       * Included for accurate
       * checking/debugging if needed.
       */
      totalIn: totalIn,

      totalOut: totalOut

    });

  });


  return {

    noStocks: noStocks,

    lowStocks: lowStocks,

    onStocks: onStocks,

    totalItems: items.length,

    totalStocks: totalStocks,

    items: items

  };

}


/* =========================================================
   SAVE BATCH STOCK ENTRY
   Entry columns:
   A = DATE
   B = TIME
   C = ITEM
   D = STOCKS
   E = STATUS
   F = REQUESTER
   G = DEPARTMENT
   ========================================================= */

function saveEntries(entries, status, requester, department) {
  status = String(status || '').trim().toUpperCase();
  requester = String(requester || '').trim();
  department = String(department || '').trim();

  if (status !== 'IN' && status !== 'OUT') {
    throw new Error('Invalid transaction type. Use IN or OUT.');
  }
  if (!requester) {
    throw new Error('Please enter the requester name before submitting.');
  }
  if (!department) {
    throw new Error('Please enter the department before submitting.');
  }
  if (!Array.isArray(entries) || !entries.length) {
    throw new Error('Please add at least one item.');
  }
  if (entries.length > 250) {
    throw new Error('A maximum of 250 rows can be submitted at once.');
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var catalog = getItemCatalog();
    var lookup = {};

    catalog.forEach(function(it) {
      if (it.code) lookup[String(it.code).trim().toLowerCase()] = it;
      if (it.name) lookup[String(it.name).trim().toLowerCase()] = it;
    });

    var grouped = {};
    var cleanEntries = [];

    // Ignore completely blank rows. Validate only rows where the user
    // entered either an item or a quantity.
    entries.forEach(function(entry, index) {
      var raw = String(entry && entry.item != null ? entry.item : '').trim();
      var quantityRaw = entry && entry.quantity != null
        ? String(entry.quantity).trim()
        : '';

      // A completely unused row is intentionally skipped.
      if (!raw && !quantityRaw) {
        return;
      }

      // A partially completed row is not silently ignored.
      if (!raw) {
        throw new Error('Please select an item on line ' + (index + 1) + '.');
      }
      if (!quantityRaw) {
        throw new Error('Please enter a quantity on line ' + (index + 1) + '.');
      }

      var qty = Number(quantityRaw);
      if (!Number.isFinite(qty) || qty <= 0) {
        throw new Error('Enter a valid positive quantity on line ' + (index + 1) + '.');
      }

      var item = lookup[raw.toLowerCase()];
      if (!item || !item.name) {
        throw new Error('Item not found in Items sheet: ' + raw);
      }

      var key = item.name.trim().toLowerCase();
      var expirationDate = String(entry.expirationDate || '').trim();
      var reason = String(entry.reason || '').trim();
      if (status === 'OUT' && !reason) throw new Error('Reason is required for Stock OUT.');
      if (expirationDate && !/^\d{4}-\d{2}-\d{2}$/.test(expirationDate)) {
        throw new Error('Invalid expiration date on line ' + (index + 1) + '.');
      }
      cleanEntries.push({
        name: item.name.trim(),
        quantity: qty,
        key: key,
        expirationDate: expirationDate,
        reason: reason
      });

      if (!grouped[key]) {
        grouped[key] = {
          name: item.name.trim(),
          quantity: 0
        };
      }
      grouped[key].quantity += qty;
    });

    // Do not write an empty submission to the Entry sheet.
    if (!cleanEntries.length) {
      throw new Error('Add at least one item and quantity before submitting.');
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Entry');
    if (!sheet) {
      throw new Error('Sheet "Entry" was not found.');
    }

    // Check combined requested quantities against current stock while locked.
    if (status === 'OUT') {
      var totals = getAllTransactionTotals();

      Object.keys(grouped).forEach(function(key) {
        var rec = totals[key];
        var available = rec
          ? Number(rec.totalIn || 0) - Number(rec.totalOut || 0)
          : 0;

        if (grouped[key].quantity > available) {
          throw new Error(
            'Not enough stocks for "' + grouped[key].name +
            '". Available: ' + available +
            ', requested: ' + grouped[key].quantity + '.'
          );
        }
      });
    }

    var now = new Date();
    var tz = Session.getScriptTimeZone();
    var time = Utilities.formatDate(now, tz, 'HH:mm:ss');

    // A:G remain unchanged. Expiration and reason are appended in H:I.
    var rows = cleanEntries.map(function(e) {
      var expiry = '';
      if (e.expirationDate) {
        var p = e.expirationDate.split('-');
        expiry = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
      }
      return [now, time, e.name, e.quantity, status, requester, department,
        expiry, e.reason];
    });

    var firstRow = sheet.getLastRow() + 1;
    sheet.getRange(firstRow, 1, rows.length, 9).setValues(rows);
    if (rows.length) sheet.getRange(firstRow, 8, rows.length, 1).setNumberFormat('yyyy-mm-dd');
    SpreadsheetApp.flush();

    return {
      success: true,
      message: status === 'IN'
        ? 'Stock In saved successfully.'
        : 'Stock Out saved successfully.',
      savedCount: rows.length,
      requester: requester,
      department: department
    };

  } finally {
    lock.releaseLock();
  }
}

/* =========================================================
   EXPIRATION DATE MONITORING
   Uses the Entry sheet and the Items catalog.
   Displays only Stock IN quantities; does not subtract OUT.
   Rows with the same item and expiration date are combined.
   ========================================================= */
function getExpirationMonitoring(startDate, endDate, itemSearch) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Entry');
  if (!sheet || sheet.getLastRow() < 2) return [];

  var from = String(startDate || '').trim();
  var to = String(endDate || '').trim();
  var query = String(itemSearch || '').trim().toLowerCase();
  var validItems = {};
  getItemCatalog().forEach(function(it) {
    if (it.name) validItems[String(it.name).trim().toLowerCase()] = String(it.name).trim();
  });

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(9, sheet.getLastColumn())).getValues();

  // Build matching IN quantities. If a selected date range has no matches,
  // automatically fall back to all expiration dates for the searched item(s).
  // Stock OUT is never subtracted from these quantities.
  function collectGroups(applyDateRange) {
    var grouped = {};
    data.forEach(function(r) {
      var item = String(r[2] || '').trim();
      var qty = Number(r[3]) || 0;
      var status = String(r[4] || '').trim().toUpperCase();
      var exp = r[7];
      if (status !== 'IN' || !item || qty <= 0 || !(exp instanceof Date) || isNaN(exp.getTime())) return;
      var canonical = validItems[item.toLowerCase()];
      if (!canonical || (query && canonical.toLowerCase().indexOf(query) === -1)) return;
      var expStr = Utilities.formatDate(exp, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (applyDateRange && from && expStr < from) return;
      if (applyDateRange && to && expStr > to) return;
      var key = canonical.toLowerCase() + '|' + expStr;
      if (!grouped[key]) grouped[key] = {expirationDate: expStr, item: canonical, quantity: 0};
      grouped[key].quantity += qty;
    });
    return grouped;
  }

  var grouped = collectGroups(true);
  // No results for the selected dates? Show the available list in earliest-to-latest order.
  if (!Object.keys(grouped).length && (from || to)) {
    grouped = collectGroups(false);
  }

  return Object.keys(grouped).map(function(k) { return grouped[k]; }).sort(function(a, b) {
    return a.expirationDate.localeCompare(b.expirationDate) || a.item.localeCompare(b.item);
  });
}

/* Backward-compatible single-entry endpoint.
   Department is now required. */
function saveEntry(item, stocks, status, requester, department) {
  var result = saveEntries(
    [{item: item, quantity: stocks}],
    status,
    requester,
    department
  );

  var matched = findItem(item);

  return {
    success: true,
    message: 'Entry saved successfully.',
    item: matched.name || String(item || ''),
    stocksLeft: matched.found ? getStocksLeft(matched.name) : 0,
    savedCount: result.savedCount
  };
}


/* =========================================================
   GET STOCKS REPORT
   ========================================================= */

function getReportData(
  startDate,
  endDate
) {

  var ss =
    SpreadsheetApp.getActiveSpreadsheet();

  var sheet =
    ss.getSheetByName('Entry');

  var itemSheet =
    ss.getSheetByName('Items');


  if (!sheet) {

    throw new Error(
      'Sheet "Entry" was not found.'
    );

  }


  if (!itemSheet) {

    throw new Error(
      'Sheet "Items" was not found.'
    );

  }


  if (
    !startDate ||
    !endDate
  ) {

    return [];

  }


  var startParts =
    String(startDate).split('-');

  var endParts =
    String(endDate).split('-');


  if (
    startParts.length !== 3 ||
    endParts.length !== 3
  ) {

    throw new Error(
      'Invalid report dates.'
    );

  }


  var start =
    new Date(

      Number(startParts[0]),

      Number(startParts[1]) - 1,

      Number(startParts[2]),

      0,
      0,
      0,
      0

    );


  var end =
    new Date(

      Number(endParts[0]),

      Number(endParts[1]) - 1,

      Number(endParts[2]),

      23,
      59,
      59,
      999

    );


  if (start > end) {

    throw new Error(
      'Date From cannot be later than Date To.'
    );

  }


  /*
   * CREATE EVERY DATE IN RANGE
   */
  var dates = [];


  var currentDate =
    new Date(

      Number(endParts[0]),

      Number(endParts[1]) - 1,

      Number(endParts[2]),

      0,
      0,
      0,
      0

    );


  var startOnly =
    new Date(

      Number(startParts[0]),

      Number(startParts[1]) - 1,

      Number(startParts[2]),

      0,
      0,
      0,
      0

    );


  while (
    currentDate >= startOnly
  ) {

    dates.push({

      key:
        Utilities.formatDate(

          currentDate,

          Session.getScriptTimeZone(),

          'yyyy-MM-dd'

        ),

      label:
        Utilities.formatDate(

          currentDate,

          Session.getScriptTimeZone(),

          'MM/dd/yyyy'

        )

    });


    currentDate.setDate(
      currentDate.getDate() - 1
    );

  }


  /*
   * GET ITEM CATALOG
   */
  var itemLastRow =
    itemSheet.getLastRow();


  var itemMap = {};

  var itemList = [];


  if (itemLastRow >= 2) {

    var itemData =
      itemSheet
        .getRange(
          2,
          1,
          itemLastRow - 1,
          3
        )
        .getValues();


    itemData.forEach(function(row) {

      var code =
        String(row[0] || '').trim();

      var name =
        String(row[1] || '').trim();


      if (!name) {
        return;
      }


      var key =
        name.toLowerCase();


      if (!itemMap[key]) {

        itemMap[key] = {

          code: code,

          name: name,

          beginningStock: 0,

          daily: {}

        };


        itemList.push(
          itemMap[key]
        );

      }

    });

  }


  /*
   * GET ALL TRANSACTIONS
   */
  var lastRow =
    sheet.getLastRow();


  if (lastRow >= 2) {

    var data =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          5
        )
        .getValues();


    var lookup =
      getItemLookup();


    data.forEach(function(row) {

      var rawDate =
        row[0];


      var transactionDate;


      if (
        rawDate instanceof Date &&
        !isNaN(rawDate.getTime())
      ) {

        transactionDate =
          new Date(
            rawDate.getTime()
          );

      } else {

        transactionDate =
          new Date(rawDate);

      }


      if (
        !transactionDate ||
        isNaN(
          transactionDate.getTime()
        )
      ) {

        return;

      }


      var transactionItem =
        String(row[2] || '').trim();


      var quantity =
        Number(row[3]) || 0;


      var status =
        String(row[4] || '')
          .trim()
          .toUpperCase();


      if (!transactionItem) {
        return;
      }


      if (
        status !== 'IN' &&
        status !== 'OUT'
      ) {

        return;

      }


      /*
       * Resolve old CODE entries and
       * NAME entries to the same item.
       */
      var transactionKey =
        transactionItem.toLowerCase();


      var catalogItem =
        lookup[transactionKey];


      var canonicalName;


      if (catalogItem) {

        canonicalName =
          catalogItem.name;

      } else {

        canonicalName =
          transactionItem;

      }


      var key =
        canonicalName
          .trim()
          .toLowerCase();


      /*
       * If transaction item was not in catalog,
       * still include it in report.
       */
      if (!itemMap[key]) {

        itemMap[key] = {

          code:
            catalogItem
              ? catalogItem.code
              : '',

          name:
            canonicalName,

          beginningStock: 0,

          daily: {}

        };


        itemList.push(
          itemMap[key]
        );

      }


      /*
       * BEFORE REPORT RANGE
       *
       * Used as BEGINNING STOCK.
       */
      if (
        transactionDate < start
      ) {

        if (status === 'IN') {

          itemMap[key]
            .beginningStock +=
              quantity;

        }


        if (status === 'OUT') {

          itemMap[key]
            .beginningStock -=
              quantity;

        }


        return;

      }


      /*
       * INSIDE REPORT RANGE
       */
      if (
        transactionDate <= end
      ) {

        var dateKey =
          Utilities.formatDate(

            transactionDate,

            Session.getScriptTimeZone(),

            'yyyy-MM-dd'

          );


        if (
          !itemMap[key].daily[dateKey]
        ) {

          itemMap[key].daily[dateKey] = {

            in: 0,

            out: 0

          };

        }


        if (status === 'IN') {

          itemMap[key]
            .daily[dateKey]
            .in += quantity;

        }


        if (status === 'OUT') {

          itemMap[key]
            .daily[dateKey]
            .out += quantity;

        }

      }

    });

  }


  /*
   * BUILD REPORT ROWS
   */
  var rows = [];


  itemList.forEach(function(item) {

    var total =
      Number(
        item.beginningStock
      ) || 0;


    dates.forEach(function(date) {

      var daily =
        item.daily[date.key] || {

          in: 0,

          out: 0

        };


      total +=
        Number(daily.in) || 0;


      total -=
        Number(daily.out) || 0;

    });


    rows.push({

      code:
        item.code,

      name:
        item.name,

      beginningStock:
        item.beginningStock,

      daily:
        item.daily,

      total:
        total

    });

  });


  /*
   * REPORT TOTALS
   */
  var totals = {

    beginningStock: 0,

    daily: {},

    total: 0

  };


  rows.forEach(function(row) {

    totals.beginningStock +=
      Number(
        row.beginningStock
      ) || 0;


    totals.total +=
      Number(
        row.total
      ) || 0;

  });


  dates.forEach(function(date) {

    totals.daily[date.key] = {

      in: 0,

      out: 0

    };


    rows.forEach(function(row) {

      var daily =
        row.daily[date.key];


      if (!daily) {
        return;
      }


      totals.daily[date.key].in +=
        Number(daily.in) || 0;


      totals.daily[date.key].out +=
        Number(daily.out) || 0;

    });

  });


  return {

    dates:
      dates,

    rows:
      rows,

    totals:
      totals

  };

}


/* =========================================================
   PARSE REPORT DATE/TIME
   ========================================================= */

function parseReportDateTime(
  dateText,
  timeText
) {

  var dateParts =
    String(dateText).split('/');


  var year =
    Number(dateParts[2]);


  var month =
    Number(dateParts[0]) - 1;


  var day =
    Number(dateParts[1]);


  var hours = 0;

  var minutes = 0;

  var seconds = 0;


  if (timeText) {

    var timeParts =
      String(timeText).split(':');


    hours =
      Number(timeParts[0]) || 0;


    minutes =
      Number(timeParts[1]) || 0;


    seconds =
      Number(timeParts[2]) || 0;

  }


  return new Date(

    year,

    month,

    day,

    hours,

    minutes,

    seconds

  ).getTime();

}


/* =========================================================
   EMPTY REPORT
   ========================================================= */

function emptyReport() {

  return [];

}