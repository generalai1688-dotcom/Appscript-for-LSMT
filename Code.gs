// =========================================
// CONFIG
// =========================================

const SHEET_ORDER = "Order";
const SHEET_CUSTOMER = "Customer_Summary";

const CACHE_TIME = 30;


// =========================================
// WEB APP
// =========================================

function doGet() {

  return HtmlService
    .createTemplateFromFile("Dashboard")
    .evaluate()
    .setTitle("សាមគ្គីទាន់ចិត្ត បោកអ៊ុត");

}


// =========================================
// INCLUDE HTML
// =========================================

function include(filename) {

  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();

}


// =========================================
// CLEAN TEXT
// =========================================

function cleanText(value){

  return String(value || "")
    .toLowerCase()
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim();

}


// =========================================
// FORMAT CUSTOMER AGE
// =========================================

function formatCustomerAge(firstDate){

  if(!firstDate) return "0 Months";

  const now = new Date();

  const months =

    (
      (now.getFullYear() - firstDate.getFullYear()) * 12
    )

    +

    (
      now.getMonth() - firstDate.getMonth()
    );

  const years =
    Math.floor(months / 12);

  const remainMonths =
    months % 12;

  if(years <= 0){

    return remainMonths + " Months";

  }

  return years + " Years " + remainMonths + " Months";

}


// =========================================
// GET MANAGEMENT DATA
// =========================================

function getManagementData(
  startDate,
  endDate,
  statusFilter
){

  return getDashboardData(
    startDate,
    endDate,
    statusFilter
  );

}


// =========================================
// GET DASHBOARD DATA
// =========================================

function getDashboardData(
  startDate,
  endDate,
  statusFilter
){

  // =====================================
  // DEFAULT FILTER
  // =====================================

  statusFilter =
    statusFilter || "All";


  // =====================================
  // DEFAULT DATE = CURRENT MONTH
  // =====================================

  const today = new Date();

  const firstDay = new Date(
    today.getFullYear(),
    today.getMonth(),
    1
  );

  if(!startDate){

    startDate = Utilities.formatDate(
      firstDay,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );

  }

  if(!endDate){

    endDate = Utilities.formatDate(
      today,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );

  }

  // =====================================
  // CACHE
  // =====================================

  const cache =
    CacheService.getScriptCache();

  const cacheKey =
    startDate + "_" +
    endDate + "_" +
    statusFilter;

  const cached =
    cache.get(cacheKey);

  if(cached){

    return JSON.parse(cached);

  }

  const ss =
    SpreadsheetApp.getActive();

  // =====================================
  // ORDER SHEET
  // =====================================

  const orderSheet =
    ss.getSheetByName(SHEET_ORDER);

  const orderData =
    orderSheet.getDataRange().getValues();

  const orderHeaders =
    orderData[0].map(h => cleanText(h));

  const orders =
    orderData.slice(1);

  // =====================================
  // CUSTOMER SUMMARY
  // =====================================

  const customerSheet =
    ss.getSheetByName(SHEET_CUSTOMER);

  const customerData =
    customerSheet.getDataRange().getValues();

  const customerHeaders =
    customerData[0].map(h => cleanText(h));

  const customerRows =
    customerData.slice(1);

  const totalShops =
  customerRows.length;

  const ccol = name =>
    customerHeaders.indexOf(cleanText(name));

  const customerMapInfo = {};

  customerRows.forEach(r => {

    const shop =
      cleanText(r[ccol("shop name")]);

    if(!shop) return;

    customerMapInfo[shop] = {

      area:
        r[ccol("area")] || "Unknown",

      address:
        r[ccol("address")] || "",

      type:
        r[ccol("customer type")] || "General"

    };

  });

  // =====================================
  // ORDER COLUMN
  // =====================================

  const col = name =>
    orderHeaders.indexOf(cleanText(name));

  const orderDateCol = col("date");
  const customerCol = col("display shop");
  const driverCol = col("driver name");
  const qtyCol = col("total qty");
  const stainQtyCol = col("total stain qty");
  const revenueCol = col("grand total");
  const invoiceCol = col("invoice id");
  const invoiceDateCol = col("invoice date");
  const driverOutCol = col("driver out");
  const paidCol = col("amount paid");
  const paymentDateCol = col("payment date");
  const statusCol = col("payment status");
  const remainCol = col("remaining bal");
  const adjustmentCol = col("adjustment amount");

  // =====================================
  // KPI
  // =====================================

  let forecastRevenue = 0;
  let outstandingDebt = 0;
  let totalOrders = 0;
  let totalQty = 0;
  let todayRevenue = 0;
  let todayQty = 0;
  let todayOrders = 0;
  let paidAmount = 0;
  let profit = 0;
  let totalAdjustment = 0;

  let holdingAlert = 0;
  let debtAlert = 0;

  let totalLaundryDays = 0;
  let completedLaundry = 0;

  let customerMap = {};
  let revenueByDate = {};
  let driverMap = {};
  let areaMap = {};

  let latestOrders = [];
  let unpaidInvoices = [];

  const uniqueDays =
    new Set();

  let start = null;
let end = null;

if(startDate){

  start = new Date(startDate + "T00:00:00");

}

if(endDate){

  end = new Date(endDate + "T23:59:59");

}

// =====================================
// DATE ONLY FILTER
// =====================================

const startOnlyDate = start
  ? new Date(
      start.getFullYear(),
      start.getMonth(),
      start.getDate()
    )
  : null;

const endOnlyDate = end
  ? new Date(
      end.getFullYear(),
      end.getMonth(),
      end.getDate()
    )
  : null;

  // =====================================
  // LOOP ORDERS
  // =====================================

  orders.forEach(r => {

    if(!r[0]) return;

    const orderDate =
      new Date(r[orderDateCol]);

    const todayStr = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );

  const orderStr = Utilities.formatDate(
    orderDate,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );


    if(isNaN(orderDate)) return;

    const inDateRange =

  (!start || orderDate >= start)

  &&

  (!end || orderDate <= end);

    const customerRaw =
      r[customerCol] || "";

    const customer =
      cleanText(customerRaw);

    const customerInfo =
      customerMapInfo[customer] || {};

    const area =
      customerInfo.area || "Unknown";

    const customerType =
      customerInfo.type || "General";

    const driver =
      String(r[driverCol] || "").trim();

    const driverOut =
      String(r[driverOutCol] || "").trim();

    const qty =
      Number(r[qtyCol]) || 0;

    const stainQty =
      Number(r[stainQtyCol]) || 0;

    const revenue =
      Number(r[revenueCol]) || 0;

    const paid =
      Number(r[paidCol]) || 0;
    
    let paymentDate = null;

if(r[paymentDateCol]){

  paymentDate = new Date(r[paymentDateCol]);

  if(isNaN(paymentDate)){

    paymentDate = null;

  }

}

    const adjustment =
      Number(r[adjustmentCol]) || 0;

    const status =
      String(r[statusCol] || "").trim();

    const remaining =
      Number(r[remainCol]) || 0;

    const invoiceID =
      r[invoiceCol];

    const invoiceDate =
      r[invoiceDateCol]
        ? new Date(r[invoiceDateCol])
        : null;

    // =====================================
    // BUSINESS LOGIC
    // =====================================

    const actualProfit =
      paid + remaining;

// =====================================
// KPI (MONTH TO DATE)
// =====================================

if(inDateRange){

  // OPERATION KPI

  forecastRevenue += revenue;

  profit += actualProfit;

  totalQty += qty;

  totalOrders++;

  if(orderStr == todayStr){

    todayRevenue += revenue;

    todayQty += qty;

    todayOrders++;

  }

}

// =====================================
// FINANCIAL KPI (ALL TIME)
// =====================================

// =====================================
// PAID AMOUNT BY PAYMENT DATE
// =====================================

if(paymentDate && paid > 0){

  const paymentOnlyDate = new Date(

    paymentDate.getFullYear(),
    paymentDate.getMonth(),
    paymentDate.getDate()
  );



  const paymentInRange =

    (!startOnlyDate || paymentOnlyDate >= startOnlyDate)

    &&

    (!endOnlyDate || paymentOnlyDate <= endOnlyDate);

  if(paymentInRange){

    paidAmount += paid;

  }

}

// =====================================
// ADJUSTMENT BY PAYMENT DATE
// =====================================

if(paymentDate && adjustment > 0){

  const adjustmentOnlyDate = new Date(

    paymentDate.getFullYear(),
    paymentDate.getMonth(),
    paymentDate.getDate()

  );

  const adjustmentInRange =

    (!startOnlyDate || adjustmentOnlyDate >= startOnlyDate)

    &&

    (!endOnlyDate || adjustmentOnlyDate <= endOnlyDate);

  if(adjustmentInRange){

    totalAdjustment += adjustment;

  }

}

// =====================================
// OUTSTANDING DEBT
// ALL MONTH
// =====================================

if(
  status != "Paid"
  &&
  remaining > 0
){

  outstandingDebt += remaining;

}

    // =====================================
// CUSTOMER MAP
// MONTH TO DATE ONLY
// =====================================

if(inDateRange){

  if(!customerMap[customer]){

    customerMap[customer] = {

      displayName: customerRaw,
      revenue:0,
      profit:0,
      qty:0,
      stainQty:0,
      orders:0,
      type:customerType,
      area:area,
      firstDate:orderDate

    };

  }

  customerMap[customer].revenue += revenue;

  customerMap[customer].profit += actualProfit;

  customerMap[customer].qty += qty;

  customerMap[customer].stainQty += stainQty;

  customerMap[customer].orders++;

}

    // =====================================
    // DRIVER MAP
    // =====================================

    if(!driverMap[driver]){

      driverMap[driver] = {

        revenue:0,
        qty:0,
        tripsIn:0,
        tripsOut:0,
        profit:0

      };

    }

    driverMap[driver].revenue += revenue;
    driverMap[driver].qty += qty;
    driverMap[driver].profit += actualProfit;

    if(driver){

      driverMap[driver].tripsIn++;

    }

    if(driverOut){

      driverMap[driverOut] =
        driverMap[driverOut] || {

          revenue:0,
          qty:0,
          tripsIn:0,
          tripsOut:0,
          profit:0

        };

      driverMap[driverOut].tripsOut++;

    }

    // =====================================
    // AREA MAP
    // =====================================

    if(!areaMap[area]){

      areaMap[area] = {

        revenue:0,
        qty:0,
        profit:0,
        shops:new Set(),
        vip:0,
        orders:0

      };

    }

    areaMap[area].revenue += revenue;
    areaMap[area].qty += qty;
    areaMap[area].profit += actualProfit;
    areaMap[area].orders++;

    areaMap[area].shops.add(customer);

    if(customerType == "VIP"){

      areaMap[area].vip++;

    }

    // =====================================
    // REVENUE TREND
    // =====================================

    const day =

      Utilities.formatDate(

        orderDate,

        Session.getScriptTimeZone(),

        "yyyy-MM-dd"

      );

    uniqueDays.add(day);

    revenueByDate[day] =

      (revenueByDate[day] || 0)

      +

      revenue;

    // =====================================
    // AVG LAUNDRY TIME
    // =====================================

    if(invoiceDate){

      const days =

        Math.floor(

          (invoiceDate - orderDate)

          /

          (1000*60*60*24)

        );

      totalLaundryDays += days;

      completedLaundry++;

    }

    // =====================================
    // HOLDING ALERT
    // =====================================

    if(!invoiceDate){

      const holdingDays =

        Math.floor(

          (new Date() - orderDate)

          /

          (1000*60*60*24)

        );

      if(holdingDays >= 3){

        holdingAlert++;

      }

    }

    // =====================================
    // UNPAID INVOICE
    // =====================================

    if(
      status != "Paid"
      &&
      remaining > 0
    ){

      let debtDays = 0;

      if(invoiceDate){

        debtDays =

          Math.floor(

            (new Date() - invoiceDate)

            /

            (1000*60*60*24)

          );

      }

      if(debtDays >= 3){

        debtAlert++;

      }

      unpaidInvoices.push({

        invoiceID,

        invoiceDate:
          invoiceDate
          ? Utilities.formatDate(
              invoiceDate,
              Session.getScriptTimeZone(),
              "yyyy-MM-dd hh:mm a"
            )
          : "",

        customer: customerRaw,

        driverOut,

        amount: remaining,

        debtDays

      });

    }

    // =====================================
    // FILTER
    // =====================================

    let include = true;

    if(statusFilter == "Partial"){

      include =
        status == "Partial";

    }

    if(statusFilter == "Unpaid"){

      include =
        status == "Unpaid";

    }

    // =====================================
    // LATEST ORDERS
    // =====================================

    if(
      include &&
      status != "Paid"
    ){

      latestOrders.push({

        date:

          Utilities.formatDate(

            orderDate,

            Session.getScriptTimeZone(),

            "yyyy-MM-dd hh:mm a"

          ),

        shop: customerRaw,

        driver: driver,

        qty: qty,

        amount: revenue,

        status: status

      });

    }

  });

  // =====================================
  // TOP CUSTOMERS
  // =====================================

  const topCustomers =

    Object.entries(customerMap)

    .map(([name,data]) => {

      let stainPercent = 0;

if(data.qty > 0){

  stainPercent =
    (data.stainQty / data.qty) * 100;

}

let risk = "Low";

if(stainPercent > 7){

  risk = "High";

}

else if(stainPercent >= 5){

  risk = "Medium";

}

      return {

        name:data.displayName,

        revenue:data.revenue,

        profit:data.profit,

        qty:data.qty,

        avgQty:

          (
            data.qty / data.orders
          ).toFixed(0),

        orders:data.orders,

        type:data.type,

        area:data.area,

        stainPercent:
          stainPercent.toFixed(1),

        customerLifetime:
          formatCustomerAge(
            data.firstDate
          ),

        risk:risk

      };

    })

    .sort((a,b)=>
      b.revenue - a.revenue
    )

    .slice(0,15);

  // =====================================
  // TOP AREA
  // =====================================

  const topAreas =

    Object.entries(areaMap)

    .map(([name,data]) => {

      let recommendation = "Monitor";

      if(data.qty >= 5000){

        recommendation =
          "Open Mini Branch";

      }

      return {

        area:name,

        revenue:data.revenue,

        qty:data.qty,

        profit:data.profit,

        shops:data.shops.size,

        orders:data.orders,

        vip:data.vip,

        recommendation

      };

    })

    .sort((a,b)=>
      b.revenue - a.revenue
    );

  // =====================================
  // AVG
  // =====================================

  const totalDays =
    uniqueDays.size || 1;

  const avgRevenuePerDay =
    forecastRevenue / totalDays;

  const avgQtyPerDay =
    totalQty / totalDays;

  const avgLaundryTime =

    completedLaundry > 0

    ?

      (
        totalLaundryDays
        /
        completedLaundry
      ).toFixed(1)

    :

      0;

  // =====================================
  // RESULT
  // =====================================

  const result = {

    forecastRevenue,
    todayRevenue,
    todayQty,
    todayOrders,

    outstandingDebt,

    totalOrders,

    totalQty,

    avgRevenuePerDay,

    avgQtyPerDay,

    paidAmount,

    profit,

    totalAdjustment,

    activeShops:
      Object.keys(customerMap).length,
      totalShops:
      totalShops,

    holdingAlert,

    debtAlert,

    avgLaundryTime,

    revenueByDate,

    topCustomers,

    topDrivers:

      Object.entries(driverMap)

      .sort((a,b)=>
        b[1].revenue
        - a[1].revenue
      )

      .slice(0,20),

    topAreas,

    latestOrders:

      latestOrders
        .slice(-20)
        .reverse(),

    unpaidInvoices:

      unpaidInvoices

        .sort((a,b)=>
          b.debtDays
          - a.debtDays
        )

        .slice(0,20)

  };

  // =====================================
  // SAVE CACHE
  // =====================================

  cache.put(
    cacheKey,
    JSON.stringify(result),
    CACHE_TIME
  );

  return result;

}


// =========================================
// CLEAR CACHE
// =========================================

function clearDashboardCache(){

  CacheService
    .getScriptCache()
    .removeAll([]);

}


// =========================================
// REBUILD CUSTOMER SUMMARY
// =========================================

function rebuildCustomerSummary(){

  const ss =
    SpreadsheetApp.getActive();

  const customerSheet =
    ss.getSheetByName("Customer");

  const summarySheet =
    ss.getSheetByName("Customer_Summary");

  const data =
    customerSheet.getDataRange().getValues();

  const headers =
    data[0].map(h => cleanText(h));

  const col = name =>
    headers.indexOf(cleanText(name));

  const shopCol = col("shop name");
  const phoneCol = col("phone");
  const addressCol = col("address");
  const areaCol = col("area");
  const gpsCol = col("gps");
  const typeCol = col("customer type");

  const output = [];

  output.push([

    "Shop Name",
    "Phone",
    "Address",
    "Area",
    "GPS",
    "Customer Type",
    "Revenue",
    "Qty",
    "Stain Qty",
    "Stain %",
    "Orders",
    "Customer Lifetime",
    "Risk"

  ]);

  const used = {};

  for(let i=1;i<data.length;i++){

    const row = data[i];

    const shop =
      String(row[shopCol] || "").trim();

    if(!shop) continue;

    const key =
      cleanText(shop);

    if(used[key]) continue;

    used[key] = true;

    output.push([

      shop,
      row[phoneCol] || "",
      row[addressCol] || "",
      row[areaCol] || "",
      row[gpsCol] || "",
      row[typeCol] || "General",
      0,
      0,
      0,
      "0%",
      0,
      "0 Months",
      "Low"

    ]);

  }

  summarySheet.clearContents();

  summarySheet

    .getRange(
      1,
      1,
      output.length,
      output[0].length
    )

    .setValues(output);

}


// =========================================
// REBUILD DRIVER SUMMARY
// =========================================

function rebuildDriverSummary(){

  const ss =
    SpreadsheetApp.getActive();

  const orderSheet =
    ss.getSheetByName("Order");

  const driverSheet =
    ss.getSheetByName("Driver_Summary");

  const data =
    orderSheet.getDataRange().getValues();

  const headers =
    data[0].map(h => cleanText(h));

  const col = name =>
    headers.indexOf(cleanText(name));

  const driverCol = col("driver name");
  const qtyCol = col("total qty");
  const revenueCol = col("grand total");

  const map = {};

  for(let i=1;i<data.length;i++){

    const row = data[i];

    const driver =
      String(row[driverCol] || "").trim();

    if(!driver) continue;

    const qty =
      Number(row[qtyCol]) || 0;

    const revenue =
      Number(row[revenueCol]) || 0;

    if(!map[driver]){

      map[driver] = {

        qty:0,
        revenue:0,
        trips:0

      };

    }

    map[driver].qty += qty;
    map[driver].revenue += revenue;
    map[driver].trips++;

  }

  const output = [];

  output.push([
    "Driver",
    "Trips",
    "Qty",
    "Forecast Revenue"
  ]);

  Object.keys(map).forEach(driver => {

    output.push([

      driver,
      map[driver].trips,
      map[driver].qty,
      map[driver].revenue

    ]);

  });

  driverSheet.clearContents();

  driverSheet

    .getRange(
      1,
      1,
      output.length,
      output[0].length
    )

    .setValues(output);

}


// =========================================
// REBUILD DAILY SUMMARY
// =========================================

function rebuildDailySummary(){

  const ss =
    SpreadsheetApp.getActive();

  const orderSheet =
    ss.getSheetByName("Order");

  const dailySheet =
    ss.getSheetByName("Daily_Summary");

  const data =
    orderSheet.getDataRange().getValues();

  const headers =
    data[0].map(h => cleanText(h));

  const col = name =>
    headers.indexOf(cleanText(name));

  const dateCol = col("date");
  const qtyCol = col("total qty");
  const revenueCol = col("grand total");

  const map = {};

  for(let i=1;i<data.length;i++){

    const row = data[i];

    const rawDate =
      row[dateCol];

    if(!rawDate) continue;

    const date =

      Utilities.formatDate(

        new Date(rawDate),

        Session.getScriptTimeZone(),

        "yyyy-MM-dd"

      );

    const qty =
      Number(row[qtyCol]) || 0;

    const revenue =
      Number(row[revenueCol]) || 0;

    if(!map[date]){

      map[date] = {

        qty:0,
        revenue:0,
        orders:0

      };

    }

    map[date].qty += qty;
    map[date].revenue += revenue;
    map[date].orders++;

  }

  const output = [];

  output.push([
    "Date",
    "Orders",
    "Qty",
    "Forecast Revenue"
  ]);

  Object.keys(map).forEach(date => {

    output.push([

      date,
      map[date].orders,
      map[date].qty,
      map[date].revenue

    ]);

  });

  dailySheet.clearContents();

  dailySheet

    .getRange(
      1,
      1,
      output.length,
      output[0].length
    )

    .setValues(output);

}


// =========================================
// AUTO REFRESH
// =========================================

function refreshAllSummary(){

  rebuildCustomerSummary();

  rebuildDriverSummary();

  rebuildDailySummary();

  clearDashboardCache();

  Logger.log("ALL SUMMARY REFRESHED");

}

function onEdit(e){

  const sheet = e.source.getSheetByName("Order");

  const range = e.range;

  const row = range.getRow();

  const col = range.getColumn();

  if(sheet.getName() != "Order") return;

  const headers = sheet
    .getRange(1,1,1,sheet.getLastColumn())
    .getValues()[0]
    .map(h => String(h).trim().toLowerCase());

  const paidCol =
    headers.indexOf("amount paid") + 1;

  const paymentDateCol =
    headers.indexOf("payment date") + 1;

  if(
    col == paidCol
    &&
    row > 1
  ){

    const paidValue =
      sheet.getRange(row, paidCol).getValue();

    const paymentCell =
      sheet.getRange(row, paymentDateCol);

    if(paidValue && !paymentCell.getValue()){

      paymentCell.setValue(new Date());

    }

    if(!paidValue){

      paymentCell.clearContent();

    }

  }

}