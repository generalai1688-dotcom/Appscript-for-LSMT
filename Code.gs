const SHEET_ORDER = "Order";
const SHEET_CUSTOMER = "Customer_Summary";


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
// FORMAT CUSTOMER AGE
// =========================================
function formatCustomerAge(firstDate){

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

  return (
    years
    + " Years "
    + remainMonths
    + " Months"
  );

}


// =========================================
// DASHBOARD DATA
// =========================================
function getDashboardData(
  startDate,
  endDate,
  statusFilter
) {

  const ss =
    SpreadsheetApp.getActive();

  // =====================================
  // ORDER SHEET
  // =====================================

  const orderSheet =
    ss.getSheetByName(SHEET_ORDER);

  const orderData =
    orderSheet
      .getDataRange()
      .getValues();

  const orderHeaders =
    orderData[0].map(h =>
      String(h).trim()
    );

  const orders =
    orderData.slice(1);

  // =====================================
  // CUSTOMER SUMMARY
  // =====================================

  const customerSheet =
    ss.getSheetByName(
      SHEET_CUSTOMER
    );

  const customerData =
    customerSheet
      .getDataRange()
      .getValues();

  const customerHeaders =
    customerData[0];

  const customerRows =
    customerData.slice(1);

  const customerMapInfo = {};

  const ccol = name =>
    customerHeaders.indexOf(name);

  customerRows.forEach(r => {

    const customer =
      String(
        r[ccol("Customer")] || ""
      ).trim();

    if(!customer) return;

    customerMapInfo[customer] = {

      area:
        r[ccol("Area")] || "Unknown",

      address:
        r[ccol("Address")] || "",

      type:
        r[ccol("Customer Type")] || "General"

    };

  });

  // =====================================
  // COLUMN
  // =====================================

  const col = name =>
    orderHeaders.indexOf(name);

  const orderDateCol =
    col("Date");

  const customerCol =
    col("Display Shop");

  const driverCol =
    col("Driver Name");

  const qtyCol =
    col("Total Qty");

  const stainQtyCol =
    col("Total Stain Qty");

  const revenueCol =
    col("Grand Total");

  const stainCostCol =
    col("Total Stain Cost");

  const deliveryCol =
    col("Delivery Charge");

  const invoiceCol =
    col("Invoice ID");

  const invoiceDateCol =
    col("Invoice Date");

  const driverOutCol =
    col("Driver Out");

  const paidCol =
    col("Amount Paid");

  const statusCol =
    col("Payment Status");

  const remainCol =
    col("Remaining Bal");

  // =====================================
  // KPI
  // =====================================

  let totalRevenue = 0;
  let totalDebt = 0;
  let totalOrders = 0;
  let totalQty = 0;
  let paidAmount = 0;
  let netProfit = 0;

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

  let start =
    startDate
      ? new Date(startDate)
      : null;

  let end =
    endDate
      ? new Date(endDate)
      : null;

  // =====================================
  // LOOP
  // =====================================

  orders.forEach(r => {

    if(!r[0]) return;

    const orderDate =
      new Date(r[orderDateCol]);

    if(isNaN(orderDate)) return;

    if(start && orderDate < start) return;

    if(end && orderDate > end) return;

    const customer =
      String(
        r[customerCol] || ""
      ).trim();

    const customerInfo =
      customerMapInfo[customer] || {};

    const area =
      customerInfo.area || "Unknown";

    const customerType =
      customerInfo.type || "General";

    const driver =
      String(
        r[driverCol] || ""
      ).trim();

    const driverOut =
      String(
        r[driverOutCol] || ""
      ).trim();

    const qty =
      Number(r[qtyCol]) || 0;

    const stainQty =
      Number(r[stainQtyCol]) || 0;

    const revenue =
      Number(r[revenueCol]) || 0;

    const stainCost =
      Number(r[stainCostCol]) || 0;

    const deliveryCost =
      Number(r[deliveryCol]) || 0;

    const paid =
      Number(r[paidCol]) || 0;

    const status =
      String(
        r[statusCol] || ""
      ).trim();

    const remaining =
      Number(r[remainCol]) || 0;

    const invoiceID =
      r[invoiceCol];

    const invoiceDate =
      r[invoiceDateCol]
        ? new Date(r[invoiceDateCol])
        : null;

    // =====================================
    // PROFIT
    // =====================================

    const profit =
      revenue
      - stainCost
      - deliveryCost;

    netProfit += profit;

    // =====================================
    // TOTAL
    // =====================================

    totalRevenue += revenue;
    totalDebt += remaining;
    totalQty += qty;
    totalOrders++;
    paidAmount += paid;

    // =====================================
    // CUSTOMER
    // =====================================

    if(!customerMap[customer]){

      customerMap[customer] = {

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
    customerMap[customer].profit += profit;
    customerMap[customer].qty += qty;
    customerMap[customer].stainQty += stainQty;
    customerMap[customer].orders++;

    // =====================================
    // DRIVER
    // =====================================

    if(!driverMap[driver]){

      driverMap[driver] = {

        revenue:0,
        qty:0,
        trips:0,
        profit:0

      };

    }

    driverMap[driver].revenue += revenue;
    driverMap[driver].qty += qty;
    driverMap[driver].trips++;
    driverMap[driver].profit += profit;

    // =====================================
    // AREA
    // =====================================

    if(!areaMap[area]){

      areaMap[area] = {

        revenue:0,
        qty:0,
        profit:0,
        shops:new Set(),
        vip:0

      };

    }

    areaMap[area].revenue += revenue;
    areaMap[area].qty += qty;
    areaMap[area].profit += profit;

    areaMap[area]
      .shops
      .add(customer);

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

    revenueByDate[day] =
      (revenueByDate[day] || 0)
      + revenue;

    // =====================================
    // AVG LAUNDRY
    // =====================================

    if(invoiceDate){

      const days =
        Math.floor(
          (invoiceDate - orderDate)
          / (1000*60*60*24)
        );

      totalLaundryDays += days;
      completedLaundry++;

    }

    // =====================================
    // HOLDING
    // =====================================

    if(!invoiceDate){

      const holdingDays =
        Math.floor(
          (new Date() - orderDate)
          / (1000*60*60*24)
        );

      if(holdingDays >= 5){

        holdingAlert++;

      }

    }

    // =====================================
    // DEBT
    // =====================================

    if(
      invoiceDate &&
      status != "Paid"
    ){

      const debtDays =
        Math.floor(
          (new Date() - invoiceDate)
          / (1000*60*60*24)
        );

      if(debtDays >= 7){

        debtAlert++;

      }

      unpaidInvoices.push({

        invoiceID,
        customer,
        driverOut,
        amount:remaining,
        debtDays

      });

    }

    // =====================================
    // FILTER
    // =====================================

    let include = true;

    if(statusFilter == "Paid"){
      include = status == "Paid";
    }

    if(statusFilter == "Partial"){
      include = status == "Partial";
    }

    if(statusFilter == "Unpaid"){
      include = status == "Unpaid";
    }

    // =====================================
    // LATEST
    // =====================================

    if(include){

      latestOrders.push({

        date:
          Utilities.formatDate(
            orderDate,
            Session.getScriptTimeZone(),
            "yyyy-MM-dd hh:mm a"
          ),

        shop:customer,
        driver:driver,
        qty:qty,
        amount:revenue,
        status:status

      });

    }

  });

  // =====================================
  // CUSTOMERS
  // =====================================

  const topCustomers =

    Object.entries(customerMap)

    .map(([name,data]) => {

      return {

        name:name,

        revenue:data.revenue,

        profit:data.profit,

        qty:data.qty,

        avgQty:
          (
            data.qty
            / data.orders
          ).toFixed(0),

        orders:data.orders,

        type:data.type,

        area:data.area,

        stainPercent:
          (
            (
              data.stainQty
              / data.qty
            ) * 100
          ).toFixed(1),

        customerAge:
          formatCustomerAge(
            data.firstDate
          )

      };

    })

    .sort((a,b)=>
      b.revenue - a.revenue
    )

    .slice(0,20);

  // =====================================
  // AREA ANALYTICS
  // =====================================

  const topAreas =

    Object.entries(areaMap)

    .map(([name,data]) => {

      let recommendation =
        "Monitor";

      if(
        data.qty >= 5000
      ){

        recommendation =
          "Open Mini Branch";

      }

      return {

        area:name,

        revenue:data.revenue,

        qty:data.qty,

        profit:data.profit,

        shops:
          data.shops.size,

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

  const avgLaundryTime =

    completedLaundry > 0

    ? (
        totalLaundryDays
        / completedLaundry
      ).toFixed(1)

    : 0;

  // =====================================
  // RETURN
  // =====================================

  return {

    totalRevenue,
    totalDebt,
    totalOrders,
    totalQty,
    paidAmount,
    netProfit,

    activeShops:
      Object.keys(customerMap).length,

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
        .slice(-30)
        .reverse(),

    unpaidInvoices:
      unpaidInvoices
        .sort((a,b)=>
          b.debtDays
          - a.debtDays
        )
        .slice(0,20)

  };

}