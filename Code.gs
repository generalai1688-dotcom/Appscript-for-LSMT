const SHEET_ORDER = "Order";


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
// GET DASHBOARD DATA
// =========================================
function getDashboardData(
  startDate,
  endDate,
  statusFilter
) {

  const ss =
    SpreadsheetApp.getActive();

  const sheet =
    ss.getSheetByName(SHEET_ORDER);

  const allData =
    sheet.getDataRange().getValues();

  const rawHeaders =
    allData[0];

  // IMPORTANT
  // AUTO REMOVE SPACE
  const headers =
    rawHeaders.map(h =>
      String(h).trim()
    );

  const data =
    allData.slice(1);

  // =====================================
  // COLUMN INDEX
  // =====================================

  const col = name =>
    headers.indexOf(name);

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

  const invoiceIDCol =
    col("Invoice ID");

  const invoiceDateCol =
    col("Invoice Date");

  const paidCol =
    col("Amount Paid");

  const statusCol =
    col("Payment Status");

  const remainingCol =
    col("Remaining Bal");

  // =====================================
  // DEBUG
  // =====================================

  Logger.log(headers);

  // =====================================
  // KPI
  // =====================================

  let totalRevenue = 0;
  let totalDebt = 0;
  let totalOrders = 0;
  let totalQty = 0;
  let paidAmount = 0;

  let customerMap = {};
  let revenueByDate = {};
  let driverMap = {};

  let latestOrders = [];

  let overdueOrders = [];
  let unpaidInvoices = [];

  let holdingAlert = 0;
  let debtAlert = 0;

  let totalLaundryDays = 0;
  let completedLaundry = 0;

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

  data.forEach(r => {

    if (!r[0]) return;

    const orderDate =
      new Date(r[orderDateCol]);

    if (isNaN(orderDate)) return;

    if (start && orderDate < start) return;

    if (end && orderDate > end) return;

    const invoiceID =
      r[invoiceIDCol];

    const invoiceDate =
      r[invoiceDateCol]
        ? new Date(r[invoiceDateCol])
        : null;

    const customer =
      String(r[customerCol] || "").trim();

    const driver =
      String(r[driverCol] || "").trim();

    const qty =
      Number(r[qtyCol]) || 0;

    const stainQty =
      Number(r[stainQtyCol]) || 0;

    const revenue =
      Number(r[revenueCol]) || 0;

    const paid =
      Number(r[paidCol]) || 0;

    const status =
      String(r[statusCol] || "").trim();

    const remaining =
      Number(r[remainingCol]) || 0;

    // KPI
    totalRevenue += revenue;
    totalDebt += remaining;
    totalQty += qty;
    totalOrders++;
    paidAmount += paid;

    // CUSTOMER
    if(!customerMap[customer]){

      customerMap[customer] = {

        revenue:0,
        qty:0,
        stainQty:0,
        firstDate:orderDate

      };

    }

    customerMap[customer].revenue += revenue;
    customerMap[customer].qty += qty;
    customerMap[customer].stainQty += stainQty;

    if(
      orderDate <
      customerMap[customer].firstDate
    ){

      customerMap[customer].firstDate =
        orderDate;

    }

    // REVENUE TREND
    const day =
      Utilities.formatDate(
        orderDate,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd"
      );

    revenueByDate[day] =
      (revenueByDate[day] || 0)
      + revenue;

    // DRIVER
    if(!driverMap[driver]){

      driverMap[driver] = {

        revenue:0,
        qty:0,
        tripsIn:0,
        tripsOut:0,
        totalTrips:0

      };

    }

    driverMap[driver].revenue += revenue;
    driverMap[driver].qty += qty;
    driverMap[driver].tripsIn++;
    driverMap[driver].totalTrips++;

    if(invoiceDate){

      driverMap[driver].tripsOut++;

    }

    // LAUNDRY TIME
    let turnaroundDays = "";

    if(invoiceDate){

      turnaroundDays =
        Math.floor(
          (invoiceDate - orderDate)
          / (1000*60*60*24)
        );

      totalLaundryDays += turnaroundDays;
      completedLaundry++;

    }

    // HOLDING
    let holdingDays = "";

    if(!invoiceDate){

      holdingDays =
        Math.floor(
          (new Date() - orderDate)
          / (1000*60*60*24)
        );

      if(holdingDays >= 5){

        holdingAlert++;

      }

      if(holdingDays >= 3){

        overdueOrders.push({

          orderDate:
            Utilities.formatDate(
              orderDate,
              Session.getScriptTimeZone(),
              "yyyy-MM-dd"
            ),

          shop:customer,
          driver:driver,
          qty:qty,
          holdingDays:holdingDays

        });

      }

    }

    // DEBT
    let debtDays = "";

    if(
      invoiceDate &&
      status != "Paid"
    ){

      debtDays =
        Math.floor(
          (new Date() - invoiceDate)
          / (1000*60*60*24)
        );

      if(debtDays >= 7){

        debtAlert++;

      }

      unpaidInvoices.push({

        invoiceID:invoiceID,
        customer:customer,
        driver:driver,
        amount:remaining,
        debtDays:debtDays

      });

    }

    // FILTER
    let include = true;

    if(statusFilter == "Paid"){
      include = status == "Paid";
    }

    if(statusFilter == "Unpaid"){
      include = status != "Paid";
    }

    if(statusFilter == "Pending"){
      include = !invoiceDate;
    }

    // LATEST
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
        status:status,
        turnaroundDays,
        holdingDays,
        debtDays

      });

    }

  });

  // =====================================
  // CUSTOMER ANALYTICS
  // =====================================

  const topCustomers =

    Object.entries(customerMap)

    .map(([name,data]) => {

      const stainPercent =

        data.qty > 0

        ? (
            data.stainQty
            / data.qty
          ) * 100

        : 0;

      return {

        name:name,

        revenue:data.revenue,

        qty:data.qty,

        stainQty:data.stainQty,

        stainPercent:
          stainPercent.toFixed(1),

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

    activeShops:
      Object.keys(customerMap).length,

    holdingAlert,
    debtAlert,
    avgLaundryTime,

    topCustomers,

    topDrivers:
      Object.entries(driverMap)

      .sort((a,b)=>
        b[1].revenue
        - a[1].revenue
      )

      .slice(0,20),

    revenueByDate,

    latestOrders:
      latestOrders
        .slice(-30)
        .reverse(),

    overdueOrders:
      overdueOrders
        .sort((a,b)=>
          b.holdingDays
          - a.holdingDays
        )
        .slice(0,20),

    unpaidInvoices:
      unpaidInvoices
        .sort((a,b)=>
          b.debtDays
          - a.debtDays
        )
        .slice(0,20)

  };

}