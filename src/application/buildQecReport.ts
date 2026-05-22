import {
  addShareColumns,
  aggregateMetricRows,
  buildCustomerSections,
  filterTransactionsThroughMonth,
  totalMetricRow,
  buildMetricRow
} from "../domain/reportCalculations";
import { compareMonths, reportPeriodMonths, type MonthKey } from "../domain/month";
import { UNMAPPED_SEGMENT, type QecReport, type SourceParseResult, type SourceTransaction, type MetricRow } from "../domain/entities";
import { DSR_ORDER, lookupSegment } from "../domain/customerMapping";

export function buildQecReport(source: SourceParseResult, reportMonth: MonthKey): QecReport {
  if (!source.availableMonths.includes(reportMonth)) {
    throw new Error("Tháng báo cáo không có trong file Data nguồn.");
  }

  const periodMonths = reportPeriodMonths(reportMonth);
  const transactions = filterTransactionsThroughMonth(source.transactions, reportMonth, periodMonths);
  const currentYear = Number(reportMonth.slice(0, 4));
  const previousYear = currentYear - 1;

  const EXCLUDED_CUSTOMERS = [
    "Hộ kinh doanh Vũ Thị Hòa (WS Ngọc Trâm)",
    "WS Như Trần Ngọc (cô Xuyến)",
    "Khách lẻ (Cảnh)",
    "Quầy thuốc Hoài Tính",
    "Chị Châu",
    "Bác sĩ Nguyễn Văn Hậu"
  ];

  const qecBaseRows = aggregateMetricRows(
    transactions,
    periodMonths,
    reportMonth,
    (transaction) => {
      if (
        transaction.customer === "Nhà thuốc An Khang (Kho AK Đà Nẵng)" ||
        transaction.customer === "Nhà thuốc An Khang (Hub)"
      ) {
        return "MT";
      }
      if (transaction.customer === "Nhà thuốc An Khang (Kho AK Bến Tre)") {
        return "OTC";
      }
      return lookupSegment(transaction.customer);
    },
    (transaction) => {
      if (EXCLUDED_CUSTOMERS.includes(transaction.customer)) {
        return 0;
      }
      return transaction.revenue;
    },
    "segment"
  );

  const STATIC_SEGMENTS = ["B&M", "WS", "Ecom", "ETC", "IDP", "KEY", "OTC", "MT", "AK"];
  const qecMap = new Map<string, typeof qecBaseRows[0]>();
  qecBaseRows.forEach((row) => {
    qecMap.set(row.label, row);
  });

  const orderedQecRows = STATIC_SEGMENTS.map((seg) => {
    return qecMap.get(seg) ?? buildMetricRow(seg, {}, periodMonths, reportMonth);
  });

  // Bổ sung dòng Unmapped vào cuối bảng QEC review nếu phân khúc này có doanh số phát sinh
  const unmappedRow = qecMap.get(UNMAPPED_SEGMENT);
  if (unmappedRow && (unmappedRow.previousYearTotal > 0 || unmappedRow.currentYearTotal > 0)) {
    orderedQecRows.push(unmappedRow);
  }

  const qecRowsWithTotal = [...addShareColumns(orderedQecRows), totalMetricRow("Total", orderedQecRows, periodMonths, reportMonth)];

  // Group quantity from actual transactions
  const qtyMap = new Map<string, Record<MonthKey, number>>();
  for (const transaction of transactions) {
    const values = qtyMap.get(transaction.product) ?? {};
    values[transaction.month] = (values[transaction.month] ?? 0) + transaction.quantity;
    qtyMap.set(transaction.product, values);
  }

  // SKU Quantity Rows: Chỉ lấy sản phẩm thực tế phát sinh giao dịch trong file input
  const skuQuantityRows: MetricRow[] = [];
  for (const [prodName, values] of qtyMap.entries()) {
    const row = buildMetricRow(prodName, values, periodMonths, reportMonth);
    if (row.previousYearTotal > 0 || row.currentYearTotal > 0) {
      skuQuantityRows.push(row);
    }
  }
  skuQuantityRows.sort((a, b) => b.currentYearTotal - a.currentYearTotal);

  // Group revenue from actual transactions
  const revMap = new Map<string, Record<MonthKey, number>>();
  for (const transaction of transactions) {
    const values = revMap.get(transaction.product) ?? {};
    values[transaction.month] = (values[transaction.month] ?? 0) + transaction.revenue;
    revMap.set(transaction.product, values);
  }

  // SKU Revenue Rows: Chỉ lấy sản phẩm thực tế phát sinh giao dịch trong file input
  const skuRevenueRows: MetricRow[] = [];
  for (const [prodName, values] of revMap.entries()) {
    const row = buildMetricRow(prodName, values, periodMonths, reportMonth);
    if (row.previousYearTotal > 0 || row.currentYearTotal > 0) {
      skuRevenueRows.push(row);
    }
  }
  skuRevenueRows.sort((a, b) => b.currentYearTotal - a.currentYearTotal);


  const nonExcludedTransactions = transactions.filter(
    (t) => !EXCLUDED_CUSTOMERS.includes(t.customer)
  );

  const customerRevenueSections = buildCustomerSections(
    nonExcludedTransactions,
    periodMonths,
    reportMonth,
    (transaction: SourceTransaction) => transaction.revenue
  );

  const customerQuantitySections = buildCustomerSections(
    nonExcludedTransactions,
    periodMonths,
    reportMonth,
    (transaction: SourceTransaction) => transaction.quantity
  );

  const dsrBaseRows = aggregateMetricRows(
    transactions,
    periodMonths,
    reportMonth,
    (transaction) => transaction.dsr,
    (transaction) => {
      if (EXCLUDED_CUSTOMERS.includes(transaction.customer)) {
        return 0;
      }
      return transaction.revenue;
    },
    "segment"
  );
  dsrBaseRows.sort((a, b) => {
    const ia = DSR_ORDER.indexOf(a.label);
    const ib = DSR_ORDER.indexOf(b.label);
    return (ia === -1 ? 999 : ia) - (ib === -1 ? 999 : ib);
  });
  const dsrRows = [...dsrBaseRows, totalMetricRow("Total", dsrBaseRows, periodMonths, reportMonth)];

  // Group customer revenues from non-excluded transactions
  const customerValuesMap = new Map<string, Record<MonthKey, number>>();
  for (const transaction of transactions) {
    if (EXCLUDED_CUSTOMERS.includes(transaction.customer)) {
      continue;
    }
    const customer = transaction.customer;
    const values = customerValuesMap.get(customer) ?? {};
    values[transaction.month] = (values[transaction.month] ?? 0) + transaction.revenue;
    customerValuesMap.set(customer, values);
  }

  // Customer Base Rows: Chỉ lấy những khách hàng thực tế phát sinh giao dịch trong file input
  const customerBaseRows: MetricRow[] = [];
  for (const [customerName, values] of customerValuesMap.entries()) {
    const row = buildMetricRow(customerName, values, periodMonths, reportMonth);
    if (row.previousYearTotal > 0 || row.currentYearTotal > 0) {
      customerBaseRows.push(row);
    }
  }
  customerBaseRows.sort((a, b) => a.label.localeCompare(b.label, "vi"));


  const customers = new Set(transactions.map((transaction) => transaction.customer).filter(Boolean));
  const products = new Set(transactions.map((transaction) => transaction.product).filter(Boolean));
  const unmappedCustomers = Array.from(
    new Set(
      transactions
        .filter((transaction) => transaction.segment === UNMAPPED_SEGMENT)
        .map((transaction) => transaction.customer)
        .filter(Boolean)
    )
  ).sort((a, b) => a.localeCompare(b, "vi"));

  return {
    reportMonth,
    periodMonths,
    previousYear,
    currentYear,
    summary: {
      transactionCount: transactions.length,
      customerCount: customers.size,
      productCount: products.size,
      unmappedCustomerCount: unmappedCustomers.length,
      unmappedCustomers,
      availableMonths: [...source.availableMonths].sort(compareMonths)
    },
    qecRows: qecRowsWithTotal,
    dsrRows,
    customerBaseRows,
    skuRevenueRows,
    skuQuantityRows,
    customerRevenueSections,
    customerQuantitySections
  };
}
