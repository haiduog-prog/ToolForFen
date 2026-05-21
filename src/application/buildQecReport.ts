import {
  addShareColumns,
  aggregateMetricRows,
  buildCustomerSections,
  filterTransactionsThroughMonth,
  totalMetricRow,
  buildMetricRow
} from "../domain/reportCalculations";
import { compareMonths, reportPeriodMonths, type MonthKey } from "../domain/month";
import { UNMAPPED_SEGMENT, type QecReport, type SourceParseResult, type SourceTransaction } from "../domain/entities";
import { DSR_ORDER, STATIC_CUSTOMERS } from "../domain/customerMapping";

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
      return transaction.segment || UNMAPPED_SEGMENT;
    },
    (transaction) => {
      if (EXCLUDED_CUSTOMERS.includes(transaction.customer)) {
        return 0;
      }
      return transaction.revenue;
    },
    "segment"
  );
  const qecRowsWithTotal = [...addShareColumns(qecBaseRows), totalMetricRow("Total", qecBaseRows, periodMonths, reportMonth)];

  const skuRevenueRows = aggregateMetricRows(
    transactions,
    periodMonths,
    reportMonth,
    (transaction) => transaction.product,
    (transaction) => transaction.revenue,
    "total"
  );

  const skuQuantityRows = aggregateMetricRows(
    transactions,
    periodMonths,
    reportMonth,
    (transaction) => transaction.product,
    (transaction) => transaction.quantity,
    "total"
  );

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

  // Construct MetricRow for all 295 static customers in order
  const customerBaseRows = STATIC_CUSTOMERS.map((customerName) => {
    const values = customerValuesMap.get(customerName) ?? {};
    return buildMetricRow(customerName, values, periodMonths, reportMonth);
  });


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
