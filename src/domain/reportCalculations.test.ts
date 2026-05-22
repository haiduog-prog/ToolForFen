import { describe, expect, it } from "vitest";
import { buildMetricRow } from "./reportCalculations";
import type { MonthKey } from "./month";

describe("buildMetricRow", () => {
  it("calculates trend, ytd growth and month growth from the selected report month", () => {
    const periodMonths = [
      "2024-01",
      "2024-02",
      "2024-03",
      "2024-04",
      "2024-05",
      "2024-06",
      "2024-07",
      "2024-08",
      "2024-09",
      "2024-10",
      "2024-11",
      "2024-12",
      "2025-01",
      "2025-02",
      "2025-03",
      "2025-04"
    ] as MonthKey[];

    const values = Object.fromEntries(periodMonths.map((month, index) => [month, index + 1])) as Record<MonthKey, number>;
    const row = buildMetricRow("B&M", values, periodMonths, "2025-04");

    expect(row.previousYearTotal).toBe(78);
    expect(row.currentYearTotal).toBe(58);
    expect(row.p3m).toBe(15);
    expect(row.p6m).toBe(13.5);
    expect(row.p9m).toBe(12);
    expect(row.trend).toBeCloseTo(30 / 25.5);
    expect(row.ifytd).toBeCloseTo(58 / 10);
    expect(row.icytd).toBeCloseTo(58 / 10);
    expect(row.iya).toBe(16 / 4);
  });
});

describe("writeReportWorkbook", () => {
  it("generates Excel file without crashing", async () => {
    const { writeReportWorkbook } = await import("../infrastructure/excel/writeReportWorkbook");
    const dummyReport = {
      reportMonth: "2025-04" as MonthKey,
      periodMonths: ["2024-01", "2025-04"] as MonthKey[],
      previousYear: 2024,
      currentYear: 2025,
      summary: {
        transactionCount: 0,
        customerCount: 0,
        productCount: 0,
        unmappedCustomerCount: 0,
        unmappedCustomers: [],
        availableMonths: ["2025-04"] as MonthKey[]
      },
      qecRows: [],
      dsrRows: [],
      customerBaseRows: [],
      skuRevenueRows: [],
      skuQuantityRows: [],
      customerRevenueSections: [],
      customerQuantitySections: []
    };

    const blob = await writeReportWorkbook(dummyReport, []);
    expect(blob).toBeInstanceOf(Blob);
    expect(blob.size).toBeGreaterThan(0);
  });
});

describe("readSourceWorkbook", () => {
  it("parses datamau.xlsx successfully", async () => {
    const fs = await import("fs");
    const { readSourceWorkbook } = await import("../infrastructure/excel/readSourceWorkbook");
    const buffer = fs.readFileSync("datamau.xlsx");
    const mockFile = {
      name: "datamau.xlsx",
      arrayBuffer: async () => {
        const ab = new ArrayBuffer(buffer.length);
        const view = new Uint8Array(ab);
        for (let i = 0; i < buffer.length; ++i) {
          view[i] = buffer[i];
        }
        return ab;
      }
    };

    const result = await readSourceWorkbook(mockFile as any);
    console.log("TEST OUTPUT - parsed transactions:", result.transactions.length);
    console.log("TEST OUTPUT - warnings:", result.warnings);
    console.log("TEST OUTPUT - availableMonths:", result.availableMonths);
    expect(result.transactions.length).toBeGreaterThan(0);
  });
});


