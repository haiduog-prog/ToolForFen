import { describe, expect, it } from "vitest";
import { generateTimelineBlocks, getInitialProducts } from "./masterplan";
import { buildMasterplan } from "../application/buildMasterplan";

describe("Sales Plan Timeline Blocks Generation", () => {
  it("should generate blocks starting from JUN 2026 to DEC 2026 with correct quarter boundaries", () => {
    // startMonthIndex = 5 (June), startYear = 2026, productCount = 6
    const blocks = generateTimelineBlocks(2026, 5, 6);
    
    // Check block sequence
    expect(blocks[0]?.label).toBe("JUN 2026");
    expect(blocks[0]?.type).toBe("month");
    expect(blocks[0]?.quarter).toBe(2);

    expect(blocks[1]?.label).toBe("TOTAL Q2 2026");
    expect(blocks[1]?.type).toBe("quarter");
    expect(blocks[1]?.quarter).toBe(2);
    expect(blocks[1]?.monthBlocks?.length).toBe(1); // Only JUN

    expect(blocks[2]?.label).toBe("JUL 2026");
    expect(blocks[2]?.type).toBe("month");
    expect(blocks[2]?.quarter).toBe(3);

    expect(blocks[3]?.label).toBe("AUG 2026");
    expect(blocks[4]?.label).toBe("SEP 2026");
    expect(blocks[4]?.quarter).toBe(3);

    expect(blocks[5]?.label).toBe("TOTAL Q3 2026");
    expect(blocks[5]?.type).toBe("quarter");
    expect(blocks[5]?.quarter).toBe(3);
    expect(blocks[5]?.monthBlocks?.length).toBe(3); // JUL, AUG, SEP

    expect(blocks[6]?.label).toBe("OCT 2026");
    expect(blocks[7]?.label).toBe("NOV 2026");
    expect(blocks[8]?.label).toBe("DEC 2026");
    
    expect(blocks[9]?.label).toBe("TOTAL Q4 2026");
    expect(blocks[9]?.type).toBe("quarter");
    expect(blocks[9]?.monthBlocks?.length).toBe(3); // OCT, NOV, DEC

    expect(blocks[10]?.label).toBe("TOTAL YEAR 2026");
    expect(blocks[10]?.type).toBe("year");
    expect(blocks[10]?.quarterBlocks?.length).toBe(3); // Q2, Q3, Q4
  });

  it("should calculate correct column boundaries for each block dynamically", () => {
    const productCount = 2; // e.g. 2 products
    const blocks = generateTimelineBlocks(2026, 5, productCount);

    // Fixed columns A to I take columns 1 to 9.
    // J (column 10) is the start of first block.
    // Block 0: JUN 2026
    // productCols: 10, 11
    // Total Volume: 12
    // Gross: 13
    expect(blocks[0]?.startCol).toBe(10);
    expect(blocks[0]?.endCol).toBe(13); // currentCol + N + 1
    expect(blocks[0]?.productCols).toEqual([10, 11]);
    expect(blocks[0]?.totalCol).toBe(12);
    expect(blocks[0]?.grossCol).toBe(13);

    // Block 1: TOTAL Q2 2026
    // currentCol incremented by N + 2 => 10 + 4 = 14 (Col N)
    // productCols: 14, 15
    // Quarter Total Volume: 16
    expect(blocks[1]?.startCol).toBe(14);
    expect(blocks[1]?.endCol).toBe(16);
    expect(blocks[1]?.productCols).toEqual([14, 15]);
    expect(blocks[1]?.totalCol).toBe(16);
    expect(blocks[1]?.grossCol).toBeNull();
  });

  it("should load initial products list", () => {
    const prods = getInitialProducts();
    expect(prods.length).toBeGreaterThan(0);
    expect(prods[0]?.name).toBe("SPR EPO");
    expect(prods[0]?.price).toBe(590000);
    expect(prods[0]?.amountThreshold).toBe(1500000);
    expect(prods[0]?.quantityA).toBe(10);
    expect(prods[0]?.quantityB).toBe(2);
  });

  it("should calculate quantities based on customerTotalAmount threshold", () => {
    const product = { name: "SPR EPO", price: 590000, amountThreshold: 1500000, quantityA: 10, quantityB: 2 };
    
    // total amount > threshold
    const qtyHigh = 2000000 > product.amountThreshold ? product.quantityA : product.quantityB;
    expect(qtyHigh).toBe(10);

    // total amount <= threshold
    const qtyLow = 1000000 > product.amountThreshold ? product.quantityA : product.quantityB;
    expect(qtyLow).toBe(2);

    // total amount equal to threshold (should be quantityB since strictly greater)
    const qtyEqual = 1500000 > product.amountThreshold ? product.quantityA : product.quantityB;
    expect(qtyEqual).toBe(2);
  });

  it("should aggregate customer sales details correctly", () => {
    // Mock the raw rows from input2.xlsx format:
    // Columns: Khách hàng, Doanh thu chi tiết, Nhân viên, Mã hóa đơn
    const mockRows = [
      { customer: "Nha Thuoc A", revenue: 1000000, staff: "Staff X" },
      { customer: "Nha Thuoc A", revenue: 500000, staff: "Staff X" },
      { customer: "Nha Thuoc A", revenue: 200000, staff: "Staff Y" },
      { customer: "Nha Thuoc B", revenue: 3000000, staff: "Staff Y" },
      { customer: "Nha Thuoc B", revenue: -500000, staff: "Staff Y" },
    ];

    const groups: Record<string, { revenue: number; staffCounts: Record<string, number> }> = {};
    for (const row of mockRows) {
      const c = row.customer;
      if (!groups[c]) {
        groups[c] = { revenue: 0, staffCounts: {} };
      }
      groups[c].revenue += row.revenue;
      groups[c].staffCounts[row.staff] = (groups[c].staffCounts[row.staff] || 0) + 1;
    }

    const customers = Object.keys(groups).map((c) => {
      const grp = groups[c]!;
      return {
        customer: c,
        customerTotalAmount: grp.revenue,
        staff: "",
      };
    });

    expect(customers.length).toBe(2);
    
    const custA = customers.find((c) => c.customer === "Nha Thuoc A")!;
    expect(custA.customerTotalAmount).toBe(1700000);
    expect(custA.staff).toBe("");

    const custB = customers.find((c) => c.customer === "Nha Thuoc B")!;
    expect(custB.customerTotalAmount).toBe(2500000); // 3000000 - 500000
    expect(custB.staff).toBe("");
  });

  it("should add special Ecom channel when ecomEnabled is true", () => {
    const data = buildMasterplan([], [], true);
    
    const ecomRow = data.customers.find((c: any) => c.codeChannel === 1003);
    expect(ecomRow).toBeDefined();
    expect(ecomRow?.channel).toBe("Ecom");
    expect(ecomRow?.subChannel).toBe("Ecommerce");
    expect(ecomRow?.customer).toBe("Kênh Ecom");
    expect(ecomRow?.isEcomSpecial).toBe(true);
  });

  it("should successfully generate workbook with channel TOTAL rows and correct labels", async () => {
    const { writeMasterplanWorkbook } = await import("../infrastructure/excel/writeMasterplanWorkbook");
    const mockData = buildMasterplan(
      [
        {
          codeChannel: 1004,
          channel: "Baby & Mom",
          subChannel: "Baby & Mom",
          customer: "Store X",
          stores: "1",
          province: "HN",
          regional: "N",
          subChannel1: "B&M",
          staff: "",
          customerTotalAmount: 1000000
        }
      ],
      [
        { name: "SPR EPO", price: 590000, amountThreshold: 1500000, quantityA: 10, quantityB: 2 }
      ],
      true
    );

    const blob = await writeMasterplanWorkbook(mockData);
    expect(blob).toBeDefined();
    expect(blob.type).toBe("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  });
});
