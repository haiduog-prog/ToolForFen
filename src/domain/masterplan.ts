export interface SalesPlanProduct {
  name: string;
  price: number;
  amountThreshold: number;
  quantityA: number;
  quantityB: number;
  ecomInitialQty?: number;
}

export interface SalesPlanCustomer {
  codeChannel: number;
  channel: string;
  subChannel: string;
  customer: string;
  stores: string;
  province: string;
  regional: string;
  subChannel1: string;
  staff: string;
  customerTotalAmount: number;
  isEcomSpecial?: boolean;
}

export interface ChannelOption {
  code: number;
  channel: string;
  subChannel: string;
}

export const CHANNEL_OPTIONS: ChannelOption[] = [
  { code: 1001, channel: "MTC", subChannel: "LC, PMC, AK" },
  { code: 1002, channel: "MTC", subChannel: "Health & Beauty" },
  { code: 1003, channel: "Ecom", subChannel: "Ecommerce" },
  { code: 1004, channel: "Baby & Mom", subChannel: "Baby & Mom" },
  { code: 1005, channel: "OTC", subChannel: "Key/Chain" },
  { code: 1006, channel: "ETC", subChannel: "ETC" },
  { code: 1007, channel: "OTC", subChannel: "IDP" }
];

export type BlockType = "month" | "quarter" | "year";

export interface TimelineBlock {
  type: BlockType;
  label: string; // e.g. "JUN 2026", "TOTAL Q2 2026", "TOTAL YEAR 2026"
  quarter: number | null; // 1..4 or null
  year: number;
  startCol: number; // 1-based column index in Excel
  endCol: number; // 1-based column index in Excel
  productCols: number[]; // Product column indexes
  totalCol: number; // Column index for Total Volume
  grossCol: number | null; // Column index for Gross Sales +VAT (null for quarter)
  monthBlocks?: TimelineBlock[]; // For quarter blocks, months included in this quarter
  quarterBlocks?: TimelineBlock[]; // For year block, quarters included in this year
}

export interface MasterplanData {
  products: SalesPlanProduct[];
  customers: SalesPlanCustomer[];
  timelineBlocks: TimelineBlock[];
  startYear: number;
  startMonthIndex: number; // 0-based month index (e.g. 5 for June)
}

const MONTH_NAMES = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];

/**
 * Returns the timeline blocks starting from the month after current month until December of the same year.
 */
export function generateTimelineBlocks(
  startYear: number,
  startMonthIndex: number, // 0-based (e.g. 5 for June)
  productCount: number
): TimelineBlock[] {
  const blocks: TimelineBlock[] = [];
  const quarterMonthsMap: Record<number, TimelineBlock[]> = {};
  
  let currentCol = 10; // J column is 10

  for (let m = startMonthIndex; m <= 11; m++) {
    const quarter = Math.ceil((m + 1) / 3);
    const productCols: number[] = [];
    for (let i = 0; i < productCount; i++) {
      productCols.push(currentCol + i);
    }

    const monthBlock: TimelineBlock = {
      type: "month",
      label: `${MONTH_NAMES[m]} ${startYear}`,
      quarter,
      year: startYear,
      startCol: currentCol,
      endCol: currentCol + productCount + 1,
      productCols,
      totalCol: currentCol + productCount,
      grossCol: currentCol + productCount + 1,
    };
    blocks.push(monthBlock);

    if (!quarterMonthsMap[quarter]) {
      quarterMonthsMap[quarter] = [];
    }
    quarterMonthsMap[quarter]!.push(monthBlock);
    currentCol += productCount + 2;

    const nextMonth = m + 1;
    const isLastInQuarter = nextMonth > 11 || Math.ceil((nextMonth + 1) / 3) !== quarter;
    if (isLastInQuarter) {
      const qProductCols: number[] = [];
      for (let i = 0; i < productCount; i++) {
        qProductCols.push(currentCol + i);
      }

      const quarterBlock: TimelineBlock = {
        type: "quarter",
        label: `TOTAL Q${quarter} ${startYear}`,
        quarter,
        year: startYear,
        startCol: currentCol,
        endCol: currentCol + productCount,
        productCols: qProductCols,
        totalCol: currentCol + productCount,
        grossCol: null,
        monthBlocks: quarterMonthsMap[quarter]!,
      };
      blocks.push(quarterBlock);
      currentCol += productCount + 1;
    }
  }

  // Create Year block at the end
  const quarterBlocks = blocks.filter((b) => b.type === "quarter");
  const yearProductCols: number[] = [];
  for (let i = 0; i < productCount; i++) {
    yearProductCols.push(currentCol + i);
  }

  const yearBlock: TimelineBlock = {
    type: "year",
    label: `TOTAL YEAR ${startYear}`,
    quarter: null,
    year: startYear,
    startCol: currentCol,
    endCol: currentCol + productCount + 1,
    productCols: yearProductCols,
    totalCol: currentCol + productCount,
    grossCol: currentCol + productCount + 1,
    quarterBlocks,
  };
  blocks.push(yearBlock);

  return blocks;
}

export function getInitialProducts(): SalesPlanProduct[] {
  return [
    { name: "SPR EPO", price: 590000, amountThreshold: 1500000, quantityA: 10, quantityB: 2 },
    { name: "SPR Collagen", price: 590000, amountThreshold: 2000000, quantityA: 8, quantityB: 1 },
    { name: "SPR Vitamin C 500", price: 750000, amountThreshold: 1000000, quantityA: 15, quantityB: 3 },
    { name: "SPR Milky Calcium", price: 745000, amountThreshold: 3000000, quantityA: 12, quantityB: 0 },
    { name: "SPR Kids Fish Oil", price: 745000, amountThreshold: 1200000, quantityA: 6, quantityB: 1 },
  ];
}
