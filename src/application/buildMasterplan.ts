import * as xlsx from "xlsx";
import {
  generateTimelineBlocks,
  getInitialProducts,
  type MasterplanData,
  type SalesPlanProduct,
  type SalesPlanCustomer,
  CHANNEL_OPTIONS,
} from "../domain/masterplan";

export function normalizeHeader(value: unknown): string {
  if (value == null) return "";
  return String(value)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/[^a-z0-9]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Parses the uploaded customer list Excel file using SheetJS and flexible heuristic column mapping.
 */
export async function parseCustomerExcel(
  file: File,
  optionsOrIdx?: number | {
    mode?: "normalized" | "sales_detail";
    totalAmountColIdx?: number;
    channelCode?: number;
  }
): Promise<SalesPlanCustomer[]> {
  const arrayBuffer = await file.arrayBuffer();
  const data = new Uint8Array(arrayBuffer);
  const workbook = xlsx.read(data, { type: "array" });

  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error("File Excel rỗng hoặc không hợp lệ.");
  }

  const worksheet = workbook.Sheets[sheetName]!;
  const rows = xlsx.utils.sheet_to_json<any[]>(worksheet, { header: 1, defval: "" });

  let mode: "normalized" | "sales_detail" = "normalized";
  let totalAmountColIdx: number | undefined;
  let channelCode: number = 1004;

  if (typeof optionsOrIdx === "number") {
    totalAmountColIdx = optionsOrIdx;
  } else if (optionsOrIdx && typeof optionsOrIdx === "object") {
    mode = optionsOrIdx.mode ?? "normalized";
    totalAmountColIdx = optionsOrIdx.totalAmountColIdx;
    channelCode = optionsOrIdx.channelCode ?? 1004;
  }

  // 1. Detect header row by looking for unique customer keywords
  let headerRowIdx = -1;
  let custCol = -1;
  let codeCol = -1;
  let chanCol = -1;
  let subCol = -1;
  let storeCol = -1;
  let provCol = -1;
  let regCol = -1;
  let sub1Col = -1;
  let staffCol = -1;
  let totalAmountCol = -1;

  let detailRevCol = -1;
  let invoiceCol = -1;

  const maxHeaderSearch = Math.min(25, rows.length);
  for (let r = 0; r < maxHeaderSearch; r++) {
    const row = rows[r];
    if (!row) continue;

    const normalized = row.map((cell) => normalizeHeader(cell));

    if (mode === "sales_detail") {
      custCol = findColumnIndex(normalized, ["khach hang"]);
      detailRevCol = findColumnIndex(normalized, ["doanh thu chi tiet"]);
      staffCol = findColumnIndex(normalized, ["nhan vien"]);
      invoiceCol = findColumnIndex(normalized, ["ma hoa don"]);

      if (custCol !== -1 && detailRevCol !== -1 && staffCol !== -1 && invoiceCol !== -1) {
        headerRowIdx = r;
        break;
      }
    } else {
      custCol = findColumnIndex(normalized, ["customer", "khach hang", "ten nha thuoc", "nha thuoc"]);
      codeCol = findColumnIndex(normalized, ["code channel", "ma kenh", "channel code", "ma"]);
      chanCol = findColumnIndex(normalized, ["channel", "kenh", "kenh ban hang"]);

      if (custCol !== -1 && codeCol !== -1 && chanCol !== -1) {
        headerRowIdx = r;
        subCol = findColumnIndex(normalized, ["sub channel", "kenh phu", "phan kenh"]);
        storeCol = findColumnIndex(normalized, ["stores", "cua hang", "so cua hang", "store"]);
        provCol = findColumnIndex(normalized, ["province", "tinh", "tinh thanh", "dia phuong"]);
        regCol = findColumnIndex(normalized, ["regional", "vung", "khu vuc", "region"]);
        sub1Col = findColumnIndex(normalized, ["sub channel 1", "subchannel 1", "sub channel1"]);
        staffCol = findColumnIndex(normalized, ["staff", "nhan vien", "nguoi thuc hien", "dsr", "staff name"]);
        
        totalAmountCol = findColumnIndex(normalized, [
          "tong tien", "doanh thu", "revenue", "amount", "tong mua", "da mua", "sales", 
          "customer total amount", "total sales", "total amount", "tong gia tri", "thanh tien"
        ]);
        break;
      }
    }
  }

  if (headerRowIdx === -1) {
    if (mode === "sales_detail") {
      throw new Error(
        "Không nhận diện được tiêu đề bảng doanh số bán hàng chi tiết. Cần có ít nhất các cột: Khách hàng, Doanh thu chi tiết, Nhân viên và Mã hóa đơn."
      );
    } else {
      throw new Error(
        "Không nhận diện được tiêu đề bảng khách hàng. Cần có ít nhất các cột: Code channel, Channel, và Customer."
      );
    }
  }

  const customers: SalesPlanCustomer[] = [];

  if (mode === "sales_detail") {
    // Group rows by customer
    const groups: Record<string, {
      revenue: number;
      staffCounts: Record<string, number>;
    }> = {};

    for (let i = headerRowIdx + 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue;

      const hasValue = row.some((cell) => cell !== null && cell !== undefined && String(cell).trim() !== "");
      if (!hasValue) continue;

      const customer = String(row[custCol] || "").trim();
      if (!customer) continue;

      const revStr = String(row[detailRevCol] || "");
      const rev = parseFloat(revStr.replace(/[^0-9.-]/g, "")) || 0;
      const staff = String(row[staffCol] || "").trim();

      if (!groups[customer]) {
        groups[customer] = { revenue: 0, staffCounts: {} };
      }
      groups[customer].revenue += rev;
      if (staff) {
        groups[customer].staffCounts[staff] = (groups[customer].staffCounts[staff] || 0) + 1;
      }
    }

    const channelOpt = CHANNEL_OPTIONS.find((c) => c.code === channelCode) ?? CHANNEL_OPTIONS[3]!; // default 1004

    for (const customer in groups) {
      const grp = groups[customer]!;
      
      // Get most frequent staff
      let maxCount = -1;
      let mostFrequentStaff = "";
      for (const staff in grp.staffCounts) {
        if (grp.staffCounts[staff] > maxCount) {
          maxCount = grp.staffCounts[staff];
          mostFrequentStaff = staff;
        }
      }

      customers.push({
        codeChannel: channelCode,
        channel: channelOpt.channel,
        subChannel: channelOpt.subChannel,
        customer: customer,
        stores: "1",
        province: "",
        regional: "",
        subChannel1: channelOpt.channel === "Baby & Mom" ? "B&M" : channelOpt.channel,
        staff: "",
        customerTotalAmount: grp.revenue,
      });
    }

  } else {
    // Normalized Customer List
    if (totalAmountColIdx !== undefined && totalAmountColIdx !== -1) {
      totalAmountCol = totalAmountColIdx;
    }

    for (let i = headerRowIdx + 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue;

      const hasValue = row.some((cell) => cell !== null && cell !== undefined && String(cell).trim() !== "");
      if (!hasValue) continue;

      const codeChannel = parseInt(row[codeCol]) || 1001;
      const channel = String(row[chanCol] || "").trim();
      const subChannel = subCol !== -1 ? String(row[subCol] || "").trim() : "";
      const customer = String(row[custCol] || "").trim();
      const stores = storeCol !== -1 ? String(row[storeCol] || "").trim() : "1";
      const province = provCol !== -1 ? String(row[provCol] || "").trim() : "";
      const regional = regCol !== -1 ? String(row[regCol] || "").trim() : "";
      const subChannel1 = sub1Col !== -1 ? String(row[sub1Col] || "").trim() : "";
      const staff = staffCol !== -1 ? String(row[staffCol] || "").trim() : "";

      let customerTotalAmount = 0;
      if (totalAmountCol !== -1 && row[totalAmountCol] !== undefined && row[totalAmountCol] !== "") {
        const parsedVal = parseFloat(String(row[totalAmountCol]).replace(/[^0-9.-]/g, ""));
        customerTotalAmount = isNaN(parsedVal) ? 0 : parsedVal;
      }

      if (customer) {
        customers.push({
          codeChannel,
          channel,
          subChannel,
          customer,
          stores,
          province,
          regional,
          subChannel1,
          staff,
          customerTotalAmount,
        });
      }
    }
  }

  return customers;
}

export async function extractExcelHeaders(file: File): Promise<{
  headers: string[];
  matchedIdx: number;
  mode: "normalized" | "sales_detail";
}> {
  const arrayBuffer = await file.arrayBuffer();
  const data = new Uint8Array(arrayBuffer);
  const workbook = xlsx.read(data, { type: "array" });

  const sheetName = workbook.SheetNames[0];
  if (!sheetName) return { headers: [], matchedIdx: -1, mode: "normalized" };

  const worksheet = workbook.Sheets[sheetName]!;
  const rows = xlsx.utils.sheet_to_json<any[]>(worksheet, { header: 1, defval: "" });

  const maxHeaderSearch = Math.min(25, rows.length);
  for (let r = 0; r < maxHeaderSearch; r++) {
    const row = rows[r];
    if (!row) continue;

    const normalized = row.map((cell) => normalizeHeader(cell));

    // 1. Check if sales_detail
    const hasCust = findColumnIndex(normalized, ["khach hang"]) !== -1;
    const hasDetailRev = findColumnIndex(normalized, ["doanh thu chi tiet"]) !== -1;
    const hasStaff = findColumnIndex(normalized, ["nhan vien"]) !== -1;
    const hasInvoice = findColumnIndex(normalized, ["ma hoa don"]) !== -1;

    if (hasCust && hasDetailRev && hasStaff && hasInvoice) {
      return {
        headers: row.map((cell) => String(cell || "").trim()),
        matchedIdx: -1,
        mode: "sales_detail",
      };
    }

    // 2. Check if normalized
    const custCol = findColumnIndex(normalized, ["customer", "khach hang", "ten nha thuoc", "nha thuoc"]);
    const codeCol = findColumnIndex(normalized, ["code channel", "ma kenh", "channel code", "ma"]);
    const chanCol = findColumnIndex(normalized, ["channel", "kenh", "kenh ban hang"]);

    if (custCol !== -1 && codeCol !== -1 && chanCol !== -1) {
      const originalHeaders = row.map((cell) => String(cell || "").trim());
      const totalAmountCol = findColumnIndex(normalized, [
        "tong tien", "doanh thu", "revenue", "amount", "tong mua", "da mua", "sales", 
        "customer total amount", "total sales", "total amount", "tong gia tri", "thanh tien"
      ]);
      return {
        headers: originalHeaders,
        matchedIdx: totalAmountCol,
        mode: "normalized",
      };
    }
  }

  return { headers: [], matchedIdx: -1, mode: "normalized" };
}

function findColumnIndex(headers: string[], keys: string[]): number {
  for (const k of keys) {
    const idx = headers.indexOf(k);
    if (idx !== -1) return idx;
  }
  for (const k of keys) {
    const idx = headers.findIndex((h) => h.includes(k));
    if (idx !== -1) return idx;
  }
  return -1;
}

/**
 * Builds the complete Masterplan data.
 * If no custom customers are passed, generates mock customer data with proper channel codes.
 */
export function buildMasterplan(
  customCustomers: SalesPlanCustomer[] | null,
  customProducts?: SalesPlanProduct[],
  ecomEnabled?: boolean
): MasterplanData {
  const products = customProducts ?? [];
  let customers = customCustomers ?? [];

  if (ecomEnabled) {
    const hasEcom = customers.some((c) => c.codeChannel === 1003);
    if (!hasEcom) {
      customers = [
        ...customers,
        {
          codeChannel: 1003,
          channel: "Ecom",
          subChannel: "Ecommerce",
          customer: "Kênh Ecom",
          stores: "1",
          province: "",
          regional: "",
          subChannel1: "E-Commerce",
          staff: "",
          customerTotalAmount: 0,
          isEcomSpecial: true,
        }
      ];
    }
  }

  // Rules: start month index = current month index + 1
  const today = new Date();
  const startMonthIndex = (today.getMonth() + 1) % 12;
  const startYear = today.getMonth() === 11 ? today.getFullYear() + 1 : today.getFullYear();

  const timelineBlocks = generateTimelineBlocks(startYear, startMonthIndex, products.length);

  return {
    products,
    customers,
    timelineBlocks,
    startYear,
    startMonthIndex,
  };
}

