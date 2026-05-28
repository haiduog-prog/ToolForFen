# 📋 Masterplan Sales Plan — Kiến thức tổng hợp

Tài liệu này tổng hợp toàn bộ kiến trúc, logic nghiệp vụ và quy tắc kỹ thuật của module **Masterplan Sales Plan Export** — công cụ tạo file kế hoạch bán hàng Excel từ dữ liệu khách hàng đầu vào.

---

## 🏗️ 1. Kiến trúc tổng quan

### Luồng dữ liệu chính:
```
File Excel khách hàng (input)
    ↓ parseCustomerExcel()
SalesPlanCustomer[] (domain)
    ↓ buildMasterplan()
MasterplanData (domain)
    ↓ writeMasterplanWorkbook()
File Excel output (Blob → download)
```

### Các file chính:
| Tầng | File | Chức năng |
|---|---|---|
| Domain | [`masterplan.ts`](file:///d:/Work/ToolsForFen/src/domain/masterplan.ts) | Interface, type definitions, timeline block generation, channel options |
| Application | [`buildMasterplan.ts`](file:///d:/Work/ToolsForFen/src/application/buildMasterplan.ts) | Parse Excel input, build masterplan data, thêm Ecom customer |
| Infrastructure | [`writeMasterplanWorkbook.ts`](file:///d:/Work/ToolsForFen/src/infrastructure/excel/writeMasterplanWorkbook.ts) | Ghi file Excel output với ExcelJS |
| Presentation | [`MasterplanPanel.tsx`](file:///d:/Work/ToolsForFen/src/presentation/MasterplanPanel.tsx) | UI React: upload file, quản lý sản phẩm, cấu hình Ecom, export |

---

## 📊 2. Cấu trúc Domain Models

### SalesPlanProduct
```typescript
interface SalesPlanProduct {
  name: string;           // Tên sản phẩm
  price: number;          // Giá bán (VND)
  amountThreshold: number; // Ngưỡng tổng tiền đã mua
  quantityA: number;      // SL khi vượt ngưỡng
  quantityB: number;      // SL khi dưới/bằng ngưỡng
  ecomInitialQty?: number; // SL ban đầu Ecom (optional)
}
```

### SalesPlanCustomer
```typescript
interface SalesPlanCustomer {
  codeChannel: number;     // Mã kênh (1001-1007)
  channel: string;         // Tên kênh (MTC, Ecom, Baby & Mom, OTC, ETC)
  subChannel: string;      // Kênh phụ
  customer: string;        // Tên khách hàng
  stores: string;          // Số cửa hàng
  province: string;
  regional: string;
  subChannel1: string;
  staff: string;
  customerTotalAmount: number; // Tổng tiền đã mua
  isEcomSpecial?: boolean;     // Đánh dấu khách Ecom đặc biệt
}
```

### CHANNEL_OPTIONS (Các kênh bán hàng)
| Code | Channel | Sub-Channel |
|------|---------|-------------|
| 1001 | MTC | LC, PMC, AK |
| 1002 | MTC | Health & Beauty |
| 1003 | Ecom | Ecommerce |
| 1004 | Baby & Mom | Baby & Mom |
| 1005 | OTC | Key/Chain |
| 1006 | ETC | ETC |
| 1007 | OTC | IDP |

---

## ⏰ 3. Timeline Blocks — Cơ chế tạo block thời gian

### Quy tắc khởi tạo:
- **Tháng bắt đầu** = tháng hiện tại + 1 (nếu tháng 12 → tháng 1 năm sau)
- **Tháng kết thúc** = tháng 12 cùng năm
- Mỗi block type có cấu trúc cột khác nhau

### 3 loại TimelineBlock:

#### Month Block
- Cột sản phẩm (N cột) + Total Volume (1 cột) + Gross Sales +VAT (1 cột)
- **Total = N + 2 cột**

#### Quarter Block
- Cột sản phẩm (N cột) + Quarter Total Volume (1 cột)
- **KHÔNG có Gross Sales +VAT**
- **Total = N + 1 cột**
- Quarter được tạo **ngay cả khi không đủ 3 tháng** (ví dụ Q2 chỉ có JUN)

#### Year Block
- Cột sản phẩm (N cột) + Total Volume (1 cột) + Gross Sales +VAT (1 cột)
- **Total = N + 2 cột**

### Ví dụ layout (2 sản phẩm, bắt đầu JUN):
```
J-L: JUN block (J,K = products, L = total, M = gross)
N-P: Q2 block (N,O = products, P = total)
Q-S: JUL block ...
T-V: AUG block ...
W-Y: SEP block ...
Z-AB: Q3 block ...
... tiếp tục đến DEC + Q4 + YEAR
```

---

## 🔢 4. Công thức Excel — Chi tiết cho mỗi loại dòng

### 4.1. Dòng khách hàng thường (customer row)

#### Month Block:
- **Product quantity**: Dựa trên `customerTotalAmount` so với `amountThreshold`
  - Nếu `totalAmount > threshold` → dùng `quantityA`
  - Nếu `totalAmount <= threshold` → dùng `quantityB`
- **Total Volume**: `=SUM(firstProduct:lastProduct)` trong block tháng
- **Gross Sales +VAT**: `=SUMPRODUCT($firstProd$priceRow:$lastProd$priceRow, firstProd{row}:lastProd{row})`
  - Giá nằm ở dòng `timelineHeadersStartRow + 2` (row 17 dynamic)
  - Dùng **absolute row** cho giá: `$J$17:$K$17`

#### Quarter Block:
- **Product columns**: Tổng các cột product tương ứng của các tháng trong quý
  - Ví dụ: `=J18+Q18+T18` (tổng 3 tháng JUL+AUG+SEP cho product 1)
- **Quarter Total Volume**: `=SUM(firstQProd:lastQProd)` trong block quý

#### Year Block:
- **Product columns**: Tổng các cột product tương ứng của các quý
  - Ví dụ: `=N18+Z18+AL18` (tổng Q2+Q3+Q4 cho product 1)
- **Total Volume**: `=SUM(firstYProd:lastYProd)` trong block năm
- **Gross Sales +VAT**: `=SUMPRODUCT(prices, yearQuantities)`

### 4.2. Dòng TOTAL kênh (channel group total)
- Tự động chèn sau nhóm khách hàng cùng kênh
- Mọi cột numeric: `=SUM(colLetter{groupStart}:colLetter{groupEnd})`
- Cột D: Label dạng `TOTAL E-COMMERCE`, `TOTAL BABY & MOM`, `TOTAL OTC`, `TOTAL MTC`, `TOTAL ETC`
- **Styling**: Fill xanh nhạt (`FFDDEBF7`), font bold, border đầy đủ

### 4.3. Dòng Summary kênh (rows 5-11 dynamic)
- Dùng công thức **SUMIF** để tổng theo mã kênh:
  ```
  =SUMIF($A${dataStart}:$A${dataEnd}, $A{row}, {col}${dataStart}:{col}${dataEnd})
  ```
- Áp dụng cho tất cả cột product, total volume, gross sales

### 4.4. Dòng Nationwide
- `=SUM(channelsStartRow:lastChannelRow)` cho mỗi cột numeric

---

## 🛒 5. Kênh Ecom — Logic đặc biệt

### Cơ chế kích hoạt:
- Toggle checkbox "Thêm kênh Ecom" trên UI (`ecomEnabled` state)
- Khi enabled, tự động inject 1 customer đặc biệt:
  ```typescript
  {
    codeChannel: 1003,
    channel: "Ecom",
    subChannel: "Ecommerce",
    customer: "Kênh Ecom",
    isEcomSpecial: true, // flag đánh dấu
    customerTotalAmount: 0, // không dùng threshold
  }
  ```

### Công thức tăng trưởng Ecom:
- **Tháng đầu tiên**: Dùng `ecomInitialQty` của mỗi sản phẩm (giá trị cố định)
- **Các tháng tiếp theo**: Tăng 20% so với tháng trước
  ```
  =prevMonthCol{row} * 1.2
  ```
- Ví dụ: Nếu tháng 6 = 100 → tháng 7 = `=J13*1.2` = 120 → tháng 8 = `=Q13*1.2` = 144

### UI cho Ecom:
- Nằm trong section riêng **"Cấu hình Kênh Ecom"** trên sidebar
- KHÔNG nằm chung với form thêm sản phẩm
- Khi enabled, hiển thị thêm cột "SL đầu Ecom" trong bảng sản phẩm

---

## 📁 6. Đọc file khách hàng (Input Parsing)

### 2 chế độ đọc file:

#### Mode "normalized" (Danh sách khách hàng chuẩn):
- Cần tối thiểu 3 cột: `Code channel`, `Channel`, `Customer`
- Cột tùy chọn: `Sub-channel`, `Stores`, `Province`, `Regional`, `Sub-Channel 1`, `Staff`, `Tổng tiền`
- Tự động nhận diện cột tổng tiền qua từ khóa: `tong tien`, `doanh thu`, `revenue`, `amount`, `sales`...

#### Mode "sales_detail" (Báo cáo bán hàng chi tiết):
- Cần 4 cột: `Khách hàng`, `Doanh thu chi tiết`, `Nhân viên`, `Mã hóa đơn`
- Tự động group theo khách hàng, tổng doanh thu
- Channel code lấy từ dropdown trên UI (mặc định 1004 - Baby & Mom)

### Heuristic Header Matching:
- Quét 25 dòng đầu tìm dòng tiêu đề
- Chuẩn hóa text: NFD, bỏ dấu, lowercase, trim
- So khớp bằng keyword list (hỗ trợ cả tiếng Việt và tiếng Anh)

---

## 📤 7. Cấu trúc file Excel Output

### Sheet "DetailPlan_Template" — Layout các dòng:
```
Row 1:  Title bar (merged, navy blue)
Row 2:  Subtitle / Rules
Row 3:  Fixed headers (A-I: Code channel → STAFF)
Row 4:  Spacer
Row 5+: Summary rows cho từng kênh (SUMIF formulas)
Row N:  Nationwide row (SUM of channel summaries)
Row N+1, N+2: Spacer
Row N+3: Timeline block labels (merged, color-coded)
Row N+4: Product names + Total Volume + Gross Sales headers
Row N+5: Product prices (yellow fill)
Row N+6+: Customer data rows (grouped by channel)
         → TOTAL row sau mỗi nhóm kênh (blue fill)
```

> **Lưu ý quan trọng**: Các row number (channelsStartRow, nationwideRow, timelineHeadersStartRow, dataStartRow) là **DYNAMIC** — phụ thuộc vào số lượng kênh unique trong dữ liệu khách hàng.

### Sheet "Products":
- Bảng tham chiếu sản phẩm: Name, Price, Amount Threshold, Quantity A, Quantity B

### Sheet "Mapping":
- Metadata mô tả cấu trúc file cho việc debug và tích hợp

---

## 🎨 8. Bảng màu và Styling

| Tên | ARGB | Mục đích |
|-----|------|----------|
| TITLE_FILL | `FF1F4E78` | Dòng title, label tháng (navy blue, chữ trắng) |
| SUBHEADER_FILL | `FFE7E6E6` | Headers cố định, summary rows |
| YELLOW_TOTAL_FILL | `FFFFF2CC` | Dòng giá sản phẩm (row 17) |
| GREEN_TOTAL_FILL | `FFD9EAD3` | Cột Total Volume, Gross Sales header |
| PURPLE_QUARTER_FILL | `FF6A329F` | Label block quý (chữ trắng) |
| TEAL_YEAR_FILL | `FF0F6B5F` | Label block năm (chữ trắng) |
| BLUE_TOTAL_FILL | `FFDDEBF7` | Dòng TOTAL kênh |

### Styling rules:
- Font: Calibri, size 11
- Border: Thin, color `FFD9D9D9`
- Number format tiền: `#,##0`
- Freeze pane: Cố định 9 cột đầu (A-I) và header rows

---

## 🐛 9. Vấn đề đã biết (Known Issues)

### Giá trị = 32 xuất hiện rải rác trong output
- **Triệu chứng**: Nhiều ô trong file Excel xuất ra có giá trị `32` không mong muốn, đặc biệt ở các cột Province, Regional, Staff của dòng TOTAL kênh (row 14 trong ảnh) và các dòng Ecom.
- **Nguyên nhân tiềm năng**: Có thể liên quan đến việc `row16.height = 32` bị ExcelJS xử lý sai, hoặc giá trị `" "` (space, ASCII 32) bị convert thành số khi ghi ô.
- **Trạng thái**: Chưa fix, cần debug thêm bằng script [`debug_32.ts`](file:///d:/Work/ToolsForFen/debug_32.ts).

---

## 🧪 10. Quy tắc kỹ thuật quan trọng

1. **Không có sản phẩm mặc định**: Products bắt đầu là mảng rỗng `[]`, người dùng tự thêm qua form.
2. **Không có khách hàng mặc định**: Customers bắt đầu rỗng `[]`, phải upload file Excel.
3. **Kênh bán hàng lấy từ file input**: Không fix sẵn — extract unique channels từ danh sách khách hàng.
4. **Ecom là luồng riêng**: Toggle + quantity inputs tách biệt với form thêm sản phẩm.
5. **Dynamic row calculation**: `channelsStartRow = 5`, `nationwideRow = channelsStartRow + numChannels`, các row khác tính tiếp theo.
6. **SUMIF cho summary rows**: Dùng SUMIF match `codeChannel` (cột A) thay vì SUM range cố định.
7. **Sorted & grouped customers**: Khách hàng được sort theo `codeChannel` → `channel` → `subChannel`, group liên tiếp, chèn TOTAL row sau mỗi group.
8. **Column letter helper**: Hàm `getColLetter(colNumber)` chuyển số cột 1-based thành letter (1→A, 27→AA).

---

## 🔧 11. UI States trong MasterplanPanel

| State | Type | Mô tả |
|-------|------|-------|
| `products` | `SalesPlanProduct[]` | Danh sách sản phẩm (user-defined) |
| `customers` | `SalesPlanCustomer[]` | Khách hàng từ file input |
| `ecomEnabled` | `boolean` | Toggle kênh Ecom |
| `fileType` | `"normalized" \| "sales_detail"` | Chế độ đọc file |
| `targetChannelCode` | `number` | Mã kênh mặc định (cho sales_detail mode) |
| `excelHeaders` | `string[]` | Headers đọc từ file (cho chọn cột tổng tiền) |
| `selectedTotalAmountColIdx` | `number` | Index cột tổng tiền đã chọn |
| `editingIndex` | `number \| null` | Index sản phẩm đang inline edit |

---

## 📝 12. Format tiền tệ trên UI

```typescript
// Format hiển thị: 1,500,000
function formatMoneyInput(value: string): string {
  const numeric = value.replace(/\D/g, "");
  if (!numeric) return "";
  return Number(numeric).toLocaleString("en-US");
}

// Parse từ hiển thị về số: "1,500,000" → 1500000
function parseMoneyInput(value: string): number {
  const numeric = value.replace(/\D/g, "");
  return parseFloat(numeric) || 0;
}
```

Áp dụng cho: Price, Amount Threshold khi nhập và hiển thị trên form sản phẩm.
