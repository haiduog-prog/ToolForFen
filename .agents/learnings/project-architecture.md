# 🏗️ Clean Architecture & DDD in QEC Export Builder

Tài liệu này giúp AI hiểu sâu sắc về kiến trúc phân tầng của dự án và các quy tắc phụ thuộc nghiêm ngặt giữa chúng.

## 📐 Nguyên lý Dependency Rule (Quy tắc phụ thuộc)
Trong dự án này, luồng phụ thuộc đi từ ngoài vào trong. Tầng bên trong **tuyệt đối không** được biết hoặc import bất kỳ thành phần nào từ tầng bên ngoài.

```
Presentation (React UI, App.tsx)
    │
    ▼
Application (Use Cases, buildQecReport.ts)
    │
    ▼
Domain (Entities, Month, Calculations) ◄── Infrastructure (Excel, exceljs)
```

---

## 📂 Chi tiết các tầng thư mục

### 1. Tầng Nghiệp vụ - Domain Layer (`src/domain/`)
*   **Mục tiêu**: Nơi chứa toàn bộ tri thức nghiệp vụ cốt lõi của doanh nghiệp. Tầng này phải hoàn toàn "sạch" (Pure TypeScript), không phụ thuộc vào React, DOM, Network, hay bất kỳ thư viện bên ngoài nào (ngoại trừ các helper nội bộ).
*   **Các file chính**:
    *   [`entities.ts`](file:///d:/Work/ToolsForFen/src/domain/entities.ts): Định nghĩa các interface và kiểu dữ liệu nghiệp vụ:
        *   `SourceTransaction`: Cấu trúc một giao dịch thô đọc từ file Excel.
        *   `MetricRow`: Cấu trúc một hàng dữ liệu phân tích chi tiết của một Segment/SKU/Khách hàng, chứa doanh số/số lượng của từng tháng trong chuỗi thời gian cùng các chỉ số tổng hợp.
        *   `QecReport`: Báo cáo hoàn chỉnh gồm các mảng `MetricRow` cho QEC, SKU, Customer Revenue và Customer Quantity.
    *   [`month.ts`](file:///d:/Work/ToolsForFen/src/domain/month.ts): Xử lý toàn bộ logic chuỗi thời gian bán hàng.
        *   Định nghĩa `MonthKey` dưới dạng chuỗi `YYYY-MM` (ví dụ: `2025-04`).
        *   Các hàm tạo dải tháng báo cáo, xác định tháng năm trước cùng kỳ, tính toán dải trung bình trượt (Trailing months).
    *   [`reportCalculations.ts`](file:///d:/Work/ToolsForFen/src/domain/reportCalculations.ts): Thuật toán tính toán chỉ số (Trend, Share, YTD, IYA). Chứa hàm `safeRatio` để phòng tránh lỗi chia cho 0.
    *   [`reportCalculations.test.ts`](file:///d:/Work/ToolsForFen/src/domain/reportCalculations.test.ts): Vitest unit tests bao phủ toàn bộ công thức toán học.

### 2. Tầng Ứng dụng - Application Layer (`src/application/`)
*   **Mục tiêu**: Điều phối luồng xử lý (Use Case). Nhận dữ liệu đầu vào, gọi các Domain service để tính toán và trả về kết quả cấu trúc hoàn chỉnh cho UI.
*   **Các file chính**:
    *   [`buildQecReport.ts`](file:///d:/Work/ToolsForFen/src/application/buildQecReport.ts): Nhận dữ liệu giao dịch thô đã phân tích từ Infrastructure, cùng với tháng báo cáo do người dùng chọn. Thực hiện các bước:
        1.  Tính toán dải tháng cần báo cáo (từ tháng 1 năm trước đến tháng báo cáo hiện tại).
        2.  Lọc các giao dịch nằm trong dải tháng này.
        3.  Gọi các hàm tổng hợp của Domain để tạo dữ liệu cho 4 bảng: QEC (nhóm theo Segment), SKU (nhóm theo SKU, sắp xếp giảm dần), Customer Revenue (nhóm theo Khách hàng + SKU, tính doanh thu), Customer Quantity (nhóm theo Khách hàng + SKU, tính số lượng).
        4.  Thống kê tổng quan và lọc ra các giao dịch có thông tin chưa map (`Unmapped`).

### 3. Tầng Cơ sở hạ tầng - Infrastructure Layer (`src/infrastructure/`)
*   **Mục tiêu**: Giao tiếp trực tiếp với môi trường ngoài (file Excel). Tầng này sử dụng thư viện `exceljs` để đọc và ghi dữ liệu.
*   **Các file chính**:
    *   [`excelCellUtils.ts`](file:///d:/Work/ToolsForFen/src/infrastructure/excel/excelCellUtils.ts): Các tiện ích xử lý ô Excel (trích xuất giá trị số từ ô chứa công thức, RichText, Object) và hàm chuẩn hóa chuỗi phục vụ khớp cột động.
    *   [`readSourceWorkbook.ts`](file:///d:/Work/ToolsForFen/src/infrastructure/excel/readSourceWorkbook.ts): Đọc file Excel nguồn được tải lên, sử dụng Heuristic Header Matching để tự động map cột.
    *   [`writeReportWorkbook.ts`](file:///d:/Work/ToolsForFen/src/infrastructure/excel/writeReportWorkbook.ts): Tạo workbook mới, thiết kế styles, merge cells, định dạng số (`numFmt`), đóng băng cột và hàng tiêu đề, rồi trả về Blob để tải xuống.

### 4. Tầng Giao diện - Presentation Layer (`src/presentation/`)
*   **Mục tiêu**: Hiển thị UI và xử lý các tương tác của người dùng.
*   **Các file chính**:
    *   [`App.tsx`](file:///d:/Work/ToolsForFen/src/presentation/App.tsx): State holder chính của ứng dụng. Quản lý trạng thái upload, tính toán báo cáo, chuyển đổi tabs xem trước, hiển thị các thông báo/cảnh báo dữ liệu và kích hoạt download file.
    *   [`styles.css`](file:///d:/Work/ToolsForFen/src/presentation/styles.css): Định nghĩa toàn bộ design system bằng CSS Variables, layout Responsive Flexbox/Grid, hiệu ứng micro-animations, glassmorphism cao cấp.

### 5. Tầng Dùng chung - Shared Layer (`src/shared/`)
*   **Mục tiêu**: Chứa các utility helpers đơn giản dùng chung ở mọi tầng.
*   **Các file chính**:
    *   [`formatters.ts`](file:///d:/Work/ToolsForFen/src/shared/formatters.ts): Helper định dạng tiền tệ, tỷ lệ phần trăm hiển thị trên UI và hàm trigger tải file Blob phía Client.
