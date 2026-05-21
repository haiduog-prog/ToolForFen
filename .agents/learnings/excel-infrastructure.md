# 📊 Excel Infrastructure (Read & Write Operations)

Tài liệu này chi tiết cơ chế hoạt động của tầng cơ sở hạ tầng (Infrastructure) xử lý tệp Excel nhằm giúp AI hiểu cách đọc và ghi dữ liệu ExcelJS chuẩn xác.

---

## 🔎 1. Cơ chế đọc và tự động nhận diện cột (Heuristic Column Matching)
Nằm trong tệp [`readSourceWorkbook.ts`](file:///d:/Work/ToolsForFen/src/infrastructure/excel/readSourceWorkbook.ts).

### Quy trình nhận diện:
1.  **Chuẩn hóa văn bản (`normalizeHeader`)**:
    Dọn dẹp chuỗi bằng cách tách dấu tiếng Việt (NFD Unicode), loại bỏ dấu (`replace(/[\u0300-\u036f]/g, "")`), giảm nhiều khoảng trắng thành 1, trim và chuyển về chữ thường.
2.  **Quét dòng tiêu đề (Header row detection)**:
    Duyệt qua 20 dòng đầu tiên của worksheet làm ứng viên. Đọc tất cả các ô trong dòng đó qua bộ lọc chuẩn hóa.
3.  **So khớp từ khóa**:
    Nhận diện dòng tiêu đề chính xác khi chứa đủ 4 cột bắt buộc:
    *   **Nhà Thuốc / Khách hàng**: chứa `nha thuoc` hoặc `customer`
    *   **Tên sản phẩm**: chứa `ten san pham` hoặc `product`
    *   **Số lượng**: chứa `so luong` hoặc `quantity`
    *   **Doanh thu**: chứa `doanh thu` hoặc `revenue`
4.  **Các cột tùy chọn được map thêm**:
    *   `Đơn giá` (`don gia` hoặc `unit price`)
    *   `Thành tiền` (`thanh tien` hoặc `amount`)
    *   `Segment` (`segment` hoặc `phan khuc` hoặc `kenh`)
    *   `DSR` (`nguoi thuc hien` hoặc `dsr` hoặc `nhan vien`)
    *   `Ngày` (`ngay` hoặc `date`)
5.  **Cột thời gian (Month Column Detection)**:
    Lọc ra các cột có tên chứa `ngay`, `month`, hoặc `thang`. Nếu không có, quét toàn bộ cột.
    Với mỗi cột, hệ thống quét 50 dòng tiếp theo để thử phân tích giá trị tháng (`parseMonthFromCell`). Cột nào trả về số lượng tháng hợp lệ nhiều nhất sẽ được chọn làm **Month Column**.

---

## 📥 2. Cơ chế ghi và định dạng Excel (ExcelJS Write Operations)
Nằm trong tệp [`writeReportWorkbook.ts`](file:///d:/Work/ToolsForFen/src/infrastructure/excel/writeReportWorkbook.ts).

Ứng dụng xuất ra một file Excel có cấu trúc 5 worksheets. Các kỹ thuật cấu trúc và thiết kế bao gồm:

### Cấu trúc 5 Sheets:
1.  **`QEC review`**: Bảng tổng hợp các Kênh/Segment và DSR.
    *   *Đóng băng (Freeze)*: Cố định dòng 1 (tiêu đề) và cột 1 (Segment).
    *   Chứa bảng Segment ở trên và bảng DSR xếp chồng ở dưới, cách nhau 1 dòng trống.
2.  **`SKU review`**: Top 50 sản phẩm đóng góp doanh số cao nhất.
    *   *Đóng băng (Freeze)*: Cố định 2 dòng đầu (Dòng 1 là index cột dạng số `1, 2, 3...` giống mẫu `chuan.xlsx`, Dòng 2 là tiêu đề cột) và cố định 2 cột đầu (Code, Brand).
3.  **`SKU - Customer review`**: Chi tiết doanh thu của từng SKU theo từng nhà thuốc.
    *   *Cấu trúc Section*: Nhóm theo Khách hàng. Hàng tên khách hàng sẽ được merge hết tất cả cột và áp dụng `SECTION_FILL` (màu xanh nhạt `#FFE8F1F2`). Hàng tổng của khách hàng áp dụng `TOTAL_FILL` (màu vàng nhạt `#FFFFF3CD`).
    *   *Đóng băng (Freeze)*: Cố định dòng 1 và 2 cột đầu (Customer, Name SKU).
4.  **`SKU customer review`**: Chi tiết số lượng SKU tiêu thụ của từng nhà thuốc (cấu trúc tương tự `SKU - Customer review`).
5.  **`Data nguồn`**: Bản sao của toàn bộ dữ liệu giao dịch thô đã lọc.

### Bảng màu Corporate (Corporate Theme Color Palette):
*   **Header Fill**: `#FF183B56` (Màu xanh dương đậm Navy, chữ trắng bold, căn giữa).
*   **Section Merged Fill**: `#FFE8F1F2` (Màu xanh pastel nhạt, chữ xanh Navy bold).
*   **Total Row Fill**: `#FFFFF3CD` (Màu vàng nhạt, chữ bold).
*   **Borders**: Viền mảnh (`thin`) màu `#FFD8DEE4`.

### Định dạng số nguyên thủy (`numFmt`):
Để đảm bảo người dùng có thể thực hiện phép tính (Sum, Average) trên file Excel xuất ra, mọi giá trị số phải được ghi dưới dạng kiểu `Number` và áp dụng định dạng hiển thị:
*   **Định dạng tiền tệ (`MONEY_FORMAT`)**: `#,##0;[Red]-#,##0;"-"` (Phân cách hàng ngàn, số âm màu đỏ, số 0 hiển thị dấu gạch ngang `-`).
*   **Định dạng số lượng (`QUANTITY_FORMAT`)**: `#,##0.00;[Red]-#,##0.00;"-"` (Hiển thị 2 chữ số thập phân).
*   **Định dạng tỷ lệ / phần trăm (`RATIO_FORMAT`)**: `0.00%;[Red]-0.00%;"-"`.

### Tiện ích căn chỉnh tự động (`autoFit`):
Hàm tự động quét qua toàn bộ các ô trong từng cột, đo lường độ dài ký tự thực tế (đối với ô chứa đối tượng RichText hoặc Date sẽ có xử lý riêng) để thiết lập độ rộng cột `column.width` động (tối thiểu `10`, tối đa `48`), tránh tình trạng chữ bị khuất hoặc xuất hiện lỗi `###` trên Excel.
