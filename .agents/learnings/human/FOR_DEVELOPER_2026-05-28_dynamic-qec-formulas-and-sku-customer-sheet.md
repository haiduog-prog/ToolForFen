# ☕ Coffee Talk: Giải thích về Hệ thống Báo cáo Động QEC & SKU Customer Review

Chào cậu! Hôm nay chúng ta đã có một phiên làm việc cực kỳ năng suất và thú vị. Chúng ta đã giải quyết xong hai vấn đề cực kỳ hóc búa liên quan đến xuất báo cáo Excel và Web UI Preview. Hãy cùng ngồi lại, làm một ngụm cà phê và điểm qua xem chúng ta đã thay đổi những gì và tại sao lại làm như thế nhé!

---

## 1. Câu chuyện về các cột IFYTD, ICYTD, IYA hiển thị dấu gạch ngang `-`

### Vấn đề thực tế
Cậu mở file Excel xuất ra cho tháng báo cáo `2026-05` (sử dụng tệp dữ liệu thô `input.xlsx` của khách hàng) và thấy các cột `IFYTD`, `ICYTD`, `IYA` đều bị hiển thị là `-` (tương ứng với giá trị 0). 
Khi chúng mình debug sâu vào dữ liệu thô, hóa ra file `input.xlsx` **chỉ có dữ liệu từ tháng 12/2025 đến tháng 05/2026**. Do đó, toàn bộ doanh thu cùng kỳ năm ngoái (tháng 05/2025) hay YTD năm ngoái (tháng 01/2025 đến tháng 05/2025) đều bằng 0 tròn trĩnh!
Vì backend của tụi mình tính toán bằng TypeScript và ghi cứng giá trị tĩnh vào Excel, nên khi mẫu số bằng 0, tỷ lệ chia ra đương nhiên bằng 0 và hiển thị thành `-`.

### Giải pháp "Động" cao cấp
Khách hàng muốn các ô này **phải tự động tính toán**. Thế là tụi mình đã chuyển đổi toàn bộ các cột tính toán của sheet `QEC review` sang **Công thức Excel thực thụ**:
- Tụi mình dùng `=SUM()` cho tổng năm ngoái và năm nay.
- Dùng `=AVERAGE()` cho các cột trung bình trượt `P3M`, `P6M`, `P9M`.
- Dùng `=IFERROR()` bọc ngoài các phép chia tỉ lệ (`TREND`, `IFYTD`, `ICYTD`, `IYA`) để tránh lỗi chia cho 0 (`#DIV/0!`).

Bây giờ, khi khách hàng mở file Excel lên, Excel sẽ tự động recalculate. Nếu sau này họ bổ sung dữ liệu năm ngoái vào sheet Data nguồn, các ô này sẽ lập tức nhảy số chính xác mà không cần xuất lại báo cáo! Cực kỳ chuyên nghiệp đúng không?

---

## 2. Điền số thứ tự cho cột Customer Code

Trong sheet `QEC review`, ở bảng `Customer Detail` dưới cùng, cột đầu tiên mang tên `Customer code` bị bỏ trống hoàn toàn. Tụi mình đã nhanh trí điền số thứ tự tự động tăng dần (`1`, `2`, `3`...) cho từng khách hàng phát sinh doanh số. Điều này giúp bảng báo cáo trông sạch sẽ, dễ theo dõi và người dùng có thể đếm tổng số dòng khách hàng trực tiếp một cách nhanh chóng.

---

## 3. Sheet báo cáo đảo ngược hoàn toàn mới: SKU review - Customer

Đây là tính năng lớn nhất hôm nay! Khách hàng muốn có một sheet báo cáo giống hệt "SKU - Customer review" nhưng có logic **ngược lại**:
- Combobox chọn ở ô `B1` sẽ là **Sản phẩm (Name SKU)**.
- Bảng hiển thị ở cột D từ dòng 4 trở đi sẽ là danh sách các **Khách hàng** mua sản phẩm đó.
- Khi đổi sản phẩm chọn ở combobox, toàn bộ doanh số tháng và chỉ số của khách hàng sẽ tự động recalculate theo sản phẩm đó.

### Tụi mình đã hiện thực hóa nó như thế nào?

1. **Domain Layer**:
   - Tụi mình thêm interface `SkuCustomerSection` và các trường `skuCustomerRevenueSections`, `skuCustomerQuantitySections` vào kiểu dữ liệu `QecReport` trong [entities.ts](file:///d:/work/ToolForFen/src/domain/entities.ts).
   - Viết hàm `buildSkuCustomerSections` trong [reportCalculations.ts](file:///d:/work/ToolForFen/src/domain/reportCalculations.ts) để gom nhóm giao dịch theo sản phẩm, sau đó thống kê theo từng khách hàng.
   - Cập nhật [buildQecReport.ts](file:///d:/work/ToolForFen/src/application/buildQecReport.ts) để tính toán 2 danh mục mới này.

2. **Excel Export**:
   - Tụi mình viết thêm hàm `addProductWorksheet` trong [writeReportWorkbook.ts](file:///d:/work/ToolForFen/src/infrastructure/excel/writeReportWorkbook.ts) để ghi 2 sheet mới vào file Excel: **`SKU review - Customer`** (tiền tệ) và **`SKU review customer`** (số lượng).
   - Bổ sung cột ghi danh sách sản phẩm duy nhất vào cột B của sheet Config ẩn để làm nguồn dữ liệu cho dropdown sản phẩm ở ô `B1`.
   - Sử dụng công thức `SUMIFS` động tại mỗi cột tháng:
     `=SUMIFS('Data nguồn'!{dataCol}:{dataCol}, 'Data nguồn'!D:D, $D{currentRowNum}, 'Data nguồn'!F:F, $B$1, 'Data nguồn'!C:C, ${colLetter}$3)`
     Công thức này sẽ tự động lọc dữ liệu thô theo tên khách hàng ở cột D và tên sản phẩm được lựa chọn ở ô `B1`!

3. **Web UI Preview**:
   - Tụi mình thêm component `SkuCustomerPreview` và bổ sung 2 tab mới trên thanh TabBar của Web UI trong [App.tsx](file:///d:/work/ToolForFen/src/presentation/App.tsx) để hiển thị báo cáo mới này trực quan ngay trên trình duyệt, đồng bộ hoàn hảo trải nghiệm với Excel!

---

## 4. Bài học bỏ túi cho lập trình viên

- **Tư duy động trong Excel**: Đừng bao giờ ghi cứng giá trị tĩnh cho các cột tính toán nếu có thể viết công thức. Công thức Excel giúp file xuất ra "sống" và tương tác tốt hơn rất nhiều.
- **Config sheet ẩn**: Việc tách các danh sách dropdown validation ra một config sheet ẩn là cách làm cực kỳ sạch sẽ và chuyên nghiệp để quản lý data source trong ExcelJS.
- **Unit test**: Khi thêm sheet mới, hãy chú ý các file test so sánh (như `generateReportAndCompare.test.ts`) để cấu hình cho khớp, tránh làm gãy bộ test tự động của dự án.

Phiên làm việc hôm nay thực sự đã nâng tầm báo cáo QEC lên một đẳng cấp mới về độ linh hoạt và trải nghiệm người dùng! Chúc cậu code vui vẻ và hẹn gặp lại ở ly cà phê tiếp theo nhé! ☕
