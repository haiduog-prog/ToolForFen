# 🧮 Hệ thống Báo cáo Động QEC & SKU Customer Review

Tài liệu này đánh giá toàn diện các thay đổi kỹ thuật, kiến trúc tính toán và bài học kinh nghiệm thu được sau phiên tối ưu hóa hệ thống báo cáo QEC (bao gồm tối ưu công thức Excel và bổ sung sheet báo cáo đảo ngược).

---

## 1. Bối cảnh & Yêu cầu thay đổi

1. **Lỗi tính toán chỉ số tích lũy (IFYTD, ICYTD, IYA)**:
   - *Hiện tượng*: Các cột chỉ số tích lũy trong sheet `QEC review` hiển thị dấu gạch ngang `-` (giá trị tĩnh bằng 0) khi dữ liệu đầu vào `input.xlsx` không có giao dịch năm ngoái.
   - *Nguyên nhân*: Các giá trị được ghi cứng dưới dạng giá trị tĩnh từ backend. Khi mẫu số bằng 0 (không có giao dịch năm ngoái), backend trả về 0 tĩnh, làm mất tính năng tự recalculate của Excel.
2. **Cột Customer Code bị rỗng**:
   - Cột `Customer code` (Cột 1) của bảng Customer Detail trước đây bị bỏ trống hoàn toàn, gây thiếu sót thông tin trực quan.
3. **Báo cáo đảo ngược mới (SKU review - Customer)**:
   - Yêu cầu bổ sung sheet báo cáo mới cho phép chọn **Sản phẩm (Name SKU)** trong combobox bộ lọc và hiển thị danh sách các **Khách hàng** mua sản phẩm đó kèm theo doanh số tháng và chỉ số tích lũy, áp dụng trên cả Web UI Preview và file Excel xuất ra.

---

## 2. Đánh giá các giải pháp kỹ thuật đã triển khai

### A. Chuyển đổi QEC Review sang Công thức Excel động 100%
- **Giải pháp**: Thay vì ghi các giá trị tính toán tĩnh (`rowData.p3m`, `rowData.ifytd`, v.v.), chúng ta chuyển đổi toàn bộ các ô tính toán của Segment, Region, DSR, và Customer Detail thành công thức Excel động:
  - **Year Y-1 Total / Year Y Total**: Dùng công thức `=SUM({tháng_bắt_đầu}:{tháng_kết_thúc})`.
  - **P3M / P6M / P9M**: Dùng công thức `=AVERAGE({tháng_bắt_đầu}:{tháng_kết_thúc})`.
  - **TREND**: Dùng `=IFERROR((P3M_Col * 2) / (P6M_Col + P9M_Col), 0)`.
  - **IFYTD / ICYTD**: Dùng `=IFERROR(CY_Total / SUM(G{R}:{tháng_báo_cáo_năm_ngoái}{R}), 0)`.
  - **IYA**: Dùng `=IFERROR({tháng_báo_cáo_năm_nay}{R} / {tháng_báo_cáo_năm_ngoái}{R}, 0)`.
- **Dòng Grand Total**:
  - Dòng Total của mỗi bảng cũng được ghi bằng công thức `SUM` dọc từ các dòng dữ liệu phía trên để tự động cộng dồn và cập nhật đồng bộ khi người dùng thao tác.
- **Đánh giá**: File Excel xuất ra hoạt động cực kỳ linh hoạt và chuyên nghiệp, có khả năng tự tính toán lại chính xác khi người dùng mở file bằng Microsoft Excel, giải quyết triệt để lỗi hiển thị tĩnh.

### B. Bổ sung Số thứ tự cho Customer Code
- **Giải pháp**: Tự động điền số thứ tự tăng dần (`1, 2, 3...`) vào cột `Customer code` (Cột 1) của bảng Customer Detail cho từng khách hàng phát sinh giao dịch.
- **Đánh giá**: Khắc phục hoàn toàn việc để trống cột đầu tiên, giúp người dùng dễ dàng theo dõi và đếm tổng số lượng khách hàng trực tiếp trên file Excel.

### C. Kiến trúc Báo cáo Đảo ngược SKU review - Customer
- **Domain Layer**:
  - Bổ sung kiểu `SkuCustomerSection` và tích hợp vào `QecReport` để quản lý độc lập 2 danh mục mới (Revenue và Quantity).
  - Viết hàm `buildSkuCustomerSections` trong `reportCalculations.ts` để thực hiện gom nhóm giao dịch theo sản phẩm, sau đó thống kê theo từng khách hàng.
- **Excel Infrastructure**:
  - Thêm hàm `addProductWorksheet` trong `writeReportWorkbook.ts`.
  - Tạo Dropdown chọn sản phẩm tại ô `B1` (trỏ tới Product List duy nhất trong sheet Config ẩn).
  - Thiết lập công thức `SUMIFS` lọc dữ liệu thô theo khách hàng (cột D) và sản phẩm được chọn (ô `B1`):
    `=SUMIFS('Data nguồn'!{dataCol}:{dataCol}, 'Data nguồn'!D:D, $D${currentRowNum}, 'Data nguồn'!F:F, $B$1, 'Data nguồn'!C:C, ${colLetter}$3)`
- **Presentation Layer (Web UI)**:
  - Thêm 2 tab mới trên thanh TabBar: **"SKU review - Customer"** và **"SKU review customer"**.
  - Xây dựng component `SkuCustomerPreview` hỗ trợ lựa chọn sản phẩm và hiển thị khách hàng tương ứng một cách trực quan, đồng bộ trải nghiệm với Excel.
- **Đánh giá**: Đây là một tính năng mở rộng cực kỳ hoàn chỉnh và đồng bộ từ tầng Domain, Data, Excel Export cho đến Web UI, mang lại giá trị phân tích dữ liệu rất cao cho người dùng.

---

## 3. Khuyến nghị Kỹ thuật cho Lập trình viên

1. **Luôn ưu tiên Công thức Excel cho các cột tính toán**:
   - Khi thiết kế xuất Excel, tuyệt đối tránh ghi các giá trị tĩnh đối với các cột cộng dồn, trung bình trượt hay tỉ lệ tăng trưởng. Việc dùng công thức Excel giúp file xuất ra giữ được tính năng tự recalculate của bảng tính.
2. **Tối ưu hóa Dropdown Validation**:
   - Dropdown list validation (`dataValidation`) nên được trỏ tới các dải ô động trong sheet Config ẩn (ví dụ: `Config!$B$2:$B$${length + 1}`) để đảm bảo danh sách lựa chọn luôn được cập nhật đầy đủ và đồng bộ.
3. **Quản lý Unit Test khi mở rộng Sheet**:
   - Khi bổ sung sheet mới vào workbook, cần cấu hình các file so sánh Excel (như `generateReportAndCompare.test.ts`) một cách khôn ngoan (chỉ so sánh các sheet critical cần khớp tuyệt đối hoặc bỏ qua các sheet phụ) để tránh việc gãy unit test do lệch cấu trúc sheet.
