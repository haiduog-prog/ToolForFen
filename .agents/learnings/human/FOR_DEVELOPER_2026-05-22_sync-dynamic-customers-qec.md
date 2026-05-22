# Coffee Talk: Giải mã hành trình Tinh gọn dữ liệu & Đồng bộ giao diện QEC Export Builder
*Ngày viết: 22-05-2026*

Chào bạn, kéo ghế lại đây làm ly cà phê nhé! Hôm nay tôi và bạn đã cùng nhau đi qua một hành trình giải quyết bài toán xử lý Excel và thiết kế Web UI rất thú vị. Tôi viết lại tài liệu này để lưu trữ những bài học đắt giá mà hai chúng ta đã đúc kết được trong suốt ngày hôm nay. Đọc qua để thấy chúng ta đã biến một đống hỗn độn thành một hệ thống tinh gọn và đẹp đẽ như thế nào nhé!

---

## 1. Approach & Reasoning (Cách tiếp cận & Logic suy nghĩ)

Khi bắt đầu, chúng ta đối mặt với một loạt yêu cầu:
1. File Excel đầu vào (`input.xlsx`) có cấu trúc cột không cố định và chữ hoa chữ thường lung dung.
2. File Excel xuất ra (`QEC_2025-04.xlsx`) có các sheet review chứa quá nhiều dòng rỗng của các sản phẩm và khách hàng không phát sinh giao dịch.
3. Người dùng muốn xem trước sheet "Data nguồn" ngay trên Web UI để đối chiếu dữ liệu.
4. Danh sách khách hàng trên tab QEC của giao diện Web bị lệch so với Excel và vẫn đang dùng dữ liệu tĩnh cũ.

**Điểm xuất phát:** Tôi không nhảy vào code ngay mà đi nghiên cứu kiến trúc dự án (`project-architecture.md`) và hạ tầng Excel (`excel-infrastructure.md`). Tôi nhận thấy dự án đang đi theo mô hình **Clean Architecture & DDD** rất nghiêm ngặt.

**Logic suy nghĩ:**
*   **Với Header Matching:** Tôi quyết định xây dựng một hàm chuẩn hóa `normalizeHeaderName` dọn sạch dấu tiếng Việt, đưa về chữ thường và bỏ khoảng trắng thừa trước khi so khớp cột. Điều này giúp hệ thống "chấp hết" các biến thể tiêu đề do người dùng nhập tay.
*   **Với Sản phẩm & Khách hàng rỗng:** Thay vì import danh sách tĩnh khổng lồ gồm 52 sản phẩm và hơn 300 khách hàng để rồi vẽ ra hàng ngàn dòng `0` vô nghĩa, tôi chuyển hướng sang **thu thập động**. Nghĩa là: *Chỉ những SKU và Khách hàng thực sự phát sinh giao dịch trong file tải lên mới được đưa vào báo cáo.*
*   **Với Giao diện Web:** Để mang lại trải nghiệm WOW, tôi quyết định không chỉ hiển thị bảng Segment đơn điệu trên tab QEC, mà mang nguyên vẹn cấu trúc 3 tầng của Excel (Segment Review, DSR Review, Customer Detail Review) lên web, sử dụng chung component hiển thị để đảm bảo dữ liệu hiển thị đồng bộ 100%.

---

## 2. Roads Not Taken (Những con đường không đi)

Hiểu được tại sao ta từ bỏ một giải pháp đôi khi còn quan trọng hơn là biết tại sao ta chọn nó. Đây là những con đường chúng ta đã né tránh:

*   **Sử dụng `exceljs` để đọc file Excel nguồn:**
    *   *Ý tưởng ban đầu:* Dùng `exceljs` cho cả đọc lẫn ghi để đồng bộ thư viện.
    *   *Lý do bỏ qua:* `exceljs` thường bị lỗi crash `Cannot read properties of undefined (reading 'styles')` khi cố đọc các file Excel có định dạng style phức tạp được xuất ra từ các phần mềm kế toán cũ.
    *   *Giải pháp thay thế:* Giữ nguyên việc đọc bằng `xlsx` (SheetJS) - vốn cực kỳ trâu bò và nhẹ, chỉ dùng `exceljs` ở chiều ghi (write) để tận dụng khả năng định dạng border/color tuyệt đẹp của nó.
*   **Giữ nguyên 52 dòng sản phẩm tĩnh cho từng khách hàng:**
    *   *Ý tưởng ban đầu:* Giữ nguyên danh sách tĩnh vì "sợ" lệch cấu trúc mẫu.
    *   *Lý do bỏ qua:* Nếu giữ danh sách tĩnh, một khách hàng chỉ mua 1 sản phẩm vẫn phải cõng thêm 51 dòng trống trị giá `0` trong sheet `SKU - Customer review`. Nhân lên với hơn 200 khách hàng, chúng ta sẽ có một file Excel nặng tới hơn **15,000 dòng**, trong đó 90% là dòng rỗng. File rất lag và xấu.
    *   *Giải pháp thay thế:* Lọc động sản phẩm. Kích thước file Excel xuất ra giảm xuống chỉ còn **1,687 dòng**, mở lên mượt mà và tập trung ngay vào sản phẩm thực tế phát sinh doanh thu.

---

## 3. How Things Connect (Mối liên kết giữa các mảnh ghép)

Hãy tưởng tượng luồng dữ liệu của chúng ta hoạt động như một nhà máy lọc dầu:
1.  **Dữ liệu thô (`xlsx` ở Infrastructure)** đổ vào -> Trải qua bộ lọc chuẩn hóa tiêu đề (`normalizeHeaderName`) để nhận diện cột động.
2.  **Bộ điều phối (`buildQecReport` ở Application)** tiếp nhận -> Lọc bỏ giao dịch của nhóm khách hàng loại trừ (`EXCLUDED_CUSTOMERS`) -> Trích xuất danh sách khách hàng và sản phẩm động -> Tính toán chuỗi thời gian (YTD, P3M, Trend) -> Trả về cấu trúc báo cáo chuẩn `QecReport`.
3.  **Đầu ra:**
    *   *Ghi file Excel (`writeReportWorkbook` ở Infrastructure)* dùng `report.customerBaseRows` để ghi ra bảng Customer Detail ở dòng 28 của sheet `QEC review`.
    *   *Giao diện Web UI (`App.tsx` ở Presentation)* cũng dùng chính `report.customerBaseRows` để vẽ nên bảng hiển thị động khớp từng con số với Excel.

---

## 4. Tools & Methods (Công cụ & Phương pháp)

*   **Vitest (Unit Testing):** Chúng tôi dùng Vitest để chạy test nghiệp vụ. Khi thay đổi cấu trúc sản phẩm và khách hàng động, các test case tự động chạy lại để đảm bảo công thức trung bình trượt và Trend không bị chia cho 0 (`safeRatio`).
*   **Client-side Pagination (Phân trang phía Client):** Với sheet "Data nguồn" có thể lên tới hàng ngàn dòng, việc render trực tiếp lên DOM sẽ gây lag trình duyệt. Giải pháp phân trang 50 dòng/trang kết hợp bộ chọn trang mượt mà giúp giao diện cực kỳ mượt và chuẩn UX.

---

## 5. Tradeoffs (Sự đánh đổi)

Mọi quyết định thiết kế đều là một sự thỏa hiệp:
*   **Đánh đổi tính "giống hệt nguyên bản" lấy "sự thực tế":**
    *   *Hy sinh:* File Excel xuất ra sẽ không có đủ 52 SKU hay 300 khách hàng tĩnh như file chuẩn mẫu của họ.
    *   *Đạt được:* File Excel nhẹ hơn 10 lần, số liệu tập trung 100% vào thực tế kinh doanh phát sinh, không có dòng rỗng vô nghĩa.
*   **Đánh đổi hiệu năng render lấy trải nghiệm trực quan:**
    *   *Hy sinh:* Tab QEC review trên Web UI giờ đây render tới 3 bảng thay vì 1 bảng như trước.
    *   *Đạt được:* Người dùng nhìn phát biết ngay dữ liệu khớp với Excel hay chưa mà không cần phải mở file Excel lên đối chiếu từng dòng.

---

## 6. Mistakes & Dead Ends (Sai lầm & Ngõ cụt)

Trong quá trình làm, có một "ngõ cụt" nhỏ khi xử lý bảng chữ cái tiếng Việt:
*   *Vấn đề:* Khi sort danh sách khách hàng động bằng hàm `.sort()`, các khách hàng bắt đầu bằng chữ tiếng Việt có dấu (như *Chị Nhung*, *Đỗ Văn Chính*) bị nhảy xuống cuối bảng do mã Unicode của ký tự tiếng Việt lớn hơn ký tự Latinh tiêu chuẩn.
*   *Cách khắc phục:* Tôi đã thay thế bằng hàm so sánh ngôn ngữ chuyên dụng:
    `customerBaseRows.sort((a, b) => a.label.localeCompare(b.label, "vi"))`
    Kết quả là danh sách khách hàng được sắp xếp chuẩn chỉ theo bảng chữ cái tiếng Việt cực kỳ tự nhiên.

---

## 7. Future Pitfalls (Cạm bẫy cần tránh trong tương lai)

*   **Cẩn thận với `EXCLUDED_CUSTOMERS`:** Hiện tại danh sách các khách hàng loại trừ khỏi doanh số (như *Chị Châu*, *Khách lẻ (Cảnh)*,...) đang được viết cứng (hardcoded) trong `buildQecReport.ts`. Trong tương lai, nếu danh sách khách hàng loại trừ này thay đổi, developer phải vào sửa trực tiếp trong code.
    *   *Lời khuyên:* Sau này nếu có điều kiện, hãy đưa danh sách loại trừ này thành một bảng cấu hình trên UI để người dùng tự tick chọn.
*   **Lỗi format ngày tháng trong Data nguồn:** Định dạng ngày của Excel rất đa dạng (chuỗi chữ, số Serial của Excel, Date Object). Hãy luôn sử dụng utility `parseMonthFromCell` để bóc tách ngày tháng an toàn.

---

## 8. Expert vs Beginner (Đẳng cấp của Expert)

*   **Beginner** sẽ cố gắng import toàn bộ file `exceljs` nặng nề vào bundle chính của React, khiến trang web tải rất chậm khi mở lần đầu.
*   **Expert** sẽ sử dụng **Dynamic Import** (như cách chúng ta làm ở `writeReportWorkbook.ts`):
    `const ExcelJSModule = await import("exceljs/dist/exceljs.min.js");`
    Keep bundle nhẹ (chỉ khoảng 500KB) và tải thư viện ExcelJS bất đồng bộ khi cần.

---

## 9. Transferable Lessons (Bài học rút ra cho các dự án khác)

*   **Không tin tưởng vào cấu trúc file Excel của người dùng:** Người dùng luôn có cách làm sai định dạng tiêu đề cột (viết hoa, viết thường, thêm dấu cách, thêm dấu tiếng Việt). Hãy luôn áp dụng phương pháp **Heuristic Header Matching** và **Normalization** trước khi đọc dữ liệu.
*   **Sử dụng CSS Variables để thiết kế đồng bộ:** Việc định nghĩa toàn bộ bảng màu, spacing và border-radius trong `:root` giúp chúng ta dễ dàng tùy biến hoặc thêm các thành phần giao diện mới (như cụm 3 bảng trong tab QEC hôm nay) mà không sợ bị lệch tông màu hay kích thước.

---

Hy vọng tài liệu "trò chuyện cà phê" này giúp bạn hiểu sâu sắc và trọn vẹn những gì chúng ta đã làm ngày hôm nay. Chúc bạn code vui vẻ và hẹn gặp lại ở ly cà phê tiếp theo nhé! ☕
