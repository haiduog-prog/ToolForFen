# Coffee Talk: Giải mã bộ chọn Khách hàng (Customer Selector) thông minh và bài toán Trực quan hóa số liệu
*Ngày viết: 28-05-2026*

Chào bạn! Hãy rót một tách cà phê và ngồi xuống đây nhé. Hôm nay, chúng ta vừa giải quyết một vấn đề cực kỳ nhức nhối về mặt UI/UX mà bất kỳ dự án báo cáo dữ liệu lớn nào cũng hay mắc phải: **Làm sao hiển thị hàng trăm khách hàng với hàng chục sản phẩm đi kèm mà không làm vỡ giao diện hoặc khiến người dùng "khóc thét" vì cuộn chuột mỏi tay?**

Dưới đây là những bài học đắt giá mà tôi và bạn vừa cùng nhau đúc kết được thông qua việc xây dựng bộ chọn Khách hàng phản ứng (Reactive Customer Selector) siêu mượt mà.

---

## 1. Approach & Reasoning (Cách tiếp cận & Logic suy nghĩ)

**Vấn đề ban đầu:**
Trước đây, tab *Customer revenue* và *Customer quantity* của chúng ta dùng giải pháp hiển thị "phẳng" khá thô sơ: ghép (flatMap) 12 khách hàng đầu tiên, mỗi khách hàng chỉ hiển thị tối đa 6 sản phẩm. Đây là một sự giới hạn nửa vời, không cho phép xem đầy đủ danh sách sản phẩm và cũng không thể xem các khách hàng phía sau.

**Logic suy nghĩ:**
*   **Ý tưởng của người dùng:** "sửa 2 sheet SKU - Customer review này chỉ hiển thị 1 Customer 1 lần thôi bấm vào tên Customer thì có thể chọn các KH khác để hiển thị lên".
*   **Với Web UI:** Ý tưởng này hoàn hảo! Nó giúp biến bảng dữ liệu dài vô tận thành một màn hình tương tác thông minh. Chúng ta chỉ hiển thị **1 khách hàng duy nhất tại một thời điểm**.
*   **Thiết kế UI:** Phía trên bảng, chúng ta đặt một thanh điều khiển (`customer-selector-bar`) gồm:
    1.  Một thẻ `<select>` chứa toàn bộ tên khách hàng thực tế có trong file data nguồn tải lên.
    2.  Bộ đôi nút điều hướng nhanh **Trước (Prev)** và **Sau (Next)** bằng các icon `ChevronLeft` và `ChevronRight` để người dùng có thể nhấp chuột chuyển đổi nhanh mà không cần mở dropdown.
*   **Thiết kế Bảng số liệu:** Khi khách hàng được chọn, bảng sẽ hiển thị **toàn bộ sản phẩm (SKU) thực tế** của khách hàng đó (không giới hạn 6 dòng nữa). Ở cuối bảng, chúng ta hiển thị dòng **TOTAL** riêng biệt của khách hàng đó (lấy trực tiếp từ `currentSection.total`), khớp 100% với số liệu tổng của khách hàng đó trong Excel.

---

## 2. Roads Not Taken (Những con đường không đi)

*   **Sử dụng Dropdown / bộ chọn động trực tiếp trong file Excel xuất ra:**
    *   *Lý do bỏ qua:* Excel là một file tĩnh dùng để lưu trữ và gửi đi cho đối tác hoặc in ấn. Một file Excel báo cáo cần giữ đầy đủ thông tin của tất cả khách hàng để khi người dùng in ấn hoặc phân tích thì không bị thiếu dữ liệu. Nếu chúng ta chỉ hiển thị 1 khách hàng và ẩn đi các khách hàng khác bằng VBA hoặc công thức lọc Excel, file sẽ không còn tính chất báo cáo tổng hợp toàn cục nữa, đồng thời gây khó khăn cho việc in ấn hàng loạt.
    *   *Chiến lược tốt nhất:* **Đồng bộ tối ưu trên Web UI và Giữ nguyên định dạng chuẩn trên Excel**. Web UI đóng vai trò là "màn hình tương tác sống động" giúp tra cứu nhanh cực kỳ mượt mà, còn file Excel đóng vai trò là "bản báo cáo đầy đủ, toàn vẹn dữ liệu".

---

## 3. How Things Connect (Mối liên kết hệ thống)

Dữ liệu của chúng ta được tổ chức cực kỳ sạch sẽ trong `QecReport` (tầng Domain):
*   `customerRevenueSections` và `customerQuantitySections` là một mảng các `CustomerSection`.
*   Mỗi `CustomerSection` chứa:
    *   `customer: string` (tên khách hàng)
    *   `rows: MetricRow[]` (mảng các sản phẩm của khách hàng đó)
    *   `total: MetricRow` (dòng tổng cộng của riêng khách hàng đó)

Nhờ cấu trúc dữ liệu DDD tuyệt đẹp này, trên React Web UI (`App.tsx`), chúng ta chỉ cần dùng `useState` để lưu tên khách hàng đang được chọn (`selectedCustomer`), sau đó dùng hàm `.find()` để bốc chính xác `CustomerSection` tương ứng ra hiển thị. Việc vẽ dòng Total ở cuối bảng chỉ đơn giản là render thêm một hàng `tr` sử dụng dữ liệu từ `currentSection.total` cực kỳ gọn gàng!

---

## 4. Tradeoffs (Sự đánh đổi)

*   **Độ mượt mà của Web Preview vs Tính toàn vẹn của Excel:**
    *   *Đánh đổi:* Chúng ta chọn cách thể hiện khác nhau trên hai môi trường (Dropdown trên Web và Danh sách xếp chồng trên Excel).
    *   *Kết quả:* Sự lựa chọn này mang lại trải nghiệm tối ưu nhất cho cả hai thế giới: Web UI trực quan, linh hoạt; Excel lưu trữ toàn vẹn, dễ in ấn.

---

## 5. Expert vs Beginner (Đẳng cấp của Expert)

Một **Beginner** khi làm tính năng này thường sẽ:
*   Chỉ đặt dropdown `<select>` mà quên mất nút Prev/Next, bắt người dùng phải click mở dropdown mỗi khi muốn chuyển sang khách hàng tiếp theo (rất mỏi tay).
*   Quên mất dòng TOTAL của riêng khách hàng đó ở cuối bảng.
*   Không cố định cột sản phẩm (`sticky-col`), khiến khi cuộn ngang bảng trên điện thoại hoặc màn hình nhỏ, người dùng không biết dòng số liệu đó thuộc về sản phẩm nào.

Một **Expert** (như cách chúng ta vừa làm):
*   **Tích hợp bộ đôi nút Prev/Next** thông minh, tự động disabled khi ở đầu hoặc cuối danh sách.
*   **Hiển thị dòng TOTAL riêng biệt** của khách hàng ở dòng cuối cùng của bảng với màu sắc đậm đà nổi bật (`total-row`).
*   **Cố định cột sản phẩm** (`sticky-col` kết hợp `text-left`), giúp giao diện xem trước cực kỳ mượt mà, responsive tốt trên mọi kích thước màn hình.

---

Hy vọng những đúc kết trên đây giúp bạn hiểu sâu sắc giải pháp thiết kế UI/UX tuyệt vời này. Hẹn gặp lại bạn ở ly cà phê tiếp theo nhé! ☕
