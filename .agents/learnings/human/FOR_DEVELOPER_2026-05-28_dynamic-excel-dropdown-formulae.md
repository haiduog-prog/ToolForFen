# Coffee Talk: Tạo dựng file Excel Phản ứng động (Reactive Spreadsheet) bằng Dropdown và công thức SUMIFS
*Ngày viết: 28-05-2026*

Chào bạn! Rót thêm một tách trà ấm nhé. Hôm nay, tôi và bạn vừa thực hiện một sự cải tiến mang tính đột phá và cực kỳ đẳng cấp trong nghệ thuật xử lý bảng tính: **Biến một file Excel báo cáo tĩnh cồng kềnh thành một ứng dụng bảng tính phản ứng động (Reactive Spreadsheet) thực thụ ngay trong file tải xuống!**

Hãy cùng ngồi phân tích những bài học đắt giá, những công thức đỉnh cao mà chúng ta đã dùng để chinh phục bài toán này nhé.

---

## 1. Approach & Reasoning (Cách tiếp cận & Logic suy nghĩ)

**Mong muốn của người dùng:**
Người dùng muốn trong file Excel xuất ra, hai sheet `SKU - Customer review` và `SKU customer review` "chỉ hiển thị 1 Customer 1 lần thôi, khi bấm vào tên Customer thì có thể chọn các KH khác để hiển thị lên". Họ đã gửi ảnh chụp màn hình Excel để làm rõ rằng họ thực sự muốn điều này **ở ngay trong file Excel tải về**, chứ không chỉ dừng lại ở giao diện Web preview.

**Giải pháp kỹ thuật của chúng ta:**
Để giải quyết bài toán tương tác này trong Excel mà không cần dùng đến VBA (vốn hay bị các trình duyệt và phần mềm diệt virus chặn), chúng ta đã kết hợp **Data Validation Dropdown** và **Hệ thống công thức Excel động tiêu chuẩn** (`SUMIFS`, `SUM`, `AVERAGE`, `IFERROR`):

1.  **Dropdown chọn khách hàng (Data Validation)**:
    *   Chúng ta tạo ra một sheet phụ ẩn tên là `"Config"`. Ghi danh sách tất cả các khách hàng thực tế vào cột A của sheet này. Ẩn sheet đi bằng thuộc tính `configSheet.state = "hidden"`.
    *   Tại ô `B1` của sheet hiển thị, chúng ta thiết lập Data Validation kiểu `list` liên kết tới danh sách khách hàng ở sheet `"Config"`.
2.  **Công thức SUMIFS động**:
    *   Cột sản phẩm (cột D) sẽ chứa danh sách toàn bộ các sản phẩm thực tế phát sinh giao dịch. Bảng của chúng ta chỉ dài khoảng 30 dòng (gọn gàng hơn 100 lần so với 15,000 dòng xếp chồng trước đây!).
    *   Với mỗi tháng (từ cột E đến cột T), thay vì ghi giá trị tĩnh, chúng ta viết công thức `=SUMIFS`:
        *   *Công thức:* `=SUMIFS('Data nguồn'!J:J, 'Data nguồn'!D:D, $B$1, 'Data nguồn'!F:F, $D[row], 'Data nguồn'!C:C, [Col]$3)`
        *   Nghĩa là: Lọc doanh thu từ sheet `Data nguồn` thỏa mãn 3 điều kiện: Nhà thuốc khớp với ô `$B$1` (khách hàng đang chọn), Tên sản phẩm khớp với cột `$D` của dòng hiện tại, và Tháng khớp với ô tiêu đề tháng ở dòng 3!
3.  **Hệ thống công thức lũy kế & trung bình trượt**:
    *   *CY 2024 (năm trước):* `=SUM(E[row]:P[row])`
    *   *CY 2025 (năm nay):* `=SUM(Q[row]:[EndCol][row])`
    *   *P3M/P6M/P9M:* `=AVERAGE(...)` của các tháng tương ứng.
    *   *TREND:* `=IFERROR((X[row]*2)/(Y[row]+Z[row]), 0)`
4.  **Dòng TOTAL phản ứng**:
    *   Dòng TOTAL ở dòng cuối cùng cũng chứa công thức `=SUM(...)`. Khi người dùng thay đổi dropdown khách hàng ở ô `B1`, dòng TOTAL này cũng tự động cập nhật ngay lập tức!

---

## 2. Roads Not Taken (Những con đường không đi)

*   **Sử dụng công thức mảng động `=FILTER` trong Office 365:**
    *   *Ý tưởng:* Dùng `=FILTER('Data nguồn'!A:Z, 'Data nguồn'!D:D = B1)` để Excel tự động lọc dữ liệu.
    *   *Lý do bỏ qua:* Công thức `=FILTER` chỉ hỗ trợ trên các phiên bản Excel mới (Office 365 hoặc Excel 2021 trở lên). Nếu người dùng mở bằng các phiên bản Excel cũ (2019, 2016, 2013), bảng tính sẽ bị lỗi `#NAME?`. Đồng thời, ExcelJS không tự tính toán kết quả công thức mảng động này khi mở file.
    *   *Giải pháp thay thế:* Sử dụng `=SUMIFS` kết hợp bảng sản phẩm cố định. Đây là công thức Excel kinh điển, hoạt động **100% trên mọi phiên bản Excel** từ cổ chí kim và cực kỳ ổn định.

---

## 3. How Things Connect (Mối liên kết hệ thống)

Hãy chú ý cách các sheet kết nối với nhau thông qua công thức Excel:
```
[Sheet Config (Ẩn)] ──(Danh sách khách hàng)──> [Dropdown Validation ô B1]
                                                          │ (Chọn Khách hàng)
                                                          ▼
[Sheet Data nguồn (Dữ liệu thô)] ──(SUMIFS)──> [Bảng số liệu sheet SKU - Customer review]
```
Sự liên kết thông qua công thức Excel này giúp file Excel tự vận hành độc lập sau khi tải xuống. Người dùng không cần phải kết nối internet hay chạy thêm bất kỳ công cụ nào khác. Bản thân Excel sẽ tự tính toán lại toàn bộ bảng tính mỗi khi người dùng chọn khách hàng khác ở ô `B1`.

---

## 4. Công thức Excel đắt giá (Formula Cheatsheet)

Dưới đây là cách chúng ta xây dựng công thức động bằng JavaScript trong `writeReportWorkbook.ts` để Excel tự hiểu:

*   **Helper lấy chữ cái cột:**
    ```typescript
    function getColumnLetter(colNumber: number): string {
      let temp = colNumber;
      let letter = "";
      while (temp > 0) {
        const modulo = (temp - 1) % 26;
        letter = String.fromCharCode(65 + modulo) + letter;
        temp = Math.floor((temp - modulo) / 26);
      }
      return letter;
    }
    ```
*   **Data Validation Dropdown:**
    ```typescript
    r1.getCell(2).dataValidation = {
      type: "list",
      allowBlank: true,
      formulae: [`Config!$A$2:$A$${customerNames.length + 1}`]
    };
    ```
*   **Công thức SUMIFS cho các tháng:**
    ```typescript
    excelRow.getCell(5 + mIdx).value = {
      formula: `SUMIFS('Data nguồn'!${dataCol}:${dataCol}, 'Data nguồn'!D:D, $B$1, 'Data nguồn'!F:F, $D${currentRowNum}, 'Data nguồn'!C:C, ${colLetter}$3)`
    };
    ```

---

## 5. Future Pitfalls (Cạm bẫy cần tránh trong tương lai)

*   **Tên sheet có dấu cách hoặc ký tự đặc biệt:**
    *   *Cạm bẫy:* Khi viết công thức Excel tham chiếu đến tên sheet, nếu tên sheet có dấu cách (như `'Data nguồn'`), bắt buộc phải bọc trong dấu nháy đơn `'`. Nếu viết là `Data nguồn!...` mà không có dấu nháy đơn, Excel sẽ báo lỗi công thức.
    *   *Lời khuyên:* Hãy luôn viết dấu nháy đơn cho tên sheet trong công thức nếu tên sheet có dấu cách tiếng Việt (như `'Data nguồn'`). Với sheet cấu hình `"Config"`, vì không có dấu cách nên viết `Config!...` là an toàn.
*   **Data Validation formulae không có dấu `=` ở đầu:**
    *   *Cạm bẫy:* Khi viết data validation cho list trong ExcelJS, chuỗi formula trong `formulae` **không được phép** chứa dấu `=` ở đầu. Ví dụ: `formulae: ['Config!$A$2:$A$20']` là đúng. Nếu viết `formulae: ['=Config!$A$2:$A$20']`, Excel có thể không nhận diện được vùng dữ liệu.

---

## 6. Expert vs Beginner (Đẳng cấp của Expert)

*   **Beginner** sẽ viết cứng toàn bộ dữ liệu tĩnh của tất cả khách hàng xếp chồng lên nhau, làm phình to file Excel lên hàng chục ngàn dòng và tạo cảm giác nghiệp dư. Hoặc họ cố gắng viết các công thức mảng động phức tạp chỉ chạy được trên Office 365 bản quyền.
*   **Expert** (như cách chúng ta vừa làm) biết cách biến file Excel thành một ứng dụng phản ứng thực thụ bằng cách sử dụng **sheet phụ ẩn `"Config"`, Data Validation Dropdown, và công thức `=SUMIFS` kinh điển**. Bảng tính vừa nhẹ (chỉ 34 dòng), vừa mượt, vừa chạy trên mọi phiên bản Excel của người dùng, mang lại trải nghiệm **WOW** thực sự đẳng cấp!

---

Hy vọng những đúc kết kỹ thuật này giúp bạn tự tin làm chủ nghệ thuật lập trình Excel cao cấp. Hẹn gặp lại bạn ở ly cà phê tiếp theo nhé! ☕
