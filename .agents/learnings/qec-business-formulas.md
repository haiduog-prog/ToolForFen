# 🧮 QEC Business Formulas & Month Logic

Tài liệu này chi tiết các chỉ số kinh doanh và cơ chế xử lý thời gian bán hàng cần thiết để AI có thể hiểu, kiểm thử và xây dựng báo cáo một cách chính xác.

---

## 📅 Cơ chế chuỗi thời gian kinh doanh (`src/domain/month.ts`)
Hệ thống sử dụng kiểu định danh `MonthKey` dạng chuỗi `YYYY-MM` (ví dụ: `2025-04`) làm đơn vị thời gian cơ sở.

*   **`reportPeriodMonths(reportMonth)`**: Tạo dải tháng phân tích kéo dài từ tháng 1 năm trước ($Y-1$) đến tháng báo cáo năm nay ($Y$).
    *   *Ví dụ*: Nếu tháng báo cáo là `2025-04`, dải tháng sẽ là từ `2024-01` đến `2025-04` (tổng cộng 16 tháng).
*   **`trailingMonths(reportMonth, count)`**: Trích xuất dải trung bình trượt với `count` tháng gần nhất.
    *   *Ví dụ*: Với `reportMonth = 2025-04`, `trailingMonths(reportMonth, 3)` sẽ trả về `['2025-02', '2025-03', '2025-04']`.
*   **`ytdMonths(year, throughMonthNumber)`**: Tạo dải tháng tích lũy YTD của năm được chỉ định.
    *   *Ví dụ*: `ytdMonths(2025, 4)` trả về từ `2025-01` đến `2025-04`.

---

## 🧮 Các công thức tính toán chỉ số (`src/domain/reportCalculations.ts`)

Mỗi dòng phân tích (`MetricRow`) chứa một bản đồ `monthValues: Record<MonthKey, number>` lưu doanh số của từng tháng trong dải báo cáo. Từ bản đồ này, các chỉ số sau được tính toán:

### 1. Previous Year Total (Tổng năm trước)
*   **Ý nghĩa**: Tổng doanh số trong toàn bộ năm trước đó ($Y-1$).
*   **Công thức**:
    $$\text{Previous Year Total} = \sum_{m \in \text{Tháng thuộc năm } Y-1} \text{Doanh số}_m$$

### 2. Current Year Total (Tổng năm nay)
*   **Ý nghĩa**: Tổng doanh số lũy kế từ tháng 1 đến tháng báo cáo của năm nay ($Y$).
*   **Công thức**:
    $$\text{Current Year Total} = \sum_{m \in \text{Tháng từ } 01 \text{ đến } \text{tháng báo cáo năm } Y} \text{Doanh số}_m$$

### 3. Share (Tỷ trọng)
*   **Ý nghĩa**: Đóng góp doanh số của một dòng (Segment/SKU) so với tổng thể toàn bảng.
*   **Công thức**:
    $$\text{Previous Year Share} = \frac{\text{Previous Year Total của dòng}}{\text{Tổng Previous Year Total của toàn bảng}}$$
    $$\text{Current Year Share} = \frac{\text{Current Year Total của dòng}}{\text{Tổng Current Year Total của toàn bảng}}$$

### 4. P3M / P6M / P9M (Trung bình trượt)
*   **Ý nghĩa**: Trung bình cộng doanh số của 3, 6, hoặc 9 tháng gần nhất (tính lùi từ tháng báo cáo).
*   **Công thức** (với $k \in \{3, 6, 9\}$):
    $$\text{PkM} = \frac{1}{k} \sum_{i=0}^{k-1} \text{Doanh số}_{\text{AddMonths}(\text{ReportMonth}, -i)}$$

### 5. TREND (Xu hướng ngắn hạn so với trung/dài hạn)
*   **Ý nghĩa**: Đánh giá biến động doanh số ngắn hạn so với trung và dài hạn. Chỉ số $> 1.0$ biểu thị xu hướng tăng trưởng tốt.
*   **Công thức**:
    $$\text{TREND} = \frac{\text{P3M} \times 2}{\text{P6M} + \text{P9M}}$$
*   *Lưu ý*: Sử dụng hàm an toàn `safeRatio` để tránh lỗi chia cho 0 nếu mẫu số $\text{P6M} + \text{P9M} = 0$.

### 6. IFYTD & ICYTD (Chỉ số tăng trưởng tích lũy lũy kế YTD)
*   **Ý nghĩa**: So sánh tổng doanh thu tích lũy từ đầu năm tới tháng báo cáo hiện tại so với cùng kỳ năm trước.
*   **Công thức**:
    $$\text{IFYTD} = \text{ICYTD} = \frac{\text{Tổng doanh thu tích lũy từ tháng 1 } \rightarrow \text{ tháng báo cáo năm nay}}{\text{Tổng doanh thu tích lũy từ tháng 1 } \rightarrow \text{ tháng báo cáo năm trước}}$$

### 7. IYA (Tăng trưởng tháng báo cáo so với cùng kỳ năm ngoái)
*   **Ý nghĩa**: Tăng trưởng của riêng tháng báo cáo hiện tại so với đúng tháng đó năm ngoái.
*   **Công thức**:
    $$\text{IYA} = \frac{\text{Doanh thu tháng báo cáo năm nay}}{\text{Doanh thu tháng báo cáo năm trước}}$$

---

## ⚠️ Quy tắc an toàn số học
1.  **Tránh chia cho 0 (`safeRatio`)**:
    ```typescript
    export function safeRatio(numerator: number, denominator: number): number {
      return denominator === 0 ? 0 : numerator / denominator;
    }
    ```
2.  **Khớp dữ liệu không đầy đủ**: Khi tính toán cho các tháng chưa có giao dịch, giá trị mặc định phải là `0` thay vì `undefined` hay `NaN`.
