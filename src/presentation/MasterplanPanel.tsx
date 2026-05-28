import { useEffect, useState, useMemo, useRef } from "react";
import { Download, FileSpreadsheet, Loader2, Plus, Trash2, Sliders, DollarSign, Tag, Users, Upload, Edit, Save, X, BarChart3 } from "lucide-react";
import { buildMasterplan, parseCustomerExcel, extractExcelHeaders } from "../application/buildMasterplan";
import { type MasterplanData, type SalesPlanProduct, type SalesPlanCustomer, CHANNEL_OPTIONS } from "../domain/masterplan";
import { formatNumber } from "../shared/formatters";
import * as xlsx from "xlsx";

function formatMoneyInput(value: string): string {
  const numeric = value.replace(/\D/g, "");
  if (!numeric) return "";
  return Number(numeric).toLocaleString("en-US");
}

function parseMoneyInput(value: string): number {
  const numeric = value.replace(/\D/g, "");
  return parseFloat(numeric) || 0;
}

interface MasterplanPanelProps {
  source: any; // ignored or synced
  onExport: (data: MasterplanData) => void;
  isExporting: boolean;
}

export function MasterplanPanel({ onExport, isExporting }: MasterplanPanelProps) {
  // Products state (initialized with empty array by default)
  const [products, setProducts] = useState<SalesPlanProduct[]>([]);

  // Customers state (defaults to sample mock data)
  const [customers, setCustomers] = useState<SalesPlanCustomer[]>([]);
  const [isReadingCustomers, setIsReadingCustomers] = useState(false);
  const [custFile, setCustFile] = useState<File | null>(null);
  const [error, setError] = useState<string | null>(null);

  // File type and target channel configuration
  const [fileType, setFileType] = useState<"normalized" | "sales_detail">("normalized");
  const [targetChannelCode, setTargetChannelCode] = useState<number>(1004);
  const [originalRowCount, setOriginalRowCount] = useState<number>(0);

  // Dynamic headers read from the Excel file (used in normalized mode)
  const [excelHeaders, setExcelHeaders] = useState<string[]>([]);
  const [selectedTotalAmountColIdx, setSelectedTotalAmountColIdx] = useState<number>(-1);

  // New product input states
  const [newProdName, setNewProdName] = useState("");
  const [newProdPrice, setNewProdPrice] = useState("");
  const [newProdThreshold, setNewProdThreshold] = useState("");
  const [newProdQtyA, setNewProdQtyA] = useState("");
  const [newProdQtyB, setNewProdQtyB] = useState("");
  const [prodError, setProdError] = useState<string | null>(null);

  // Ecom configuration state
  const [ecomEnabled, setEcomEnabled] = useState(false);

  // Inline editing state
  const [editingIndex, setEditingIndex] = useState<number | null>(null);
  const [editName, setEditName] = useState("");
  const [editPrice, setEditPrice] = useState("");
  const [editThreshold, setEditThreshold] = useState("");
  const [editQtyA, setEditQtyA] = useState("");
  const [editQtyB, setEditQtyB] = useState("");

  const fileInputRef = useRef<HTMLInputElement | null>(null);

  // Build timeline and masterplan data dynamically
  const planData: MasterplanData = useMemo(() => {
    return buildMasterplan(customers, products, ecomEnabled);
  }, [customers, products, ecomEnabled]);

  // Handle customer file upload
  const handleCustomerFileChange = async (file: File | null) => {
    if (!file) return;
    setCustFile(file);
    setIsReadingCustomers(true);
    setError(null);
    setExcelHeaders([]);
    setSelectedTotalAmountColIdx(-1);
    setOriginalRowCount(0);

    try {
      const { headers, matchedIdx, mode } = await extractExcelHeaders(file);
      setExcelHeaders(headers);
      setSelectedTotalAmountColIdx(matchedIdx);
      setFileType(mode);

      const parsed = await parseCustomerExcel(file, {
        mode,
        totalAmountColIdx: matchedIdx === -1 ? undefined : matchedIdx,
        channelCode: targetChannelCode,
      });

      // Count original raw rows
      const arrayBuffer = await file.arrayBuffer();
      const data = new Uint8Array(arrayBuffer);
      const workbook = xlsx.read(data, { type: "array" });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]!];
      if (worksheet) {
        const rows = xlsx.utils.sheet_to_json<any[]>(worksheet, { header: 1 });
        setOriginalRowCount(rows.length > 0 ? rows.length - 1 : 0);
      }

      setCustomers(parsed);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "Lỗi đọc file khách hàng.");
      setCustomers([]);
      setExcelHeaders([]);
      setSelectedTotalAmountColIdx(-1);
      setOriginalRowCount(0);
    } finally {
      setIsReadingCustomers(false);
      if (fileInputRef.current) {
        fileInputRef.current.value = "";
      }
    }
  };

  // Re-parse when manually switching file type or channel in the UI
  const handleManualSettingsChange = async (type: "normalized" | "sales_detail", code: number) => {
    if (!custFile) return;
    setIsReadingCustomers(true);
    setError(null);
    try {
      const parsed = await parseCustomerExcel(custFile, {
        mode: type,
        totalAmountColIdx: selectedTotalAmountColIdx === -1 ? undefined : selectedTotalAmountColIdx,
        channelCode: code,
      });
      setCustomers(parsed);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "Lỗi đọc file khách hàng.");
      setCustomers([]);
    } finally {
      setIsReadingCustomers(false);
    }
  };

  const handleFileTypeChange = (type: "normalized" | "sales_detail") => {
    setFileType(type);
    void handleManualSettingsChange(type, targetChannelCode);
  };

  const handleChannelCodeChange = (code: number) => {
    setTargetChannelCode(code);
    void handleManualSettingsChange(fileType, code);
  };

  const handleColumnIndexChange = async (idx: number) => {
    if (!custFile) return;
    setSelectedTotalAmountColIdx(idx);
    setIsReadingCustomers(true);
    setError(null);
    try {
      const parsed = await parseCustomerExcel(custFile, {
        mode: fileType,
        totalAmountColIdx: idx === -1 ? undefined : idx,
        channelCode: targetChannelCode,
      });
      setCustomers(parsed);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "Lỗi đọc file khách hàng.");
      setCustomers([]);
    } finally {
      setIsReadingCustomers(false);
    }
  };

  // Add dynamic product
  const handleAddProduct = (e: React.FormEvent) => {
    e.preventDefault();
    setProdError(null);

    if (!newProdName.trim()) {
      setProdError("Tên sản phẩm không được để trống.");
      return;
    }
    if (!newProdPrice.trim()) {
      setProdError("Giá bán không được để trống.");
      return;
    }
    const price = parseMoneyInput(newProdPrice);
    if (price < 0) {
      setProdError("Giá bán phải lớn hơn hoặc bằng 0.");
      return;
    }
    if (!newProdThreshold.trim()) {
      setProdError("Ngưỡng tiền mua không được để trống.");
      return;
    }
    const threshold = parseMoneyInput(newProdThreshold);
    if (threshold < 0) {
      setProdError("Ngưỡng tiền mua phải lớn hơn hoặc bằng 0.");
      return;
    }
    const qtyA = parseFloat(newProdQtyA);
    if (isNaN(qtyA) || qtyA < 0) {
      setProdError("Số lượng A phải là số lớn hơn hoặc bằng 0.");
      return;
    }
    const qtyB = parseFloat(newProdQtyB);
    if (isNaN(qtyB) || qtyB < 0) {
      setProdError("Số lượng B phải là số lớn hơn hoặc bằng 0.");
      return;
    }

    const newProd: SalesPlanProduct = {
      name: newProdName.trim(),
      price,
      amountThreshold: threshold,
      quantityA: qtyA,
      quantityB: qtyB,
      ecomInitialQty: 0,
    };

    setProducts([...products, newProd]);
    setNewProdName("");
    setNewProdPrice("");
    setNewProdThreshold("");
    setNewProdQtyA("");
    setNewProdQtyB("");
  };

  // Delete product
  const handleDeleteProduct = (index: number) => {
    if (products.length <= 1) {
      setProdError("Kế hoạch Sales Plan phải có ít nhất 1 sản phẩm.");
      return;
    }
    setProducts(products.filter((_, idx) => idx !== index));
    if (editingIndex === index) {
      setEditingIndex(null);
    }
  };

  // Start inline editing
  const startEdit = (index: number) => {
    const prod = products[index]!;
    setEditingIndex(index);
    setEditName(prod.name);
    setEditPrice(formatMoneyInput(String(prod.price)));
    setEditThreshold(formatMoneyInput(String(prod.amountThreshold)));
    setEditQtyA(String(prod.quantityA));
    setEditQtyB(String(prod.quantityB));
  };

  // Save inline editing
  const saveEdit = (index: number) => {
    setProdError(null);
    if (!editName.trim()) {
      setProdError("Tên sản phẩm không được để trống.");
      return;
    }
    if (!editPrice.trim()) {
      setProdError("Giá bán không được để trống.");
      return;
    }
    const price = parseMoneyInput(editPrice);
    if (price < 0) {
      setProdError("Giá bán phải lớn hơn hoặc bằng 0.");
      return;
    }
    if (!editThreshold.trim()) {
      setProdError("Ngưỡng tiền mua không được để trống.");
      return;
    }
    const threshold = parseMoneyInput(editThreshold);
    if (threshold < 0) {
      setProdError("Ngưỡng tiền mua phải lớn hơn hoặc bằng 0.");
      return;
    }
    const qtyA = parseFloat(editQtyA);
    if (isNaN(qtyA) || qtyA < 0) {
      setProdError("Số lượng A phải là số lớn hơn hoặc bằng 0.");
      return;
    }
    const qtyB = parseFloat(editQtyB);
    if (isNaN(qtyB) || qtyB < 0) {
      setProdError("Số lượng B phải là số lớn hơn hoặc bằng 0.");
      return;
    }

    const prod = products[index]!;
    const updated = [...products];
    updated[index] = {
      name: editName.trim(),
      price,
      amountThreshold: threshold,
      quantityA: qtyA,
      quantityB: qtyB,
      ecomInitialQty: prod.ecomInitialQty,
    };
    setProducts(updated);
    setEditingIndex(null);
  };

  const cancelEdit = () => {
    setEditingIndex(null);
    setProdError(null);
  };

  // Calculate timeline start labels
  const startMonthName = useMemo(() => {
    const idx = planData.startMonthIndex;
    const names = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
    return `${names[idx]} ${planData.startYear}`;
  }, [planData]);

  // Compute aggregated sales preview data
  const top10Customers = useMemo(() => {
    return [...customers]
      .sort((a, b) => b.customerTotalAmount - a.customerTotalAmount)
      .slice(0, 10);
  }, [customers]);

  const maxRevenue = useMemo(() => {
    if (top10Customers.length === 0) return 1;
    return top10Customers[0]!.customerTotalAmount || 1;
  }, [top10Customers]);

  const totalAmountSum = useMemo(() => {
    return customers.reduce((sum, c) => sum + c.customerTotalAmount, 0);
  }, [customers]);

  return (
    <main className="masterplan-layout">
      {/* SIDEBAR CONFIGURATION */}
      <aside className="control-panel glass" style={{ maxHeight: "calc(100vh - 100px)", overflowY: "auto" }}>
        {/* EXCEL CUSTOMER IMPORT */}
        <div className="panel-section">
          <span className="section-label">
            <FileSpreadsheet size={16} /> 1. File Khách Hàng
          </span>

          {/* FILE TYPE SELECTOR */}
          <div className="form-group" style={{ marginBottom: 12 }}>
            <label className="field-label" style={{ fontSize: 12, fontWeight: 600 }}>Loại file đầu vào:</label>
            <div style={{ display: "flex", gap: 6, marginTop: 4 }}>
              <button
                type="button"
                className={`view-btn ${fileType === "normalized" ? "active" : ""}`}
                style={{ flex: 1, padding: "6px", fontSize: 11, minHeight: 30, borderRadius: 4 }}
                onClick={() => handleFileTypeChange("normalized")}
              >
                File chuẩn hóa sẵn
              </button>
              <button
                type="button"
                className={`view-btn ${fileType === "sales_detail" ? "active" : ""}`}
                style={{ flex: 1, padding: "6px", fontSize: 11, minHeight: 30, borderRadius: 4 }}
                onClick={() => handleFileTypeChange("sales_detail")}
              >
                Báo cáo bán hàng chi tiết
              </button>
            </div>
          </div>

          {/* DYNAMIC DEFAULT TARGET CHANNEL SELECTOR */}
          {fileType === "sales_detail" && (
            <div className="form-group" style={{ marginBottom: 12 }}>
              <label className="field-label" htmlFor="channel-select" style={{ fontSize: 12, fontWeight: 600 }}>
                Kênh bán hàng mặc định:
              </label>
              <select
                id="channel-select"
                value={targetChannelCode}
                onChange={(e) => handleChannelCodeChange(parseInt(e.target.value))}
                className="input-field"
                style={{ width: "100%", padding: "8px", fontSize: 13, border: "1px solid var(--color-border)", borderRadius: "6px" }}
              >
                {CHANNEL_OPTIONS.map((ch) => (
                  <option key={ch.code} value={ch.code}>
                    {ch.code} | {ch.channel} | {ch.subChannel}
                  </option>
                ))}
              </select>
            </div>
          )}

          <input
            ref={fileInputRef}
            className="file-input"
            type="file"
            accept=".xlsx,.xlsm"
            onChange={(e) => void handleCustomerFileChange(e.target.files?.[0] ?? null)}
          />

          <button
            className="secondary-button full-width"
            type="button"
            onClick={() => fileInputRef.current?.click()}
            disabled={isReadingCustomers}
            style={{ marginBottom: 12 }}
          >
            {isReadingCustomers ? <Loader2 className="spin" /> : <Upload />}
            {isReadingCustomers ? "Đang đọc..." : "Upload File Khách Hàng"}
          </button>

          {fileType === "normalized" && excelHeaders.length > 0 && (
            <div className="form-group" style={{ marginBottom: 12 }}>
              <label className="field-label" htmlFor="col-select" style={{ fontSize: 12, fontWeight: 600 }}>
                Chọn cột Tổng tiền đã mua:
              </label>
              <select
                id="col-select"
                value={selectedTotalAmountColIdx}
                onChange={(e) => void handleColumnIndexChange(parseInt(e.target.value))}
                className="input-field"
                style={{ width: "100%", padding: "8px", fontSize: 13, border: "1px solid var(--color-border)", borderRadius: "6px" }}
              >
                <option value="-1">-- Tự động nhận diện --</option>
                {excelHeaders.map((header, idx) => (
                  <option key={idx} value={idx}>
                    Cột {String.fromCharCode(65 + idx)}: {header || `(Cột trống ${idx + 1})`}
                  </option>
                ))}
              </select>
            </div>
          )}

          {custFile ? (
            <div className="file-state" style={{ marginTop: 8 }}>
              <div>
                <strong>{custFile.name}</strong>
                <span className="text-small font-mono text-muted">{customers.length} khách hàng đã nạp</span>
              </div>
            </div>
          ) : (
            <span className="field-hint">
              💡 Chưa chọn file. Vui lòng upload file để hiển thị và xuất kế hoạch.
            </span>
          )}

          {error && (
            <div className="field-hint warning" style={{ marginTop: 10 }}>
              ⚠️ {error}
            </div>
          )}
        </div>

        {/* ECOM CHANNEL CONFIGURATION */}
        <div className="panel-section" style={{ borderTop: "1px dashed rgba(0,0,0,0.1)", paddingTop: 14 }}>
          <span className="section-label" style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <Sliders size={16} /> Cấu hình Kênh Ecom
          </span>
          <div className="form-group" style={{ marginTop: 8 }}>
            <label style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer", fontSize: 13, fontWeight: 600 }}>
              <input
                type="checkbox"
                checked={ecomEnabled}
                onChange={(e) => setEcomEnabled(e.target.checked)}
                style={{ width: 18, height: 18, accentColor: "var(--color-primary)", cursor: "pointer" }}
              />
              <span>Thêm kênh Ecom (1003 | Ecom)</span>
            </label>
            <span className="field-hint" style={{ marginTop: 6, display: "block", fontSize: 11, color: "var(--color-text-muted)" }}>
              💡 Khi chọn, kênh Ecom tự động chèn vào kế hoạch với công thức tăng trưởng 20% mỗi tháng tiếp theo.
            </span>
          </div>

          {ecomEnabled && products.length > 0 && (
            <div style={{ marginTop: 12, borderTop: "1px dashed rgba(0,0,0,0.06)", paddingTop: 10 }}>
              <label className="field-label" style={{ fontSize: 12, fontWeight: 700, marginBottom: 8, display: "block", color: "var(--color-primary)" }}>
                Số lượng Month 1 của từng sản phẩm:
              </label>
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                {products.map((prod, idx) => (
                  <div key={idx} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, padding: "2px 0" }}>
                    <span style={{ fontSize: 12, fontWeight: 500, color: "var(--color-text-secondary)", flex: 1, textOverflow: "ellipsis", overflow: "hidden", whiteSpace: "nowrap" }} title={prod.name}>
                      {prod.name}
                    </span>
                    <input
                      type="number"
                      placeholder="0"
                      min="0"
                      value={prod.ecomInitialQty ?? ""}
                      onChange={(e) => {
                        const val = parseFloat(e.target.value);
                        const updated = [...products];
                        updated[idx] = {
                          ...prod,
                          ecomInitialQty: isNaN(val) ? undefined : val
                        };
                        setProducts(updated);
                      }}
                      className="input-field"
                      style={{ width: 80, padding: "4px 8px", fontSize: 12, minHeight: 28, textAlign: "center", border: "1px solid var(--color-border)", borderRadius: "6px" }}
                    />
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>

        {/* DYNAMIC PRODUCT DEFINITION FORM */}
        <div className="panel-section">
          <span className="section-label">
            <Plus size={16} /> 2. Thêm Sản phẩm
          </span>

          <form onSubmit={handleAddProduct} className="product-form">
            {prodError && <div className="field-hint warning" style={{ marginBottom: 10 }}>{prodError}</div>}

            <div className="form-group">
              <label className="field-label" htmlFor="p-name">Tên sản phẩm</label>
              <input
                id="p-name"
                type="text"
                placeholder="Ví dụ: SPR Milky Calcium"
                value={newProdName}
                onChange={(e) => setNewProdName(e.target.value)}
                className="input-field"
              />
            </div>

            <div className="form-group">
              <label className="field-label" htmlFor="p-price">Giá bán sản phẩm</label>
              <div className="input-with-suffix">
                <input
                  id="p-price"
                  type="text"
                  placeholder="745,000"
                  value={newProdPrice}
                  onChange={(e) => setNewProdPrice(formatMoneyInput(e.target.value))}
                />
                <span className="suffix">đ</span>
              </div>
            </div>

            <div className="form-group">
              <label className="field-label" htmlFor="p-threshold">Ngưỡng tiền mua (Amount threshold)</label>
              <div className="input-with-suffix">
                <input
                  id="p-threshold"
                  type="text"
                  placeholder="1,500,000"
                  value={newProdThreshold}
                  onChange={(e) => setNewProdThreshold(formatMoneyInput(e.target.value))}
                />
                <span className="suffix">đ</span>
              </div>
            </div>

            <div style={{ display: "flex", gap: 8 }}>
              <div className="form-group" style={{ flex: 1 }}>
                <label className="field-label" htmlFor="p-qtya">SL A (Vượt ngưỡng)</label>
                <input
                  id="p-qtya"
                  type="number"
                  placeholder="10"
                  value={newProdQtyA}
                  onChange={(e) => setNewProdQtyA(e.target.value)}
                  className="input-field"
                />
              </div>

              <div className="form-group" style={{ flex: 1 }}>
                <label className="field-label" htmlFor="p-qtyb">SL B (Dưới/Bằng)</label>
                <input
                  id="p-qtyb"
                  type="number"
                  placeholder="2"
                  value={newProdQtyB}
                  onChange={(e) => setNewProdQtyB(e.target.value)}
                  className="input-field"
                />
              </div>
            </div>

            <button type="submit" className="secondary-button full-width" style={{ marginTop: 8 }}>
              <Plus size={16} /> Thêm sản phẩm
            </button>
          </form>
        </div>

        {/* METADATA SUMMARY & EXPORT BUTTON */}
        <div className="panel-section">
          <span className="section-label">
            <DollarSign size={16} /> 3. Xuất Sales Plan
          </span>

          <div className="kpi-card-stack" style={{ marginBottom: 16 }}>
            <div className="kpi-card">
              <span className="kpi-title">Tháng Bắt đầu</span>
              <strong className="kpi-value highlighted">{startMonthName}</strong>
              <div className="kpi-badge primary">Quy tắc M+1</div>
            </div>

            <div className="kpi-card">
              <span className="kpi-title">Tổng số khách hàng</span>
              <strong className="kpi-value">{customers.length}</strong>
              <div className="kpi-badge font-mono">Khách hàng</div>
            </div>
          </div>

          <button
            className="primary-button full-width"
            type="button"
            onClick={() => onExport(planData)}
            disabled={isExporting || products.length === 0}
          >
            {isExporting ? <Loader2 className="spin" /> : <Download />}
            {isExporting ? "Đang tạo Excel..." : "Xuất Excel Sales Plan"}
          </button>
        </div>
      </aside>

      {/* WORKSPACE PREVIEW MAIN PANEL */}
      <section className="preview-panel workspace-main glass" style={{ maxHeight: "calc(100vh - 100px)", overflowY: "auto" }}>
        
        {/* PARSE PREVIEW KPI SUMMARY AND CHART RANKING */}
        {custFile && (
          <div className="panel-section" style={{ margin: "16px", padding: "16px", border: "1px solid var(--color-border)", borderRadius: "12px", background: "rgba(255,255,255,0.02)", boxShadow: "0 4px 20px rgba(0,0,0,0.05)" }}>
            <span className="section-label" style={{ marginBottom: 12, display: "flex", alignItems: "center", gap: 6 }}>
              <BarChart3 size={15} /> 📊 Kết quả Phân tích & Gom nhóm
            </span>
            
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12, marginBottom: 16 }}>
              <div className="kpi-card" style={{ background: "rgba(30,58,138,0.03)", border: "1px solid rgba(30,58,138,0.08)" }}>
                <span className="kpi-title">Số dòng thô đọc được</span>
                <strong className="kpi-value">{formatNumber(originalRowCount)} dòng</strong>
                <span className="text-small text-muted" style={{ fontSize: 10, marginTop: 4 }}>Dữ liệu giao dịch gốc</span>
              </div>
              <div className="kpi-card" style={{ background: "rgba(15,107,95,0.03)", border: "1px solid rgba(15,107,95,0.08)" }}>
                <span className="kpi-title">Khách hàng unique</span>
                <strong className="kpi-value highlighted" style={{ color: "var(--color-primary)" }}>{customers.length} khách hàng</strong>
                <span className="text-small text-muted" style={{ fontSize: 10, marginTop: 4 }}>Đã gom nhóm thành công</span>
              </div>
              <div className="kpi-card" style={{ background: "rgba(22,163,74,0.03)", border: "1px solid rgba(22,163,74,0.08)" }}>
                <span className="kpi-title">Tổng doanh thu chi tiết</span>
                <strong className="kpi-value" style={{ color: "#16a34a" }}>{formatNumber(totalAmountSum)} đ</strong>
                <span className="text-small text-muted" style={{ fontSize: 10, marginTop: 4 }}>Cộng dồn Doanh thu chi tiết</span>
              </div>
            </div>

            {/* TOP 10 RANKING CHART */}
            {top10Customers.length > 0 && (
              <div style={{ marginTop: 16, borderTop: "1px dashed var(--color-border)", paddingTop: 16 }}>
                <h4 style={{ fontSize: 13, fontWeight: 600, color: "var(--color-text-primary)", marginBottom: 12 }}>🏆 Top 10 Khách hàng có Doanh thu cao nhất</h4>
                <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                  {top10Customers.map((c, idx) => {
                    const percent = Math.min(100, Math.max(0, (c.customerTotalAmount / maxRevenue) * 100));
                    return (
                      <div key={idx} style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, fontWeight: 500 }}>
                          <span style={{ color: "var(--color-text-secondary)" }}>{idx + 1}. {c.customer}</span>
                          <span className="font-mono bold" style={{ color: "#16a34a" }}>{formatNumber(c.customerTotalAmount)} đ</span>
                        </div>
                        <div style={{ width: "100%", height: 8, background: "rgba(226,232,240,0.4)", borderRadius: 4, overflow: "hidden" }}>
                          <div style={{ width: `${percent}%`, height: "100%", background: "linear-gradient(90deg, #10b981, #059669)", borderRadius: 4 }} />
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </div>
        )}

        {/* PRODUCT LIST MANAGEMENT */}
        <div className="preview-head" style={{ borderBottom: "none" }}>
          <div>
            <span className="section-label">Danh mục Sản phẩm Động ({products.length})</span>
            <h2>Danh sách sản phẩm & Công thức SL theo Ngưỡng Doanh Số</h2>
          </div>
        </div>

        <div style={{ padding: "0 16px 16px" }}>
          <div className="table-wrap notion-spreadsheet" style={{ maxHeight: 250 }}>
            <table className="data-table">
              <thead>
                <tr>
                  <th style={{ width: 40 }}>STT</th>
                  <th className="text-left" style={{ minWidth: 200 }}>Tên sản phẩm & Công thức</th>
                  <th>Giá bán</th>
                  <th>Ngưỡng mua</th>
                  <th>SL A (Vượt)</th>
                  <th>SL B (Dưới/Bằng)</th>
                  {ecomEnabled && <th style={{ background: "#eff6ff", color: "#1e40af" }}>SL đầu Ecom</th>}
                  <th style={{ width: 100 }}>Hành động</th>
                </tr>
              </thead>
              <tbody>
                {products.map((p, idx) => (
                  <tr key={idx}>
                    <td className="font-mono text-muted text-center">{idx + 1}</td>
                    {editingIndex === idx ? (
                      <>
                        <td className="text-left editable-cell">
                          <input
                            type="text"
                            className="cell-input"
                            value={editName}
                            onChange={(e) => setEditName(e.target.value)}
                          />
                        </td>
                        <td className="editable-cell">
                          <input
                            type="text"
                            className="cell-input text-right font-mono"
                            value={editPrice}
                            onChange={(e) => setEditPrice(formatMoneyInput(e.target.value))}
                          />
                        </td>
                        <td className="editable-cell">
                          <input
                            type="text"
                            className="cell-input text-right font-mono"
                            value={editThreshold}
                            onChange={(e) => setEditThreshold(formatMoneyInput(e.target.value))}
                          />
                        </td>
                        <td className="editable-cell">
                          <input
                            type="number"
                            className="cell-input text-center font-mono"
                            value={editQtyA}
                            onChange={(e) => setEditQtyA(e.target.value)}
                          />
                        </td>
                        <td className="editable-cell">
                          <input
                            type="number"
                            className="cell-input text-center font-mono"
                            value={editQtyB}
                            onChange={(e) => setEditQtyB(e.target.value)}
                          />
                        </td>
                        {ecomEnabled && (
                          <td style={{ background: "#f8fafc", color: "#64748b", textAlign: "center" }} className="font-mono text-center">
                            {p.ecomInitialQty ?? 0}
                          </td>
                        )}
                        <td>
                          <div style={{ display: "flex", gap: 4, justifyContent: "center" }}>
                            <button
                              type="button"
                              className="nav-btn"
                              onClick={() => saveEdit(idx)}
                              style={{ height: 26, width: 26, color: "#16a34a", border: "none" }}
                            >
                              <Save size={14} />
                            </button>
                            <button
                              type="button"
                              className="nav-btn"
                              onClick={cancelEdit}
                              style={{ height: 26, width: 26, color: "#64748b", border: "none" }}
                            >
                              <X size={14} />
                            </button>
                          </div>
                        </td>
                      </>
                    ) : (
                      <>
                        <td className="text-left bold text-muted">
                          {p.name}
                          <div className="text-small text-muted" style={{ fontWeight: "normal", marginTop: 4, fontSize: 11 }}>
                            💡 Nếu tổng mua KH &gt; <strong style={{ color: "#1e3a8a" }}>{formatNumber(p.amountThreshold)} đ</strong> thì SL = <strong style={{ color: "#0f6b5f" }}>{p.quantityA}</strong>, ngược lại SL = <strong style={{ color: "#ef4444" }}>{p.quantityB}</strong>
                          </div>
                        </td>
                        <td className="font-mono text-right">{formatNumber(p.price)} đ</td>
                        <td className="font-mono text-right">{formatNumber(p.amountThreshold)} đ</td>
                        <td className="font-mono text-center bold" style={{ color: "#0f6b5f" }}>{p.quantityA}</td>
                        <td className="font-mono text-center bold" style={{ color: "#ef4444" }}>{p.quantityB}</td>
                        {ecomEnabled && (
                          <td className="font-mono text-center bold" style={{ background: "#f8fafc", color: "#3b82f6" }}>
                            {p.ecomInitialQty ?? 0}
                          </td>
                        )}
                        <td>
                          <div style={{ display: "flex", gap: 4, justifyContent: "center" }}>
                            <button
                              type="button"
                              className="nav-btn"
                              onClick={() => startEdit(idx)}
                              title="Sửa sản phẩm"
                              style={{ height: 26, width: 26, color: "#0e7c7b", border: "none" }}
                            >
                              <Edit size={14} />
                            </button>
                            <button
                              type="button"
                              className="nav-btn"
                              onClick={() => handleDeleteProduct(idx)}
                              title="Xóa sản phẩm"
                              style={{ height: 26, width: 26, color: "#ef4444", border: "none" }}
                            >
                              <Trash2 size={14} />
                            </button>
                          </div>
                        </td>
                      </>
                    )}
                  </tr>
                ))}
                {products.length === 0 && (
                  <tr>
                    <td colSpan={7} className="text-center text-muted" style={{ padding: 16 }}>
                      Chưa có sản phẩm. Vui lòng thêm sản phẩm ở form "2. Thêm Sản phẩm" bên trái!
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>

        {/* CUSTOMERS PREVIEW TABLE */}
        <div className="preview-head" style={{ paddingTop: 8 }}>
          <div>
            <span className="section-label">Danh sách khách hàng nạp được</span>
            <h2>Dữ liệu khách hàng & Doanh số input (Áp số lượng dòng 18 trở đi)</h2>
          </div>
          <div className="customer-info-badge font-mono" style={{ background: custFile ? "#ccfbf1" : "#f1f5f9", color: custFile ? "#0d9488" : "#64748b" }}>
            {custFile ? `✅ Khớp: ${customers.length} khách hàng` : "💡 Chưa nạp file khách hàng"}
          </div>
        </div>

        <div className="planning-grid-container">
          <div className="table-wrap notion-spreadsheet" style={{ maxHeight: 280 }}>
            <table className="data-table">
              <thead>
                <tr>
                  <th style={{ width: 40 }}>STT</th>
                  <th>Code channel (A)</th>
                  <th>Channel (B)</th>
                  <th>Sub-channel (C)</th>
                  <th className="text-left">Customer (D)</th>
                  <th>Stores (E)</th>
                  <th>Province (F)</th>
                  <th>Regional (G)</th>
                  <th>Sub-Channel 1 (H)</th>
                  <th>STAFF (I)</th>
                  <th style={{ background: "#f0fdf4", color: "#166534", minWidth: 140 }}>Tổng tiền đã mua</th>
                </tr>
              </thead>
              <tbody>
                {planData.customers.map((c, idx) => (
                  <tr key={idx}>
                    <td className="font-mono text-muted text-center">{idx + 18}</td>
                    <td className="font-mono bold">{c.codeChannel}</td>
                    <td>{c.channel}</td>
                    <td>{c.subChannel}</td>
                    <td className="text-left bold text-muted">{c.customer}</td>
                    <td>{c.stores}</td>
                    <td>{c.province}</td>
                    <td>{c.regional}</td>
                    <td>{c.subChannel1}</td>
                    <td>{c.staff}</td>
                    <td className="font-mono bold text-right" style={{ background: "#f9fefb", color: "#16a34a" }}>
                      {formatNumber(c.customerTotalAmount)} đ
                    </td>
                  </tr>
                ))}
                {planData.customers.length === 0 && (
                  <tr>
                    <td colSpan={11} className="text-center text-muted" style={{ padding: 24 }}>
                      Chưa nạp khách hàng. Vui lòng upload file khách hàng ở bên trái!
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>

        <div className="planning-tips">
          ⚠️ <strong>Quy tắc phân bổ số lượng Sales Plan:</strong> Số lượng mỗi sản phẩm của từng khách hàng sẽ được tính toán trực tiếp từ cột **Tổng tiền đã mua** của file input theo công thức ngưỡng riêng của sản phẩm đó.
        </div>
      </section>
    </main>
  );
}
