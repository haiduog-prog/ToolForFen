import { AlertTriangle, CheckCircle2, Download, FileSpreadsheet, Loader2, Upload } from "lucide-react";
import { useMemo, useRef, useState } from "react";
import { buildQecReport } from "../application/buildQecReport";
import { type MetricRow, type QecReport, type SourceParseResult } from "../domain/entities";
import { displayMonth, type MonthKey } from "../domain/month";
import { readSourceWorkbook } from "../infrastructure/excel/readSourceWorkbook";
import { writeReportWorkbook } from "../infrastructure/excel/writeReportWorkbook";
import { downloadBlob, formatNumber, formatRatio } from "../shared/formatters";

type PreviewTab = "qec" | "sku" | "customerRevenue" | "customerQuantity";

export function App() {
  const inputRef = useRef<HTMLInputElement | null>(null);
  const [source, setSource] = useState<SourceParseResult | null>(null);
  const [reportMonth, setReportMonth] = useState<MonthKey | "">("");
  const [activeTab, setActiveTab] = useState<PreviewTab>("qec");
  const [isReading, setIsReading] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const report = useMemo(() => {
    if (!source || !reportMonth) {
      return null;
    }

    return buildQecReport(source, reportMonth);
  }, [source, reportMonth]);

  async function handleFile(file: File | null) {
    if (!file) {
      return;
    }

    setIsReading(true);
    setError(null);

    try {
      const parsed = await readSourceWorkbook(file);
      setSource(parsed);
      setReportMonth(parsed.availableMonths.at(-1) ?? "");
      setActiveTab("qec");
    } catch (caught) {
      setSource(null);
      setReportMonth("");
      setError(caught instanceof Error ? caught.message : "Không đọc được file Excel.");
    } finally {
      setIsReading(false);
      if (inputRef.current) {
        inputRef.current.value = "";
      }
    }
  }

  async function handleExport() {
    if (!report || !source) {
      return;
    }

    setIsExporting(true);
    setError(null);

    try {
      const blob = await writeReportWorkbook(report, source.transactions);
      downloadBlob(blob, `QEC_${displayMonth(report.reportMonth).replace(/\s/g, "")}.xlsx`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "Không export được file Excel.");
    } finally {
      setIsExporting(false);
    }
  }

  return (
    <main className="app-shell">
      <section className="topbar">
        <div className="brand">
          <FileSpreadsheet aria-hidden="true" />
          <div>
            <h1>QEC Export Builder</h1>
            <p>Excel review export</p>
          </div>
        </div>

        <div className="actions">
          <input
            ref={inputRef}
            className="file-input"
            type="file"
            accept=".xlsx,.xlsm"
            onChange={(event) => void handleFile(event.target.files?.[0] ?? null)}
          />
          <button className="secondary-button" type="button" onClick={() => inputRef.current?.click()} disabled={isReading}>
            {isReading ? <Loader2 className="spin" aria-hidden="true" /> : <Upload aria-hidden="true" />}
            {isReading ? "Đang đọc" : "Upload Excel"}
          </button>
          <button className="primary-button" type="button" onClick={() => void handleExport()} disabled={!report || isExporting}>
            {isExporting ? <Loader2 className="spin" aria-hidden="true" /> : <Download aria-hidden="true" />}
            Export
          </button>
        </div>
      </section>

      {error ? (
        <section className="notice error-notice">
          <AlertTriangle aria-hidden="true" />
          <span>{error}</span>
        </section>
      ) : null}

      <section className="workspace">
        <aside className="control-panel">
          <div className="panel-section">
            <span className="section-label">Nguồn dữ liệu</span>
            {source ? (
              <div className="file-state">
                <CheckCircle2 aria-hidden="true" />
                <div>
                  <strong>{source.fileName}</strong>
                  <span>{formatNumber(source.transactions.length)} dòng giao dịch</span>
                </div>
              </div>
            ) : (
              <div className="empty-state">
                <FileSpreadsheet aria-hidden="true" />
                <span>Chưa có file</span>
              </div>
            )}
          </div>

          <div className="panel-section">
            <label className="field-label" htmlFor="report-month">
              Tháng báo cáo
            </label>
            <select
              id="report-month"
              value={reportMonth}
              onChange={(event) => setReportMonth(event.target.value as MonthKey)}
              disabled={!source || source.availableMonths.length === 0}
            >
              {!source ? <option value="">Upload file trước</option> : null}
              {source?.availableMonths.map((month) => (
                <option key={month} value={month}>
                  {displayMonth(month)}
                </option>
              ))}
            </select>
          </div>

          {report ? <Summary report={report} /> : null}

          {source?.warnings.length ? (
            <div className="panel-section">
              <span className="section-label">Cảnh báo</span>
              <div className="warning-list">
                {source.warnings.map((warning) => (
                  <div className="warning-item" key={warning}>
                    <AlertTriangle aria-hidden="true" />
                    <span>{warning}</span>
                  </div>
                ))}
              </div>
            </div>
          ) : null}
        </aside>

        <section className="preview-panel">
          {report ? (
            <>
              <div className="preview-head">
                <div>
                  <span className="section-label">Preview</span>
                  <h2>{displayMonth(report.reportMonth)}</h2>
                </div>
                <TabBar activeTab={activeTab} onChange={setActiveTab} />
              </div>

              {activeTab === "qec" ? (
                <MetricTable
                  rows={report.qecRows}
                  report={report}
                  labelHeader="Segment"
                  showShares
                  valuePrecision={0}
                />
              ) : null}

              {activeTab === "sku" ? (
                <MetricTable
                  rows={report.skuRevenueRows.slice(0, 50)}
                  report={report}
                  labelHeader="BRAND_OF_PRODUCT (PCS)"
                  valuePrecision={0}
                />
              ) : null}

              {activeTab === "customerRevenue" ? (
                <CustomerPreview report={report} kind="revenue" />
              ) : null}

              {activeTab === "customerQuantity" ? (
                <CustomerPreview report={report} kind="quantity" />
              ) : null}
            </>
          ) : (
            <div className="start-panel">
              <FileSpreadsheet aria-hidden="true" />
              <h2>Upload Data nguồn</h2>
              <p>Ứng dụng sẽ nhận diện các tháng có trong file và tạo workbook review.</p>
            </div>
          )}
        </section>
      </section>
    </main>
  );
}

function Summary({ report }: { report: QecReport }) {
  const items = [
    ["Giao dịch", formatNumber(report.summary.transactionCount)],
    ["Khách hàng", formatNumber(report.summary.customerCount)],
    ["Sản phẩm", formatNumber(report.summary.productCount)],
    ["Unmapped", formatNumber(report.summary.unmappedCustomerCount)]
  ];

  return (
    <div className="panel-section">
      <span className="section-label">Tổng quan</span>
      <div className="summary-grid">
        {items.map(([label, value]) => (
          <div className="summary-cell" key={label}>
            <span>{label}</span>
            <strong>{value}</strong>
          </div>
        ))}
      </div>
      {report.summary.unmappedCustomers.length > 0 ? (
        <details className="unmapped-details">
          <summary>Danh sách Unmapped</summary>
          <div>{report.summary.unmappedCustomers.slice(0, 40).join(", ")}</div>
        </details>
      ) : null}
    </div>
  );
}

function TabBar({ activeTab, onChange }: { activeTab: PreviewTab; onChange: (tab: PreviewTab) => void }) {
  const tabs: Array<[PreviewTab, string]> = [
    ["qec", "QEC"],
    ["sku", "SKU"],
    ["customerRevenue", "Customer revenue"],
    ["customerQuantity", "Customer quantity"]
  ];

  return (
    <div className="tabs" role="tablist" aria-label="Preview sheets">
      {tabs.map(([key, label]) => (
        <button
          key={key}
          type="button"
          role="tab"
          aria-selected={activeTab === key}
          className={activeTab === key ? "tab active" : "tab"}
          onClick={() => onChange(key)}
        >
          {label}
        </button>
      ))}
    </div>
  );
}

function MetricTable({
  rows,
  report,
  labelHeader,
  showShares = false,
  valuePrecision
}: {
  rows: MetricRow[];
  report: QecReport;
  labelHeader: string;
  showShares?: boolean;
  valuePrecision: number;
}) {
  return (
    <div className="table-wrap">
      <table className="data-table">
        <thead>
          <tr>
            <th>{labelHeader}</th>
            {report.periodMonths.map((month) => (
              <th key={month}>{displayMonth(month)}</th>
            ))}
            {showShares ? (
              <>
                <th>Share {report.previousYear}</th>
                <th>Share {report.currentYear}</th>
              </>
            ) : null}
            <th>{report.previousYear}</th>
            <th>{report.currentYear}</th>
            <th>P3M</th>
            <th>P6M</th>
            <th>P9M</th>
            <th>TREND</th>
            <th>IFYTD</th>
            <th>ICYTD</th>
            <th>IYA</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((row) => (
            <tr key={row.label} className={row.label === "Total" ? "total-row" : undefined}>
              <th>{row.label}</th>
              {report.periodMonths.map((month) => (
                <td key={month}>{formatNumber(row.monthValues[month] ?? 0, valuePrecision)}</td>
              ))}
              {showShares ? (
                <>
                  <td>{formatRatio(row.previousYearShare ?? 0)}</td>
                  <td>{formatRatio(row.currentYearShare ?? 0)}</td>
                </>
              ) : null}
              <td>{formatNumber(row.previousYearTotal, valuePrecision)}</td>
              <td>{formatNumber(row.currentYearTotal, valuePrecision)}</td>
              <td>{formatNumber(row.p3m, valuePrecision)}</td>
              <td>{formatNumber(row.p6m, valuePrecision)}</td>
              <td>{formatNumber(row.p9m, valuePrecision)}</td>
              <td>{formatRatio(row.trend)}</td>
              <td>{formatRatio(row.ifytd)}</td>
              <td>{formatRatio(row.icytd)}</td>
              <td>{formatRatio(row.iya)}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function CustomerPreview({ report, kind }: { report: QecReport; kind: "revenue" | "quantity" }) {
  const sections = kind === "revenue" ? report.customerRevenueSections : report.customerQuantitySections;
  const rows = sections.slice(0, 12).flatMap((section) => {
    return section.rows.slice(0, 6).map((row) => ({
      customer: section.customer,
      row
    }));
  });

  return (
    <div className="table-wrap">
      <table className="data-table">
        <thead>
          <tr>
            <th>Customer</th>
            <th>Name SKU</th>
            {report.periodMonths.map((month) => (
              <th key={month}>{displayMonth(month)}</th>
            ))}
            <th>{report.previousYear}</th>
            <th>{report.currentYear}</th>
            <th>P3M</th>
            <th>P6M</th>
            <th>P9M</th>
            <th>TREND</th>
          </tr>
        </thead>
        <tbody>
          {rows.map(({ customer, row }) => (
            <tr key={`${customer}-${row.label}`}>
              <th>{customer}</th>
              <th>{row.label}</th>
              {report.periodMonths.map((month) => (
                <td key={month}>{formatNumber(row.monthValues[month] ?? 0, kind === "quantity" ? 2 : 0)}</td>
              ))}
              <td>{formatNumber(row.previousYearTotal, kind === "quantity" ? 2 : 0)}</td>
              <td>{formatNumber(row.currentYearTotal, kind === "quantity" ? 2 : 0)}</td>
              <td>{formatNumber(row.p3m, kind === "quantity" ? 2 : 0)}</td>
              <td>{formatNumber(row.p6m, kind === "quantity" ? 2 : 0)}</td>
              <td>{formatNumber(row.p9m, kind === "quantity" ? 2 : 0)}</td>
              <td>{formatRatio(row.trend)}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
