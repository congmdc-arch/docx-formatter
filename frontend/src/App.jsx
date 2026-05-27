import { useState, useCallback, useRef } from "react";

const API_BASE = "http://localhost:8000";

const STYLE_RULES = [
  { label: "Font", value: "Times New Roman" },
  { label: "Cỡ chữ thân", value: "14pt" },
  { label: "Line spacing", value: "1.5" },
  { label: "Heading 1", value: "14pt, HOA, đậm, căn giữa" },
  { label: "Heading 2", value: "14pt, HOA, đậm, căn trái" },
  { label: "Heading 3", value: "14pt, đậm, căn trái" },
  { label: "Heading 4", value: "14pt, đậm+nghiêng" },
  { label: "Tên hình", value: "13pt, nghiêng, căn giữa, bên dưới" },
  { label: "Tên bảng", value: "13pt, đậm, căn giữa, bên trên" },
  { label: "Đánh số hình", value: "Hình X.Y. theo chương" },
  { label: "Đánh số bảng", value: "Bảng X.Y. theo chương" },
  { label: "Margin", value: "T3.5 D3.0 T3.5 P2.0 cm" },
  { label: "Indent đầu dòng", value: "1.25 cm" },
  { label: "Font bảng biểu", value: "Times New Roman 13pt" },
];

export default function App() {
  const [file, setFile] = useState(null);
  const [status, setStatus] = useState("idle");
  const [log, setLog] = useState([]);
  const [inspect, setInspect] = useState(null);
  const [downloadUrl, setDownloadUrl] = useState(null);
  const [downloadName, setDownloadName] = useState("");
  const fileInputRef = useRef();

  const addLog = (msg, type = "info") =>
    setLog((prev) => [...prev, { msg, type }]);

  const handleFile = useCallback(async (f) => {
    if (!f?.name.endsWith(".docx")) {
      addLog("❌ Chỉ hỗ trợ .docx", "error");
      return;
    }
    setFile(f);
    setLog([]);
    setInspect(null);
    setDownloadUrl(null);
    setStatus("inspecting");

    // Inspect trước
    try {
      addLog(`📄 Đọc file: ${f.name}`);
      const fd1 = new FormData();
      fd1.append("file", f);
      const r1 = await fetch(`${API_BASE}/inspect`, { method: "POST", body: fd1 });
      if (!r1.ok) throw new Error(await r1.text());
      const info = await r1.json();
      setInspect(info);
      addLog(`✅ Phân tích xong — ${info.paragraphs_total} đoạn, ${info.tables} bảng`);
    } catch (e) {
      addLog(`⚠️ Inspect lỗi: ${e.message} — tiếp tục format`, "warn");
    }

    // Format
    setStatus("formatting");
    addLog("🔨 Đang chuẩn hóa định dạng...");
    try {
      const fd2 = new FormData();
      fd2.append("file", f);
      const r2 = await fetch(`${API_BASE}/format`, { method: "POST", body: fd2 });
      if (!r2.ok) {
        const err = await r2.json();
        throw new Error(err.detail || "Format lỗi");
      }
      const blob = await r2.blob();
      const url = URL.createObjectURL(blob);
      const name = f.name.replace(".docx", "_formatted.docx");
      setDownloadUrl(url);
      setDownloadName(name);
      setStatus("done");
      addLog(`✅ Hoàn thành! File sẵn sàng tải về.`, "success");
    } catch (e) {
      setStatus("error");
      addLog(`❌ ${e.message}`, "error");
    }
  }, []);

  const onDrop = useCallback(
    (e) => {
      e.preventDefault();
      handleFile(e.dataTransfer.files[0]);
    },
    [handleFile]
  );

  const [dragging, setDragging] = useState(false);

  const STATUS_COLOR = {
    idle: "#475569", inspecting: "#f59e0b",
    formatting: "#3b82f6", done: "#22c55e", error: "#ef4444",
  };
  const STATUS_LABEL = {
    idle: "Chờ file", inspecting: "Đang phân tích...",
    formatting: "Đang chuẩn hóa...", done: "Hoàn thành", error: "Lỗi",
  };
  const isProcessing = ["inspecting", "formatting"].includes(status);

  return (
    <div style={s.app}>
      <style>{css}</style>

      {/* Header */}
      <div style={s.header}>
        <div>
          <div style={s.headerTitle}>Đề Án Formatter</div>
          <div style={s.headerSub}>Chuẩn hóa định dạng đề án tốt nghiệp • Times New Roman</div>
        </div>
        <div style={{ ...s.statusPill, background: STATUS_COLOR[status] + "22", color: STATUS_COLOR[status], border: `1px solid ${STATUS_COLOR[status]}44` }}>
          <span style={{ ...s.statusDot, background: STATUS_COLOR[status], animation: isProcessing ? "pulse 1s infinite" : "none" }} />
          {STATUS_LABEL[status]}
        </div>
      </div>

      <div style={s.body}>
        {/* Left panel */}
        <div style={s.left}>
          {/* Drop zone */}
          <div
            style={{ ...s.dropZone, ...(dragging ? s.dropZoneActive : {}) }}
            onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
            onDragLeave={() => setDragging(false)}
            onDrop={(e) => { setDragging(false); onDrop(e); }}
            onClick={() => fileInputRef.current?.click()}
          >
            <input ref={fileInputRef} type="file" accept=".docx" style={{ display: "none" }}
              onChange={(e) => handleFile(e.target.files[0])} />
            <div style={s.dropIcon}>{isProcessing ? "⚙️" : "📄"}</div>
            <div style={s.dropLabel}>
              {file ? file.name : "Kéo thả hoặc click để chọn file .docx"}
            </div>
            <div style={s.dropHint}>
              {file ? `${(file.size / 1024).toFixed(0)} KB` : "Chỉ hỗ trợ .docx"}
            </div>
          </div>

          {/* Download button */}
          {downloadUrl && (
            <a href={downloadUrl} download={downloadName} style={s.downloadBtn} className="dl-btn">
              ⬇ Tải file đã chuẩn hóa
            </a>
          )}

          {/* Inspect result */}
          {inspect && (
            <div style={s.card}>
              <div style={s.cardTitle}>Thông tin file gốc</div>
              <div style={s.infoGrid}>
                <InfoRow label="Tổng đoạn văn" value={inspect.paragraphs_total} />
                <InfoRow label="Số bảng" value={inspect.tables} />
                <InfoRow label="Trang" value={`${inspect.page?.width_cm} × ${inspect.page?.height_cm} cm`} />
                <InfoRow label="Margin T/D/T/P" value={`${inspect.page?.margin_top}/${inspect.page?.margin_bottom}/${inspect.page?.margin_left}/${inspect.page?.margin_right} cm`} />
              </div>
              {inspect.paragraphs_sample?.length > 0 && (
                <>
                  <div style={{ ...s.cardTitle, marginTop: 12 }}>Cấu trúc heading</div>
                  <div style={s.paraList}>
                    {inspect.paragraphs_sample
                      .filter(p => p.style.startsWith('Heading'))
                      .slice(0, 10)
                      .map((p, i) => (
                        <div key={i} style={s.paraItem}>
                          <span style={{ ...s.styleBadge, background: headingColor(p.style) }}>{p.style}</span>
                          <span style={s.paraText}>{p.text}</span>
                        </div>
                      ))}
                  </div>
                </>
              )}
            </div>
          )}
        </div>

        {/* Right panel */}
        <div style={s.right}>
          {/* Log */}
          <div style={s.card}>
            <div style={s.cardTitle}>Process Log</div>
            <div style={s.logBox}>
              {log.length === 0
                ? <div style={s.logEmpty}>Upload file để bắt đầu...</div>
                : log.map((l, i) => (
                  <div key={i} style={{ ...s.logLine, color: logColor(l.type) }}>
                    {l.msg}
                  </div>
                ))}
            </div>
          </div>

          {/* Rules */}
          <div style={s.card}>
            <div style={s.cardTitle}>Format Rules</div>
            <div style={s.ruleList}>
              {STYLE_RULES.map((r, i) => (
                <div key={i} style={s.ruleRow}>
                  <span style={s.ruleLabel}>{r.label}</span>
                  <span style={s.ruleValue}>{r.value}</span>
                </div>
              ))}
            </div>
            <div style={s.ruleNote}>
              Chỉnh rules tại <code style={s.code}>backend/main.py → STYLE_CONFIG</code>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

// ---- helpers ----
function InfoRow({ label, value }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", padding: "4px 0", borderBottom: "1px solid #1e2533" }}>
      <span style={{ color: "#64748b", fontSize: 12 }}>{label}</span>
      <span style={{ color: "#cbd5e1", fontSize: 12, fontFamily: "monospace" }}>{value}</span>
    </div>
  );
}

function headingColor(style) {
  const map = {
    "Heading 1": "#7c3aed", "Heading 2": "#2563eb",
    "Heading 3": "#0891b2", "Heading 4": "#059669", "Heading 5": "#d97706",
  };
  return map[style] || "#475569";
}
function logColor(type) {
  return { success: "#22c55e", error: "#ef4444", warn: "#f59e0b", info: "#94a3b8" }[type] || "#94a3b8";
}

// ---- styles ----
const s = {
  app: { minHeight: "100vh", background: "#0f1117", color: "#e2e8f0", fontFamily: "'IBM Plex Mono', monospace", display: "flex", flexDirection: "column" },
  header: { padding: "20px 28px", borderBottom: "1px solid #1e2533", display: "flex", alignItems: "center", justifyContent: "space-between" },
  headerTitle: { fontSize: 20, fontWeight: 700, color: "#f1f5f9", fontFamily: "Georgia, serif", letterSpacing: "-0.3px" },
  headerSub: { fontSize: 11, color: "#475569", marginTop: 3, letterSpacing: "0.5px" },
  statusPill: { display: "flex", alignItems: "center", gap: 7, padding: "5px 12px", borderRadius: 20, fontSize: 12 },
  statusDot: { width: 7, height: 7, borderRadius: "50%", flexShrink: 0 },
  body: { display: "grid", gridTemplateColumns: "1fr 340px", gap: 0, flex: 1 },
  left: { padding: "24px 24px 24px 28px", display: "flex", flexDirection: "column", gap: 16, borderRight: "1px solid #1e2533" },
  right: { padding: "24px 28px 24px 24px", display: "flex", flexDirection: "column", gap: 16 },
  dropZone: { border: "1.5px dashed #2d3748", borderRadius: 8, padding: "36px 20px", textAlign: "center", cursor: "pointer", background: "#13161f", transition: "all 0.2s" },
  dropZoneActive: { borderColor: "#4f6ef7", background: "#141929" },
  dropIcon: { fontSize: 30, marginBottom: 10 },
  dropLabel: { fontSize: 13, color: "#94a3b8", wordBreak: "break-all" },
  dropHint: { fontSize: 11, color: "#475569", marginTop: 5 },
  downloadBtn: { display: "flex", alignItems: "center", justifyContent: "center", padding: "12px", background: "#16a34a", color: "white", borderRadius: 6, fontSize: 13, fontWeight: 600, textDecoration: "none", fontFamily: "'IBM Plex Mono', monospace" },
  card: { background: "#13161f", border: "1px solid #1e2533", borderRadius: 8, padding: 16 },
  cardTitle: { fontSize: 10, color: "#475569", textTransform: "uppercase", letterSpacing: "1.2px", marginBottom: 10 },
  infoGrid: { display: "flex", flexDirection: "column", gap: 0 },
  paraList: { display: "flex", flexDirection: "column", gap: 5, marginTop: 4 },
  paraItem: { display: "flex", alignItems: "flex-start", gap: 8 },
  styleBadge: { fontSize: 9, padding: "2px 6px", borderRadius: 3, color: "white", whiteSpace: "nowrap", flexShrink: 0, marginTop: 1 },
  paraText: { fontSize: 11, color: "#64748b", lineHeight: 1.4, wordBreak: "break-word" },
  logBox: { background: "#0a0d14", borderRadius: 6, padding: "10px 12px", minHeight: 80, maxHeight: 200, overflowY: "auto", display: "flex", flexDirection: "column", gap: 4 },
  logEmpty: { fontSize: 11, color: "#2d3748", fontStyle: "italic" },
  logLine: { fontSize: 11, lineHeight: 1.5 },
  ruleList: { display: "flex", flexDirection: "column", gap: 0 },
  ruleRow: { display: "flex", justifyContent: "space-between", alignItems: "center", padding: "4px 0", borderBottom: "1px solid #1e2533" },
  ruleLabel: { fontSize: 11, color: "#64748b" },
  ruleValue: { fontSize: 11, color: "#4f6ef7", fontFamily: "monospace" },
  ruleNote: { marginTop: 10, fontSize: 10, color: "#475569", lineHeight: 1.6 },
  code: { background: "#0a0d14", padding: "1px 4px", borderRadius: 3, color: "#94a3b8" },
};

const css = `
  @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&display=swap');
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: #0f1117; }
  @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.3} }
  .dl-btn:hover { background: #15803d !important; }
  ::-webkit-scrollbar { width: 4px; }
  ::-webkit-scrollbar-thumb { background: #2d3748; border-radius: 2px; }
`;
