import { useState, useCallback, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  PieChart, Pie, Cell, Tooltip, ResponsiveContainer, Legend,
  BarChart, Bar, XAxis, YAxis, CartesianGrid,
  LineChart, Line, Area, AreaChart
} from "recharts";

// ─── Theme ───────────────────────────────────────────────────────────────────
const T = {
  bg: "#080b14",
  surface: "#0d1117",
  card: "#111827",
  border: "#1f2937",
  accent: "#22d3ee",     // cyan
  accent2: "#a78bfa",     // violet
  green: "#34d399",
  yellow: "#fbbf24",
  red: "#f87171",
  text: "#f1f5f9",
  muted: "#64748b",
  subtle: "#1e293b",
};

const CATEGORY_COLORS = [
  "#22d3ee", "#a78bfa", "#34d399", "#fbbf24", "#f87171", "#fb923c",
  "#e879f9", "#38bdf8", "#4ade80", "#facc15", "#f472b6", "#60a5fa",
  "#2dd4bf", "#c084fc", "#86efac", "#fde68a", "#fca5a5", "#fed7aa",
];

const CATEGORY_GROUPS = {
  "Boliglån": { group: "Bolig", icon: "🏠" },
  "Husleie": { group: "Bolig", icon: "🏠" },
  "Kommunale avgifter": { group: "Bolig", icon: "🏠" },
  "Strøm": { group: "Bolig", icon: "⚡" },
  "Vann": { group: "Bolig", icon: "💧" },
  "Interiør": { group: "Bolig", icon: "🛋️" },
  "Parkering": { group: "Transport", icon: "🚗" },
  "Elbillading": { group: "Transport", icon: "🔋" },
  "Offentlig transport": { group: "Transport", icon: "🚌" },
  "Dagligvarer": { group: "Mat & Drikke", icon: "🛒" },
  "Kantine": { group: "Mat & Drikke", icon: "🍽️" },
  "Restaurant": { group: "Mat & Drikke", icon: "🍽️" },
  "Kiosk": { group: "Mat & Drikke", icon: "☕" },
  "Strømmetjenester": { group: "Abonnement", icon: "📺" },
  "Internettjenester": { group: "Abonnement", icon: "🌐" },
  "Musikk": { group: "Abonnement", icon: "🎵" },
  "Treningssenter": { group: "Helse & Fritid", icon: "💪" },
  "Apotek": { group: "Helse & Fritid", icon: "💊" },
  "Sportsutstyr": { group: "Helse & Fritid", icon: "⚽" },
  "Klær og sko": { group: "Klær", icon: "👗" },
  "Frisør": { group: "Klær", icon: "✂️" },
  "Studielån": { group: "Lån & Sparing", icon: "🎓" },
  "Overføring": { group: "Lån & Sparing", icon: "💰" },
  "Spill": { group: "Underholdning", icon: "🎮" },
  "Elektronikk": { group: "Elektronikk", icon: "💻" },
  "Forsikring": { group: "Forsikring", icon: "🛡️" },
  "Fagforeningskontingent": { group: "Diverse", icon: "📋" },
  "Kontingent": { group: "Diverse", icon: "📋" },
  "Betaling": { group: "Diverse", icon: "💳" },
  "Gave": { group: "Diverse", icon: "🎁" },
  "Diverse": { group: "Diverse", icon: "📦" },
  "VISA varekjøp": { group: "Varekjøp", icon: "💳" },
  "Varekjøp": { group: "Varekjøp", icon: "🛍️" },
};

function getCategoryIcon(cat) {
  return CATEGORY_GROUPS[cat]?.icon ?? "💡";
}

// ─── Format helpers ───────────────────────────────────────────────────────────
function fmtCurrency(n) {
  return new Intl.NumberFormat("nb-NO", { style: "currency", currency: "NOK", maximumFractionDigits: 0 }).format(n);
}

function fmtDateTime(d) {
  if (!d) return "—";
  const dt = d instanceof Date ? d : new Date(d);
  if (isNaN(dt)) return String(d);
  const pad = n => String(n).padStart(2, "0");
  return `${pad(dt.getDate())}.${pad(dt.getMonth() + 1)}.${dt.getFullYear()} ${pad(dt.getHours())}:${pad(dt.getMinutes())}:${pad(dt.getSeconds())}`;
}

function fmtDate(d) {
  if (!d) return "—";
  const dt = d instanceof Date ? d : new Date(d);
  if (isNaN(dt)) return String(d);
  const pad = n => String(n).padStart(2, "0");
  return `${pad(dt.getDate())}.${pad(dt.getMonth() + 1)}.${dt.getFullYear()}`;
}

// ─── Parse xlsx ───────────────────────────────────────────────────────────────
function parseXlsx(buffer) {
  const wb = XLSX.read(buffer, { type: "array", cellDates: true });
  let allTransactions = [];

  wb.SheetNames.forEach(sheetName => {
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: false, dateNF: "dd.mm.yyyy" });
    if (rows.length < 2) return;

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length < 5) continue;
      const bokfort = row[0] ? new Date(row[0]) : null;
      const type = row[2] ? String(row[2]).trim() : "Ukjent";
      const beskrivelse = row[3] ? String(row[3]).trim() : "";
      const utAvKonto = parseFloat(String(row[4]).replace(",", ".")) || 0;
      if (utAvKonto === 0) continue;
      allTransactions.push({
        month: sheetName,
        date: bokfort,
        type,
        description: beskrivelse,
        amount: Math.abs(utAvKonto),
        raw: utAvKonto,
      });
    }
  });

  return { transactions: allTransactions, months: wb.SheetNames };
}

// ─── Upload Screen ────────────────────────────────────────────────────────────
function UploadScreen({ onFile }) {
  const [dragging, setDragging] = useState(false);

  const handleDrop = useCallback(e => {
    e.preventDefault();
    setDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) readFile(file);
  }, []);

  const readFile = file => {
    const reader = new FileReader();
    reader.onload = e => onFile(e.target.result, file.name);
    reader.readAsArrayBuffer(file);
  };

  return (
    <div style={{ minHeight: "100vh", background: T.bg, display: "flex", alignItems: "center", justifyContent: "center", padding: "20px", fontFamily: "'DM Mono', 'Courier New', monospace" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Mono:ital,wght@0,300;0,400;0,500;1,400&family=Syne:wght@400;600;700;800&display=swap');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { background: ${T.bg}; }
        ::-webkit-scrollbar { width: 4px; }
        ::-webkit-scrollbar-track { background: ${T.surface}; }
        ::-webkit-scrollbar-thumb { background: ${T.border}; border-radius: 2px; }
        .upload-area { transition: all 0.3s ease; }
        .upload-area:hover { border-color: ${T.accent} !important; box-shadow: 0 0 40px rgba(34,211,238,0.15) !important; }
      `}</style>

      <div style={{ maxWidth: "480px", width: "100%", textAlign: "center" }}>
        {/* Logo/Title */}
        <div style={{ marginBottom: "48px" }}>
          <div style={{ fontSize: "48px", marginBottom: "12px" }}>◈</div>
          <h1 style={{ fontFamily: "'Syne', sans-serif", fontSize: "clamp(28px,6vw,36px)", fontWeight: 800, color: T.text, letterSpacing: "-0.02em", marginBottom: "8px" }}>
            Min Økonomi
          </h1>
          <p style={{ color: T.muted, fontSize: "14px", letterSpacing: "0.05em" }}>PERSONLIG FINANSDASHBOARD</p>
        </div>

        {/* Upload zone */}
        <div
          className="upload-area"
          onDragOver={e => { e.preventDefault(); setDragging(true); }}
          onDragLeave={() => setDragging(false)}
          onDrop={handleDrop}
          style={{
            border: `2px dashed ${dragging ? T.accent : T.border}`,
            borderRadius: "16px",
            padding: "48px 32px",
            background: dragging ? "rgba(34,211,238,0.05)" : T.card,
            cursor: "pointer",
            boxShadow: dragging ? `0 0 40px rgba(34,211,238,0.15)` : "none",
          }}
        >
          <div style={{ fontSize: "40px", marginBottom: "16px" }}>📊</div>
          <p style={{ color: T.text, fontSize: "16px", fontWeight: 500, marginBottom: "8px", fontFamily: "'Syne', sans-serif" }}>
            Dra og slipp din Excel-fil
          </p>
          <p style={{ color: T.muted, fontSize: "13px", marginBottom: "24px" }}>
            .xlsx fil med transaksjoner
          </p>
          <label style={{ cursor: "pointer" }}>
            <input type="file" accept=".xlsx" style={{ display: "none" }} onChange={e => e.target.files[0] && readFile(e.target.files[0])} />
            <span style={{
              background: `linear-gradient(135deg, ${T.accent}, ${T.accent2})`,
              color: "#000",
              padding: "12px 28px",
              borderRadius: "8px",
              fontSize: "13px",
              fontWeight: 600,
              fontFamily: "'Syne', sans-serif",
              letterSpacing: "0.04em",
            }}>
              VELG FIL
            </span>
          </label>
        </div>

        <p style={{ color: T.muted, fontSize: "11px", marginTop: "20px" }}>
          Data behandles lokalt i din nettleser · Ingen data sendes til server
        </p>
      </div>
    </div>
  );
}

// ─── KPI Card ─────────────────────────────────────────────────────────────────
function KpiCard({ label, value, sub, color, icon }) {
  return (
    <div style={{
      background: T.card,
      border: `1px solid ${T.border}`,
      borderRadius: "12px",
      padding: "16px 20px",
      display: "flex",
      flexDirection: "column",
      gap: "4px",
      position: "relative",
      overflow: "hidden",
    }}>
      <div style={{ position: "absolute", top: 0, left: 0, right: 0, height: "2px", background: `linear-gradient(90deg, ${color}, transparent)` }} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <span style={{ fontSize: "11px", color: T.muted, letterSpacing: "0.06em", fontWeight: 500 }}>{label}</span>
        <span style={{ fontSize: "20px" }}>{icon}</span>
      </div>
      <div style={{ fontSize: "clamp(18px,4vw,24px)", fontWeight: 700, color: T.text, fontFamily: "'Syne', sans-serif", letterSpacing: "-0.02em" }}>{value}</div>
      {sub && <div style={{ fontSize: "11px", color: T.muted }}>{sub}</div>}
    </div>
  );
}

// ─── Custom Tooltip ───────────────────────────────────────────────────────────
function CustomTooltip({ active, payload, label }) {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: "8px", padding: "10px 14px", fontSize: "12px", fontFamily: "'DM Mono', monospace" }}>
      {label && <p style={{ color: T.muted, marginBottom: "4px" }}>{label}</p>}
      {payload.map((p, i) => (
        <p key={i} style={{ color: p.color || T.accent }}>
          {p.name ? `${p.name}: ` : ""}{fmtCurrency(p.value)}
        </p>
      ))}
    </div>
  );
}

// ─── Category Donut ───────────────────────────────────────────────────────────
function CategoryDonut({ data, onSelect, selected }) {
  const top10 = data.slice(0, 10);
  const rest = data.slice(10);
  const chartData = rest.length > 0
    ? [...top10, { type: "Andre", total: rest.reduce((s, x) => s + x.total, 0) }]
    : top10;

  return (
    <ResponsiveContainer width="100%" height={260}>
      <PieChart>
        <Pie
          data={chartData}
          cx="50%"
          cy="50%"
          innerRadius={65}
          outerRadius={110}
          dataKey="total"
          nameKey="type"
          paddingAngle={2}
          onClick={d => onSelect(d.type === selected ? null : d.type)}
        >
          {chartData.map((entry, i) => (
            <Cell
              key={i}
              fill={CATEGORY_COLORS[i % CATEGORY_COLORS.length]}
              opacity={selected && selected !== entry.type ? 0.35 : 1}
              stroke={selected === entry.type ? "#fff" : "transparent"}
              strokeWidth={2}
              style={{ cursor: "pointer" }}
            />
          ))}
        </Pie>
        <Tooltip content={<CustomTooltip />} />
      </PieChart>
    </ResponsiveContainer>
  );
}

// ─── Transaction Row ──────────────────────────────────────────────────────────
function TxRow({ tx, idx }) {
  const [open, setOpen] = useState(false);
  const color = CATEGORY_COLORS[idx % CATEGORY_COLORS.length];
  return (
    <div
      onClick={() => setOpen(o => !o)}
      style={{
        padding: "12px 16px",
        borderBottom: `1px solid ${T.border}`,
        cursor: "pointer",
        transition: "background 0.15s",
        background: open ? T.subtle : "transparent",
      }}
    >
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: "8px" }}>
        <div style={{ display: "flex", alignItems: "center", gap: "10px", minWidth: 0 }}>
          <span style={{ fontSize: "18px", flexShrink: 0 }}>{getCategoryIcon(tx.type)}</span>
          <div style={{ minWidth: 0 }}>
            <div style={{ color: T.text, fontSize: "13px", fontWeight: 500, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
              {tx.description}
            </div>
            <div style={{ color: T.muted, fontSize: "11px", display: "flex", gap: "8px", marginTop: "2px" }}>
              <span style={{ color, opacity: 0.85 }}>{tx.type}</span>
              <span>{fmtDate(tx.date)}</span>
            </div>
          </div>
        </div>
        <div style={{ flexShrink: 0, textAlign: "right" }}>
          <div style={{ color: T.red, fontSize: "14px", fontWeight: 600, fontFamily: "'DM Mono', monospace" }}>
            -{fmtCurrency(tx.amount)}
          </div>
          <div style={{ color: T.muted, fontSize: "10px" }}>{tx.month}</div>
        </div>
      </div>
      {open && (
        <div style={{ marginTop: "10px", padding: "10px", background: T.surface, borderRadius: "8px", fontSize: "11px", color: T.muted, fontFamily: "'DM Mono', monospace" }}>
          <div><span style={{ color: T.accent }}>Tid:</span> {fmtDateTime(tx.date)}</div>
          <div><span style={{ color: T.accent }}>Kategori:</span> {tx.type}</div>
          <div><span style={{ color: T.accent }}>Beløp:</span> {fmtCurrency(tx.amount)}</div>
        </div>
      )}
    </div>
  );
}

// ─── Dashboard ────────────────────────────────────────────────────────────────
function Dashboard({ transactions, months, fileName, onReset }) {
  const [selectedMonth, setSelectedMonth] = useState("all");
  const [search, setSearch] = useState("");
  const [selectedCategory, setSelectedCategory] = useState(null);
  const [activeTab, setActiveTab] = useState("overview"); // overview | transactions

  // Filter by month
  const monthFiltered = useMemo(() =>
    selectedMonth === "all" ? transactions : transactions.filter(t => t.month === selectedMonth),
    [transactions, selectedMonth]
  );

  // Category breakdown
  const categoryData = useMemo(() => {
    const map = {};
    monthFiltered.forEach(t => {
      map[t.type] = (map[t.type] || 0) + t.amount;
    });
    return Object.entries(map)
      .map(([type, total]) => ({ type, total, icon: getCategoryIcon(type) }))
      .sort((a, b) => b.total - a.total);
  }, [monthFiltered]);

  // Monthly totals
  const monthlyData = useMemo(() => {
    const map = {};
    transactions.forEach(t => {
      map[t.month] = (map[t.month] || 0) + t.amount;
    });
    return months.map(m => ({ month: m, total: map[m] || 0 }));
  }, [transactions, months]);

  // KPIs
  const totalSpent = monthFiltered.reduce((s, t) => s + t.amount, 0);
  const topCategory = categoryData[0];
  const avgTx = monthFiltered.length ? totalSpent / monthFiltered.length : 0;
  const biggestTx = monthFiltered.reduce((max, t) => t.amount > max ? t.amount : max, 0);

  // Filtered transactions
  const filtered = useMemo(() => {
    let txs = monthFiltered;
    if (selectedCategory) txs = txs.filter(t => t.type === selectedCategory);
    if (search.trim()) {
      const q = search.toLowerCase();
      txs = txs.filter(t => t.description.toLowerCase().includes(q) || t.type.toLowerCase().includes(q));
    }
    return [...txs].sort((a, b) => (b.date || 0) - (a.date || 0));
  }, [monthFiltered, selectedCategory, search]);

  const tabStyle = (tab) => ({
    padding: "10px 20px",
    background: activeTab === tab ? T.accent : "transparent",
    color: activeTab === tab ? "#000" : T.muted,
    border: "none",
    borderRadius: "8px",
    cursor: "pointer",
    fontSize: "12px",
    fontWeight: 600,
    fontFamily: "'Syne', sans-serif",
    letterSpacing: "0.04em",
    transition: "all 0.2s",
  });

  return (
    <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'DM Mono', 'Courier New', monospace", color: T.text }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Mono:ital,wght@0,300;0,400;0,500;1,400&family=Syne:wght@400;600;700;800&display=swap');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        ::-webkit-scrollbar { width: 4px; height: 4px; }
        ::-webkit-scrollbar-track { background: ${T.surface}; }
        ::-webkit-scrollbar-thumb { background: ${T.border}; border-radius: 2px; }
        input::placeholder { color: ${T.muted}; }
        input { outline: none; }
        button:focus { outline: none; }
        .month-btn:hover { border-color: ${T.accent} !important; color: ${T.accent} !important; }
        .cat-row:hover { background: ${T.subtle} !important; }
      `}</style>

      {/* Header */}
      <div style={{ background: T.surface, borderBottom: `1px solid ${T.border}`, padding: "12px 16px", display: "flex", justifyContent: "space-between", alignItems: "center", position: "sticky", top: 0, zIndex: 100 }}>
        <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
          <span style={{ fontSize: "20px" }}>◈</span>
          <div>
            <div style={{ fontFamily: "'Syne', sans-serif", fontWeight: 800, fontSize: "15px", letterSpacing: "-0.01em" }}>Min Økonomi</div>
            <div style={{ fontSize: "10px", color: T.muted, letterSpacing: "0.04em" }}>{fileName}</div>
          </div>
        </div>
        <button onClick={onReset} style={{ background: "transparent", border: `1px solid ${T.border}`, color: T.muted, padding: "6px 12px", borderRadius: "6px", cursor: "pointer", fontSize: "11px", fontFamily: "'DM Mono', monospace" }}>
          Bytt fil
        </button>
      </div>

      {/* Month selector */}
      <div style={{ padding: "12px 16px", overflowX: "auto", display: "flex", gap: "8px", background: T.surface, borderBottom: `1px solid ${T.border}` }}>
        {["all", ...months].map(m => (
          <button
            key={m}
            className="month-btn"
            onClick={() => setSelectedMonth(m)}
            style={{
              padding: "6px 14px",
              borderRadius: "20px",
              border: `1px solid ${selectedMonth === m ? T.accent : T.border}`,
              background: selectedMonth === m ? `rgba(34,211,238,0.15)` : "transparent",
              color: selectedMonth === m ? T.accent : T.muted,
              fontSize: "11px",
              fontWeight: 500,
              cursor: "pointer",
              whiteSpace: "nowrap",
              transition: "all 0.15s",
              fontFamily: "'DM Mono', monospace",
            }}
          >
            {m === "all" ? "Alle måneder" : m}
          </button>
        ))}
      </div>

      {/* Tab nav */}
      <div style={{ padding: "12px 16px", display: "flex", gap: "8px", background: T.surface, borderBottom: `1px solid ${T.border}` }}>
        <button style={tabStyle("overview")} onClick={() => setActiveTab("overview")}>Oversikt</button>
        <button style={tabStyle("transactions")} onClick={() => setActiveTab("transactions")}>Transaksjoner</button>
      </div>

      <div style={{ padding: "16px", maxWidth: "900px", margin: "0 auto" }}>

        {activeTab === "overview" && (
          <>
            {/* KPI grid */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: "10px", marginBottom: "20px" }}>
              <KpiCard label="TOTAL BRUKT" value={fmtCurrency(totalSpent)} sub={`${monthFiltered.length} transaksjoner`} color={T.red} icon="💸" />
              <KpiCard label="STØRSTE UTGIFT" value={fmtCurrency(biggestTx)} sub="enkelt transaksjon" color={T.yellow} icon="⚡" />
              <KpiCard label="SNITT PER TX" value={fmtCurrency(avgTx)} sub="gjennomsnitt" color={T.accent} icon="📊" />
              <KpiCard label="TOPP KATEGORI" value={topCategory?.type ?? "—"} sub={topCategory ? fmtCurrency(topCategory.total) : ""} color={T.accent2} icon="🏆" />
            </div>

            {/* Category donut */}
            <div style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: "12px", padding: "16px", marginBottom: "16px" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "4px" }}>
                <h3 style={{ fontFamily: "'Syne', sans-serif", fontSize: "13px", fontWeight: 700, color: T.text, letterSpacing: "0.04em" }}>FORBRUK PER KATEGORI</h3>
                {selectedCategory && (
                  <button onClick={() => setSelectedCategory(null)} style={{ background: "transparent", border: "none", color: T.accent, fontSize: "11px", cursor: "pointer", fontFamily: "'DM Mono', monospace" }}>
                    × Fjern filter
                  </button>
                )}
              </div>
              <p style={{ fontSize: "10px", color: T.muted, marginBottom: "12px" }}>Klikk på sektor for å filtrere</p>
              <CategoryDonut data={categoryData} onSelect={setSelectedCategory} selected={selectedCategory} />

              {/* Legend */}
              <div style={{ display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: "4px", marginTop: "8px" }}>
                {categoryData.slice(0, 10).map((cat, i) => {
                  const pct = (cat.total / totalSpent * 100).toFixed(1);
                  const isSelected = selectedCategory === cat.type;
                  return (
                    <div
                      key={cat.type}
                      className="cat-row"
                      onClick={() => setSelectedCategory(isSelected ? null : cat.type)}
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: "6px",
                        padding: "5px 8px",
                        borderRadius: "6px",
                        cursor: "pointer",
                        background: isSelected ? T.subtle : "transparent",
                        opacity: selectedCategory && !isSelected ? 0.5 : 1,
                        transition: "all 0.15s",
                      }}
                    >
                      <div style={{ width: "8px", height: "8px", borderRadius: "50%", background: CATEGORY_COLORS[i % CATEGORY_COLORS.length], flexShrink: 0 }} />
                      <div style={{ minWidth: 0, flex: 1 }}>
                        <div style={{ fontSize: "10px", color: T.text, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                          {cat.icon} {cat.type}
                        </div>
                        <div style={{ fontSize: "9px", color: T.muted }}>{pct}%</div>
                      </div>
                      <div style={{ fontSize: "10px", color: CATEGORY_COLORS[i % CATEGORY_COLORS.length], flexShrink: 0, fontWeight: 600 }}>
                        {fmtCurrency(cat.total)}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Monthly bar chart */}
            {months.length > 1 && (
              <div style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: "12px", padding: "16px", marginBottom: "16px" }}>
                <h3 style={{ fontFamily: "'Syne', sans-serif", fontSize: "13px", fontWeight: 700, color: T.text, letterSpacing: "0.04em", marginBottom: "16px" }}>MÅNEDLIG FORBRUK</h3>
                <ResponsiveContainer width="100%" height={180}>
                  <BarChart data={monthlyData} barSize={40}>
                    <CartesianGrid strokeDasharray="3 3" stroke={T.border} vertical={false} />
                    <XAxis dataKey="month" tick={{ fill: T.muted, fontSize: 11 }} axisLine={false} tickLine={false} />
                    <YAxis tick={{ fill: T.muted, fontSize: 10 }} axisLine={false} tickLine={false} tickFormatter={v => `${(v / 1000).toFixed(0)}k`} />
                    <Tooltip content={<CustomTooltip />} />
                    <Bar dataKey="total" name="Forbruk" fill={T.accent} radius={[4, 4, 0, 0]} opacity={0.85} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            )}

            {/* Top transactions */}
            <div style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: "12px", padding: "16px", marginBottom: "16px" }}>
              <h3 style={{ fontFamily: "'Syne', sans-serif", fontSize: "13px", fontWeight: 700, color: T.text, letterSpacing: "0.04em", marginBottom: "12px" }}>TOPP 5 UTGIFTER</h3>
              {[...monthFiltered].sort((a, b) => b.amount - a.amount).slice(0, 5).map((tx, i) => (
                <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: i < 4 ? `1px solid ${T.border}` : "none", gap: "8px" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: "8px", minWidth: 0 }}>
                    <span style={{ fontSize: "16px" }}>{getCategoryIcon(tx.type)}</span>
                    <div style={{ minWidth: 0 }}>
                      <div style={{ fontSize: "12px", color: T.text, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{tx.description}</div>
                      <div style={{ fontSize: "10px", color: T.muted }}>{tx.type} · {fmtDate(tx.date)}</div>
                    </div>
                  </div>
                  <div style={{ color: T.red, fontSize: "13px", fontWeight: 600, flexShrink: 0, fontFamily: "'DM Mono', monospace" }}>
                    -{fmtCurrency(tx.amount)}
                  </div>
                </div>
              ))}
            </div>

            {/* Category details breakdown */}
            <div style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: "12px", padding: "16px" }}>
              <h3 style={{ fontFamily: "'Syne', sans-serif", fontSize: "13px", fontWeight: 700, color: T.text, letterSpacing: "0.04em", marginBottom: "12px" }}>ALLE KATEGORIER</h3>
              {categoryData.map((cat, i) => {
                const pct = (cat.total / totalSpent * 100);
                return (
                  <div key={cat.type} style={{ marginBottom: "10px" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "4px" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                        <span style={{ fontSize: "14px" }}>{cat.icon}</span>
                        <span style={{ fontSize: "12px", color: T.text }}>{cat.type}</span>
                      </div>
                      <div style={{ display: "flex", gap: "10px", alignItems: "center" }}>
                        <span style={{ fontSize: "10px", color: T.muted }}>{pct.toFixed(1)}%</span>
                        <span style={{ fontSize: "12px", color: CATEGORY_COLORS[i % CATEGORY_COLORS.length], fontWeight: 600 }}>{fmtCurrency(cat.total)}</span>
                      </div>
                    </div>
                    <div style={{ height: "3px", background: T.border, borderRadius: "2px", overflow: "hidden" }}>
                      <div style={{ height: "100%", width: `${pct}%`, background: CATEGORY_COLORS[i % CATEGORY_COLORS.length], borderRadius: "2px", transition: "width 0.5s ease" }} />
                    </div>
                  </div>
                );
              })}
            </div>
          </>
        )}

        {activeTab === "transactions" && (
          <>
            {/* Search + category filter */}
            <div style={{ display: "flex", gap: "8px", marginBottom: "12px" }}>
              <input
                value={search}
                onChange={e => setSearch(e.target.value)}
                placeholder="🔍 Søk transaksjoner..."
                style={{
                  flex: 1,
                  background: T.card,
                  border: `1px solid ${T.border}`,
                  borderRadius: "8px",
                  padding: "10px 14px",
                  color: T.text,
                  fontSize: "13px",
                  fontFamily: "'DM Mono', monospace",
                }}
              />
              {selectedCategory && (
                <button onClick={() => setSelectedCategory(null)} style={{
                  background: "transparent",
                  border: `1px solid ${T.accent}`,
                  color: T.accent,
                  padding: "8px 12px",
                  borderRadius: "8px",
                  cursor: "pointer",
                  fontSize: "11px",
                  fontFamily: "'DM Mono', monospace",
                  whiteSpace: "nowrap",
                }}>
                  {selectedCategory} ×
                </button>
              )}
            </div>

            {/* Category quick filters */}
            <div style={{ display: "flex", gap: "6px", overflowX: "auto", marginBottom: "12px", paddingBottom: "4px" }}>
              {categoryData.slice(0, 8).map((cat, i) => (
                <button
                  key={cat.type}
                  onClick={() => setSelectedCategory(selectedCategory === cat.type ? null : cat.type)}
                  style={{
                    padding: "5px 10px",
                    borderRadius: "16px",
                    border: `1px solid ${selectedCategory === cat.type ? CATEGORY_COLORS[i % CATEGORY_COLORS.length] : T.border}`,
                    background: selectedCategory === cat.type ? `${CATEGORY_COLORS[i % CATEGORY_COLORS.length]}22` : "transparent",
                    color: selectedCategory === cat.type ? CATEGORY_COLORS[i % CATEGORY_COLORS.length] : T.muted,
                    fontSize: "10px",
                    cursor: "pointer",
                    whiteSpace: "nowrap",
                    fontFamily: "'DM Mono', monospace",
                    transition: "all 0.15s",
                  }}
                >
                  {cat.icon} {cat.type}
                </button>
              ))}
            </div>

            {/* Count */}
            <div style={{ fontSize: "11px", color: T.muted, marginBottom: "8px", textAlign: "right" }}>
              {filtered.length} transaksjoner · {fmtCurrency(filtered.reduce((s, t) => s + t.amount, 0))} totalt
            </div>

            {/* Transaction list */}
            <div style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: "12px", overflow: "hidden" }}>
              {filtered.length === 0 ? (
                <div style={{ padding: "40px", textAlign: "center", color: T.muted, fontSize: "13px" }}>
                  Ingen transaksjoner funnet
                </div>
              ) : (
                filtered.map((tx, i) => <TxRow key={i} tx={tx} idx={categoryData.findIndex(c => c.type === tx.type)} />)
              )}
            </div>
          </>
        )}
      </div>

      {/* Bottom padding for mobile */}
      <div style={{ height: "32px" }} />
    </div>
  );
}

// ─── Root App ─────────────────────────────────────────────────────────────────
export default function App() {
  const [data, setData] = useState(null);
  const [fileName, setFileName] = useState("");

  const handleFile = useCallback((buffer, name) => {
    const parsed = parseXlsx(buffer);
    setData(parsed);
    setFileName(name || "transaksjoner.xlsx");
  }, []);

  useEffect(() => {
    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.register('/sw.js')
        .then(registration => {
          registration.addEventListener('updatefound', () => {
            const newWorker = registration.installing;
            newWorker.addEventListener('statechange', () => {
              if (newWorker.state === 'installed' && navigator.serviceWorker.controller) {
                // New content is available, notify user
                if (confirm('New version available! Would you like to update?')) {
                  window.location.reload();
                }
              }
            });
          });
        })
        .catch(error => console.log('Service worker registration failed:', error));
    }
  }, []);

  if (!data) return <UploadScreen onFile={handleFile} />;

  return (
    <Dashboard
      transactions={data.transactions}
      months={data.months}
      fileName={fileName}
      onReset={() => setData(null)}
    />
  );
}