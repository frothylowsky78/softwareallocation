import { useState, useCallback, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import _ from "lodash";

// ─── Utility helpers ────────────────────────────────────────────────
const normalizeEmail = (e) => (e || "").toString().trim().toLowerCase();
const fmt = (n) => (n == null || isNaN(n) ? "—" : n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
const fmtPct = (n) => (n == null || isNaN(n) ? "—" : (n * 100).toFixed(2) + "%");

const guessColumn = (headers, patterns) => {
  const h = headers.map((x) => (x || "").toString().toLowerCase());
  for (const p of patterns) {
    const idx = h.findIndex((x) => x.includes(p));
    if (idx >= 0) return idx;
  }
  return -1;
};

const parseSheet = (wb, sheetName) => {
  const ws = wb.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
};

// ─── Styles ─────────────────────────────────────────────────────────
const COLORS = {
  bg: "#0f1117",
  surface: "#181b24",
  surfaceAlt: "#1e2230",
  border: "#2a2e3d",
  borderFocus: "#4f7cff",
  text: "#e4e6ed",
  textMuted: "#8b8fa3",
  accent: "#4f7cff",
  accentHover: "#6b91ff",
  warn: "#f5a623",
  warnBg: "rgba(245,166,35,0.1)",
  success: "#3dd68c",
  successBg: "rgba(61,214,140,0.08)",
  danger: "#ef4444",
  dangerBg: "rgba(239,68,68,0.08)",
  white: "#ffffff",
};

const baseInput = {
  background: COLORS.surfaceAlt,
  border: `1px solid ${COLORS.border}`,
  borderRadius: 6,
  color: COLORS.text,
  padding: "8px 12px",
  fontSize: 13,
  outline: "none",
  fontFamily: "'IBM Plex Mono', monospace",
};

const btn = (variant = "primary") => ({
  padding: "10px 20px",
  borderRadius: 6,
  border: variant === "ghost" ? `1px solid ${COLORS.border}` : "none",
  background: variant === "primary" ? COLORS.accent : variant === "ghost" ? "transparent" : COLORS.surfaceAlt,
  color: variant === "primary" ? COLORS.white : COLORS.text,
  fontSize: 13,
  fontWeight: 600,
  cursor: "pointer",
  fontFamily: "'IBM Plex Sans', sans-serif",
  transition: "all 0.15s ease",
});

const chip = (active) => ({
  padding: "8px 16px",
  borderRadius: 20,
  border: `1.5px solid ${active ? COLORS.accent : COLORS.border}`,
  background: active ? "rgba(79,124,255,0.12)" : "transparent",
  color: active ? COLORS.accent : COLORS.textMuted,
  fontSize: 13,
  fontWeight: 600,
  cursor: "pointer",
  transition: "all 0.15s ease",
});

// ─── Sub-components ─────────────────────────────────────────────────
function FileUpload({ label, hint, onParsed, parsed }) {
  const ref = useRef();
  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const wb = XLSX.read(ev.target.result, { type: "array" });
      onParsed({ name: file.name, wb, sheets: wb.SheetNames });
    };
    reader.readAsArrayBuffer(file);
  };
  return (
    <div
      onClick={() => !parsed && ref.current.click()}
      style={{
        border: `2px dashed ${parsed ? COLORS.success : COLORS.border}`,
        borderRadius: 10,
        padding: 28,
        textAlign: "center",
        cursor: parsed ? "default" : "pointer",
        background: parsed ? COLORS.successBg : COLORS.surface,
        transition: "all 0.2s ease",
      }}
    >
      <input ref={ref} type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} style={{ display: "none" }} />
      <div style={{ fontSize: 14, fontWeight: 600, color: parsed ? COLORS.success : COLORS.text, marginBottom: 4 }}>
        {parsed ? `✓ ${parsed.name}` : label}
      </div>
      <div style={{ fontSize: 12, color: COLORS.textMuted }}>{parsed ? `${parsed.sheets.length} sheet(s)` : hint}</div>
      {parsed && (
        <button
          onClick={(e) => { e.stopPropagation(); onParsed(null); }}
          style={{ ...btn("ghost"), marginTop: 10, padding: "4px 12px", fontSize: 11 }}
        >
          Remove
        </button>
      )}
    </div>
  );
}

function ColumnMapper({ label, headers, value, onChange }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
      <span style={{ fontSize: 12, color: COLORS.textMuted, width: 180, flexShrink: 0 }}>{label}</span>
      <select value={value} onChange={(e) => onChange(parseInt(e.target.value))} style={{ ...baseInput, flex: 1 }}>
        <option value={-1}>— select —</option>
        {headers.map((h, i) => (
          <option key={i} value={i}>{h || `(Column ${i + 1})`}</option>
        ))}
      </select>
    </div>
  );
}

function CostCenterPicker({ costCenters, value, onChange, placeholder }) {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState("");
  const ref = useRef();

  useEffect(() => {
    const handler = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, []);

  const filtered = useMemo(() => {
    if (!search) return costCenters;
    const q = search.toLowerCase();
    return costCenters.filter((cc) => cc.ccId.toLowerCase().includes(q) || cc.ccDesc.toLowerCase().includes(q));
  }, [costCenters, search]);

  return (
    <div ref={ref} style={{ position: "relative" }}>
      <input
        value={value || ""}
        onChange={(e) => { onChange(e.target.value, ""); setSearch(e.target.value); setOpen(true); }}
        onFocus={() => setOpen(true)}
        placeholder={placeholder}
        style={{ ...baseInput, width: "100%" }}
      />
      {open && filtered.length > 0 && (
        <div style={{
          position: "absolute", top: "100%", left: 0, right: 0, zIndex: 100,
          background: COLORS.surfaceAlt, border: `1px solid ${COLORS.border}`, borderRadius: 6,
          maxHeight: 180, overflowY: "auto", marginTop: 2, boxShadow: "0 8px 24px rgba(0,0,0,0.4)",
        }}>
          {filtered.slice(0, 50).map((cc) => (
            <div
              key={cc.ccId}
              onClick={() => { onChange(cc.ccId, cc.ccDesc); setSearch(""); setOpen(false); }}
              style={{
                padding: "6px 10px", cursor: "pointer", fontSize: 12,
                borderBottom: `1px solid ${COLORS.border}`,
                color: COLORS.text,
              }}
              onMouseEnter={(e) => e.currentTarget.style.background = COLORS.surface}
              onMouseLeave={(e) => e.currentTarget.style.background = "transparent"}
            >
              <span style={{ fontWeight: 600, marginRight: 8, fontFamily: "'IBM Plex Mono', monospace" }}>{cc.ccId}</span>
              <span style={{ color: COLORS.textMuted }}>{cc.ccDesc}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function DataTable({ columns, rows, maxRows = 200, emptyMsg = "No data" }) {
  if (!rows.length) return <div style={{ padding: 20, color: COLORS.textMuted, textAlign: "center", fontSize: 13 }}>{emptyMsg}</div>;
  const display = rows.slice(0, maxRows);
  return (
    <div style={{ overflowX: "auto", borderRadius: 8, border: `1px solid ${COLORS.border}` }}>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, fontFamily: "'IBM Plex Mono', monospace" }}>
        <thead>
          <tr>
            {columns.map((c, i) => (
              <th key={i} style={{ padding: "10px 12px", textAlign: "left", background: COLORS.surfaceAlt, color: COLORS.textMuted, fontWeight: 600, borderBottom: `1px solid ${COLORS.border}`, whiteSpace: "nowrap", position: "sticky", top: 0 }}>
                {c.label}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {display.map((r, ri) => (
            <tr key={ri} style={{ background: ri % 2 === 0 ? COLORS.surface : COLORS.surfaceAlt }}>
              {columns.map((c, ci) => (
                <td key={ci} style={{ padding: "8px 12px", borderBottom: `1px solid ${COLORS.border}`, color: c.color?.(r) || COLORS.text, whiteSpace: "nowrap" }}>
                  {c.render ? c.render(r, ri) : r[c.key] ?? ""}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
      {rows.length > maxRows && (
        <div style={{ padding: 8, textAlign: "center", fontSize: 11, color: COLORS.textMuted }}>
          Showing {maxRows} of {rows.length} rows
        </div>
      )}
    </div>
  );
}

// ─── Step indicators ────────────────────────────────────────────────
const STEPS = ["Upload", "Configure", "Review", "Results"];

function StepBar({ step }) {
  return (
    <div style={{ display: "flex", gap: 0, marginBottom: 32 }}>
      {STEPS.map((s, i) => (
        <div key={i} style={{ flex: 1, display: "flex", alignItems: "center" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <div style={{
              width: 28, height: 28, borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center",
              fontSize: 12, fontWeight: 700,
              background: i <= step ? COLORS.accent : COLORS.surfaceAlt,
              color: i <= step ? COLORS.white : COLORS.textMuted,
              border: `2px solid ${i <= step ? COLORS.accent : COLORS.border}`,
            }}>{i + 1}</div>
            <span style={{ fontSize: 12, fontWeight: 600, color: i <= step ? COLORS.text : COLORS.textMuted }}>{s}</span>
          </div>
          {i < STEPS.length - 1 && <div style={{ flex: 1, height: 2, background: i < step ? COLORS.accent : COLORS.border, marginLeft: 10, marginRight: 10 }} />}
        </div>
      ))}
    </div>
  );
}

// ─── Main App ───────────────────────────────────────────────────────
export default function AllocationTool() {
  const [step, setStep] = useState(0);
  const [hrFile, setHrFile] = useState(null);
  const [swFile, setSwFile] = useState(null);
  const [hrSheet, setHrSheet] = useState("");
  const [swSheet, setSwSheet] = useState("");
  const [hrHeaderRow, setHrHeaderRow] = useState(0);
  const [swHeaderRow, setSwHeaderRow] = useState(0);

  // Column mappings
  const [hrEmailCol, setHrEmailCol] = useState(-1);
  const [hrCcIdCol, setHrCcIdCol] = useState(-1);
  const [hrCcDescCol, setHrCcDescCol] = useState(-1);
  const [swEmailCol, setSwEmailCol] = useState(-1);
  const [swStatusCol, setSwStatusCol] = useState(-1);
  const [swLicenseCol, setSwLicenseCol] = useState(-1);
  const [swCcIdCol, setSwCcIdCol] = useState(-1);

  // Allocation config
  const [allocType, setAllocType] = useState("dollar"); // "dollar" | "percent"
  const [invoices, setInvoices] = useState([{ description: "", amount: "", months: "12" }]);
  const [licenseDelimiter, setLicenseDelimiter] = useState(";");
  const [rates, setRates] = useState({});
  const [softwareName, setSoftwareName] = useState("");
  const [rateCardFile, setRateCardFile] = useState(null);

  // Processing results
  const [processed, setProcessed] = useState(null);
  const [manualOverrides, setManualOverrides] = useState({});

  // ── Derived data ──────────────────────────────────────────────
  const hrData = useMemo(() => {
    if (!hrFile || !hrSheet) return [];
    return parseSheet(hrFile.wb, hrSheet);
  }, [hrFile, hrSheet]);

  const hrHeaders = useMemo(() => {
    if (!hrData.length) return [];
    return (hrData[hrHeaderRow] || []).map((x) => (x || "").toString());
  }, [hrData, hrHeaderRow]);

  const swData = useMemo(() => {
    if (!swFile || !swSheet) return [];
    return parseSheet(swFile.wb, swSheet);
  }, [swFile, swSheet]);

  const swHeaders = useMemo(() => {
    if (!swData.length) return [];
    return (swData[swHeaderRow] || []).map((x) => (x || "").toString());
  }, [swData, swHeaderRow]);

  // Auto-detect columns when headers change
  const autoDetect = useCallback(() => {
    if (hrHeaders.length) {
      const e = guessColumn(hrHeaders, ["email", "e-mail", "mail"]);
      const c = guessColumn(hrHeaders, ["ccid", "cc id", "cost center id", "cc_id"]);
      const d = guessColumn(hrHeaders, ["cost center", "cc descr", "cc description", "cost center name"]);
      if (e >= 0) setHrEmailCol(e);
      if (c >= 0) setHrCcIdCol(c);
      if (d >= 0) setHrCcDescCol(d);
    }
    if (swHeaders.length) {
      const e = guessColumn(swHeaders, ["email", "e-mail", "mail"]);
      const s = guessColumn(swHeaders, ["status"]);
      const l = guessColumn(swHeaders, ["product config", "license", "subscription", "plan"]);
      const c = guessColumn(swHeaders, ["ccid", "cc id", "cost center"]);
      if (e >= 0) setSwEmailCol(e);
      if (s >= 0) setSwStatusCol(s);
      if (l >= 0) setSwLicenseCol(l);
      if (c >= 0) setSwCcIdCol(c);
    }
  }, [hrHeaders, swHeaders]);

  // Build HR lookup map
  const hrLookup = useMemo(() => {
    if (hrEmailCol < 0 || hrCcIdCol < 0) return {};
    const map = {};
    hrData.forEach((row, i) => {
      if (i <= hrHeaderRow) return;
      const email = normalizeEmail(row[hrEmailCol]);
      if (email) {
        map[email] = {
          ccId: (row[hrCcIdCol] || "").toString().trim(),
          ccDesc: hrCcDescCol >= 0 ? (row[hrCcDescCol] || "").toString().trim() : "",
        };
      }
    });
    return map;
  }, [hrData, hrHeaderRow, hrEmailCol, hrCcIdCol, hrCcDescCol]);

  // Detect unique license types (for Fixed Dollar)
  const uniqueLicenses = useMemo(() => {
    if (swLicenseCol < 0 || allocType !== "dollar") return [];
    const set = new Set();
    swData.forEach((row, i) => {
      if (i <= swHeaderRow) return;
      const val = (row[swLicenseCol] || "").toString().trim();
      if (!val) return;
      const parts = val.split(licenseDelimiter).map((p) => p.trim()).filter(Boolean);
      parts.forEach((p) => set.add(p));
    });
    return [...set].sort();
  }, [swData, swHeaderRow, swLicenseCol, licenseDelimiter, allocType]);

  // Unique cost centers from HR (for dropdown in flagged user assignment)
  const hrCostCenters = useMemo(() => {
    const map = new Map();
    Object.values(hrLookup).forEach((v) => {
      if (v.ccId && !map.has(v.ccId)) map.set(v.ccId, { ccId: v.ccId, ccDesc: v.ccDesc });
    });
    return [...map.values()].sort((a, b) => a.ccId.localeCompare(b.ccId));
  }, [hrLookup]);

  // Combined monthly amortized from all invoices
  const totalMonthlyAmort = useMemo(() => {
    return invoices.reduce((sum, inv) => {
      const amt = parseFloat(inv.amount) || 0;
      const mo = parseFloat(inv.months) || 12;
      return sum + (amt / mo);
    }, 0);
  }, [invoices]);

  const totalInvoiceAmount = useMemo(() => {
    return invoices.reduce((sum, inv) => sum + (parseFloat(inv.amount) || 0), 0);
  }, [invoices]);

  // Rate card upload handler
  const handleRateCardUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const wb = XLSX.read(ev.target.result, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      if (rows.length < 2) return;
      const newRates = {};
      // Try to detect header row - skip rows where col 2 isn't numeric
      let startRow = 0;
      for (let i = 0; i < rows.length; i++) {
        if (rows[i].length >= 2 && !isNaN(parseFloat(rows[i][1]))) { startRow = i; break; }
        startRow = i + 1;
      }
      for (let i = startRow; i < rows.length; i++) {
        const name = (rows[i][0] || "").toString().trim();
        const rate = parseFloat(rows[i][1]);
        if (name && !isNaN(rate)) newRates[name] = rate;
      }
      setRates((prev) => ({ ...prev, ...newRates }));
      setRateCardFile(file.name);
    };
    reader.readAsArrayBuffer(file);
  };

  // ── Processing ────────────────────────────────────────────────
  const runAllocation = () => {
    const users = [];
    swData.forEach((row, i) => {
      if (i <= swHeaderRow) return;
      const email = normalizeEmail(row[swEmailCol]);
      if (!email) return;

      const status = swStatusCol >= 0 ? (row[swStatusCol] || "").toString().trim() : "Active";
      const itCcId = swCcIdCol >= 0 ? (row[swCcIdCol] || "").toString().trim() : "";
      const hrMatch = hrLookup[email];

      let ccId = hrMatch?.ccId || "";
      let ccDesc = hrMatch?.ccDesc || "";
      let source = hrMatch ? "HR" : "";

      if (!ccId && itCcId) {
        ccId = "";
        ccDesc = "";
        source = "";
      }

      const flagged = !hrMatch;

      if (allocType === "dollar") {
        const licRaw = swLicenseCol >= 0 ? (row[swLicenseCol] || "").toString().trim() : "";
        const licenses = licRaw ? licRaw.split(licenseDelimiter).map((p) => p.trim()).filter(Boolean) : [];
        const licenseCosts = licenses.map((l) => ({ name: l, rate: rates[l] || 0 }));
        const totalCost = licenseCosts.reduce((s, lc) => s + lc.rate, 0);
        users.push({ email, status, ccId, ccDesc, source, flagged, licenses, licenseCosts, totalCost, itCcId });
      } else {
        users.push({ email, status, ccId, ccDesc, source, flagged, itCcId });
      }
    });

    // For Fixed %, filter to Active only
    const activeUsers = allocType === "percent" ? users.filter((u) => u.status === "Active") : users;
    const flaggedUsers = activeUsers.filter((u) => u.flagged);
    const paidUsers = allocType === "dollar" ? activeUsers.filter((u) => u.totalCost > 0) : activeUsers;

    setProcessed({ users: activeUsers, flaggedUsers, paidUsers });
    setManualOverrides({});
    setStep(2);
  };

  // Apply overrides and compute final allocation
  const finalAllocation = useMemo(() => {
    if (!processed) return null;
    const users = processed.paidUsers.map((u) => {
      const override = manualOverrides[u.email];
      if (override) return { ...u, ccId: override.ccId, ccDesc: override.ccDesc, source: "Manual" };
      return u;
    });

    const assigned = users.filter((u) => u.ccId);
    const stillFlagged = users.filter((u) => !u.ccId);

    if (allocType === "dollar") {
      const groups = _.groupBy(assigned, "ccId");
      const summary = Object.entries(groups)
        .map(([ccId, members]) => ({
          ccId,
          ccDesc: members[0].ccDesc,
          userCount: members.length,
          paidLicenseCount: members.reduce((s, m) => s + m.licenses.length, 0),
          totalCost: members.reduce((s, m) => s + m.totalCost, 0),
        }))
        .sort((a, b) => a.ccId.localeCompare(b.ccId));
      const grandTotal = summary.reduce((s, r) => s + r.totalCost, 0);
      return { type: "dollar", users: assigned, summary, grandTotal, stillFlagged };
    } else {
      const total = assigned.length;
      const groups = _.groupBy(assigned, "ccId");
      const monthlyAmort = totalMonthlyAmort;
      const summary = Object.entries(groups)
        .map(([ccId, members]) => ({
          ccId,
          ccDesc: members[0].ccDesc,
          userCount: members.length,
          pct: members.length / total,
          monthlyCharge: (members.length / total) * monthlyAmort,
        }))
        .sort((a, b) => a.ccId.localeCompare(b.ccId));
      return { type: "percent", users: assigned, summary, totalUsers: total, monthlyAmort, stillFlagged, invoices, totalInvoiceAmount };
    }
  }, [processed, manualOverrides, allocType, totalMonthlyAmort, totalInvoiceAmount, invoices, rates]);

  // ── Excel Export ──────────────────────────────────────────────
  const exportExcel = () => {
    if (!finalAllocation) return;
    const wb = XLSX.utils.book_new();

    if (finalAllocation.type === "dollar") {
      // Tab 1: Allocation Entry for Accounting (pivot)
      const sumData = [
        ["", "", "", `${softwareName || "Software"} Cost Allocation per Month`],
        ["CCID", "Cost Center Description", "# of Paid License Holders", "Total Monthly Cost"],
        ...finalAllocation.summary.map((r) => [r.ccId, r.ccDesc, r.userCount, r.totalCost]),
        ["Grand Total", "", finalAllocation.summary.reduce((s, r) => s + r.userCount, 0), finalAllocation.grandTotal],
      ];
      const ws1 = XLSX.utils.aoa_to_sheet(sumData);
      ws1["!cols"] = [{ wch: 10 }, { wch: 50 }, { wch: 22 }, { wch: 20 }];
      XLSX.utils.book_append_sheet(wb, ws1, "AllocEntry for Accounting");

      // Tab 2: User Detail
      const maxLic = Math.max(...finalAllocation.users.map((u) => u.licenses.length), 1);
      const licHeaders = [];
      for (let i = 0; i < maxLic; i++) {
        licHeaders.push(`License ${i + 1}`, `License ${i + 1} Cost`);
      }
      const detailHeader = ["Email", "CC ID", "CC Desc", "Source", ...licHeaders, "Total Cost"];
      const detailRows = finalAllocation.users.map((u) => {
        const licCols = [];
        for (let i = 0; i < maxLic; i++) {
          licCols.push(u.licenseCosts[i]?.name || "", u.licenseCosts[i]?.rate || 0);
        }
        return [u.email, u.ccId, u.ccDesc, u.source, ...licCols, u.totalCost];
      });
      const ws2 = XLSX.utils.aoa_to_sheet([detailHeader, ...detailRows]);
      ws2["!cols"] = [{ wch: 35 }, { wch: 10 }, { wch: 50 }, { wch: 8 }, ...licHeaders.map(() => ({ wch: 18 })), { wch: 12 }];
      XLSX.utils.book_append_sheet(wb, ws2, "User Detail");

      // Tab 3: Rate Sheet
      const rateRows = [["License Type", "Monthly Rate"], ...Object.entries(rates).map(([k, v]) => [k, v])];
      const ws3 = XLSX.utils.aoa_to_sheet(rateRows);
      ws3["!cols"] = [{ wch: 45 }, { wch: 15 }];
      XLSX.utils.book_append_sheet(wb, ws3, "Rate Sheet");
    } else {
      // Tab 1: Monthly Allocation
      const monthlyAmort = finalAllocation.monthlyAmort;
      const invoiceRows = finalAllocation.invoices.map((inv) => {
        const desc = inv.description ? `${inv.description}: ` : "";
        return [`${desc}$${fmt(parseFloat(inv.amount) || 0)} over ${inv.months || 12} months → $${fmt((parseFloat(inv.amount) || 0) / (parseFloat(inv.months) || 12))}/mo`];
      });
      const sumData = [
        [`${softwareName || "Software"} Fixed % Allocation`],
        [`Total Invoice Amount: $${fmt(finalAllocation.totalInvoiceAmount)}`],
        ...invoiceRows,
        [`Combined Monthly Amortized: $${fmt(monthlyAmort)}`],
        [],
        ["CC ID", "CC Description", "User Count", "% Allocation", "Monthly Charge"],
        ...finalAllocation.summary.map((r) => [r.ccId, r.ccDesc, r.userCount, r.pct, r.monthlyCharge]),
        ["Grand Total", "", finalAllocation.totalUsers, 1, monthlyAmort],
      ];
      const ws1 = XLSX.utils.aoa_to_sheet(sumData);
      ws1["!cols"] = [{ wch: 12 }, { wch: 55 }, { wch: 12 }, { wch: 14 }, { wch: 16 }];
      XLSX.utils.book_append_sheet(wb, ws1, "Monthly Allocation");

      // Tab 2: Active Users
      const userRows = [
        ["Email", "Status", "CC ID", "CC Description", "Source"],
        ...finalAllocation.users.map((u) => [u.email, u.status || "Active", u.ccId, u.ccDesc, u.source]),
      ];
      const ws2 = XLSX.utils.aoa_to_sheet(userRows);
      ws2["!cols"] = [{ wch: 35 }, { wch: 10 }, { wch: 10 }, { wch: 55 }, { wch: 8 }];
      XLSX.utils.book_append_sheet(wb, ws2, "Active Users");
    }

    // Flagged users tab (if any)
    if (finalAllocation.stillFlagged.length > 0) {
      const flagData = [
        ["Email", "IT Cost Center (stale)", "Action Needed"],
        ...finalAllocation.stillFlagged.map((u) => [u.email, u.itCcId, "Assign cost center manually"]),
      ];
      const wsF = XLSX.utils.aoa_to_sheet(flagData);
      wsF["!cols"] = [{ wch: 35 }, { wch: 20 }, { wch: 30 }];
      XLSX.utils.book_append_sheet(wb, wsF, "Flagged - Unmatched");
    }

    XLSX.writeFile(wb, `${(softwareName || "Software").replace(/\s+/g, "_")}_Allocation_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  // ── Render ────────────────────────────────────────────────────
  return (
    <div style={{ fontFamily: "'IBM Plex Sans', sans-serif", background: COLORS.bg, color: COLORS.text, minHeight: "100vh", padding: "24px 32px" }}>
      <link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans:wght@400;500;600;700&display=swap" rel="stylesheet" />

      <div style={{ maxWidth: 960, margin: "0 auto" }}>
        {/* Header */}
        <div style={{ marginBottom: 8 }}>
          <div style={{ fontSize: 11, fontWeight: 600, color: COLORS.accent, letterSpacing: 2, textTransform: "uppercase", marginBottom: 4 }}>MEG Finance</div>
          <h1 style={{ fontSize: 22, fontWeight: 700, margin: 0, color: COLORS.text }}>STEVE</h1>
          <p style={{ fontSize: 13, color: COLORS.textMuted, marginTop: 4, marginBottom: 0 }}>Software Tool Expense Visibility Engine</p>
        </div>

        <StepBar step={step} />

        {/* ───── STEP 0: Upload ───── */}
        {step === 0 && (
          <div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, marginBottom: 24 }}>
              <div>
                <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 8 }}>HR Employee File</div>
                <FileUpload label="Drop HR Employee List" hint=".xlsx with emails and cost centers" parsed={hrFile} onParsed={setHrFile} />
                {hrFile && (
                  <div style={{ marginTop: 10 }}>
                    <span style={{ fontSize: 12, color: COLORS.textMuted, marginRight: 8 }}>Sheet:</span>
                    <select value={hrSheet} onChange={(e) => setHrSheet(e.target.value)} style={{ ...baseInput, width: "auto" }}>
                      <option value="">Select sheet</option>
                      {hrFile.sheets.map((s) => <option key={s} value={s}>{s}</option>)}
                    </select>
                  </div>
                )}
              </div>
              <div>
                <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 8 }}>Software User List</div>
                <FileUpload label="Drop Software User Export" hint=".xlsx from IT (any software)" parsed={swFile} onParsed={setSwFile} />
                {swFile && (
                  <div style={{ marginTop: 10 }}>
                    <span style={{ fontSize: 12, color: COLORS.textMuted, marginRight: 8 }}>Sheet:</span>
                    <select value={swSheet} onChange={(e) => setSwSheet(e.target.value)} style={{ ...baseInput, width: "auto" }}>
                      <option value="">Select sheet</option>
                      {swFile.sheets.map((s) => <option key={s} value={s}>{s}</option>)}
                    </select>
                  </div>
                )}
              </div>
            </div>

            {hrSheet && swSheet && (
              <div style={{ background: COLORS.surface, borderRadius: 10, padding: 20, border: `1px solid ${COLORS.border}`, marginBottom: 20 }}>
                <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 12 }}>Header Row Detection</div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
                  <div>
                    <span style={{ fontSize: 12, color: COLORS.textMuted }}>HR file header row:</span>
                    <select value={hrHeaderRow} onChange={(e) => setHrHeaderRow(parseInt(e.target.value))} style={{ ...baseInput, marginLeft: 8, width: "auto" }}>
                      {hrData.slice(0, 15).map((r, i) => (
                        <option key={i} value={i}>Row {i + 1}: {(r || []).slice(0, 3).filter(Boolean).join(" | ").slice(0, 60)}</option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <span style={{ fontSize: 12, color: COLORS.textMuted }}>Software file header row:</span>
                    <select value={swHeaderRow} onChange={(e) => setSwHeaderRow(parseInt(e.target.value))} style={{ ...baseInput, marginLeft: 8, width: "auto" }}>
                      {swData.slice(0, 15).map((r, i) => (
                        <option key={i} value={i}>Row {i + 1}: {(r || []).slice(0, 3).filter(Boolean).join(" | ").slice(0, 60)}</option>
                      ))}
                    </select>
                  </div>
                </div>
              </div>
            )}

            <div style={{ display: "flex", justifyContent: "flex-end" }}>
              <button
                disabled={!hrSheet || !swSheet}
                onClick={() => { autoDetect(); setStep(1); }}
                style={{ ...btn("primary"), opacity: !hrSheet || !swSheet ? 0.4 : 1 }}
              >
                Next: Configure →
              </button>
            </div>
          </div>
        )}

        {/* ───── STEP 1: Configure ───── */}
        {step === 1 && (
          <div>
            {/* Software Name */}
            <div style={{ background: COLORS.surface, borderRadius: 10, padding: 20, border: `1px solid ${COLORS.border}`, marginBottom: 20 }}>
              <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 12 }}>Software Details</div>
              <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 16 }}>
                <span style={{ fontSize: 12, color: COLORS.textMuted, width: 120 }}>Software Name</span>
                <input value={softwareName} onChange={(e) => setSoftwareName(e.target.value)} placeholder="e.g. Adobe, Dropbox, Zoom" style={{ ...baseInput, flex: 1, maxWidth: 300 }} />
              </div>
              <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 10 }}>Allocation Method</div>
              <div style={{ display: "flex", gap: 10 }}>
                <button style={chip(allocType === "dollar")} onClick={() => setAllocType("dollar")}>
                  Fixed Dollar Allocation
                </button>
                <button style={chip(allocType === "percent")} onClick={() => setAllocType("percent")}>
                  Fixed % Allocation
                </button>
              </div>
              <p style={{ fontSize: 11, color: COLORS.textMuted, marginTop: 8 }}>
                {allocType === "dollar"
                  ? "Charge each cost center based on per-user license costs."
                  : "Allocate a % of monthly amortized cost to each cost center by user count."}
              </p>
            </div>

            {/* Column Mapping */}
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, marginBottom: 20 }}>
              <div style={{ background: COLORS.surface, borderRadius: 10, padding: 20, border: `1px solid ${COLORS.border}` }}>
                <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 12 }}>HR File Columns</div>
                <ColumnMapper label="Email (for lookup)" headers={hrHeaders} value={hrEmailCol} onChange={setHrEmailCol} />
                <ColumnMapper label="Cost Center ID" headers={hrHeaders} value={hrCcIdCol} onChange={setHrCcIdCol} />
                <ColumnMapper label="Cost Center Description" headers={hrHeaders} value={hrCcDescCol} onChange={setHrCcDescCol} />
                <div style={{ fontSize: 11, color: COLORS.success, marginTop: 8 }}>{Object.keys(hrLookup).length} employees mapped</div>
              </div>
              <div style={{ background: COLORS.surface, borderRadius: 10, padding: 20, border: `1px solid ${COLORS.border}` }}>
                <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 12 }}>Software User Columns</div>
                <ColumnMapper label="Email" headers={swHeaders} value={swEmailCol} onChange={setSwEmailCol} />
                <ColumnMapper label="Status (opt.)" headers={swHeaders} value={swStatusCol} onChange={setSwStatusCol} />
                {allocType === "dollar" && (
                  <ColumnMapper label="License / Product Config" headers={swHeaders} value={swLicenseCol} onChange={setSwLicenseCol} />
                )}
                <ColumnMapper label="IT Cost Center (fallback)" headers={swHeaders} value={swCcIdCol} onChange={setSwCcIdCol} />
              </div>
            </div>

            {/* Allocation-specific config */}
            {allocType === "dollar" && (
              <div style={{ background: COLORS.surface, borderRadius: 10, padding: 20, border: `1px solid ${COLORS.border}`, marginBottom: 20 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
                  <div style={{ fontSize: 13, fontWeight: 600 }}>License Rate Sheet</div>
                  <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                    <label style={{ ...btn("ghost"), padding: "6px 14px", fontSize: 11, display: "inline-flex", alignItems: "center", gap: 6 }}>
                      ↑ Upload Rate Card
                      <input type="file" accept=".xlsx,.xls,.csv" onChange={handleRateCardUpload} style={{ display: "none" }} />
                    </label>
                    {rateCardFile && <span style={{ fontSize: 11, color: COLORS.success }}>✓ {rateCardFile}</span>}
                    <span style={{ fontSize: 11, color: COLORS.textMuted }}>Delimiter:</span>
                    <input value={licenseDelimiter} onChange={(e) => setLicenseDelimiter(e.target.value)} style={{ ...baseInput, width: 40, textAlign: "center" }} />
                  </div>
                </div>
                <div style={{ fontSize: 11, color: COLORS.textMuted, marginBottom: 10 }}>
                  Upload a rate card (.xlsx with License Type in column A, Monthly Rate in column B) or set rates manually below.
                </div>
                {uniqueLicenses.length === 0 && (
                  <div style={{ fontSize: 12, color: COLORS.warn, background: COLORS.warnBg, padding: 10, borderRadius: 6 }}>
                    Map the License column above to detect license types.
                  </div>
                )}
                <div style={{ maxHeight: 320, overflowY: "auto" }}>
                  {uniqueLicenses.map((lic) => (
                    <div key={lic} style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6 }}>
                      <span style={{ fontSize: 12, color: COLORS.text, flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{lic}</span>
                      <span style={{ fontSize: 11, color: COLORS.textMuted }}>$/mo</span>
                      <input
                        type="number"
                        step="0.01"
                        value={rates[lic] ?? ""}
                        onChange={(e) => setRates((r) => ({ ...r, [lic]: parseFloat(e.target.value) || 0 }))}
                        style={{ ...baseInput, width: 100, textAlign: "right" }}
                        placeholder="0.00"
                      />
                    </div>
                  ))}
                </div>
              </div>
            )}

            {allocType === "percent" && (
              <div style={{ background: COLORS.surface, borderRadius: 10, padding: 20, border: `1px solid ${COLORS.border}`, marginBottom: 20 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
                  <div style={{ fontSize: 13, fontWeight: 600 }}>Invoices</div>
                  <button
                    onClick={() => setInvoices((inv) => [...inv, { description: "", amount: "", months: "12" }])}
                    style={{ ...btn("ghost"), padding: "4px 12px", fontSize: 11 }}
                  >
                    + Add Invoice
                  </button>
                </div>
                <div style={{ fontSize: 11, color: COLORS.textMuted, marginBottom: 12 }}>
                  Add one or more invoices. The monthly amortized cost from all invoices will be combined for the allocation.
                </div>
                {invoices.map((inv, idx) => (
                  <div key={idx} style={{ display: "flex", gap: 12, marginBottom: 10, alignItems: "flex-end", background: COLORS.surfaceAlt, padding: 12, borderRadius: 8 }}>
                    <div style={{ flex: 1.5 }}>
                      <div style={{ fontSize: 11, color: COLORS.textMuted, marginBottom: 3 }}>Description (opt.)</div>
                      <input
                        value={inv.description}
                        onChange={(e) => setInvoices((arr) => arr.map((v, i) => i === idx ? { ...v, description: e.target.value } : v))}
                        style={{ ...baseInput, width: "100%" }}
                        placeholder="e.g. Dropbox Annual"
                      />
                    </div>
                    <div style={{ flex: 1 }}>
                      <div style={{ fontSize: 11, color: COLORS.textMuted, marginBottom: 3 }}>Amount ($)</div>
                      <input
                        type="number"
                        value={inv.amount}
                        onChange={(e) => setInvoices((arr) => arr.map((v, i) => i === idx ? { ...v, amount: e.target.value } : v))}
                        style={{ ...baseInput, width: "100%" }}
                        placeholder="e.g. 69600"
                      />
                    </div>
                    <div style={{ flex: 0.7 }}>
                      <div style={{ fontSize: 11, color: COLORS.textMuted, marginBottom: 3 }}>Months</div>
                      <input
                        type="number"
                        value={inv.months}
                        onChange={(e) => setInvoices((arr) => arr.map((v, i) => i === idx ? { ...v, months: e.target.value } : v))}
                        style={{ ...baseInput, width: "100%" }}
                      />
                    </div>
                    <div style={{ fontSize: 11, color: COLORS.accent, fontFamily: "'IBM Plex Mono', monospace", width: 90, textAlign: "right", paddingBottom: 2 }}>
                      {inv.amount && inv.months ? `$${fmt((parseFloat(inv.amount) || 0) / (parseFloat(inv.months) || 12))}/mo` : ""}
                    </div>
                    {invoices.length > 1 && (
                      <button
                        onClick={() => setInvoices((arr) => arr.filter((_, i) => i !== idx))}
                        style={{ background: "none", border: "none", color: COLORS.danger, cursor: "pointer", fontSize: 14, padding: "4px 8px", lineHeight: 1 }}
                        title="Remove invoice"
                      >×</button>
                    )}
                  </div>
                ))}
                {totalMonthlyAmort > 0 && (
                  <div style={{ marginTop: 12, background: COLORS.surfaceAlt, borderRadius: 8, padding: 12, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <div>
                      <div style={{ fontSize: 11, color: COLORS.textMuted }}>Combined Monthly Amortized</div>
                      {invoices.length > 1 && <div style={{ fontSize: 10, color: COLORS.textMuted }}>from {invoices.filter((i) => parseFloat(i.amount) > 0).length} invoice(s), total ${fmt(totalInvoiceAmount)}</div>}
                    </div>
                    <div style={{ fontSize: 22, fontWeight: 700, color: COLORS.accent, fontFamily: "'IBM Plex Mono', monospace" }}>
                      ${fmt(totalMonthlyAmort)}
                    </div>
                  </div>
                )}
              </div>
            )}

            <div style={{ display: "flex", justifyContent: "space-between" }}>
              <button onClick={() => setStep(0)} style={btn("ghost")}>← Back</button>
              <button
                onClick={runAllocation}
                disabled={hrEmailCol < 0 || hrCcIdCol < 0 || swEmailCol < 0}
                style={{ ...btn("primary"), opacity: hrEmailCol < 0 || hrCcIdCol < 0 || swEmailCol < 0 ? 0.4 : 1 }}
              >
                Run Allocation →
              </button>
            </div>
          </div>
        )}

        {/* ───── STEP 2: Review ───── */}
        {step === 2 && processed && (
          <div>
            {/* Stats */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12, marginBottom: 24 }}>
              {[
                { label: "Total Users", value: processed.users.length, color: COLORS.text },
                { label: allocType === "dollar" ? "Paid Users" : "Active Users", value: processed.paidUsers.length, color: COLORS.accent },
                { label: "HR Matched", value: processed.paidUsers.filter((u) => !u.flagged).length, color: COLORS.success },
                { label: "Flagged (No Match)", value: processed.flaggedUsers.length, color: processed.flaggedUsers.length > 0 ? COLORS.warn : COLORS.success },
              ].map((s, i) => (
                <div key={i} style={{ background: COLORS.surface, borderRadius: 10, padding: 16, border: `1px solid ${COLORS.border}`, textAlign: "center" }}>
                  <div style={{ fontSize: 24, fontWeight: 700, color: s.color, fontFamily: "'IBM Plex Mono', monospace" }}>{s.value}</div>
                  <div style={{ fontSize: 11, color: COLORS.textMuted, marginTop: 2 }}>{s.label}</div>
                </div>
              ))}
            </div>

            {/* Flagged users for manual assignment */}
            {processed.flaggedUsers.length > 0 && (
              <div style={{ background: COLORS.warnBg, borderRadius: 10, padding: 20, border: `1px solid ${COLORS.warn}33`, marginBottom: 24 }}>
                <div style={{ fontSize: 13, fontWeight: 600, color: COLORS.warn, marginBottom: 4 }}>
                  ⚠ {processed.flaggedUsers.filter((u) => !manualOverrides[u.email]?.ccId).length} users need cost center assignment
                </div>
                <p style={{ fontSize: 12, color: COLORS.textMuted, margin: "4px 0 12px" }}>
                  These emails weren't found in the HR file. Assign cost centers below or leave blank to exclude from allocation.
                </p>
                <div style={{ maxHeight: 400, overflowY: "auto" }}>
                  {processed.flaggedUsers.map((u) => (
                    <div key={u.email} style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6, background: COLORS.surface, padding: "8px 12px", borderRadius: 6 }}>
                      <span style={{ fontSize: 12, fontFamily: "'IBM Plex Mono', monospace", flex: 1, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{u.email}</span>
                      {u.itCcId && <span style={{ fontSize: 10, color: COLORS.textMuted, flexShrink: 0 }}>IT: {u.itCcId}</span>}
                      <div style={{ width: 280, flexShrink: 0 }}>
                        <CostCenterPicker
                          costCenters={hrCostCenters}
                          value={manualOverrides[u.email]?.ccId || ""}
                          placeholder="Search or type CC ID..."
                          onChange={(ccId, ccDesc) => setManualOverrides((o) => ({
                            ...o,
                            [u.email]: { ccId, ccDesc: ccDesc || o[u.email]?.ccDesc || "" },
                          }))}
                        />
                      </div>
                      <span style={{ fontSize: 11, color: COLORS.textMuted, width: 180, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", flexShrink: 0 }}>
                        {manualOverrides[u.email]?.ccDesc || ""}
                      </span>
                    </div>
                  ))}
                </div>
              </div>
            )}

            <div style={{ display: "flex", justifyContent: "space-between" }}>
              <button onClick={() => setStep(1)} style={btn("ghost")}>← Back</button>
              <button onClick={() => setStep(3)} style={btn("primary")}>View Results →</button>
            </div>
          </div>
        )}

        {/* ───── STEP 3: Results ───── */}
        {step === 3 && finalAllocation && (
          <div>
            {/* Summary Header */}
            <div style={{ background: COLORS.surface, borderRadius: 10, padding: 20, border: `1px solid ${COLORS.border}`, marginBottom: 20, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div>
                <div style={{ fontSize: 16, fontWeight: 700 }}>{softwareName || "Software"} Allocation</div>
                <div style={{ fontSize: 12, color: COLORS.textMuted, marginTop: 2 }}>
                  {finalAllocation.type === "dollar"
                    ? `${finalAllocation.users.length} users across ${finalAllocation.summary.length} cost centers · $${fmt(finalAllocation.grandTotal)}/mo`
                    : `${finalAllocation.totalUsers} active users across ${finalAllocation.summary.length} cost centers · $${fmt(finalAllocation.monthlyAmort)}/mo amortized${finalAllocation.invoices.length > 1 ? ` (${finalAllocation.invoices.length} invoices)` : ""}`}
                </div>
                {finalAllocation.stillFlagged.length > 0 && (
                  <div style={{ fontSize: 11, color: COLORS.warn, marginTop: 4 }}>
                    {finalAllocation.stillFlagged.length} user(s) excluded (unassigned cost center)
                  </div>
                )}
              </div>
              <button onClick={exportExcel} style={{ ...btn("primary"), display: "flex", alignItems: "center", gap: 6 }}>
                ↓ Download Excel
              </button>
            </div>

            {/* Allocation Table */}
            {finalAllocation.type === "dollar" ? (
              <DataTable
                columns={[
                  { key: "ccId", label: "CC ID" },
                  { key: "ccDesc", label: "Cost Center" },
                  { key: "userCount", label: "Users" },
                  { key: "paidLicenseCount", label: "Licenses" },
                  { key: "totalCost", label: "Monthly Cost", render: (r) => `$${fmt(r.totalCost)}` },
                ]}
                rows={finalAllocation.summary}
              />
            ) : (
              <DataTable
                columns={[
                  { key: "ccId", label: "CC ID" },
                  { key: "ccDesc", label: "Cost Center" },
                  { key: "userCount", label: "Users" },
                  { key: "pct", label: "% Allocation", render: (r) => fmtPct(r.pct) },
                  { key: "monthlyCharge", label: "Monthly Charge", render: (r) => `$${fmt(r.monthlyCharge)}` },
                ]}
                rows={finalAllocation.summary}
              />
            )}

            <div style={{ display: "flex", justifyContent: "space-between", marginTop: 24 }}>
              <button onClick={() => setStep(2)} style={btn("ghost")}>← Back to Review</button>
              <button onClick={() => { setStep(0); setProcessed(null); setHrFile(null); setSwFile(null); setHrSheet(""); setSwSheet(""); setSoftwareName(""); setRates({}); setInvoices([{ description: "", amount: "", months: "12" }]); setManualOverrides({}); setRateCardFile(null); }} style={btn("ghost")}>
                Start New Allocation
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
