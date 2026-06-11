import { useState, useEffect, useMemo, useRef } from "react";
import {
  ResponsiveContainer, PieChart, Pie, Cell, BarChart, Bar,
  ComposedChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
} from "recharts";
import {
  LayoutDashboard, Wallet, Receipt, BarChart3, FileText, Settings as SettingsIcon,
  Plus, Trash2, Pencil, Download, Search, X, HardHat, TrendingUp, TrendingDown,
  CircleDollarSign, Check, Building2, Filter, AlertTriangle, Upload, FileDown,
} from "lucide-react";
import {
  Document, Packer, Paragraph, TextRun, Table as DocTable, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, ShadingType,
} from "docx";
import seed from "./seed.json";

/* ----------------------------------------------------------------------------
   Theme
---------------------------------------------------------------------------- */
const T = {
  ink: "#1f2933", sub: "#5b6b7a", faint: "#8a97a3",
  line: "#e4e8ec", bg: "#f4f6f8", panel: "#ffffff",
  accent: "#b4690e", accentSoft: "#fbf2e6",
  positive: "#0f7b5a", positiveSoft: "#e7f4ee",
  negative: "#b3261e", negativeSoft: "#fbeceb",
  blue: "#2a5d8f", blueSoft: "#eaf1f7",
};
const CHART = ["#2a5d8f","#b4690e","#0f7b5a","#7a4ea3","#b3261e","#3a7d7d",
  "#8a6d3b","#c2792e","#4a7a3a","#9a4b6a","#6a5acd","#247ba0","#a05050","#5b6b7a"];

/* ----------------------------------------------------------------------------
   Defaults
---------------------------------------------------------------------------- */
const PERSONNEL_CAT = "Labour / Personnel";
const DEFAULT_CATEGORIES = [
  { name: PERSONNEL_CAT, subs: ["Mason", "Helper / Labourer", "Supervisor", "Electrician", "Plumber", "Painter", "Carpenter", "Steel Fixer", "Tiler"] },
  { name: "Cement", subs: [] },
  { name: "Steel / Rebar", subs: [] },
  { name: "Bricks / Blocks", subs: [] },
  { name: "Sand / Aggregate", subs: [] },
  { name: "Materials (General)", subs: [] },
  { name: "Sanitary", subs: ["Fittings", "Pipes", "Fixtures"] },
  { name: "Plumbing", subs: [] },
  { name: "Roofing", subs: [] },
  { name: "Flooring", subs: [] },
  { name: "Tiles / Marble", subs: [] },
  { name: "Kitchen", subs: ["Cabinets", "Countertop", "Fittings"] },
  { name: "Electrical", subs: ["Wiring", "Fixtures", "DB / Breakers"] },
  { name: "Paint", subs: [] },
  { name: "Woodwork / Carpentry", subs: ["Doors", "Windows", "Wardrobes"] },
  { name: "Glass / Aluminium", subs: [] },
  { name: "Plaster / POP", subs: [] },
  { name: "Excavation / Earthwork", subs: [] },
  { name: "Foundation / RCC", subs: [] },
  { name: "Waterproofing", subs: [] },
  { name: "Boundary Wall", subs: [] },
  { name: "Transport / Freight", subs: [] },
  { name: "Equipment Rental", subs: [] },
  { name: "Permits / Fees", subs: [] },
  { name: "Miscellaneous", subs: [] },
];

const DEFAULT_CONFIG = {
  projectName: "New Construction Project",
  ownerName: "Owner",
  contractorName: "Contractor",
  currency: "Rs",
  categories: DEFAULT_CATEGORIES,
  personnel: [],
};
const EMPTY = { config: DEFAULT_CONFIG, payments: [], expenses: [], reports: [] };
const STORE_KEY = "cpt:data:v2";
const PAY_MODES = ["Cash", "Bank Transfer", "Cheque", "Online", "Other"];

/* ----------------------------------------------------------------------------
   Helpers
---------------------------------------------------------------------------- */
const uid = () =>
  (typeof crypto !== "undefined" && crypto.randomUUID)
    ? crypto.randomUUID()
    : "id-" + Math.random().toString(36).slice(2) + Date.now().toString(36);
const today = () => new Date().toISOString().slice(0, 10);
const num = (v) => { const n = parseFloat(v); return isNaN(n) ? 0 : n; };
const fmtDate = (s) => {
  if (!s) return "—";
  const d = new Date(s + "T00:00:00");
  return isNaN(d) ? s : d.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
};
const makeFmt = (sym) => (n) => {
  const v = Number(n) || 0;
  const s = Math.abs(v).toLocaleString("en-US", { maximumFractionDigits: 2 });
  return `${v < 0 ? "−" : ""}${sym} ${s}`;
};
const monthKey = (s) => (s ? s.slice(0, 7) : "");
const monthLabel = (k) => {
  if (!k) return "";
  const [y, m] = k.split("-");
  return new Date(Number(y), Number(m) - 1, 1).toLocaleDateString("en-GB", { month: "short", year: "2-digit" });
};

async function storageGet(key) {
  try { if (window.storage?.get) { const r = await window.storage.get(key); return r ? r.value : null; } } catch (e) {}
  return null;
}
async function storageSet(key, value) {
  try { if (window.storage?.set) { await window.storage.set(key, value); return true; } } catch (e) {}
  return false;
}

function downloadCSV(filename, rows) {
  const esc = (c) => {
    const s = c == null ? "" : String(c);
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const csv = rows.map((r) => r.map(esc).join(",")).join("\n");
  const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; document.body.appendChild(a); a.click();
  a.remove(); setTimeout(() => URL.revokeObjectURL(url), 1500);
}

/* ----------------------------------------------------------------------------
   UI primitives
---------------------------------------------------------------------------- */
function Card({ children, style, className = "" }) {
  return (
    <div className={"rounded-xl " + className}
      style={{ background: T.panel, border: `1px solid ${T.line}`, ...style }}>
      {children}
    </div>
  );
}
function SectionTitle({ icon: Icon, children, right }) {
  return (
    <div className="flex items-center justify-between mb-3">
      <div className="flex items-center gap-2">
        {Icon && <Icon size={16} style={{ color: T.accent }} />}
        <h3 className="text-sm font-semibold tracking-wide uppercase" style={{ color: T.sub, letterSpacing: ".06em" }}>{children}</h3>
      </div>
      {right}
    </div>
  );
}
function Button({ children, onClick, variant = "primary", size = "md", icon: Icon, type = "button", disabled }) {
  const base = { borderRadius: 8, fontWeight: 600, display: "inline-flex", alignItems: "center", gap: 6, cursor: disabled ? "not-allowed" : "pointer", opacity: disabled ? .55 : 1, border: "1px solid transparent", whiteSpace: "nowrap" };
  const sizes = { sm: { padding: "5px 10px", fontSize: 12 }, md: { padding: "8px 14px", fontSize: 13 } };
  const variants = {
    primary: { background: T.ink, color: "#fff" },
    accent: { background: T.accent, color: "#fff" },
    soft: { background: T.bg, color: T.ink, border: `1px solid ${T.line}` },
    ghost: { background: "transparent", color: T.sub },
    danger: { background: T.negativeSoft, color: T.negative, border: `1px solid ${T.negative}33` },
  };
  return (
    <button type={type} onClick={onClick} disabled={disabled} style={{ ...base, ...sizes[size], ...variants[variant] }}>
      {Icon && <Icon size={size === "sm" ? 13 : 15} />}{children}
    </button>
  );
}
function Field({ label, children, hint }) {
  return (
    <label className="flex flex-col gap-1">
      <span className="text-xs font-medium" style={{ color: T.sub }}>{label}</span>
      {children}
      {hint && <span className="text-xs" style={{ color: T.faint }}>{hint}</span>}
    </label>
  );
}
const inputStyle = { padding: "8px 10px", borderRadius: 8, border: `1px solid ${T.line}`, fontSize: 13, background: "#fff", color: T.ink, width: "100%", outline: "none" };
function Input(props) { return <input {...props} style={{ ...inputStyle, ...(props.style || {}) }} />; }
function Select({ children, ...props }) {
  return <select {...props} style={{ ...inputStyle, ...(props.style || {}) }}>{children}</select>;
}
function Empty({ icon: Icon, title, hint }) {
  return (
    <div className="flex flex-col items-center justify-center text-center py-12 px-4">
      <div className="rounded-full p-3 mb-3" style={{ background: T.bg }}>
        {Icon && <Icon size={22} style={{ color: T.faint }} />}
      </div>
      <p className="text-sm font-semibold" style={{ color: T.sub }}>{title}</p>
      {hint && <p className="text-xs mt-1 max-w-xs" style={{ color: T.faint }}>{hint}</p>}
    </div>
  );
}
function ConfirmDelete({ onConfirm, onCancel }) {
  return (
    <span className="inline-flex items-center gap-1">
      <button onClick={onConfirm} title="Confirm delete" style={{ color: T.negative }} className="p-1"><Check size={15} /></button>
      <button onClick={onCancel} title="Cancel" style={{ color: T.faint }} className="p-1"><X size={15} /></button>
    </span>
  );
}

/* ----------------------------------------------------------------------------
   App
---------------------------------------------------------------------------- */
const NAV = [
  { id: "overview", label: "Overview", icon: LayoutDashboard },
  { id: "payments", label: "Owner Payments", icon: Wallet },
  { id: "expenses", label: "Expenses", icon: Receipt },
  { id: "analytics", label: "Analytics", icon: BarChart3 },
  { id: "reports", label: "Reports", icon: FileText },
  { id: "settings", label: "Settings", icon: SettingsIcon },
];

export default function App() {
  const [data, setData] = useState(EMPTY);
  const [loaded, setLoaded] = useState(false);
  const [tab, setTab] = useState("overview");
  const firstSave = useRef(true);

  useEffect(() => {
    (async () => {
      const raw = await storageGet(STORE_KEY);
      let initial = null;
      if (raw) {
        try { initial = normalize(JSON.parse(raw)); } catch (e) {}
      }
      // No saved data on this browser yet? Fall back to a baked-in seed
      // dataset if one was provided in src/seed.json.
      if (!initial && hasData(seed)) initial = normalize(seed);
      if (initial) setData(initial);
      setLoaded(true);
    })();
  }, []);

  useEffect(() => {
    if (!loaded) return;
    if (firstSave.current) { firstSave.current = false; return; }
    storageSet(STORE_KEY, JSON.stringify(data));
  }, [data, loaded]);

  const { config, payments, expenses, reports } = data;
  const fmt = useMemo(() => makeFmt(config.currency), [config.currency]);

  const totals = useMemo(() => {
    const received = payments.reduce((s, p) => s + num(p.amount), 0);
    const spent = expenses.reduce((s, e) => s + num(e.amount), 0);
    return { received, spent, balance: received - spent };
  }, [payments, expenses]);

  const setConfig = (patch) => setData((d) => ({ ...d, config: { ...d.config, ...patch } }));
  const upsertPayment = (p) => setData((d) => {
    const exists = d.payments.some((x) => x.id === p.id);
    return { ...d, payments: exists ? d.payments.map((x) => x.id === p.id ? p : x) : [p, ...d.payments] };
  });
  const removePayment = (id) => setData((d) => ({ ...d, payments: d.payments.filter((x) => x.id !== id) }));
  const upsertExpense = (e) => setData((d) => {
    const exists = d.expenses.some((x) => x.id === e.id);
    return { ...d, expenses: exists ? d.expenses.map((x) => x.id === e.id ? e : x) : [e, ...d.expenses] };
  });
  const removeExpense = (id) => setData((d) => ({ ...d, expenses: d.expenses.filter((x) => x.id !== id) }));
  const addReport = (r) => setData((d) => ({ ...d, reports: [r, ...d.reports] }));
  const removeReport = (id) => setData((d) => ({ ...d, reports: d.reports.filter((x) => x.id !== id) }));
  const replaceAll = (next) => { const n = normalize(next); if (n) { setData(n); return true; } return false; };

  if (!loaded) {
    return <div className="flex items-center justify-center" style={{ minHeight: 420, color: T.faint, fontSize: 13 }}>Loading your project ledger…</div>;
  }

  const ctx = { data, config, payments, expenses, reports, fmt, totals,
    setConfig, upsertPayment, removePayment, upsertExpense, removeExpense, addReport, removeReport, replaceAll, setTab };

  return (
    <div style={{ background: T.bg, color: T.ink, minHeight: 560, fontFamily: "system-ui, -apple-system, 'Segoe UI', Roboto, sans-serif", borderRadius: 14, overflow: "hidden", border: `1px solid ${T.line}` }}>
      <div className="flex" style={{ minHeight: 560 }}>
        {/* Sidebar */}
        <aside style={{ width: 216, background: "#fff", borderRight: `1px solid ${T.line}`, flexShrink: 0 }} className="hidden sm:flex flex-col">
          <Brand config={config} />
          <nav className="px-2 py-2 flex flex-col gap-1">
            {NAV.map((n) => {
              const active = tab === n.id;
              return (
                <button key={n.id} onClick={() => setTab(n.id)}
                  className="flex items-center gap-3 rounded-lg text-sm font-medium"
                  style={{ padding: "9px 12px", color: active ? T.ink : T.sub, background: active ? T.accentSoft : "transparent", border: active ? `1px solid ${T.accent}33` : "1px solid transparent" }}>
                  <n.icon size={17} style={{ color: active ? T.accent : T.faint }} />
                  {n.label}
                </button>
              );
            })}
          </nav>
          <div className="mt-auto p-3 text-xs" style={{ color: T.faint, borderTop: `1px solid ${T.line}` }}>
            Data is saved automatically to this dashboard.
          </div>
        </aside>

        {/* Main */}
        <main className="flex-1 min-w-0 flex flex-col">
          {/* Mobile tab bar */}
          <div className="sm:hidden flex gap-1 p-2 overflow-x-auto" style={{ background: "#fff", borderBottom: `1px solid ${T.line}` }}>
            {NAV.map((n) => (
              <button key={n.id} onClick={() => setTab(n.id)} className="flex items-center gap-1.5 rounded-lg text-xs font-medium"
                style={{ padding: "7px 10px", whiteSpace: "nowrap", color: tab === n.id ? T.accent : T.sub, background: tab === n.id ? T.accentSoft : "transparent" }}>
                <n.icon size={14} />{n.label}
              </button>
            ))}
          </div>

          <div className="p-4 sm:p-6 flex-1">
            {tab === "overview" && <Overview {...ctx} />}
            {tab === "payments" && <Payments {...ctx} />}
            {tab === "expenses" && <Expenses {...ctx} />}
            {tab === "analytics" && <Analytics {...ctx} />}
            {tab === "reports" && <Reports {...ctx} />}
            {tab === "settings" && <SettingsView {...ctx} />}
          </div>
        </main>
      </div>
    </div>
  );
}

function Brand({ config }) {
  return (
    <div className="p-4" style={{ borderBottom: `1px solid ${T.line}` }}>
      <div className="flex items-center gap-2">
        <div className="rounded-lg flex items-center justify-center" style={{ width: 34, height: 34, background: T.ink }}>
          <HardHat size={18} color="#fff" />
        </div>
        <div className="min-w-0">
          <div className="text-sm font-bold truncate" style={{ color: T.ink }}>{config.projectName || "Project"}</div>
          <div className="text-xs truncate" style={{ color: T.faint }}>Finance Ledger</div>
        </div>
      </div>
    </div>
  );
}

/* ----------------------------------------------------------------------------
   Page header
---------------------------------------------------------------------------- */
function PageHead({ title, subtitle, actions }) {
  return (
    <div className="flex items-start justify-between gap-3 mb-5 flex-wrap">
      <div>
        <h2 className="text-xl font-bold" style={{ color: T.ink }}>{title}</h2>
        {subtitle && <p className="text-sm mt-0.5" style={{ color: T.sub }}>{subtitle}</p>}
      </div>
      {actions}
    </div>
  );
}

/* ----------------------------------------------------------------------------
   Overview
---------------------------------------------------------------------------- */
function Overview({ config, payments, expenses, totals, fmt, setTab }) {
  const owed = totals.balance < 0;
  const byCat = useMemo(() => aggregate(expenses, "category"), [expenses]);
  const recent = useMemo(() => {
    const items = [
      ...payments.map((p) => ({ ...p, _type: "payment" })),
      ...expenses.map((e) => ({ ...e, _type: "expense" })),
    ].sort((a, b) => (b.date || "").localeCompare(a.date || "") || (b.createdAt || 0) - (a.createdAt || 0));
    return items.slice(0, 7);
  }, [payments, expenses]);

  return (
    <div>
      <PageHead title="Overview" subtitle={`${config.ownerName} → ${config.contractorName}`} />

      {/* Signature balance card */}
      <Card style={{ background: owed ? T.negativeSoft : T.positiveSoft, borderColor: owed ? `${T.negative}33` : `${T.positive}33` }} className="p-5 mb-4">
        <div className="flex items-center justify-between flex-wrap gap-4">
          <div>
            <div className="flex items-center gap-2 mb-1">
              {owed ? <TrendingDown size={16} style={{ color: T.negative }} /> : <CircleDollarSign size={16} style={{ color: T.positive }} />}
              <span className="text-xs font-semibold uppercase tracking-wide" style={{ color: owed ? T.negative : T.positive, letterSpacing: ".06em" }}>
                {owed ? `Owner owes ${config.contractorName}` : `Balance held by ${config.contractorName}`}
              </span>
            </div>
            <div className="text-3xl font-bold tabular-nums" style={{ color: owed ? T.negative : T.positive }}>
              {fmt(Math.abs(totals.balance))}
            </div>
            <p className="text-xs mt-1.5" style={{ color: T.sub, maxWidth: 420 }}>
              {owed
                ? `${config.contractorName} has spent more on the project than received and is owed this amount.`
                : `Unspent funds remaining with ${config.contractorName} from payments received.`}
            </p>
          </div>
          <div className="flex gap-2">
            <Button variant="accent" icon={Plus} onClick={() => setTab("payments")}>Record payment</Button>
            <Button variant="soft" icon={Plus} onClick={() => setTab("expenses")}>Add expense</Button>
          </div>
        </div>
      </Card>

      {/* KPIs */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 mb-4">
        <Kpi label="Total received" value={fmt(totals.received)} icon={Wallet} tint={T.blue} soft={T.blueSoft} sub={`${payments.length} payment${payments.length === 1 ? "" : "s"}`} />
        <Kpi label="Total spent" value={fmt(totals.spent)} icon={Receipt} tint={T.accent} soft={T.accentSoft} sub={`${expenses.length} expense${expenses.length === 1 ? "" : "s"}`} />
        <Kpi label="Spend ratio" value={totals.received > 0 ? Math.round((totals.spent / totals.received) * 100) + "%" : "—"} icon={TrendingUp} tint={T.positive} soft={T.positiveSoft} sub="of funds utilised" />
        <Kpi label="Categories used" value={byCat.length} icon={Filter} tint="#7a4ea3" soft="#f1ecf6" sub="distinct heads" />
      </div>

      <div className="grid lg:grid-cols-5 gap-4">
        <Card className="p-4 lg:col-span-2">
          <SectionTitle icon={BarChart3}>Spending by category</SectionTitle>
          {byCat.length === 0
            ? <Empty icon={BarChart3} title="No expenses yet" hint="Add expenses to see the breakdown." />
            : (
              <div style={{ height: 230 }}>
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie data={byCat.slice(0, 8)} dataKey="value" nameKey="key" innerRadius={48} outerRadius={82} paddingAngle={2}>
                      {byCat.slice(0, 8).map((_, i) => <Cell key={i} fill={CHART[i % CHART.length]} />)}
                    </Pie>
                    <Tooltip formatter={(v) => fmt(v)} contentStyle={tipStyle} />
                  </PieChart>
                </ResponsiveContainer>
              </div>
            )}
        </Card>

        <Card className="p-4 lg:col-span-3">
          <SectionTitle icon={Receipt} right={<button onClick={() => setTab("expenses")} className="text-xs font-semibold" style={{ color: T.accent }}>View all</button>}>Recent activity</SectionTitle>
          {recent.length === 0
            ? <Empty icon={Receipt} title="Nothing recorded yet" hint="Your latest payments and expenses appear here." />
            : (
              <div className="flex flex-col">
                {recent.map((r) => (
                  <div key={r.id} className="flex items-center justify-between py-2.5" style={{ borderBottom: `1px solid ${T.line}` }}>
                    <div className="flex items-center gap-3 min-w-0">
                      <div className="rounded-lg flex items-center justify-center flex-shrink-0" style={{ width: 30, height: 30, background: r._type === "payment" ? T.blueSoft : T.accentSoft }}>
                        {r._type === "payment" ? <Wallet size={15} style={{ color: T.blue }} /> : <Receipt size={15} style={{ color: T.accent }} />}
                      </div>
                      <div className="min-w-0">
                        <div className="text-sm font-medium truncate" style={{ color: T.ink }}>
                          {r._type === "payment" ? "Payment from owner" : (r.category + (r.subcategory ? ` · ${r.subcategory}` : ""))}
                        </div>
                        <div className="text-xs truncate" style={{ color: T.faint }}>
                          {fmtDate(r.date)}{r.person ? ` · ${r.person}` : ""}{r.note ? ` · ${r.note}` : ""}{r.description ? ` · ${r.description}` : ""}
                        </div>
                      </div>
                    </div>
                    <div className="text-sm font-semibold tabular-nums flex-shrink-0 ml-2" style={{ color: r._type === "payment" ? T.blue : T.ink }}>
                      {r._type === "payment" ? "+" : "−"}{fmt(num(r.amount)).replace("−", "")}
                    </div>
                  </div>
                ))}
              </div>
            )}
        </Card>
      </div>
    </div>
  );
}
function Kpi({ label, value, icon: Icon, tint, soft, sub }) {
  return (
    <Card className="p-4">
      <div className="flex items-center justify-between mb-2">
        <span className="text-xs font-medium" style={{ color: T.sub }}>{label}</span>
        <div className="rounded-lg flex items-center justify-center" style={{ width: 28, height: 28, background: soft }}>
          <Icon size={15} style={{ color: tint }} />
        </div>
      </div>
      <div className="text-xl font-bold tabular-nums" style={{ color: T.ink }}>{value}</div>
      {sub && <div className="text-xs mt-0.5" style={{ color: T.faint }}>{sub}</div>}
    </Card>
  );
}
const tipStyle = { borderRadius: 8, border: `1px solid ${T.line}`, fontSize: 12, boxShadow: "0 4px 16px rgba(0,0,0,.08)" };

/* ----------------------------------------------------------------------------
   Owner Payments
---------------------------------------------------------------------------- */
function Payments({ config, payments, fmt, totals, upsertPayment, removePayment }) {
  const blank = { date: today(), amount: "", mode: "Cash", note: "" };
  const [form, setForm] = useState(blank);
  const [editId, setEditId] = useState(null);
  const [confirmId, setConfirmId] = useState(null);

  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));
  const save = () => {
    if (num(form.amount) <= 0) return;
    const rec = { ...form, amount: num(form.amount), id: editId || uid(), createdAt: Date.now() };
    upsertPayment(rec); setForm(blank); setEditId(null);
  };
  const edit = (p) => { setForm({ date: p.date, amount: String(p.amount), mode: p.mode || "Cash", note: p.note || "" }); setEditId(p.id); };
  const sorted = [...payments].sort((a, b) => (b.date || "").localeCompare(a.date || ""));

  return (
    <div>
      <PageHead title="Owner Payments"
        subtitle={`Funds transferred from ${config.ownerName} to ${config.contractorName}`}
        actions={<div className="text-right"><div className="text-xs" style={{ color: T.faint }}>Total received</div><div className="text-lg font-bold tabular-nums" style={{ color: T.blue }}>{fmt(totals.received)}</div></div>}
      />

      <Card className="p-4 mb-4">
        <SectionTitle icon={editId ? Pencil : Plus}>{editId ? "Edit payment" : "Record a payment"}</SectionTitle>
        <div className="grid sm:grid-cols-4 gap-3">
          <Field label="Date"><Input type="date" value={form.date} onChange={(e) => set("date", e.target.value)} /></Field>
          <Field label={`Amount (${config.currency})`}><Input type="number" min="0" placeholder="0" value={form.amount} onChange={(e) => set("amount", e.target.value)} /></Field>
          <Field label="Payment method">
            <Select value={form.mode} onChange={(e) => set("mode", e.target.value)}>{PAY_MODES.map((m) => <option key={m}>{m}</option>)}</Select>
          </Field>
          <Field label="Reference / note"><Input placeholder="e.g. advance #2" value={form.note} onChange={(e) => set("note", e.target.value)} /></Field>
        </div>
        <div className="flex gap-2 mt-3">
          <Button variant="accent" icon={editId ? Check : Plus} onClick={save} disabled={num(form.amount) <= 0}>{editId ? "Update payment" : "Add payment"}</Button>
          {editId && <Button variant="ghost" onClick={() => { setForm(blank); setEditId(null); }}>Cancel</Button>}
        </div>
      </Card>

      <Card>
        <Table head={["Date", "Method", "Reference", "Amount", ""]} widths={["18%", "18%", "auto", "18%", "70px"]}>
          {sorted.length === 0
            ? <tr><td colSpan={5}><Empty icon={Wallet} title="No payments recorded" hint={`Record what ${config.ownerName} has paid to ${config.contractorName}.`} /></td></tr>
            : sorted.map((p) => (
              <tr key={p.id} style={{ borderTop: `1px solid ${T.line}` }}>
                <Td>{fmtDate(p.date)}</Td>
                <Td><Chip>{p.mode}</Chip></Td>
                <Td muted>{p.note || "—"}</Td>
                <Td bold style={{ color: T.blue }}>{fmt(num(p.amount))}</Td>
                <Td>
                  {confirmId === p.id
                    ? <ConfirmDelete onConfirm={() => { removePayment(p.id); setConfirmId(null); }} onCancel={() => setConfirmId(null)} />
                    : <RowActions onEdit={() => edit(p)} onDelete={() => setConfirmId(p.id)} />}
                </Td>
              </tr>
            ))}
        </Table>
      </Card>
    </div>
  );
}

/* ----------------------------------------------------------------------------
   Expenses
---------------------------------------------------------------------------- */
function Expenses({ config, expenses, fmt, totals, upsertExpense, removeExpense }) {
  const blank = { date: today(), category: PERSONNEL_CAT, subcategory: "", amount: "", person: "", vendor: "", mode: "Cash", description: "" };
  const [form, setForm] = useState(blank);
  const [editId, setEditId] = useState(null);
  const [confirmId, setConfirmId] = useState(null);

  const [fCat, setFCat] = useState("");
  const [fPerson, setFPerson] = useState("");
  const [q, setQ] = useState("");

  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));
  const catObj = config.categories.find((c) => c.name === form.category);
  const isPersonnel = form.category === PERSONNEL_CAT;

  const persons = useMemo(() => {
    const s = new Set(config.personnel || []);
    expenses.forEach((e) => e.person && s.add(e.person));
    return [...s].sort();
  }, [expenses, config.personnel]);

  const save = () => {
    if (num(form.amount) <= 0) return;
    const rec = { ...form, amount: num(form.amount), id: editId || uid(), createdAt: Date.now() };
    upsertExpense(rec); setForm({ ...blank, category: form.category }); setEditId(null);
  };
  const edit = (e) => {
    setForm({ date: e.date, category: e.category, subcategory: e.subcategory || "", amount: String(e.amount), person: e.person || "", vendor: e.vendor || "", mode: e.mode || "Cash", description: e.description || "" });
    setEditId(e.id);
  };

  const filtered = useMemo(() => {
    return expenses
      .filter((e) => !fCat || e.category === fCat)
      .filter((e) => !fPerson || e.person === fPerson)
      .filter((e) => {
        if (!q) return true;
        const hay = `${e.category} ${e.subcategory} ${e.person} ${e.vendor} ${e.description}`.toLowerCase();
        return hay.includes(q.toLowerCase());
      })
      .sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  }, [expenses, fCat, fPerson, q]);
  const filteredTotal = filtered.reduce((s, e) => s + num(e.amount), 0);

  return (
    <div>
      <PageHead title="Expenses"
        subtitle="Money spent by the contractor, recorded against project heads"
        actions={<div className="text-right"><div className="text-xs" style={{ color: T.faint }}>Total spent</div><div className="text-lg font-bold tabular-nums" style={{ color: T.accent }}>{fmt(totals.spent)}</div></div>}
      />

      <Card className="p-4 mb-4">
        <SectionTitle icon={editId ? Pencil : Plus}>{editId ? "Edit expense" : "Add an expense"}</SectionTitle>
        <div className="grid sm:grid-cols-3 gap-3">
          <Field label="Date"><Input type="date" value={form.date} onChange={(e) => set("date", e.target.value)} /></Field>
          <Field label="Category">
            <Select value={form.category} onChange={(e) => set("category", e.target.value) || set("subcategory", "")}>
              {config.categories.map((c) => <option key={c.name}>{c.name}</option>)}
            </Select>
          </Field>
          <Field label="Sub-category">
            {catObj && catObj.subs.length > 0
              ? <Select value={form.subcategory} onChange={(e) => set("subcategory", e.target.value)}>
                  <option value="">— optional —</option>
                  {catObj.subs.map((s) => <option key={s}>{s}</option>)}
                </Select>
              : <Input placeholder="optional" value={form.subcategory} onChange={(e) => set("subcategory", e.target.value)} />}
          </Field>
          <Field label={`Amount (${config.currency})`}><Input type="number" min="0" placeholder="0" value={form.amount} onChange={(e) => set("amount", e.target.value)} /></Field>
          <Field label={isPersonnel ? "Person / Worker" : "Person / Payee"} hint={isPersonnel ? "Used for per-person analytics" : "optional"}>
            <Input list="persons-dl" placeholder={isPersonnel ? "e.g. Aslam (Mason)" : "optional"} value={form.person} onChange={(e) => set("person", e.target.value)} />
            <datalist id="persons-dl">{persons.map((p) => <option key={p} value={p} />)}</datalist>
          </Field>
          <Field label="Vendor / supplier"><Input placeholder="optional" value={form.vendor} onChange={(e) => set("vendor", e.target.value)} /></Field>
          <Field label="Payment method">
            <Select value={form.mode} onChange={(e) => set("mode", e.target.value)}>{PAY_MODES.map((m) => <option key={m}>{m}</option>)}</Select>
          </Field>
          <div className="sm:col-span-2">
            <Field label="Description / note"><Input placeholder="optional details" value={form.description} onChange={(e) => set("description", e.target.value)} /></Field>
          </div>
        </div>
        <div className="flex gap-2 mt-3">
          <Button variant="accent" icon={editId ? Check : Plus} onClick={save} disabled={num(form.amount) <= 0}>{editId ? "Update expense" : "Add expense"}</Button>
          {editId && <Button variant="ghost" onClick={() => { setForm(blank); setEditId(null); }}>Cancel</Button>}
        </div>
      </Card>

      {/* Filters */}
      <Card className="p-3 mb-4">
        <div className="flex flex-wrap items-center gap-2">
          <div className="flex items-center gap-2 flex-1 min-w-[180px]" style={{ background: T.bg, borderRadius: 8, padding: "2px 10px", border: `1px solid ${T.line}` }}>
            <Search size={15} style={{ color: T.faint }} />
            <input placeholder="Search expenses…" value={q} onChange={(e) => setQ(e.target.value)} style={{ border: "none", background: "transparent", outline: "none", fontSize: 13, padding: "7px 0", width: "100%", color: T.ink }} />
          </div>
          <Select value={fCat} onChange={(e) => setFCat(e.target.value)} style={{ width: "auto", minWidth: 150 }}>
            <option value="">All categories</option>
            {config.categories.map((c) => <option key={c.name}>{c.name}</option>)}
          </Select>
          <Select value={fPerson} onChange={(e) => setFPerson(e.target.value)} style={{ width: "auto", minWidth: 140 }}>
            <option value="">All persons</option>
            {persons.map((p) => <option key={p}>{p}</option>)}
          </Select>
          {(fCat || fPerson || q) && <Button variant="ghost" size="sm" icon={X} onClick={() => { setFCat(""); setFPerson(""); setQ(""); }}>Clear</Button>}
          <div className="ml-auto text-sm font-semibold tabular-nums" style={{ color: T.sub }}>
            {filtered.length} item{filtered.length === 1 ? "" : "s"} · <span style={{ color: T.accent }}>{fmt(filteredTotal)}</span>
          </div>
        </div>
      </Card>

      <Card>
        <Table head={["Date", "Category", "Person / Vendor", "Method", "Amount", ""]} widths={["13%", "26%", "auto", "12%", "15%", "70px"]}>
          {filtered.length === 0
            ? <tr><td colSpan={6}><Empty icon={Receipt} title={expenses.length ? "No matches" : "No expenses recorded"} hint={expenses.length ? "Adjust the filters above." : "Add your first expense to start tracking spend."} /></td></tr>
            : filtered.map((e) => (
              <tr key={e.id} style={{ borderTop: `1px solid ${T.line}` }}>
                <Td>{fmtDate(e.date)}</Td>
                <Td><div className="font-medium" style={{ color: T.ink }}>{e.category}</div>{e.subcategory && <div className="text-xs" style={{ color: T.faint }}>{e.subcategory}</div>}</Td>
                <Td muted>{[e.person, e.vendor].filter(Boolean).join(" · ") || "—"}{e.description ? <div className="text-xs" style={{ color: T.faint }}>{e.description}</div> : null}</Td>
                <Td><Chip>{e.mode}</Chip></Td>
                <Td bold>{fmt(num(e.amount))}</Td>
                <Td>
                  {confirmId === e.id
                    ? <ConfirmDelete onConfirm={() => { removeExpense(e.id); setConfirmId(null); }} onCancel={() => setConfirmId(null)} />
                    : <RowActions onEdit={() => edit(e)} onDelete={() => setConfirmId(e.id)} />}
                </Td>
              </tr>
            ))}
        </Table>
      </Card>
    </div>
  );
}

/* ----------------------------------------------------------------------------
   Analytics
---------------------------------------------------------------------------- */
function Analytics({ config, expenses, payments, fmt }) {
  const byCat = useMemo(() => aggregate(expenses, "category"), [expenses]);
  const trend = useMemo(() => buildTrend(payments, expenses), [payments, expenses]);

  // personnel
  const [scope, setScope] = useState(PERSONNEL_CAT);
  const personExpenses = useMemo(
    () => expenses.filter((e) => e.person && (scope === "__all__" || e.category === scope)),
    [expenses, scope]
  );
  const allPersons = useMemo(() => [...new Set(personExpenses.map((e) => e.person))].sort(), [personExpenses]);
  const [selected, setSelected] = useState(null); // null => all
  const activePersons = selected || allPersons;
  const personAgg = useMemo(() => {
    const m = {};
    personExpenses.filter((e) => activePersons.includes(e.person)).forEach((e) => { m[e.person] = (m[e.person] || 0) + num(e.amount); });
    return Object.entries(m).map(([key, value]) => ({ key, value })).sort((a, b) => b.value - a.value);
  }, [personExpenses, activePersons]);
  const personTotal = personAgg.reduce((s, p) => s + p.value, 0);

  const togglePerson = (p) => {
    const base = selected || allPersons;
    const next = base.includes(p) ? base.filter((x) => x !== p) : [...base, p];
    setSelected(next.length === allPersons.length ? null : next);
  };

  return (
    <div>
      <PageHead title="Analytics" subtitle="Where the money is going" />

      <div className="grid lg:grid-cols-2 gap-4 mb-4">
        <Card className="p-4">
          <SectionTitle icon={BarChart3}>Expenditure by category</SectionTitle>
          {byCat.length === 0 ? <Empty icon={BarChart3} title="No data yet" hint="Add expenses to populate analytics." /> : (
            <div style={{ height: 300 }}>
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={byCat} layout="vertical" margin={{ left: 8, right: 16 }}>
                  <CartesianGrid horizontal={false} stroke={T.line} />
                  <XAxis type="number" tick={{ fontSize: 11, fill: T.faint }} tickFormatter={(v) => compact(v)} />
                  <YAxis type="category" dataKey="key" width={110} tick={{ fontSize: 11, fill: T.sub }} />
                  <Tooltip formatter={(v) => fmt(v)} contentStyle={tipStyle} cursor={{ fill: T.bg }} />
                  <Bar dataKey="value" radius={[0, 4, 4, 0]}>
                    {byCat.map((_, i) => <Cell key={i} fill={CHART[i % CHART.length]} />)}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          )}
        </Card>

        <Card className="p-4">
          <SectionTitle icon={TrendingUp}>Monthly cash flow</SectionTitle>
          {trend.length === 0 ? <Empty icon={TrendingUp} title="No data yet" hint="Record payments and expenses over time." /> : (
            <div style={{ height: 300 }}>
              <ResponsiveContainer width="100%" height="100%">
                <ComposedChart data={trend} margin={{ left: 4, right: 8 }}>
                  <CartesianGrid stroke={T.line} vertical={false} />
                  <XAxis dataKey="label" tick={{ fontSize: 11, fill: T.faint }} />
                  <YAxis tick={{ fontSize: 11, fill: T.faint }} tickFormatter={(v) => compact(v)} />
                  <Tooltip formatter={(v) => fmt(v)} contentStyle={tipStyle} />
                  <Legend wrapperStyle={{ fontSize: 11 }} />
                  <Bar dataKey="received" name="Received" fill={T.blue} radius={[3, 3, 0, 0]} barSize={14} />
                  <Bar dataKey="spent" name="Spent" fill={T.accent} radius={[3, 3, 0, 0]} barSize={14} />
                  <Line dataKey="balance" name="Balance" stroke={T.positive} strokeWidth={2} dot={false} />
                </ComposedChart>
              </ResponsiveContainer>
            </div>
          )}
        </Card>
      </div>

      {/* Personnel analytics */}
      <Card className="p-4">
        <div className="flex items-center justify-between flex-wrap gap-2 mb-3">
          <SectionTitle icon={HardHat}>Personnel expenditure</SectionTitle>
          <Select value={scope} onChange={(e) => { setScope(e.target.value); setSelected(null); }} style={{ width: "auto" }}>
            <option value={PERSONNEL_CAT}>Labour / Personnel only</option>
            <option value="__all__">All categories (any payee)</option>
          </Select>
        </div>

        {allPersons.length === 0 ? (
          <Empty icon={HardHat} title="No personnel data" hint="Add expenses with a person/worker name to analyse per-person spend." />
        ) : (
          <div className="grid lg:grid-cols-2 gap-5">
            <div>
              <div className="flex items-center justify-between mb-2">
                <span className="text-xs font-medium" style={{ color: T.sub }}>Filter persons</span>
                <button className="text-xs font-semibold" style={{ color: T.accent }} onClick={() => setSelected(null)}>Select all</button>
              </div>
              <div className="flex flex-wrap gap-1.5 mb-4">
                {allPersons.map((p) => {
                  const on = activePersons.includes(p);
                  return (
                    <button key={p} onClick={() => togglePerson(p)}
                      className="rounded-full text-xs font-medium"
                      style={{ padding: "5px 11px", border: `1px solid ${on ? T.accent : T.line}`, background: on ? T.accentSoft : "#fff", color: on ? T.accent : T.sub }}>
                      {p}
                    </button>
                  );
                })}
              </div>
              <div className="rounded-lg p-3" style={{ background: T.bg }}>
                <div className="text-xs" style={{ color: T.sub }}>Selected personnel spend</div>
                <div className="text-2xl font-bold tabular-nums" style={{ color: T.ink }}>{fmt(personTotal)}</div>
                <div className="text-xs mt-0.5" style={{ color: T.faint }}>{activePersons.length} of {allPersons.length} persons</div>
              </div>
              <Table head={["Person", "Spend", "Share"]} widths={["auto", "30%", "22%"]} compact>
                {personAgg.map((p) => (
                  <tr key={p.key} style={{ borderTop: `1px solid ${T.line}` }}>
                    <Td>{p.key}</Td>
                    <Td bold>{fmt(p.value)}</Td>
                    <Td muted>{personTotal > 0 ? Math.round((p.value / personTotal) * 100) + "%" : "—"}</Td>
                  </tr>
                ))}
              </Table>
            </div>
            <div style={{ height: 320 }}>
              {personAgg.length === 0 ? <Empty icon={HardHat} title="No persons selected" hint="Pick at least one person." /> : (
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={personAgg} layout="vertical" margin={{ left: 8, right: 16 }}>
                    <CartesianGrid horizontal={false} stroke={T.line} />
                    <XAxis type="number" tick={{ fontSize: 11, fill: T.faint }} tickFormatter={(v) => compact(v)} />
                    <YAxis type="category" dataKey="key" width={100} tick={{ fontSize: 11, fill: T.sub }} />
                    <Tooltip formatter={(v) => fmt(v)} contentStyle={tipStyle} cursor={{ fill: T.bg }} />
                    <Bar dataKey="value" radius={[0, 4, 4, 0]} fill={T.accent} />
                  </BarChart>
                </ResponsiveContainer>
              )}
            </div>
          </div>
        )}
      </Card>
    </div>
  );
}

/* ----------------------------------------------------------------------------
   Reports
---------------------------------------------------------------------------- */
function Reports({ config, payments, expenses, reports, totals, fmt, addReport, removeReport }) {
  const [viewing, setViewing] = useState(null);
  const [confirmId, setConfirmId] = useState(null);

  const snapshot = () => {
    const byCategory = aggregate(expenses, "category");
    const byPerson = aggregatePersons(expenses);
    return {
      id: uid(), createdAt: Date.now(),
      title: `Project Report — ${new Date().toLocaleString("en-GB", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" })}`,
      project: config.projectName, owner: config.ownerName, contractor: config.contractorName, currency: config.currency,
      totals: { ...totals },
      counts: { payments: payments.length, expenses: expenses.length },
      byCategory, byPerson,
    };
  };
  const generate = () => { const r = snapshot(); addReport(r); setViewing(r); };

  const exportLedgerCSV = () => {
    const rows = [["Type", "Date", "Category", "Sub-category", "Person", "Vendor", "Method", "Description / Note", `Amount (${config.currency})`]];
    payments.slice().sort((a, b) => (a.date || "").localeCompare(b.date || "")).forEach((p) =>
      rows.push(["Payment", p.date, "", "", "", "", p.mode || "", p.note || "", num(p.amount)]));
    expenses.slice().sort((a, b) => (a.date || "").localeCompare(b.date || "")).forEach((e) =>
      rows.push(["Expense", e.date, e.category, e.subcategory || "", e.person || "", e.vendor || "", e.mode || "", e.description || "", num(e.amount)]));
    rows.push([]);
    rows.push(["", "", "", "", "", "", "", "Total received", totals.received]);
    rows.push(["", "", "", "", "", "", "", "Total spent", totals.spent]);
    rows.push(["", "", "", "", "", "", "", totals.balance < 0 ? "Owed to contractor" : "Balance with contractor", Math.abs(totals.balance)]);
    downloadCSV(`${slug(config.projectName)}_ledger_${today()}.csv`, rows);
  };

  const reportCSV = (r) => {
    const rows = [
      ["Project Report"], ["Project", r.project], ["Owner", r.owner], ["Contractor", r.contractor],
      ["Generated", new Date(r.createdAt).toLocaleString("en-GB")], [],
      ["Summary", `Amount (${r.currency})`],
      ["Total received", r.totals.received], ["Total spent", r.totals.spent],
      [r.totals.balance < 0 ? "Owed to contractor" : "Balance with contractor", Math.abs(r.totals.balance)], [],
      ["Category", `Amount (${r.currency})`], ...r.byCategory.map((c) => [c.key, c.value]), [],
      ["Person", `Amount (${r.currency})`], ...r.byPerson.map((p) => [p.key, p.value]),
    ];
    downloadCSV(`${slug(r.project)}_report_${new Date(r.createdAt).toISOString().slice(0, 10)}.csv`, rows);
  };

  const meta = { projectName: config.projectName, ownerName: config.ownerName, contractorName: config.contractorName, currency: config.currency };
  const wordLive = () => generateWordReport({
    meta, totals,
    categories: aggregate(expenses, "category"),
    persons: aggregatePersons(expenses),
    monthly: buildTrend(payments, expenses),
    payments, expenses, generatedAt: Date.now(), includeLedger: true,
  });
  const wordSnapshot = (r) => generateWordReport({
    meta: { projectName: r.project, ownerName: r.owner, contractorName: r.contractor, currency: r.currency },
    totals: r.totals, categories: r.byCategory, persons: r.byPerson,
    monthly: [], payments: [], expenses: [], generatedAt: r.createdAt, includeLedger: false,
  });

  return (
    <div>
      <PageHead title="Reports"
        subtitle="Generate snapshots from your data and keep a record of every report"
        actions={<div className="flex gap-2 flex-wrap">
          <Button variant="soft" icon={Download} onClick={exportLedgerCSV}>Export CSV</Button>
          <Button variant="soft" icon={FileDown} onClick={wordLive}>Word report</Button>
          <Button variant="accent" icon={FileText} onClick={generate}>Generate report</Button>
        </div>}
      />

      {viewing && <ReportView r={viewing} fmt={makeFmt(viewing.currency)} onClose={() => setViewing(null)} onDownloadCSV={() => reportCSV(viewing)} onDownloadWord={() => wordSnapshot(viewing)} />}

      <Card>
        <div className="p-4" style={{ borderBottom: `1px solid ${T.line}` }}>
          <SectionTitle icon={FileText}>Generated reports</SectionTitle>
          <p className="text-xs -mt-1" style={{ color: T.faint }}>
            Each report is a saved snapshot of totals, categories and personnel at the moment it was generated. Open to review, or download as a Word document or CSV.
          </p>
        </div>
        {reports.length === 0
          ? <Empty icon={FileText} title="No reports yet" hint="Generate your first report from the current data." />
          : (
            <div>
              {reports.map((r) => {
                const owed = r.totals.balance < 0;
                return (
                  <div key={r.id} className="flex items-center justify-between gap-3 p-4 flex-wrap" style={{ borderTop: `1px solid ${T.line}` }}>
                    <div className="flex items-center gap-3 min-w-0">
                      <div className="rounded-lg flex items-center justify-center flex-shrink-0" style={{ width: 36, height: 36, background: T.accentSoft }}>
                        <FileText size={17} style={{ color: T.accent }} />
                      </div>
                      <div className="min-w-0">
                        <div className="text-sm font-semibold truncate" style={{ color: T.ink }}>{r.title}</div>
                        <div className="text-xs" style={{ color: T.faint }}>
                          {r.counts.payments} payments · {r.counts.expenses} expenses · {r.byCategory.length} categories
                        </div>
                      </div>
                    </div>
                    <div className="flex items-center gap-2 flex-wrap">
                      <span className="text-xs font-semibold px-2 py-1 rounded-md tabular-nums" style={{ background: owed ? T.negativeSoft : T.positiveSoft, color: owed ? T.negative : T.positive }}>
                        {owed ? "Owed " : "Balance "}{makeFmt(r.currency)(Math.abs(r.totals.balance))}
                      </span>
                      <Button variant="soft" size="sm" onClick={() => setViewing(r)}>Open</Button>
                      <Button variant="soft" size="sm" icon={FileDown} onClick={() => wordSnapshot(r)}>Word</Button>
                      <Button variant="soft" size="sm" icon={Download} onClick={() => reportCSV(r)}>CSV</Button>
                      {confirmId === r.id
                        ? <ConfirmDelete onConfirm={() => { removeReport(r.id); setConfirmId(null); }} onCancel={() => setConfirmId(null)} />
                        : <button onClick={() => setConfirmId(r.id)} className="p-1.5" style={{ color: T.faint }} title="Delete"><Trash2 size={15} /></button>}
                    </div>
                  </div>
                );
              })}
            </div>
          )}
      </Card>
    </div>
  );
}

function ReportView({ r, fmt, onClose, onDownloadCSV, onDownloadWord }) {
  const owed = r.totals.balance < 0;
  return (
    <Card className="p-5 mb-4" style={{ borderColor: `${T.accent}44` }}>
      <div className="flex items-start justify-between gap-3 mb-4">
        <div>
          <div className="text-xs font-semibold uppercase tracking-wide" style={{ color: T.accent, letterSpacing: ".06em" }}>Report snapshot</div>
          <h3 className="text-lg font-bold" style={{ color: T.ink }}>{r.project}</h3>
          <p className="text-xs" style={{ color: T.faint }}>{r.owner} → {r.contractor} · Generated {new Date(r.createdAt).toLocaleString("en-GB")}</p>
        </div>
        <div className="flex gap-2 flex-wrap justify-end">
          <Button variant="accent" size="sm" icon={FileDown} onClick={onDownloadWord}>Word</Button>
          <Button variant="soft" size="sm" icon={Download} onClick={onDownloadCSV}>CSV</Button>
          <button onClick={onClose} className="p-1.5" style={{ color: T.faint }}><X size={18} /></button>
        </div>
      </div>
      <div className="grid grid-cols-3 gap-3 mb-4">
        <MiniStat label="Received" value={fmt(r.totals.received)} tint={T.blue} />
        <MiniStat label="Spent" value={fmt(r.totals.spent)} tint={T.accent} />
        <MiniStat label={owed ? "Owed to contractor" : "Balance with contractor"} value={fmt(Math.abs(r.totals.balance))} tint={owed ? T.negative : T.positive} />
      </div>
      <div className="grid sm:grid-cols-2 gap-5">
        <div>
          <SectionTitle icon={BarChart3}>By category</SectionTitle>
          <SimpleBars rows={r.byCategory} fmt={fmt} />
        </div>
        <div>
          <SectionTitle icon={HardHat}>By person</SectionTitle>
          {r.byPerson.length === 0 ? <p className="text-xs" style={{ color: T.faint }}>No personnel-tagged expenses.</p> : <SimpleBars rows={r.byPerson} fmt={fmt} color={T.accent} />}
        </div>
      </div>
    </Card>
  );
}
function MiniStat({ label, value, tint }) {
  return (
    <div className="rounded-lg p-3" style={{ background: T.bg }}>
      <div className="text-xs" style={{ color: T.sub }}>{label}</div>
      <div className="text-base font-bold tabular-nums" style={{ color: tint }}>{value}</div>
    </div>
  );
}
function SimpleBars({ rows, fmt, color }) {
  const max = Math.max(...rows.map((r) => r.value), 1);
  return (
    <div className="flex flex-col gap-2">
      {rows.map((r, i) => (
        <div key={r.key}>
          <div className="flex items-center justify-between text-xs mb-0.5">
            <span style={{ color: T.sub }}>{r.key}</span>
            <span className="font-semibold tabular-nums" style={{ color: T.ink }}>{fmt(r.value)}</span>
          </div>
          <div style={{ height: 6, borderRadius: 4, background: T.line }}>
            <div style={{ height: 6, borderRadius: 4, width: `${(r.value / max) * 100}%`, background: color || CHART[i % CHART.length] }} />
          </div>
        </div>
      ))}
    </div>
  );
}

/* ----------------------------------------------------------------------------
   Settings
---------------------------------------------------------------------------- */
function SettingsView({ config, setConfig, data, replaceAll, setTab }) {
  const [newCat, setNewCat] = useState("");
  const [newPerson, setNewPerson] = useState("");
  const [confirmReset, setConfirmReset] = useState(false);
  const [importMsg, setImportMsg] = useState(null);
  const fileRef = useRef(null);

  const addCat = () => {
    const n = newCat.trim();
    if (!n || config.categories.some((c) => c.name.toLowerCase() === n.toLowerCase())) return;
    setConfig({ categories: [...config.categories, { name: n, subs: [] }] }); setNewCat("");
  };
  const removeCat = (name) => setConfig({ categories: config.categories.filter((c) => c.name !== name) });
  const addPerson = () => {
    const n = newPerson.trim();
    if (!n || (config.personnel || []).includes(n)) return;
    setConfig({ personnel: [...(config.personnel || []), n] }); setNewPerson("");
  };
  const removePerson = (n) => setConfig({ personnel: (config.personnel || []).filter((p) => p !== n) });

  const backup = () => {
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href = url; a.download = `${slug(config.projectName)}_backup_${today()}.json`; a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  };
  const onImportFile = (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(reader.result);
        const ok = replaceAll(parsed);
        setImportMsg(ok
          ? { type: "ok", text: `Loaded ${(parsed.payments || []).length} payments and ${(parsed.expenses || []).length} expenses. Saved to this browser.` }
          : { type: "err", text: "That file isn't a valid project backup." });
      } catch {
        setImportMsg({ type: "err", text: "Could not read that file — make sure it's a JSON backup exported from this app." });
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  };
  const reset = async () => { await storageSet(STORE_KEY, JSON.stringify(EMPTY)); window.location.reload(); };

  return (
    <div>
      <PageHead title="Settings" subtitle="Project details, categories and personnel" />

      <div className="grid lg:grid-cols-2 gap-4">
        <Card className="p-4">
          <SectionTitle icon={Building2}>Project details</SectionTitle>
          <div className="grid grid-cols-2 gap-3">
            <div className="col-span-2"><Field label="Project name"><Input value={config.projectName} onChange={(e) => setConfig({ projectName: e.target.value })} /></Field></div>
            <Field label="Owner name"><Input value={config.ownerName} onChange={(e) => setConfig({ ownerName: e.target.value })} /></Field>
            <Field label="Contractor name"><Input value={config.contractorName} onChange={(e) => setConfig({ contractorName: e.target.value })} /></Field>
            <Field label="Currency symbol" hint="e.g. Rs, €, $, £"><Input value={config.currency} onChange={(e) => setConfig({ currency: e.target.value })} /></Field>
          </div>
        </Card>

        <Card className="p-4">
          <SectionTitle icon={HardHat}>Personnel</SectionTitle>
          <p className="text-xs mb-3" style={{ color: T.faint }}>Saved names appear as quick suggestions when adding expenses.</p>
          <div className="flex gap-2 mb-3">
            <Input placeholder="Add a worker / person" value={newPerson} onChange={(e) => setNewPerson(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addPerson()} />
            <Button variant="accent" icon={Plus} onClick={addPerson}>Add</Button>
          </div>
          <div className="flex flex-wrap gap-1.5">
            {(config.personnel || []).length === 0 && <span className="text-xs" style={{ color: T.faint }}>No saved personnel yet.</span>}
            {(config.personnel || []).map((p) => (
              <span key={p} className="inline-flex items-center gap-1 rounded-full text-xs font-medium" style={{ padding: "5px 8px 5px 11px", background: T.bg, border: `1px solid ${T.line}`, color: T.sub }}>
                {p}<button onClick={() => removePerson(p)} style={{ color: T.faint }}><X size={13} /></button>
              </span>
            ))}
          </div>
        </Card>

        <Card className="p-4 lg:col-span-2">
          <SectionTitle icon={Filter}>Expense categories</SectionTitle>
          <div className="flex gap-2 mb-3 max-w-md">
            <Input placeholder="Add a custom category" value={newCat} onChange={(e) => setNewCat(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addCat()} />
            <Button variant="accent" icon={Plus} onClick={addCat}>Add</Button>
          </div>
          <div className="flex flex-wrap gap-1.5">
            {config.categories.map((c) => (
              <span key={c.name} className="inline-flex items-center gap-1.5 rounded-full text-xs font-medium" style={{ padding: "5px 8px 5px 11px", background: c.name === PERSONNEL_CAT ? T.accentSoft : "#fff", border: `1px solid ${c.name === PERSONNEL_CAT ? T.accent + "44" : T.line}`, color: c.name === PERSONNEL_CAT ? T.accent : T.sub }}>
                {c.name}
                {c.name !== PERSONNEL_CAT && <button onClick={() => removeCat(c.name)} style={{ color: T.faint }}><X size={13} /></button>}
              </span>
            ))}
          </div>
        </Card>

        <Card className="p-4 lg:col-span-2">
          <SectionTitle icon={Download}>Data &amp; handoff</SectionTitle>
          <p className="text-xs mb-3" style={{ color: T.faint }}>
            Download a backup to share this project, or import a backup file to load another person's data and continue from it.
            Importing replaces everything currently saved in this browser.
          </p>
          <input ref={fileRef} type="file" accept="application/json,.json" onChange={onImportFile} style={{ display: "none" }} />
          <div className="flex items-center gap-3 flex-wrap">
            <Button variant="soft" icon={Download} onClick={backup}>Download backup (JSON)</Button>
            <Button variant="accent" icon={Upload} onClick={() => fileRef.current && fileRef.current.click()}>Import backup (JSON)</Button>
            {confirmReset
              ? <span className="inline-flex items-center gap-2 text-sm" style={{ color: T.negative }}>
                  <AlertTriangle size={15} /> This deletes everything.
                  <Button variant="danger" size="sm" onClick={reset}>Yes, reset</Button>
                  <Button variant="ghost" size="sm" onClick={() => setConfirmReset(false)}>Cancel</Button>
                </span>
              : <Button variant="danger" icon={Trash2} onClick={() => setConfirmReset(true)}>Reset all data</Button>}
          </div>
          {importMsg && (
            <div className="mt-3 text-sm rounded-lg" style={{ padding: "9px 12px", background: importMsg.type === "ok" ? T.positiveSoft : T.negativeSoft, color: importMsg.type === "ok" ? T.positive : T.negative }}>
              {importMsg.text}
            </div>
          )}
        </Card>
      </div>
    </div>
  );
}

/* ----------------------------------------------------------------------------
   Small table primitives
---------------------------------------------------------------------------- */
function Table({ head, widths = [], children, compact }) {
  return (
    <div style={{ overflowX: "auto" }}>
      <table style={{ width: "100%", borderCollapse: "collapse", minWidth: compact ? 0 : 520 }}>
        <thead>
          <tr>
            {head.map((h, i) => (
              <th key={i} style={{ textAlign: i === head.length - 1 && h === "" ? "right" : (i >= head.length - 2 && /Amount|Spend|Share/.test(h) ? "right" : "left"), width: widths[i], padding: compact ? "8px 10px" : "11px 14px", fontSize: 11, fontWeight: 600, textTransform: "uppercase", letterSpacing: ".05em", color: T.faint, background: T.bg }}>
                {h}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>{children}</tbody>
      </table>
    </div>
  );
}
function Td({ children, bold, muted, style }) {
  return (
    <td style={{ padding: "11px 14px", fontSize: 13, verticalAlign: "top", textAlign: bold ? "right" : "left", fontWeight: bold ? 700 : 400, fontVariantNumeric: bold ? "tabular-nums" : "normal", color: muted ? T.sub : T.ink, ...style }}>
      {children}
    </td>
  );
}
function Chip({ children }) {
  return <span className="text-xs font-medium rounded-md" style={{ padding: "2px 8px", background: T.bg, border: `1px solid ${T.line}`, color: T.sub }}>{children}</span>;
}
function RowActions({ onEdit, onDelete }) {
  return (
    <div className="flex items-center justify-end gap-1">
      <button onClick={onEdit} title="Edit" className="p-1.5" style={{ color: T.sub }}><Pencil size={14} /></button>
      <button onClick={onDelete} title="Delete" className="p-1.5" style={{ color: T.faint }}><Trash2 size={14} /></button>
    </div>
  );
}

/* ----------------------------------------------------------------------------
   Word (.docx) report
---------------------------------------------------------------------------- */
const W = { ink: "1F2933", sub: "5B6B7A", faint: "8A97A3", accent: "B4690E", blue: "2A5D8F", line: "E4E8EC", zebra: "F7F8FA", posSoft: "E7F4EE", negSoft: "FBECEB", pos: "0F7B5A", neg: "B3261E" };
const wBorder = (() => { const b = { style: BorderStyle.SINGLE, size: 4, color: W.line }; return { top: b, bottom: b, left: b, right: b, insideHorizontal: b, insideVertical: b }; })();

function wCell(text, { bold, align, fill, color, size } = {}) {
  return new TableCell({
    shading: fill ? { type: ShadingType.CLEAR, fill, color: "auto" } : undefined,
    margins: { top: 60, bottom: 60, left: 110, right: 110 },
    children: [new Paragraph({
      alignment: align === "right" ? AlignmentType.RIGHT : align === "center" ? AlignmentType.CENTER : AlignmentType.LEFT,
      children: [new TextRun({ text: String(text), bold: !!bold, color: color || W.ink, size: size || 18 })],
    })],
  });
}
function wTable(headers, rows) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h) => wCell(h.text, { bold: true, align: h.align, fill: W.ink, color: "FFFFFF" })),
  });
  const bodyRows = rows.map((cells, i) =>
    new TableRow({ children: cells.map((c) => wCell(c.text, { ...c, fill: c.fill || (i % 2 ? W.zebra : undefined) })) }));
  return new DocTable({
    width: { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: headers.map((h) => Math.round((h.w || (100 / headers.length)) * 96)),
    borders: wBorder,
    rows: [headerRow, ...bodyRows],
  });
}
function wHeading(text) {
  return new Paragraph({ spacing: { before: 300, after: 130 }, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: W.accent } }, children: [new TextRun({ text, bold: true, size: 24, color: W.accent })] });
}
function wPara(runs, opts = {}) { return new Paragraph({ spacing: opts.spacing, alignment: opts.align, children: runs }); }
function wSpacer(after = 120) { return new Paragraph({ spacing: { after }, children: [] }); }

function generateWordReport(payload) {
  const { meta, totals, categories, persons, monthly, payments, expenses, generatedAt, includeLedger } = payload;
  const fmt = makeFmt(meta.currency);
  const owed = totals.balance < 0;
  const spent = totals.spent || 0;
  const dateStr = new Date(generatedAt || Date.now()).toLocaleString("en-GB", { day: "2-digit", month: "long", year: "numeric", hour: "2-digit", minute: "2-digit" });

  const children = [];

  // Title block
  children.push(new Paragraph({ children: [new TextRun({ text: meta.projectName || "Project", bold: true, size: 44, color: W.ink })] }));
  children.push(new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Construction Project — Finance & Analytics Report", size: 24, color: W.sub })] }));
  children.push(new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: `${meta.ownerName} (Owner)  \u2192  ${meta.contractorName} (Contractor)`, size: 19, color: W.sub })] }));
  children.push(new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: `Generated ${dateStr}`, size: 17, color: W.faint, italics: true })] }));

  // Financial summary
  children.push(wHeading("Financial Summary"));
  children.push(wTable(
    [{ text: "Metric", w: 60 }, { text: `Amount (${meta.currency})`, align: "right", w: 40 }],
    [
      [{ text: "Total received from owner" }, { text: fmt(totals.received), align: "right", bold: true }],
      [{ text: "Total spent by contractor" }, { text: fmt(totals.spent), align: "right", bold: true }],
      [{ text: owed ? `Amount owed to ${meta.contractorName}` : `Balance held by ${meta.contractorName}`, bold: true, fill: owed ? W.negSoft : W.posSoft, color: owed ? W.neg : W.pos },
       { text: fmt(Math.abs(totals.balance)), align: "right", bold: true, fill: owed ? W.negSoft : W.posSoft, color: owed ? W.neg : W.pos }],
    ]
  ));
  children.push(wPara([new TextRun({ text: owed
    ? `The contractor has spent more on the project than received and is owed ${fmt(Math.abs(totals.balance))}.`
    : `Unspent funds of ${fmt(totals.balance)} remain with the contractor. Funds utilised: ${totals.received > 0 ? Math.round((spent / totals.received) * 100) : 0}% of payments received.`,
    size: 17, color: W.sub, italics: true })], { spacing: { before: 100, after: 80 } }));

  // Category analytics
  children.push(wHeading("Expenditure by Category"));
  if (!categories.length) {
    children.push(wPara([new TextRun({ text: "No expenses recorded.", size: 18, color: W.faint })]));
  } else {
    const rows = categories.map((c) => [
      { text: c.key }, { text: fmt(c.value), align: "right" },
      { text: spent > 0 ? Math.round((c.value / spent) * 100) + "%" : "\u2014", align: "right" },
    ]);
    rows.push([{ text: "Total", bold: true, fill: W.zebra }, { text: fmt(spent), align: "right", bold: true, fill: W.zebra }, { text: "100%", align: "right", bold: true, fill: W.zebra }]);
    children.push(wTable([{ text: "Category", w: 56 }, { text: `Amount (${meta.currency})`, align: "right", w: 30 }, { text: "Share", align: "right", w: 14 }], rows));
  }

  // Personnel analytics
  children.push(wHeading("Personnel Expenditure"));
  if (!persons.length) {
    children.push(wPara([new TextRun({ text: "No personnel-tagged expenses recorded.", size: 18, color: W.faint })]));
  } else {
    const ptotal = persons.reduce((s, p) => s + p.value, 0);
    const rows = persons.map((p) => [
      { text: p.key }, { text: fmt(p.value), align: "right" },
      { text: ptotal > 0 ? Math.round((p.value / ptotal) * 100) + "%" : "\u2014", align: "right" },
    ]);
    rows.push([{ text: "Total", bold: true, fill: W.zebra }, { text: fmt(ptotal), align: "right", bold: true, fill: W.zebra }, { text: "100%", align: "right", bold: true, fill: W.zebra }]);
    children.push(wTable([{ text: "Person / Worker", w: 56 }, { text: `Amount (${meta.currency})`, align: "right", w: 30 }, { text: "Share", align: "right", w: 14 }], rows));
  }

  // Monthly cash flow
  if (monthly && monthly.length) {
    children.push(wHeading("Monthly Cash Flow"));
    const rows = monthly.map((m) => [
      { text: m.label }, { text: fmt(m.received), align: "right" }, { text: fmt(m.spent), align: "right" }, { text: fmt(m.balance), align: "right" },
    ]);
    children.push(wTable(
      [{ text: "Month", w: 28 }, { text: "Received", align: "right", w: 24 }, { text: "Spent", align: "right", w: 24 }, { text: "Running balance", align: "right", w: 24 }],
      rows
    ));
  }

  // Ledger appendix
  if (includeLedger) {
    if (payments && payments.length) {
      children.push(wHeading("Appendix \u2014 Owner Payments"));
      const rows = [...payments].sort((a, b) => (a.date || "").localeCompare(b.date || "")).map((p) => [
        { text: fmtDate(p.date) }, { text: p.mode || "\u2014" }, { text: p.note || "\u2014" }, { text: fmt(num(p.amount)), align: "right" },
      ]);
      children.push(wTable([{ text: "Date", w: 20 }, { text: "Method", w: 22 }, { text: "Reference", w: 36 }, { text: `Amount (${meta.currency})`, align: "right", w: 22 }], rows));
    }
    if (expenses && expenses.length) {
      children.push(wHeading("Appendix \u2014 Expenses"));
      const rows = [...expenses].sort((a, b) => (a.date || "").localeCompare(b.date || "")).map((e) => [
        { text: fmtDate(e.date) },
        { text: e.category + (e.subcategory ? ` \u00b7 ${e.subcategory}` : "") },
        { text: [e.person, e.vendor].filter(Boolean).join(" \u00b7 ") || "\u2014" },
        { text: fmt(num(e.amount)), align: "right" },
      ]);
      children.push(wTable([{ text: "Date", w: 16 }, { text: "Category", w: 34 }, { text: "Person / Vendor", w: 28 }, { text: `Amount (${meta.currency})`, align: "right", w: 22 }], rows));
    }
  }

  children.push(wSpacer(200));
  children.push(wPara([new TextRun({ text: "Generated by the Construction Project Finance Ledger.", size: 15, color: W.faint, italics: true })], { align: AlignmentType.CENTER }));

  const doc = new Document({
    styles: { default: { document: { run: { font: "Calibri" } } } },
    sections: [{ properties: { page: { margin: { top: 720, bottom: 720, left: 720, right: 720 } } }, children }],
  });

  Packer.toBlob(doc).then((blob) => {
    downloadBlob(`${slug(meta.projectName)}_report_${new Date(generatedAt || Date.now()).toISOString().slice(0, 10)}.docx`, blob);
  });
}

function downloadBlob(filename, blob) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

/* ----------------------------------------------------------------------------
   Data utils
---------------------------------------------------------------------------- */
function normalize(d) {
  if (!d || typeof d !== "object") return null;
  return {
    config: { ...DEFAULT_CONFIG, ...(d.config || {}) },
    payments: Array.isArray(d.payments) ? d.payments : [],
    expenses: Array.isArray(d.expenses) ? d.expenses : [],
    reports: Array.isArray(d.reports) ? d.reports : [],
  };
}
function hasData(d) {
  return !!(d && ((Array.isArray(d.payments) && d.payments.length) || (Array.isArray(d.expenses) && d.expenses.length)));
}
function aggregate(items, key) {
  const m = {};
  items.forEach((it) => { const k = it[key] || "Uncategorised"; m[k] = (m[k] || 0) + num(it.amount); });
  return Object.entries(m).map(([k, v]) => ({ key: k, value: v })).sort((a, b) => b.value - a.value);
}
function aggregatePersons(items) {
  const m = {};
  items.forEach((it) => { if (it.person) m[it.person] = (m[it.person] || 0) + num(it.amount); });
  return Object.entries(m).map(([k, v]) => ({ key: k, value: v })).sort((a, b) => b.value - a.value);
}
function buildTrend(payments, expenses) {
  const keys = new Set();
  payments.forEach((p) => p.date && keys.add(monthKey(p.date)));
  expenses.forEach((e) => e.date && keys.add(monthKey(e.date)));
  const sorted = [...keys].sort();
  let cum = 0;
  return sorted.map((k) => {
    const received = payments.filter((p) => monthKey(p.date) === k).reduce((s, p) => s + num(p.amount), 0);
    const spent = expenses.filter((e) => monthKey(e.date) === k).reduce((s, e) => s + num(e.amount), 0);
    cum += received - spent;
    return { key: k, label: monthLabel(k), received, spent, balance: cum };
  });
}
function compact(v) {
  const n = Math.abs(v);
  if (n >= 1e7) return (v / 1e7).toFixed(1) + "Cr";
  if (n >= 1e5) return (v / 1e5).toFixed(1) + "L";
  if (n >= 1e3) return (v / 1e3).toFixed(0) + "k";
  return String(v);
}
const slug = (s) => (s || "project").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 40) || "project";
