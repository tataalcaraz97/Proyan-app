import React, { useState, useEffect, useMemo, useCallback } from "react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend,
} from "recharts";
import {
  Play, Pause, Square, Plus, ClipboardList, BarChart3, Users, Download,
  Wrench, ChevronDown, ChevronUp, X, Loader2, Pencil, Trash2, Check,
} from "lucide-react";
import { doc, onSnapshot, updateDoc, setDoc } from "firebase/firestore";
import { db } from "./firebase";
import ExcelJS from "exceljs";

const WORKSHOP_DOC = doc(db, "proyan", "workshop");
const DEFAULT_MECHANICS = [
  { id: "m1", name: "Carlos Medina", code: "1001" },
  { id: "m2", name: "Julián Rivas", code: "1002" },
  { id: "m3", name: "Emanuel Soto", code: "1003" },
];

function genMechCode() {
  return String(Math.floor(1000 + Math.random() * 9000));
}

function getOrderMechanicIds(order) {
  if (Array.isArray(order.mechanicIds)) return order.mechanicIds;
  if (order.mechanicId) return [order.mechanicId];
  return [];
}

/* ---------------------------------------------------------
   TOKENS
--------------------------------------------------------- */
const COLORS = {
  bg: "#121417",
  surface: "#1B1E22",
  surfaceAlt: "#22262B",
  line: "#2E3339",
  text: "#EDEFF2",
  textDim: "#9BA3AD",
  accent: "#F2A93B", // safety amber — primary action
  accentDim: "rgba(242,169,59,0.14)",
  green: "#3FB27F",
  greenDim: "rgba(63,178,127,0.14)",
  amber: "#E8B339",
  amberDim: "rgba(232,179,57,0.14)",
  gray: "#7A828C",
  grayDim: "rgba(122,130,140,0.14)",
  teal: "#2FB6C4",
  tealDim: "rgba(47,182,196,0.14)",
  red: "#E15B4F",
  redDim: "rgba(225,91,79,0.14)",
};

const ORDER_TYPES = {
  OTN: { label: "Orden de Trabajo Normal", color: "#4C7EF3", dim: "rgba(76,126,243,0.14)" },
  OTI: { label: "Orden de Trabajo Interna", color: "#B26EF2", dim: "rgba(178,110,242,0.14)" },
  OTG: { label: "Orden de Trabajo en Garantía", color: COLORS.accent, dim: COLORS.accentDim },
};

const BRANCHES = ["Mackenna", "Río Cuarto"];

const STATUS_META = {
  pendiente: { label: "Pendiente", color: COLORS.gray, dim: COLORS.grayDim },
  en_progreso: { label: "En progreso", color: COLORS.green, dim: COLORS.greenDim },
  pausada: { label: "Pausada", color: COLORS.amber, dim: COLORS.amberDim },
  finalizada: { label: "Finalizada", color: COLORS.teal, dim: COLORS.tealDim },
};

const FONTS_IMPORT =
  "@import url('https://fonts.googleapis.com/css2?family=Oswald:wght@500;600;700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@500;700&display=swap');";

/* ---------------------------------------------------------
   HELPERS
--------------------------------------------------------- */
function uid() {
  return Math.random().toString(36).slice(2, 9);
}

function pad2(n) {
  return String(Math.max(0, Math.floor(n))).padStart(2, "0");
}

function formatClock(totalSeconds) {
  const h = Math.floor(totalSeconds / 3600);
  const m = Math.floor((totalSeconds % 3600) / 60);
  const s = Math.floor(totalSeconds % 60);
  return `${pad2(h)}:${pad2(m)}:${pad2(s)}`;
}

function formatHours(totalSeconds) {
  return (totalSeconds / 3600).toFixed(2);
}

function getOrderSeconds(order, now) {
  return (order.segments || []).reduce((acc, seg) => {
    const end = seg.end || now;
    return acc + Math.max(0, (end - seg.start) / 1000);
  }, 0);
}

function monthKey(ts) {
  const d = new Date(ts);
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}`;
}

function currentMonthKey() {
  return monthKey(Date.now());
}

function todayLabel() {
  return new Date().toLocaleDateString("es-AR", {
    weekday: "long",
    day: "numeric",
    month: "long",
  });
}

/* ---------------------------------------------------------
   SMALL UI PRIMITIVES
--------------------------------------------------------- */
function Badge({ text, color, dim }) {
  return (
    <span
      className="px-2 py-0.5 rounded text-xs font-semibold tracking-wide"
      style={{ color, backgroundColor: dim, fontFamily: "'JetBrains Mono', monospace" }}
    >
      {text}
    </span>
  );
}

function Chip({ active, onClick, children }) {
  return (
    <button
      onClick={onClick}
      className="px-3 py-1.5 rounded-full text-sm font-medium transition-colors whitespace-nowrap"
      style={{
        backgroundColor: active ? COLORS.accent : COLORS.surfaceAlt,
        color: active ? "#1A1300" : COLORS.textDim,
        border: `1px solid ${active ? COLORS.accent : COLORS.line}`,
      }}
    >
      {children}
    </button>
  );
}

function ActionButton({ icon: Icon, label, onClick, tone = "accent", full }) {
  const toneMap = {
    accent: { bg: COLORS.accent, fg: "#1A1300" },
    green: { bg: COLORS.green, fg: "#052015" },
    amber: { bg: COLORS.amber, fg: "#241900" },
    red: { bg: COLORS.red, fg: "#210000" },
    ghost: { bg: COLORS.surfaceAlt, fg: COLORS.text },
  };
  const t = toneMap[tone];
  return (
    <button
      onClick={onClick}
      className={`flex items-center justify-center gap-2 rounded-lg py-2.5 px-4 font-semibold text-sm active:scale-[0.98] transition-transform ${
        full ? "flex-1" : ""
      }`}
      style={{ backgroundColor: t.bg, color: t.fg }}
    >
      <Icon size={16} strokeWidth={2.5} />
      {label}
    </button>
  );
}

/* ---------------------------------------------------------
   TICKET CARD (order)
--------------------------------------------------------- */
function TicketCard({ order, mechanicNames, now, expanded, onToggle, actions, showMechanic }) {
  const seconds = getOrderSeconds(order, now);
  const type = ORDER_TYPES[order.type];
  const status = STATUS_META[order.status];
  const isRunning = order.status === "en_progreso";

  return (
    <div
      className="rounded-lg overflow-hidden"
      style={{ backgroundColor: COLORS.surface, borderLeft: `4px solid ${type.color}` }}
    >
      <button onClick={onToggle} className="w-full text-left px-4 py-3 flex items-center gap-3">
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-2 flex-wrap mb-1">
            <span
              className="text-sm font-bold tracking-wide"
              style={{ fontFamily: "'JetBrains Mono', monospace", color: COLORS.text }}
            >
              {order.code}
            </span>
            <Badge text={order.type} color={type.color} dim={type.dim} />
            <Badge text={status.label} color={status.color} dim={status.dim} />
            {order.branch && <Badge text={order.branch} color={COLORS.textDim} dim={COLORS.surfaceAlt} />}
          </div>
          <p className="text-sm truncate" style={{ color: COLORS.text }}>
            {order.client} · {order.vehicle}
          </p>
          {showMechanic && (
            <p className="text-xs mt-0.5 truncate" style={{ color: COLORS.textDim }}>
              Asignado a {mechanicNames && mechanicNames.length ? mechanicNames.join(", ") : "sin asignar"}
            </p>
          )}
        </div>
        <div className="text-right shrink-0">
          <div
            className="flex items-center gap-1.5 justify-end text-lg font-bold"
            style={{ fontFamily: "'JetBrains Mono', monospace", color: isRunning ? COLORS.green : COLORS.text }}
          >
            {isRunning && (
              <span className="relative flex h-2 w-2">
                <span
                  className="animate-ping absolute inline-flex h-full w-full rounded-full opacity-75"
                  style={{ backgroundColor: COLORS.green }}
                />
                <span className="relative inline-flex rounded-full h-2 w-2" style={{ backgroundColor: COLORS.green }} />
              </span>
            )}
            {formatClock(seconds)}
          </div>
          {expanded ? (
            <ChevronUp size={16} style={{ color: COLORS.textDim }} className="ml-auto mt-1" />
          ) : (
            <ChevronDown size={16} style={{ color: COLORS.textDim }} className="ml-auto mt-1" />
          )}
        </div>
      </button>

      {expanded && (
        <div className="px-4 pb-4 pt-1" style={{ borderTop: `1px dashed ${COLORS.line}` }}>
          {order.plate && (
            <p className="text-xs mt-2" style={{ color: COLORS.textDim }}>
              Chasis: <span style={{ color: COLORS.text }}>{order.plate}</span>
            </p>
          )}
          {order.description && (
            <p className="text-sm mt-2" style={{ color: COLORS.text }}>
              {order.description}
            </p>
          )}
          {order.finishNote && (
            <div className="mt-3 rounded-md p-2.5" style={{ backgroundColor: COLORS.tealDim }}>
              <p className="text-xs font-semibold mb-1" style={{ color: COLORS.teal }}>OBSERVACIÓN FINAL</p>
              <p className="text-sm" style={{ color: COLORS.text }}>{order.finishNote}</p>
            </div>
          )}
          {order.trip && (
            <div className="mt-3 rounded-md p-2.5" style={{ backgroundColor: COLORS.surfaceAlt }}>
              <p className="text-xs font-semibold mb-1" style={{ color: COLORS.textDim }}>VIAJE</p>
              <p className="text-sm" style={{ color: COLORS.text }}>
                {order.trip.origin || "?"} → {order.trip.destination || "?"}
              </p>
              <p className="text-xs mt-0.5" style={{ color: COLORS.textDim }}>
                {order.trip.km || 0} km recorridos
              </p>
            </div>
          )}
          {order.segments?.length > 0 && (
            <div className="mt-3 space-y-1">
              <p className="text-xs font-semibold" style={{ color: COLORS.textDim }}>
                REGISTRO DE TIEMPO
              </p>
              {order.segments.map((seg, i) => (
                <div key={i} className="text-xs" style={{ color: COLORS.textDim }}>
                  <div className="flex justify-between">
                    <span>
                      {new Date(seg.start).toLocaleTimeString("es-AR", { hour: "2-digit", minute: "2-digit" })}
                      {" → "}
                      {seg.end
                        ? new Date(seg.end).toLocaleTimeString("es-AR", { hour: "2-digit", minute: "2-digit" })
                        : "en curso"}
                    </span>
                    <span style={{ fontFamily: "'JetBrains Mono', monospace" }}>
                      {formatClock((seg.end || now) - seg.start ? ((seg.end || now) - seg.start) / 1000 : 0)}
                    </span>
                  </div>
                  {seg.note && (
                    <p className="mt-0.5 italic" style={{ color: COLORS.text }}>
                      "{seg.note}"
                    </p>
                  )}
                </div>
              ))}
            </div>
          )}
          {actions && <div className="flex gap-2 mt-4">{actions}</div>}
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   MAIN APP
--------------------------------------------------------- */
export default function App() {
  const [loading, setLoading] = useState(true);
  const [orders, setOrders] = useState([]);
  const [mechanics, setMechanics] = useState([]);
  const [role, setRole] = useState("admin");
  const [now, setNow] = useState(Date.now());
  const [storageOk, setStorageOk] = useState(true);

  // acceso
  const [accessCode, setAccessCode] = useState("02120512");
  const [unlocked, setUnlocked] = useState(false);
  const [codeInput, setCodeInput] = useState("");
  const [codeError, setCodeError] = useState("");
  const [loggedInMechanicId, setLoggedInMechanicId] = useState(null);

  // admin state
  const [adminTab, setAdminTab] = useState("ordenes");
  const [filterStatus, setFilterStatus] = useState("todas");
  const [filterType, setFilterType] = useState("todas");
  const [filterBranch, setFilterBranch] = useState("todas");
  const [showForm, setShowForm] = useState(false);
  const [expandedId, setExpandedId] = useState(null);
  const [reportMonth, setReportMonth] = useState(currentMonthKey());
  const [newMechName, setNewMechName] = useState("");
  const [form, setForm] = useState({
    client: "", vehicle: "", plate: "", description: "", type: "OTN", mechanicIds: [], branch: BRANCHES[0],
    hasTrip: false, tripKm: "", tripOrigin: "", tripDestination: "",
  });

  // mechanic state
  const [activeMechId, setActiveMechId] = useState("");
  const [showFinished, setShowFinished] = useState(false);

  /* ---------- load (Firestore, tiempo real) ---------- */
  useEffect(() => {
    const unsubscribe = onSnapshot(
      WORKSHOP_DOC,
      (snap) => {
        if (snap.exists()) {
          const data = snap.data();
          setMechanics(data.mechanics || DEFAULT_MECHANICS);
          setOrders(data.orders || []);
          setAccessCode(data.accessCode || "02120512");
          setActiveMechId((prev) => prev || (data.mechanics || DEFAULT_MECHANICS)[0]?.id || "");
          setForm((f) => ({
            ...f,
            mechanicIds: f.mechanicIds.length ? f.mechanicIds : [(data.mechanics || DEFAULT_MECHANICS)[0]?.id].filter(Boolean),
          }));
        } else {
          setDoc(WORKSHOP_DOC, { mechanics: DEFAULT_MECHANICS, orders: [], accessCode: "02120512" });
          setMechanics(DEFAULT_MECHANICS);
          setActiveMechId(DEFAULT_MECHANICS[0]?.id || "");
          setForm((f) => ({ ...f, mechanicIds: DEFAULT_MECHANICS[0]?.id ? [DEFAULT_MECHANICS[0].id] : [] }));
        }
        setLoading(false);
      },
      (err) => {
        console.error(err);
        setStorageOk(false);
        setLoading(false);
      }
    );
    return () => unsubscribe();
  }, []);

  /* ---------- live tick ---------- */
  useEffect(() => {
    const interval = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(interval);
  }, []);

  /* ---------- persistence (Firestore) ---------- */
  const persistOrders = useCallback(async (next) => {
    setOrders(next);
    try {
      await updateDoc(WORKSHOP_DOC, { orders: next });
    } catch (e) {
      setStorageOk(false);
    }
  }, []);

  const persistMechanics = useCallback(async (next) => {
    setMechanics(next);
    try {
      await updateDoc(WORKSHOP_DOC, { mechanics: next });
    } catch (e) {
      setStorageOk(false);
    }
  }, []);

  function checkCode(e) {
    e.preventDefault();
    const val = codeInput.trim();
    if (val === String(accessCode)) {
      setUnlocked(true);
      setLoggedInMechanicId(null);
      setRole("admin");
      setCodeError("");
      return;
    }
    const mech = mechanics.find((m) => m.code && String(m.code) === val);
    if (mech) {
      setUnlocked(true);
      setLoggedInMechanicId(mech.id);
      setRole("mecanico");
      setActiveMechId(mech.id);
      setCodeError("");
      return;
    }
    setCodeError("Código incorrecto. Pedíselo al administrador.");
  }

  const changeAccessCode = useCallback(async (newCode) => {
    setAccessCode(newCode);
    try {
      await updateDoc(WORKSHOP_DOC, { accessCode: newCode });
    } catch (e) {
      setStorageOk(false);
    }
  }, []);

  /* ---------- order actions ---------- */
  function toggleFormMechanic(id) {
    setForm((f) => ({
      ...f,
      mechanicIds: f.mechanicIds.includes(id)
        ? f.mechanicIds.filter((mid) => mid !== id)
        : [...f.mechanicIds, id],
    }));
  }

  function createOrder(e) {
    e.preventDefault();
    if (!form.client.trim() || !form.vehicle.trim() || form.mechanicIds.length === 0) return;
    const seq = orders.filter((o) => o.type === form.type).length + 1;
    const order = {
      id: uid(),
      code: `${form.type}-${String(seq).padStart(4, "0")}`,
      client: form.client.trim(),
      vehicle: form.vehicle.trim(),
      plate: form.plate.trim(),
      description: form.description.trim(),
      type: form.type,
      mechanicIds: form.mechanicIds,
      branch: form.branch,
      status: "pendiente",
      segments: [],
      createdAt: Date.now(),
      trip: form.hasTrip
        ? { km: form.tripKm, origin: form.tripOrigin.trim(), destination: form.tripDestination.trim() }
        : null,
    };
    persistOrders([order, ...orders]);
    setForm({
      client: "", vehicle: "", plate: "", description: "", type: "OTN",
      mechanicIds: mechanics[0]?.id ? [mechanics[0].id] : [], branch: form.branch,
      hasTrip: false, tripKm: "", tripOrigin: "", tripDestination: "",
    });
    setShowForm(false);
  }

  function updateOrder(id, updater) {
    const next = orders.map((o) => (o.id === id ? updater(o) : o));
    persistOrders(next);
  }

  function startOrder(id) {
    updateOrder(id, (o) => ({
      ...o,
      status: "en_progreso",
      segments: [...(o.segments || []), { start: Date.now(), end: null }],
    }));
  }

  function pauseOrder(id, note) {
    updateOrder(id, (o) => {
      const segs = [...(o.segments || [])];
      if (segs.length && !segs[segs.length - 1].end) {
        segs[segs.length - 1] = { ...segs[segs.length - 1], end: Date.now(), note: (note || "").trim() };
      }
      return { ...o, status: "pausada", segments: segs };
    });
  }

  function resumeOrder(id) {
    updateOrder(id, (o) => ({
      ...o,
      status: "en_progreso",
      segments: [...(o.segments || []), { start: Date.now(), end: null }],
    }));
  }

  function finishOrder(id, note) {
    updateOrder(id, (o) => {
      const segs = [...(o.segments || [])];
      if (segs.length && !segs[segs.length - 1].end) segs[segs.length - 1] = { ...segs[segs.length - 1], end: Date.now() };
      return { ...o, status: "finalizada", segments: segs, finishNote: (note || "").trim() };
    });
  }

  function removeMechanicFromOrders(mechId) {
    const next = orders.map((o) => {
      const ids = getOrderMechanicIds(o);
      if (!ids.includes(mechId)) return o;
      return { ...o, mechanicIds: ids.filter((id) => id !== mechId) };
    });
    persistOrders(next);
  }

  function addMechanic(e) {
    e.preventDefault();
    if (!newMechName.trim()) return;
    persistMechanics([...mechanics, { id: uid(), name: newMechName.trim(), code: genMechCode() }]);
    setNewMechName("");
  }

  function regenerateMechanicCode(id) {
    persistMechanics(mechanics.map((m) => (m.id === id ? { ...m, code: genMechCode() } : m)));
  }

  /* ---------- derived data ---------- */
  const mechanicById = useMemo(() => {
    const map = {};
    mechanics.forEach((m) => (map[m.id] = m.name));
    return map;
  }, [mechanics]);

  const filteredOrders = useMemo(() => {
    return orders.filter((o) => {
      if (filterStatus !== "todas" && o.status !== filterStatus) return false;
      if (filterType !== "todas" && o.type !== filterType) return false;
      if (filterBranch !== "todas" && o.branch !== filterBranch) return false;
      return true;
    });
  }, [orders, filterStatus, filterType, filterBranch]);

  const mechanicOrders = useMemo(() => {
    return orders
      .filter((o) => getOrderMechanicIds(o).includes(activeMechId))
      .filter((o) => (showFinished ? true : o.status !== "finalizada"))
      .sort((a, b) => b.createdAt - a.createdAt);
  }, [orders, activeMechId, showFinished]);

  const reportRows = useMemo(() => {
    const acc = {};
    mechanics.forEach((m) => {
      acc[m.id] = { name: m.name, OTN: 0, OTI: 0, OTG: 0 };
    });
    orders.forEach((o) => {
      const ids = getOrderMechanicIds(o);
      if (ids.length === 0) return;
      (o.segments || []).forEach((seg) => {
        if (!seg.end) return; // solo tiempo cerrado cuenta para el reporte
        if (monthKey(seg.start) !== reportMonth) return;
        const share = (seg.end - seg.start) / 1000 / ids.length; // partes iguales entre los asignados
        ids.forEach((id) => {
          if (acc[id]) acc[id][o.type] += share;
        });
      });
    });
    return Object.values(acc);
  }, [orders, mechanics, reportMonth]);

  const chartData = reportRows.map((r) => ({
    name: r.name.split(" ")[0],
    OTN: +formatHours(r.OTN),
    OTI: +formatHours(r.OTI),
    OTG: +formatHours(r.OTG),
  }));

  async function exportExcel() {
    const wb = new ExcelJS.Workbook();
    wb.creator = "Proyan";
    const ws = wb.addWorksheet(`Horas ${reportMonth}`);

    ws.columns = [
      { header: "Mecánico", key: "name", width: 26 },
      { header: "OTN (h)", key: "OTN", width: 12 },
      { header: "OTI (h)", key: "OTI", width: 12 },
      { header: "OTG (h)", key: "OTG", width: 12 },
      { header: "Total (h)", key: "total", width: 14 },
    ];

    const headerRow = ws.getRow(1);
    headerRow.height = 22;
    headerRow.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1B1E22" } };
      cell.alignment = { vertical: "middle", horizontal: "center" };
    });

    const grand = { OTN: 0, OTI: 0, OTG: 0 };
    reportRows.forEach((r) => {
      const total = r.OTN + r.OTI + r.OTG;
      grand.OTN += r.OTN;
      grand.OTI += r.OTI;
      grand.OTG += r.OTG;
      const row = ws.addRow({
        name: r.name,
        OTN: +formatHours(r.OTN),
        OTI: +formatHours(r.OTI),
        OTG: +formatHours(r.OTG),
        total: +formatHours(total),
      });
      row.alignment = { vertical: "middle" };
      ["OTN", "OTI", "OTG", "total"].forEach((k) => {
        row.getCell(k).numFmt = "0.00";
        row.getCell(k).alignment = { horizontal: "right" };
      });
    });

    const grandTotal = grand.OTN + grand.OTI + grand.OTG;
    const totalRow = ws.addRow({
      name: "TOTAL",
      OTN: +formatHours(grand.OTN),
      OTI: +formatHours(grand.OTI),
      OTG: +formatHours(grand.OTG),
      total: +formatHours(grandTotal),
    });
    totalRow.eachCell((cell) => {
      cell.font = { bold: true };
      cell.border = { top: { style: "thin", color: { argb: "FF888888" } } };
    });
    ["OTN", "OTI", "OTG", "total"].forEach((k) => {
      totalRow.getCell(k).numFmt = "0.00";
      totalRow.getCell(k).alignment = { horizontal: "right" };
    });

    ws.eachRow((row) => {
      row.eachCell((cell) => {
        cell.border = {
          ...cell.border,
          left: { style: "thin", color: { argb: "FFE0E0E0" } },
          right: { style: "thin", color: { argb: "FFE0E0E0" } },
          bottom: cell.border?.bottom || { style: "thin", color: { argb: "FFE0E0E0" } },
        };
      });
    });

    ws.views = [{ state: "frozen", ySplit: 1 }];

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `reporte_horas_${reportMonth}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
  }

  /* ---------------------------------------------------------
     RENDER
  --------------------------------------------------------- */
  if (loading) {
    return (
      <div className="min-h-screen flex flex-col items-center justify-center gap-4" style={{ backgroundColor: COLORS.bg }}>
        <style>{FONTS_IMPORT}</style>
        <Wrench size={30} style={{ color: COLORS.accent }} />
        <h1
          className="text-2xl tracking-wide uppercase"
          style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 700, color: COLORS.text }}
        >
          Proyan
        </h1>
        <Loader2 className="animate-spin" size={22} style={{ color: COLORS.accent }} />
        <p className="text-xs absolute bottom-8" style={{ color: COLORS.textDim, fontFamily: "'Inter', sans-serif" }}>
          Creado por Tomás Alcaraz
        </p>
      </div>
    );
  }

  if (!unlocked) {
    return (
      <div className="min-h-screen flex flex-col items-center justify-center gap-5 px-6" style={{ backgroundColor: COLORS.bg }}>
        <style>{FONTS_IMPORT}</style>
        <Wrench size={28} style={{ color: COLORS.accent }} />
        <h1
          className="text-xl tracking-wide uppercase"
          style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 700, color: COLORS.text }}
        >
          Proyan
        </h1>
        <form onSubmit={checkCode} className="w-full max-w-xs space-y-3">
          <Field label="Código de acceso">
            <input
              autoFocus
              value={codeInput}
              onChange={(e) => setCodeInput(e.target.value)}
              placeholder="Pedíselo al administrador"
              style={{ ...inputStyle, width: "100%", textAlign: "center", letterSpacing: 2 }}
              inputMode="numeric"
            />
          </Field>
          {codeError && (
            <p className="text-xs text-center" style={{ color: COLORS.red }}>{codeError}</p>
          )}
          <ActionButton icon={Check} label="Ingresar" tone="accent" full onClick={checkCode} />
        </form>
        <p className="text-xs absolute bottom-8" style={{ color: COLORS.textDim }}>
          Creado por Tomás Alcaraz
        </p>
      </div>
    );
  }

  return (
    <div className="min-h-screen pb-10" style={{ backgroundColor: COLORS.bg, fontFamily: "'Inter', sans-serif" }}>
      <style>{FONTS_IMPORT}</style>

      {/* HEADER */}
      <div style={{ backgroundColor: COLORS.surface, borderBottom: `1px solid ${COLORS.line}` }}>
        <div className="max-w-md mx-auto px-4 pt-5 pb-4">
          <div className="flex items-center gap-2 mb-1">
            <Wrench size={20} style={{ color: COLORS.accent }} />
            <h1
              className="text-xl tracking-wide uppercase"
              style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, color: COLORS.text }}
            >
              Proyan
            </h1>
          </div>
          <p className="text-xs capitalize" style={{ color: COLORS.textDim }}>
            Control de órdenes · {todayLabel()}
          </p>

          {loggedInMechanicId ? (
            <p className="text-sm mt-4 font-medium" style={{ color: COLORS.accent }}>
              Ingresaste como {mechanicById[loggedInMechanicId] || "mecánico"}
            </p>
          ) : (
            <div className="flex mt-4 rounded-lg overflow-hidden" style={{ border: `1px solid ${COLORS.line}` }}>
              {["admin", "mecanico"].map((r) => (
                <button
                  key={r}
                  onClick={() => setRole(r)}
                  className="flex-1 py-2 text-sm font-semibold transition-colors"
                  style={{
                    backgroundColor: role === r ? COLORS.accent : "transparent",
                    color: role === r ? "#1A1300" : COLORS.textDim,
                  }}
                >
                  {r === "admin" ? "Panel Admin" : "Panel Mecánico"}
                </button>
              ))}
            </div>
          )}
          {!storageOk && (
            <p className="text-xs mt-2" style={{ color: COLORS.red }}>
              No se pudo sincronizar con el almacenamiento — los cambios quedan solo en esta sesión.
            </p>
          )}
        </div>
      </div>

      <div className="max-w-md mx-auto px-4 mt-4">
        {role === "admin" ? (
          <AdminPanel
            adminTab={adminTab}
            setAdminTab={setAdminTab}
            filterStatus={filterStatus}
            setFilterStatus={setFilterStatus}
            filterType={filterType}
            setFilterType={setFilterType}
            filterBranch={filterBranch}
            setFilterBranch={setFilterBranch}
            showForm={showForm}
            setShowForm={setShowForm}
            form={form}
            setForm={setForm}
            createOrder={createOrder}
            toggleFormMechanic={toggleFormMechanic}
            filteredOrders={filteredOrders}
            expandedId={expandedId}
            setExpandedId={setExpandedId}
            mechanicById={mechanicById}
            mechanics={mechanics}
            now={now}
            reportMonth={reportMonth}
            setReportMonth={setReportMonth}
            reportRows={reportRows}
            chartData={chartData}
            exportExcel={exportExcel}
            newMechName={newMechName}
            setNewMechName={setNewMechName}
            addMechanic={addMechanic}
            orders={orders}
            persistMechanics={persistMechanics}
            regenerateMechanicCode={regenerateMechanicCode}
            removeMechanicFromOrders={removeMechanicFromOrders}
            accessCode={accessCode}
            changeAccessCode={changeAccessCode}
          />
        ) : (
          <MechanicPanel
            mechanics={mechanics}
            activeMechId={activeMechId}
            setActiveMechId={setActiveMechId}
            mechanicOrders={mechanicOrders}
            showFinished={showFinished}
            setShowFinished={setShowFinished}
            expandedId={expandedId}
            setExpandedId={setExpandedId}
            now={now}
            startOrder={startOrder}
            pauseOrder={pauseOrder}
            resumeOrder={resumeOrder}
            finishOrder={finishOrder}
            loggedInMechanicId={loggedInMechanicId}
          />
        )}
      </div>

      <p className="text-center text-[11px] mt-8" style={{ color: COLORS.textDim }}>
        Proyan · Creado por Tomás Alcaraz
      </p>
    </div>
  );
}

/* ---------------------------------------------------------
   ADMIN PANEL
--------------------------------------------------------- */
function AdminPanel(props) {
  const {
    adminTab, setAdminTab, filterStatus, setFilterStatus, filterType, setFilterType,
    filterBranch, setFilterBranch,
    showForm, setShowForm, form, setForm, createOrder, toggleFormMechanic, filteredOrders, expandedId, setExpandedId,
    mechanicById, mechanics, now, reportMonth, setReportMonth, reportRows, chartData, exportExcel,
    newMechName, setNewMechName, addMechanic, orders, persistMechanics, accessCode, changeAccessCode,
    regenerateMechanicCode, removeMechanicFromOrders,
  } = props;

  const [editingId, setEditingId] = useState(null);
  const [editingName, setEditingName] = useState("");
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [codeDraft, setCodeDraft] = useState(accessCode);
  const [codeSaved, setCodeSaved] = useState(false);

  function startEdit(m) {
    setEditingId(m.id);
    setEditingName(m.name);
  }

  function saveEdit(id) {
    const name = editingName.trim();
    if (!name) return;
    persistMechanics(mechanics.map((m) => (m.id === id ? { ...m, name } : m)));
    setEditingId(null);
    setEditingName("");
  }

  function confirmDelete(id) {
    persistMechanics(mechanics.filter((m) => m.id !== id));
    removeMechanicFromOrders(id);
    setConfirmDeleteId(null);
  }

  function saveCode(e) {
    e.preventDefault();
    if (!codeDraft.trim()) return;
    changeAccessCode(codeDraft.trim());
    setCodeSaved(true);
    setTimeout(() => setCodeSaved(false), 2000);
  }

  return (
    <div>
      <div className="flex gap-2 mb-4">
        {[
          { id: "ordenes", label: "Órdenes", icon: ClipboardList },
          { id: "reportes", label: "Reportes", icon: BarChart3 },
          { id: "equipo", label: "Equipo", icon: Users },
        ].map((t) => (
          <button
            key={t.id}
            onClick={() => setAdminTab(t.id)}
            className="flex-1 flex flex-col items-center gap-1 py-2 rounded-lg text-xs font-semibold"
            style={{
              backgroundColor: adminTab === t.id ? COLORS.surfaceAlt : "transparent",
              color: adminTab === t.id ? COLORS.accent : COLORS.textDim,
              border: `1px solid ${adminTab === t.id ? COLORS.accent : COLORS.line}`,
            }}
          >
            <t.icon size={16} />
            {t.label}
          </button>
        ))}
      </div>

      {adminTab === "ordenes" && (
        <div>
          <div className="flex gap-2 overflow-x-auto pb-1 mb-2">
            {["todas", "pendiente", "en_progreso", "pausada", "finalizada"].map((s) => (
              <Chip key={s} active={filterStatus === s} onClick={() => setFilterStatus(s)}>
                {s === "todas" ? "Todas" : STATUS_META[s].label}
              </Chip>
            ))}
          </div>
          <div className="flex gap-2 overflow-x-auto pb-1 mb-3">
            {["todas", "OTN", "OTI", "OTG"].map((t) => (
              <Chip key={t} active={filterType === t} onClick={() => setFilterType(t)}>
                {t}
              </Chip>
            ))}
          </div>
          <div className="flex gap-2 overflow-x-auto pb-1 mb-3">
            {["todas", ...BRANCHES].map((b) => (
              <Chip key={b} active={filterBranch === b} onClick={() => setFilterBranch(b)}>
                {b === "todas" ? "Todas las sucursales" : b}
              </Chip>
            ))}
          </div>

          <button
            onClick={() => setShowForm((v) => !v)}
            className="w-full mb-3 flex items-center justify-center gap-2 rounded-lg py-2.5 font-semibold text-sm"
            style={{ backgroundColor: showForm ? COLORS.surfaceAlt : COLORS.accent, color: showForm ? COLORS.text : "#1A1300" }}
          >
            {showForm ? <X size={16} /> : <Plus size={16} />}
            {showForm ? "Cancelar" : "Nueva orden"}
          </button>

          {showForm && (
            <form
              onSubmit={createOrder}
              className="rounded-lg p-4 mb-4 space-y-3"
              style={{ backgroundColor: COLORS.surface, border: `1px solid ${COLORS.line}` }}
            >
              <Field label="Cliente">
                <input
                  value={form.client}
                  onChange={(e) => setForm({ ...form, client: e.target.value })}
                  className="w-full"
                  style={inputStyle}
                  placeholder="Nombre del cliente"
                  required
                />
              </Field>
              <Field label="Vehículo">
                <input
                  value={form.vehicle}
                  onChange={(e) => setForm({ ...form, vehicle: e.target.value })}
                  className="w-full"
                  style={inputStyle}
                  placeholder="Marca / modelo"
                  required
                />
              </Field>
              <Field label="Chasis">
                <input
                  value={form.plate}
                  onChange={(e) => setForm({ ...form, plate: e.target.value })}
                  className="w-full"
                  style={inputStyle}
                  placeholder="Número de chasis"
                />
              </Field>
              <Field label="Descripción del trabajo">
                <textarea
                  value={form.description}
                  onChange={(e) => setForm({ ...form, description: e.target.value })}
                  className="w-full"
                  style={{ ...inputStyle, minHeight: 70 }}
                  placeholder="Detalle de la reparación"
                />
              </Field>
              <Field label="Sucursal">
                <div className="flex gap-2">
                  {BRANCHES.map((b) => (
                    <button
                      type="button"
                      key={b}
                      onClick={() => setForm({ ...form, branch: b })}
                      className="flex-1 py-2 rounded-md text-xs font-bold"
                      style={{
                        backgroundColor: form.branch === b ? COLORS.accentDim : COLORS.surfaceAlt,
                        color: form.branch === b ? COLORS.accent : COLORS.textDim,
                        border: `1px solid ${form.branch === b ? COLORS.accent : COLORS.line}`,
                      }}
                    >
                      {b}
                    </button>
                  ))}
                </div>
              </Field>
              <Field label="Tipo de orden">
                <div className="flex gap-2">
                  {Object.keys(ORDER_TYPES).map((t) => (
                    <button
                      type="button"
                      key={t}
                      onClick={() => setForm({ ...form, type: t })}
                      className="flex-1 py-2 rounded-md text-xs font-bold"
                      style={{
                        backgroundColor: form.type === t ? ORDER_TYPES[t].dim : COLORS.surfaceAlt,
                        color: form.type === t ? ORDER_TYPES[t].color : COLORS.textDim,
                        border: `1px solid ${form.type === t ? ORDER_TYPES[t].color : COLORS.line}`,
                      }}
                    >
                      {t}
                    </button>
                  ))}
                </div>
              </Field>
              <Field label="Mecánicos asignados (podés elegir más de uno)">
                <div className="space-y-1.5">
                  {mechanics.map((m) => {
                    const checked = form.mechanicIds.includes(m.id);
                    return (
                      <label
                        key={m.id}
                        className="flex items-center gap-2 text-sm rounded-md px-2 py-1.5"
                        style={{ backgroundColor: checked ? COLORS.accentDim : "transparent", color: COLORS.text }}
                      >
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => toggleFormMechanic(m.id)}
                          style={{ accentColor: COLORS.accent }}
                        />
                        {m.name}
                      </label>
                    );
                  })}
                </div>
              </Field>

              <label className="flex items-center gap-2 text-sm" style={{ color: COLORS.text }}>
                <input
                  type="checkbox"
                  checked={form.hasTrip}
                  onChange={(e) => setForm({ ...form, hasTrip: e.target.checked })}
                  style={{ accentColor: COLORS.accent }}
                />
                Esta orden incluye un viaje
              </label>

              {form.hasTrip && (
                <div className="space-y-3 rounded-md p-3" style={{ backgroundColor: COLORS.surfaceAlt }}>
                  <Field label="Km recorridos">
                    <input
                      type="number"
                      min="0"
                      value={form.tripKm}
                      onChange={(e) => setForm({ ...form, tripKm: e.target.value })}
                      className="w-full"
                      style={inputStyle}
                      placeholder="0"
                    />
                  </Field>
                  <Field label="Lugar de inicio">
                    <input
                      value={form.tripOrigin}
                      onChange={(e) => setForm({ ...form, tripOrigin: e.target.value })}
                      className="w-full"
                      style={inputStyle}
                      placeholder="Punto de partida"
                    />
                  </Field>
                  <Field label="Destino">
                    <input
                      value={form.tripDestination}
                      onChange={(e) => setForm({ ...form, tripDestination: e.target.value })}
                      className="w-full"
                      style={inputStyle}
                      placeholder="Punto de llegada"
                    />
                  </Field>
                </div>
              )}

              <ActionButton icon={Plus} label="Crear orden" tone="accent" full onClick={createOrder} />
            </form>
          )}

          <div className="space-y-2">
            {filteredOrders.length === 0 && (
              <EmptyState text="No hay órdenes que coincidan con estos filtros." />
            )}
            {filteredOrders.map((o) => (
              <TicketCard
                key={o.id}
                order={o}
                mechanicNames={getOrderMechanicIds(o).map((id) => mechanicById[id]).filter(Boolean)}
                now={now}
                expanded={expandedId === o.id}
                onToggle={() => setExpandedId(expandedId === o.id ? null : o.id)}
                showMechanic
              />
            ))}
          </div>
        </div>
      )}

      {adminTab === "reportes" && (
        <div>
          <Field label="Mes">
            <input
              type="month"
              value={reportMonth}
              onChange={(e) => setReportMonth(e.target.value)}
              style={inputStyle}
              className="w-full"
            />
          </Field>

          {orders.length === 0 ? (
            <EmptyState text="Todavía no hay órdenes cargadas para reportar." />
          ) : (
            <>
              <div className="rounded-lg p-3 mt-3" style={{ backgroundColor: COLORS.surface, border: `1px solid ${COLORS.line}` }}>
                <ResponsiveContainer width="100%" height={Math.max(160, reportRows.length * 44)}>
                  <BarChart data={chartData} layout="vertical" margin={{ left: 0, right: 10 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke={COLORS.line} horizontal={false} />
                    <XAxis type="number" stroke={COLORS.textDim} fontSize={11} unit="h" />
                    <YAxis type="category" dataKey="name" stroke={COLORS.textDim} fontSize={11} width={70} />
                    <Tooltip
                      contentStyle={{ backgroundColor: COLORS.surfaceAlt, border: `1px solid ${COLORS.line}`, fontSize: 12 }}
                      labelStyle={{ color: COLORS.text }}
                    />
                    <Legend wrapperStyle={{ fontSize: 11 }} />
                    <Bar dataKey="OTN" stackId="a" fill={ORDER_TYPES.OTN.color} radius={[0, 0, 0, 0]} />
                    <Bar dataKey="OTI" stackId="a" fill={ORDER_TYPES.OTI.color} />
                    <Bar dataKey="OTG" stackId="a" fill={ORDER_TYPES.OTG.color} radius={[0, 4, 4, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>

              <div className="rounded-lg mt-3 overflow-hidden" style={{ border: `1px solid ${COLORS.line}` }}>
                <table className="w-full text-xs" style={{ borderCollapse: "collapse" }}>
                  <thead>
                    <tr style={{ backgroundColor: COLORS.surfaceAlt, color: COLORS.textDim }}>
                      <th className="text-left py-2 px-2 font-semibold">Mecánico</th>
                      <th className="text-right py-2 px-2 font-semibold" style={{ color: ORDER_TYPES.OTN.color }}>OTN</th>
                      <th className="text-right py-2 px-2 font-semibold" style={{ color: ORDER_TYPES.OTI.color }}>OTI</th>
                      <th className="text-right py-2 px-2 font-semibold" style={{ color: ORDER_TYPES.OTG.color }}>OTG</th>
                      <th className="text-right py-2 px-2 font-semibold" style={{ color: COLORS.text }}>Total</th>
                    </tr>
                  </thead>
                  <tbody>
                    {reportRows.map((r, i) => {
                      const total = r.OTN + r.OTI + r.OTG;
                      return (
                        <tr key={i} style={{ borderTop: `1px solid ${COLORS.line}`, color: COLORS.text }}>
                          <td className="py-2 px-2">{r.name}</td>
                          <td className="text-right py-2 px-2" style={{ fontFamily: "'JetBrains Mono', monospace" }}>{formatHours(r.OTN)}</td>
                          <td className="text-right py-2 px-2" style={{ fontFamily: "'JetBrains Mono', monospace" }}>{formatHours(r.OTI)}</td>
                          <td className="text-right py-2 px-2" style={{ fontFamily: "'JetBrains Mono', monospace" }}>{formatHours(r.OTG)}</td>
                          <td className="text-right py-2 px-2 font-bold" style={{ fontFamily: "'JetBrains Mono', monospace" }}>{formatHours(total)}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              <button
                onClick={exportExcel}
                className="w-full mt-3 flex items-center justify-center gap-2 rounded-lg py-2.5 text-sm font-semibold"
                style={{ backgroundColor: COLORS.surfaceAlt, color: COLORS.text, border: `1px solid ${COLORS.line}` }}
              >
                <Download size={16} />
                Exportar Excel
              </button>
              <p className="text-[11px] mt-2 text-center" style={{ color: COLORS.textDim }}>
                Solo se contabilizan tramos de tiempo ya cerrados (pausados o finalizados).
              </p>
            </>
          )}
        </div>
      )}

      {adminTab === "equipo" && (
        <div>
          <form onSubmit={addMechanic} className="flex gap-2 mb-4">
            <input
              value={newMechName}
              onChange={(e) => setNewMechName(e.target.value)}
              placeholder="Nombre del mecánico"
              style={{ ...inputStyle, flex: 1 }}
            />
            <button
              type="submit"
              className="px-4 rounded-md font-semibold text-sm"
              style={{ backgroundColor: COLORS.accent, color: "#1A1300" }}
            >
              Agregar
            </button>
          </form>
          <p className="text-xs mb-3" style={{ color: COLORS.textDim }}>
            Cada mecánico tiene su propio código (el número naranja). Con ese código entra directo a su panel — no ve Admin ni puede elegir ser otro mecánico.
          </p>
          <div className="space-y-2">
            {mechanics.length === 0 && <EmptyState text="Todavía no hay mecánicos cargados." />}
            {mechanics.map((m) => {
              const count = orders.filter((o) => getOrderMechanicIds(o).includes(m.id) && o.status !== "finalizada").length;
              const isEditing = editingId === m.id;
              const isConfirmingDelete = confirmDeleteId === m.id;
              return (
                <div
                  key={m.id}
                  className="rounded-lg px-4 py-3"
                  style={{ backgroundColor: COLORS.surface, border: `1px solid ${COLORS.line}` }}
                >
                  {isConfirmingDelete ? (
                    <div>
                      <p className="text-sm" style={{ color: COLORS.text }}>
                        {count > 0
                          ? `${m.name} tiene ${count} orden(es) activa(s). Si lo eliminás, esas órdenes quedan sin mecánico asignado.`
                          : `¿Eliminar a ${m.name} del equipo?`}
                      </p>
                      <div className="flex gap-2 mt-2">
                        <ActionButton icon={Trash2} label="Eliminar" tone="red" full onClick={() => confirmDelete(m.id)} />
                        <ActionButton icon={X} label="Cancelar" tone="ghost" full onClick={() => setConfirmDeleteId(null)} />
                      </div>
                    </div>
                  ) : (
                    <div className="flex items-center justify-between gap-2">
                      {isEditing ? (
                        <input
                          autoFocus
                          value={editingName}
                          onChange={(e) => setEditingName(e.target.value)}
                          onKeyDown={(e) => {
                            if (e.key === "Enter") saveEdit(m.id);
                            if (e.key === "Escape") setEditingId(null);
                          }}
                          style={{ ...inputStyle, flex: 1, padding: "6px 10px" }}
                        />
                      ) : (
                        <div className="min-w-0">
                          <span style={{ color: COLORS.text }} className="text-sm font-medium block truncate">{m.name}</span>
                          <span className="text-xs" style={{ color: COLORS.textDim }}>{count} activas</span>
                          <div className="flex items-center gap-1.5 mt-1">
                            <span
                              className="text-xs font-bold px-1.5 py-0.5 rounded"
                              style={{ fontFamily: "'JetBrains Mono', monospace", color: COLORS.accent, backgroundColor: COLORS.accentDim }}
                            >
                              {m.code || "sin código"}
                            </span>
                            <button
                              onClick={() => regenerateMechanicCode(m.id)}
                              className="text-xs underline"
                              style={{ color: COLORS.textDim }}
                            >
                              Regenerar
                            </button>
                          </div>
                        </div>
                      )}
                      <div className="flex items-center gap-1 shrink-0">
                        {isEditing ? (
                          <>
                            <button
                              onClick={() => saveEdit(m.id)}
                              className="p-2 rounded-md"
                              style={{ backgroundColor: COLORS.greenDim, color: COLORS.green }}
                              aria-label="Guardar"
                            >
                              <Check size={14} />
                            </button>
                            <button
                              onClick={() => setEditingId(null)}
                              className="p-2 rounded-md"
                              style={{ backgroundColor: COLORS.surfaceAlt, color: COLORS.textDim }}
                              aria-label="Cancelar"
                            >
                              <X size={14} />
                            </button>
                          </>
                        ) : (
                          <>
                            <button
                              onClick={() => startEdit(m)}
                              className="p-2 rounded-md"
                              style={{ backgroundColor: COLORS.surfaceAlt, color: COLORS.textDim }}
                              aria-label={`Editar ${m.name}`}
                            >
                              <Pencil size={14} />
                            </button>
                            <button
                              onClick={() => setConfirmDeleteId(m.id)}
                              className="p-2 rounded-md"
                              style={{ backgroundColor: COLORS.redDim, color: COLORS.red }}
                              aria-label={`Eliminar ${m.name}`}
                            >
                              <Trash2 size={14} />
                            </button>
                          </>
                        )}
                      </div>
                    </div>
                  )}
                </div>
              );
            })}
          </div>

          <div className="mt-6 rounded-lg p-4" style={{ backgroundColor: COLORS.surface, border: `1px solid ${COLORS.line}` }}>
            <p className="text-xs font-semibold uppercase tracking-wide mb-1" style={{ color: COLORS.textDim }}>
              Código de acceso del administrador
            </p>
            <p className="text-xs mb-3" style={{ color: COLORS.textDim }}>
              Es tu código, el único que abre el panel Admin. No se lo compartas a los mecánicos — ellos usan su propio código individual (arriba).
            </p>
            <form onSubmit={saveCode} className="flex gap-2">
              <input
                value={codeDraft}
                onChange={(e) => setCodeDraft(e.target.value)}
                style={{ ...inputStyle, flex: 1, fontFamily: "'JetBrains Mono', monospace", letterSpacing: 1 }}
                inputMode="numeric"
              />
              <button
                type="submit"
                className="px-4 rounded-md font-semibold text-sm"
                style={{ backgroundColor: COLORS.accent, color: "#1A1300" }}
              >
                Guardar
              </button>
            </form>
            {codeSaved && (
              <p className="text-xs mt-2" style={{ color: COLORS.green }}>Código actualizado.</p>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   MECHANIC PANEL
--------------------------------------------------------- */
function MechanicPanel(props) {
  const {
    mechanics, activeMechId, setActiveMechId, mechanicOrders, showFinished, setShowFinished,
    expandedId, setExpandedId, now, startOrder, pauseOrder, resumeOrder, finishOrder,
    loggedInMechanicId,
  } = props;

  const [pausingId, setPausingId] = useState(null);
  const [pauseNote, setPauseNote] = useState("");

  function requestPause(id) {
    setPausingId(id);
    setPauseNote("");
    setExpandedId(id);
  }

  function confirmPause(id) {
    pauseOrder(id, pauseNote);
    setPausingId(null);
    setPauseNote("");
  }

  function cancelPause() {
    setPausingId(null);
    setPauseNote("");
  }

  const [finishingId, setFinishingId] = useState(null);
  const [finishNote, setFinishNote] = useState("");

  function requestFinish(id) {
    setFinishingId(id);
    setFinishNote("");
    setExpandedId(id);
  }

  function confirmFinish(id) {
    finishOrder(id, finishNote);
    setFinishingId(null);
    setFinishNote("");
  }

  function cancelFinish() {
    setFinishingId(null);
    setFinishNote("");
  }

  return (
    <div>
      {loggedInMechanicId ? null : (
        <Field label="Sos">
          <select
            value={activeMechId}
            onChange={(e) => setActiveMechId(e.target.value)}
            style={inputStyle}
            className="w-full"
          >
            {mechanics.map((m) => (
              <option key={m.id} value={m.id}>{m.name}</option>
            ))}
          </select>
        </Field>
      )}

      <div className="flex items-center justify-between mt-4 mb-2">
        <p className="text-xs font-semibold uppercase tracking-wide" style={{ color: COLORS.textDim }}>
          Tus órdenes
        </p>
        <button onClick={() => setShowFinished((v) => !v)} className="text-xs font-medium" style={{ color: COLORS.accent }}>
          {showFinished ? "Ocultar finalizadas" : "Ver finalizadas"}
        </button>
      </div>

      <div className="space-y-2">
        {mechanicOrders.length === 0 && <EmptyState text="No tenés órdenes asignadas por ahora." />}
        {mechanicOrders.map((o) => {
          let actions = null;
          if (finishingId === o.id) {
            actions = (
              <div className="w-full">
                <label className="block text-xs font-semibold mb-1.5" style={{ color: COLORS.textDim }}>
                  Observación final
                </label>
                <textarea
                  autoFocus
                  value={finishNote}
                  onChange={(e) => setFinishNote(e.target.value)}
                  placeholder="Ej: trabajo finalizado, unidad probada y comprobado su correcto funcionamiento"
                  style={{ ...inputStyle, minHeight: 60, width: "100%" }}
                />
                <div className="flex gap-2 mt-2">
                  <ActionButton icon={Check} label="Confirmar finalización" tone="red" full onClick={() => confirmFinish(o.id)} />
                  <ActionButton icon={X} label="Cancelar" tone="ghost" full onClick={cancelFinish} />
                </div>
              </div>
            );
          } else if (pausingId === o.id) {
            actions = (
              <div className="w-full">
                <label className="block text-xs font-semibold mb-1.5" style={{ color: COLORS.textDim }}>
                  ¿Qué hiciste en este tramo? (breve)
                </label>
                <textarea
                  autoFocus
                  value={pauseNote}
                  onChange={(e) => setPauseNote(e.target.value)}
                  placeholder="Ej: terminé el cambio de aceite, falta revisar frenos"
                  style={{ ...inputStyle, minHeight: 60, width: "100%" }}
                />
                <div className="flex gap-2 mt-2">
                  <ActionButton icon={Pause} label="Confirmar pausa" tone="amber" full onClick={() => confirmPause(o.id)} />
                  <ActionButton icon={X} label="Cancelar" tone="ghost" full onClick={cancelPause} />
                </div>
              </div>
            );
          } else if (o.status === "pendiente") {
            actions = <ActionButton icon={Play} label="Iniciar" tone="green" full onClick={() => startOrder(o.id)} />;
          } else if (o.status === "en_progreso") {
            actions = (
              <>
                <ActionButton icon={Pause} label="Pausar" tone="amber" full onClick={() => requestPause(o.id)} />
                <ActionButton icon={Square} label="Finalizar" tone="red" full onClick={() => requestFinish(o.id)} />
              </>
            );
          } else if (o.status === "pausada") {
            actions = (
              <>
                <ActionButton icon={Play} label="Reanudar" tone="green" full onClick={() => resumeOrder(o.id)} />
                <ActionButton icon={Square} label="Finalizar" tone="red" full onClick={() => requestFinish(o.id)} />
              </>
            );
          }
          return (
            <TicketCard
              key={o.id}
              order={o}
              now={now}
              expanded={expandedId === o.id}
              onToggle={() => setExpandedId(expandedId === o.id ? null : o.id)}
              actions={actions}
            />
          );
        })}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   MISC
--------------------------------------------------------- */
const inputStyle = {
  backgroundColor: COLORS.surfaceAlt,
  border: `1px solid ${COLORS.line}`,
  borderRadius: 8,
  padding: "9px 12px",
  color: COLORS.text,
  fontSize: 14,
  outline: "none",
};

function Field({ label, children }) {
  return (
    <div>
      <label className="block text-xs font-semibold mb-1.5" style={{ color: COLORS.textDim }}>
        {label}
      </label>
      {children}
    </div>
  );
}

function EmptyState({ text }) {
  return (
    <div
      className="rounded-lg py-8 text-center text-sm"
      style={{ backgroundColor: COLORS.surface, color: COLORS.textDim, border: `1px dashed ${COLORS.line}` }}
    >
      {text}
    </div>
  );
}
