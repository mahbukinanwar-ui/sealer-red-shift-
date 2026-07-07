import React, { useState, useEffect, useMemo, useCallback } from "react";
import {
  LineChart, Line, BarChart, Bar, AreaChart, Area,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
} from "recharts";
import { Factory, Gauge, ClipboardList, History, Plus, Trash2, Save, AlertTriangle, Droplets, TrendingUp, X, Download } from "lucide-react";
import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient.js";

const MODELS = [
  { key: "yarisCR", label: "Yaris Cross" },
  { key: "yaris", label: "Yaris" },
  { key: "calya", label: "Calya" },
  { key: "avanza", label: "Avanza" },
];

const LINESTOP_FIELDS = [
  { key: "process", label: "Proses Sealer" },
  { key: "bodyShort", label: "Body Short" },
  { key: "bodyFull", label: "Body Full" },
  { key: "equipment", label: "Equipment" },
  { key: "berencana", label: "Berencana" },
];

const DEFECT_FIELDS = ["KOS", "STR", "SBB", "SBO", "STM", "SMP", "SNF", "OTHERS"];

const MATERIALS = [
  { key: "sundine2650", label: "Sundine 2650", std: 0.6315 },
  { key: "sundine2660", label: "Sundine 2660", std: 2.1207 },
  { key: "protagePV200", label: "Protage PV-200", std: 1.215 },
  { key: "sealerGrey", label: "Sealer Grey", std: 0.131 },
  { key: "pvcPlastisol", label: "PVC Plastisol", std: 1.33 },
  { key: "majunHygiet", label: "Majun Hygiet", std: 0.0992 },
  { key: "majunKaos", label: "Majun Kaos", std: 0.0015 },
  { key: "sarungTanganComet", label: "Sarung Tangan Comet", std: 0.0244 },
];

function perModelZero() {
  return Object.fromEntries(MODELS.map((m) => [m.key, 0]));
}

const DEFECT_AREAS = [
  {
    key: "inProsesSealerInspeksi",
    label: "In Proses Sealer Inspeksi",
    perModel: [
      { key: "kos", label: "KOS" },
      { key: "sbo", label: "SBO" },
      { key: "str", label: "STR" },
      { key: "silincerNG", label: "Silincer NG" },
      { key: "osPvc", label: "O/S PVC" },
    ],
    flat: [
      { key: "smp", label: "SMP" },
      { key: "strTambahan", label: "STR (tambahan)" },
      { key: "stm", label: "STM" },
    ],
  },
  {
    key: "sealerOffline",
    label: "Sealer Offline",
    perModel: [
      { key: "sbb", label: "SBB" },
      { key: "str", label: "STR" },
      { key: "smp", label: "SMP" },
    ],
    flat: [],
  },
  {
    key: "tcOffline",
    label: "T/C Offline",
    perModel: [
      { key: "kos", label: "KOS" },
      { key: "str", label: "STR" },
    ],
    flat: [
      { key: "sbb", label: "SBB" },
      { key: "sbo", label: "SBO" },
      { key: "stm", label: "STM" },
      { key: "smp", label: "SMP" },
    ],
  },
  {
    key: "outFlowPrimer",
    label: "Out Flow Primer",
    perModel: [{ key: "kos", label: "KOS" }],
    flat: [],
  },
  {
    key: "titipanBOut",
    label: "Titipan B/Out",
    perModel: [{ key: "kos", label: "KOS" }],
    flat: [],
  },
  {
    key: "chousaGoumi",
    label: "Chousa Goumi",
    perModel: [{ key: "waterLeakage", label: "Water Leakage" }],
    flat: [],
  },
];

function emptyDefectDetail() {
  const out = {};
  DEFECT_AREAS.forEach((area) => {
    out[area.key] = {};
    area.perModel.forEach((f) => (out[area.key][f.key] = perModelZero()));
    area.flat.forEach((f) => (out[area.key][f.key] = 0));
  });
  return out;
}

function emptyRecord(date) {
  return {
    id: date,
    date,
    unit: {
      planning: 0,
      actual: { yarisCR: 0, yaris: 0, calya: 0, avanza: 0 },
      reject: { yarisCR: 0, yaris: 0, calya: 0, avanza: 0 },
    },
    lineStop: { process: 0, bodyShort: 0, bodyFull: 0, equipment: 0, berencana: 0 },
    lineStopUBC: { process: 0, bodyShort: 0, bodyFull: 0, equipmentBarcode: 0, equipmentLC: 0, berencana: 0 },
    scheduleCCR: { shift: 0, planOtCcr: 0, jamKerjaOt: 0 },
    workingHour: { planning: 0, actual: 0 },
    ot: { planning: 0, actual: 0 },
    efficiency: { mh: 0, c3mh: 0 },
    defect: Object.fromEntries(DEFECT_FIELDS.map((d) => [d, 0])),
    defectDetail: emptyDefectDetail(),
    material: Object.fromEntries(
      MATERIALS.map((m) => [
        m.key,
        {
          stockAwal: 0,
          bon: 0,
          stokAkhir: 0,
          ratio: Object.fromEntries(MODELS.map((mo) => [mo.key, 0])),
        },
      ])
    ),
  };
}

function sum(obj) {
  return Object.values(obj || {}).reduce((a, b) => a + (Number(b) || 0), 0);
}

function fmt(n, d = 0) {
  const v = Number(n) || 0;
  return v.toLocaleString("id-ID", { minimumFractionDigits: d, maximumFractionDigits: d });
}

const num = (v) => (v === "" || v === null || v === undefined ? 0 : Number(v));

const TABLE = "sealer_records";

function useRecords() {
  const [records, setRecords] = useState({});
  const [loaded, setLoaded] = useState(false);
  const [error, setError] = useState(null);

  const fetchAll = useCallback(async () => {
    const { data, error: err } = await supabase.from(TABLE).select("date, data");
    if (err) {
      setError("Gagal memuat data: " + err.message);
      return;
    }
    const next = {};
    (data || []).forEach((row) => {
      next[row.date] = row.data;
    });
    setRecords(next);
  }, []);

  useEffect(() => {
    (async () => {
      await fetchAll();
      setLoaded(true);
    })();

    const channel = supabase
      .channel("sealer_records_changes")
      .on("postgres_changes", { event: "*", schema: "public", table: TABLE }, () => {
        fetchAll();
      })
      .subscribe();

    return () => {
      supabase.removeChannel(channel);
    };
  }, [fetchAll]);

  const saveRecord = useCallback(
    async (rec) => {
      setRecords((prev) => ({ ...prev, [rec.date]: rec }));
      const { error: err } = await supabase
        .from(TABLE)
        .upsert({ date: rec.date, data: rec, updated_at: new Date().toISOString() });
      if (err) setError("Gagal menyimpan data: " + err.message);
    },
    []
  );

  const deleteRecord = useCallback(async (date) => {
    setRecords((prev) => {
      const next = { ...prev };
      delete next[date];
      return next;
    });
    const { error: err } = await supabase.from(TABLE).delete().eq("date", date);
    if (err) setError("Gagal menghapus data: " + err.message);
  }, []);

  return { records, loaded, error, saveRecord, deleteRecord };
}

// ---------- shared UI atoms ----------

function HazardBar() {
  return (
    <div
      style={{
        height: 6,
        backgroundImage:
          "repeating-linear-gradient(135deg, #F5A623 0 14px, #14181C 14px 28px)",
      }}
    />
  );
}

function Panel({ title, icon: Icon, children, right }) {
  return (
    <div style={{
      background: "#1D2329", border: "1px solid #2A323A", borderRadius: 4,
      padding: "16px 18px", marginBottom: 16,
    }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          {Icon && <Icon size={16} color="#F5A623" />}

          <span style={{
            fontFamily: "'Oswald', sans-serif", fontSize: 13, letterSpacing: "0.08em",
            textTransform: "uppercase", color: "#8B95A1", fontWeight: 500,
          }}>{title}</span>
        </div>
        {right}
      </div>
      {children}
    </div>
  );
}

function Stat({ label, value, unit, accent }) {
  return (
    <div style={{ flex: 1, minWidth: 130 }}>
      <div style={{ fontFamily: "'Oswald', sans-serif", fontSize: 11, letterSpacing: "0.06em", textTransform: "uppercase", color: "#6B7480" }}>{label}</div>
      <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 26, fontWeight: 600, color: accent || "#E7EAEE", marginTop: 2 }}>
        {value}<span style={{ fontSize: 13, color: "#6B7480", marginLeft: 4 }}>{unit}</span>
      </div>
    </div>
  );
}

function NumField({ label, value, onChange, small }) {
  return (
    <label style={{ display: "block", flex: small ? "0 0 100px" : 1 }}>
      <div style={{ fontSize: 11, color: "#8B95A1", marginBottom: 4, fontFamily: "'Inter', sans-serif" }}>{label}</div>
      <input
        type="number"
        value={value}
        onChange={(e) => onChange(e.target.value)}
        style={{
          width: "100%", background: "#14181C", border: "1px solid #2A323A", borderRadius: 3,
          color: "#E7EAEE", padding: "7px 9px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 13,
          outline: "none",
        }}
        onFocus={(e) => (e.target.style.borderColor = "#F5A623")}
        onBlur={(e) => (e.target.style.borderColor = "#2A323A")}
      />
    </label>
  );
}

// ---------- Dashboard ----------

function Dashboard({ records }) {
  const [month, setMonth] = useState("ALL");
  const [material, setMaterial] = useState(MATERIALS[0].key);

  const sorted = useMemo(
    () => Object.values(records).sort((a, b) => a.date.localeCompare(b.date)),
    [records]
  );

  const filtered = useMemo(
    () => (month === "ALL" ? sorted : sorted.filter((r) => r.date.slice(0, 7) === month)),
    [sorted, month]
  );

  const months = useMemo(() => {
    const s = new Set(sorted.map((r) => r.date.slice(0, 7)));
    return Array.from(s).sort();
  }, [sorted]);

  const chartData = filtered.map((r) => {
    const actual = sum(r.unit.actual);
    const reject = sum(r.unit.reject);
    const ok = actual > 0 ? ((actual - reject) / actual) * 100 : 0;
    const mat = r.material[material] || { stockAwal: 0, bon: 0, stokAkhir: 0 };
    const pemakaian = num(mat.stockAwal) + num(mat.bon) - num(mat.stokAkhir);
    return {
      date: r.date.slice(5),
      Planning: num(r.unit.planning),
      Actual: actual,
      Reject: reject,
      OKRatio: Number(ok.toFixed(1)),
      "Line Stop": sum(r.lineStop),
      "MH/U": actual > 0 ? Number((num(r.efficiency.mh) / actual).toFixed(2)) : 0,
      "Stok Akhir": num(mat.stokAkhir),
      "Pemakaian/Unit": actual > 0 ? Number((pemakaian / actual).toFixed(4)) : 0,
    };
  });

  const totals = filtered.reduce(
    (acc, r) => {
      acc.planning += num(r.unit.planning);
      acc.actual += sum(r.unit.actual);
      acc.reject += sum(r.unit.reject);
      acc.lineStop += sum(r.lineStop);
      acc.lineStopUBC += sum(r.lineStopUBC);
      acc.defectGrandTotal +=
        sum(r.defect) +
        DEFECT_AREAS.reduce((s, area) => {
          const d = r.defectDetail?.[area.key] || {};
          let areaSum = 0;
          area.perModel.forEach((f) => (areaSum += sum(d[f.key])));
          area.flat.forEach((f) => (areaSum += num(d[f.key])));
          return s + areaSum;
        }, 0);
      return acc;
    },
    { planning: 0, actual: 0, reject: 0, lineStop: 0, lineStopUBC: 0, defectGrandTotal: 0 }
  );
  const okRatio = totals.actual > 0 ? ((totals.actual - totals.reject) / totals.actual) * 100 : 0;

  const lineStopBreakdown = [
    { name: "Process", Sealer: filtered.reduce((a, r) => a + num(r.lineStop.process), 0), UBC: filtered.reduce((a, r) => a + num(r.lineStopUBC.process), 0) },
    { name: "Body Short", Sealer: filtered.reduce((a, r) => a + num(r.lineStop.bodyShort), 0), UBC: filtered.reduce((a, r) => a + num(r.lineStopUBC.bodyShort), 0) },
    { name: "Body Full", Sealer: filtered.reduce((a, r) => a + num(r.lineStop.bodyFull), 0), UBC: filtered.reduce((a, r) => a + num(r.lineStopUBC.bodyFull), 0) },
    { name: "Equipment", Sealer: filtered.reduce((a, r) => a + num(r.lineStop.equipment), 0), UBC: filtered.reduce((a, r) => a + num(r.lineStopUBC.equipmentBarcode) + num(r.lineStopUBC.equipmentLC), 0) },
    { name: "Berencana", Sealer: filtered.reduce((a, r) => a + num(r.lineStop.berencana), 0), UBC: filtered.reduce((a, r) => a + num(r.lineStopUBC.berencana), 0) },
  ];

  if (sorted.length === 0) {
    return (
      <div style={{ textAlign: "center", padding: "80px 20px", color: "#6B7480" }}>
        <Gauge size={40} color="#3A4148" style={{ marginBottom: 12 }} />
        <div style={{ fontFamily: "'Oswald', sans-serif", fontSize: 16, color: "#8B95A1" }}>Belum ada data</div>
        <div style={{ fontSize: 13, marginTop: 6 }}>Isi data harian di tab "Input Data" untuk melihat dashboard.</div>
      </div>
    );
  }

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 16 }}>
        <select
          value={month}
          onChange={(e) => setMonth(e.target.value)}
          style={{
            background: "#1D2329", border: "1px solid #2A323A", color: "#E7EAEE",
            borderRadius: 3, padding: "6px 10px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12,
          }}
        >
          <option value="ALL">Semua Periode</option>
          {months.map((m) => (
            <option key={m} value={m}>{m}</option>
          ))}
        </select>
      </div>

      <div style={{ display: "flex", gap: 20, flexWrap: "wrap", marginBottom: 16 }}>
        <Panel title="Ringkasan" icon={Gauge} right={null}>
          <div style={{ display: "flex", gap: 24, flexWrap: "wrap" }}>
            <Stat label="Unit Planning" value={fmt(totals.planning)} unit="unit" />
            <Stat label="Unit Actual" value={fmt(totals.actual)} unit="unit" accent="#F5A623" />
            <Stat label="Total Reject" value={fmt(totals.reject)} unit="unit" accent="#E8543E" />
            <Stat label="OK Ratio" value={fmt(okRatio, 1)} unit="%" accent={okRatio >= 95 ? "#4CAF7D" : "#E8543E"} />
            <Stat label="Total Line Stop Sealer" value={fmt(totals.lineStop)} unit="menit" />
            <Stat label="Total Line Stop UBC" value={fmt(totals.lineStopUBC)} unit="menit" />
            <Stat label="Total Defect (Grand)" value={fmt(totals.defectGrandTotal)} unit="pcs" accent="#E8543E" />
          </div>
        </Panel>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
        <Panel title="Unit: Planning vs Actual" icon={TrendingUp}>
          <ResponsiveContainer width="100%" height={220}>
            <LineChart data={chartData}>
              <CartesianGrid stroke="#2A323A" strokeDasharray="3 3" />
              <XAxis dataKey="date" stroke="#6B7480" fontSize={11} />
              <YAxis stroke="#6B7480" fontSize={11} />
              <Tooltip contentStyle={{ background: "#1D2329", border: "1px solid #2A323A", fontSize: 12 }} />
              <Legend wrapperStyle={{ fontSize: 11 }} />
              <Line type="monotone" dataKey="Planning" stroke="#8B95A1" strokeWidth={2} dot={false} />
              <Line type="monotone" dataKey="Actual" stroke="#F5A623" strokeWidth={2} dot={false} />
            </LineChart>
          </ResponsiveContainer>
        </Panel>

        <Panel title="OK Ratio Harian (%)" icon={Gauge}>
          <ResponsiveContainer width="100%" height={220}>
            <AreaChart data={chartData}>
              <CartesianGrid stroke="#2A323A" strokeDasharray="3 3" />
              <XAxis dataKey="date" stroke="#6B7480" fontSize={11} />
              <YAxis stroke="#6B7480" fontSize={11} domain={[0, 100]} />
              <Tooltip contentStyle={{ background: "#1D2329", border: "1px solid #2A323A", fontSize: 12 }} />
              <Area type="monotone" dataKey="OKRatio" stroke="#4CAF7D" fill="#4CAF7D33" strokeWidth={2} />
            </AreaChart>
          </ResponsiveContainer>
        </Panel>

        <Panel title="Line Stop: Sealer vs UBC (menit)" icon={AlertTriangle}>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={lineStopBreakdown}>
              <CartesianGrid stroke="#2A323A" strokeDasharray="3 3" />
              <XAxis dataKey="name" stroke="#6B7480" fontSize={10} interval={0} angle={-15} textAnchor="end" height={50} />
              <YAxis stroke="#6B7480" fontSize={11} />
              <Tooltip contentStyle={{ background: "#1D2329", border: "1px solid #2A323A", fontSize: 12 }} />
              <Legend wrapperStyle={{ fontSize: 11 }} />
              <Bar dataKey="Sealer" fill="#E8543E" radius={[3, 3, 0, 0]} />
              <Bar dataKey="UBC" fill="#F5A623" radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </Panel>

        <Panel
          title="Stok & Pemakaian Material"
          icon={Droplets}
          right={
            <select
              value={material}
              onChange={(e) => setMaterial(e.target.value)}
              style={{
                background: "#14181C", border: "1px solid #2A323A", color: "#E7EAEE",
                borderRadius: 3, padding: "5px 8px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 11,
              }}
            >
              {MATERIALS.map((m) => (
                <option key={m.key} value={m.key}>{m.label}</option>
              ))}
            </select>
          }
        >
          <div style={{ fontSize: 11, color: "#6B7480", marginBottom: 8, fontFamily: "'IBM Plex Mono', monospace" }}>
            STD Pemakaian/Unit: <b style={{ color: "#F5A623" }}>{MATERIALS.find((m) => m.key === material)?.std}</b> kg/unit
          </div>
          <ResponsiveContainer width="100%" height={220}>
            <LineChart data={chartData}>
              <CartesianGrid stroke="#2A323A" strokeDasharray="3 3" />
              <XAxis dataKey="date" stroke="#6B7480" fontSize={11} />
              <YAxis yAxisId="left" stroke="#6B7480" fontSize={11} />
              <YAxis yAxisId="right" orientation="right" stroke="#6B7480" fontSize={11} />
              <Tooltip contentStyle={{ background: "#1D2329", border: "1px solid #2A323A", fontSize: 12 }} />
              <Legend wrapperStyle={{ fontSize: 11 }} />
              <Line yAxisId="left" type="monotone" dataKey="Stok Akhir" stroke="#F5A623" strokeWidth={2} dot={false} />
              <Line yAxisId="right" type="monotone" dataKey="Pemakaian/Unit" stroke="#5B8DEF" strokeWidth={2} dot={false} />
            </LineChart>
          </ResponsiveContainer>
        </Panel>
      </div>
    </div>
  );
}

// ---------- Input Form ----------

function InputForm({ records, onSave }) {
  const today = new Date().toISOString().slice(0, 10);
  const [date, setDate] = useState(today);
  const [rec, setRec] = useState(records[today] || emptyRecord(today));

const [rec, setRec] = useState(records[today] || emptyRecord(today));
  const [savedFlash, setSavedFlash] = useState(false);

  useEffect(() => {
    setRec(records[date] || emptyRecord(date));
  }, [date]); // eslint-disable-line

  const set = (path, value) => {
    setRec((prev) => {
      const next = JSON.parse(JSON.stringify(prev));
      let cur = next;
      const keys = path.split(".");
      for (let i = 0; i < keys.length - 1; i++) {
        if (cur[keys[i]] === undefined || cur[keys[i]] === null) cur[keys[i]] = {};
        cur = cur[keys[i]];
      }
      cur[keys[keys.length - 1]] = num(value);
      return next;
    });
  };

  const handleSave = () => {
    onSave({ ...rec, date, id: date });
    setSavedFlash(true);
    setTimeout(() => setSavedFlash(false), 1500);
  };

  const totalActual = sum(rec.unit.actual);
  const totalReject = sum(rec.unit.reject);

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 16 }}>
        <label style={{ fontFamily: "'Oswald', sans-serif", fontSize: 12, color: "#8B95A1", textTransform: "uppercase" }}>Tanggal</label>
        <input
          type="date"
          value={date}
          onChange={(e) => setDate(e.target.value)}
          style={{
            background: "#1D2329", border: "1px solid #2A323A", color: "#E7EAEE",
            borderRadius: 3, padding: "7px 10px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 13,
          }}
        />
        {records[date] && (
          <span style={{ fontSize: 11, color: "#F5A623", fontFamily: "'Inter', sans-serif" }}>● data sudah ada — akan ditimpa saat disimpan</span>
        )}
      </div>

      <Panel title="Unit Produksi" icon={Factory}>
        <div style={{ marginBottom: 12 }}>
          <NumField label="Planning (total unit)" value={rec.unit.planning} onChange={(v) => set("unit.planning", v)} />
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 10, marginBottom: 6 }}>
          {MODELS.map((m) => (
            <NumField key={m.key} label={`Actual — ${m.label}`} value={rec.unit.actual[m.key]} onChange={(v) => set(`unit.actual.${m.key}`, v)} />
          ))}
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 10 }}>
          {MODELS.map((m) => (
            <NumField key={m.key} label={`Reject — ${m.label}`} value={rec.unit.reject[m.key]} onChange={(v) => set(`unit.reject.${m.key}`, v)} />
          ))}
        </div>
        <div style={{ marginTop: 10, fontSize: 12, color: "#8B95A1", fontFamily: "'IBM Plex Mono', monospace" }}>
          Total Actual: <b style={{ color: "#F5A623" }}>{totalActual}</b> &nbsp;|&nbsp;
          Total Reject: <b style={{ color: "#E8543E" }}>{totalReject}</b> &nbsp;|&nbsp;
          OK Ratio: <b style={{ color: "#4CAF7D" }}>{totalActual > 0 ? fmt(((totalActual - totalReject) / totalActual) * 100, 1) : 0}%</b>
        </div>
      </Panel>

      <Panel title="Line Stop Sealer (menit)" icon={AlertTriangle}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr 1fr", gap: 10 }}>
          {LINESTOP_FIELDS.map((f) => (
            <NumField key={f.key} label={f.label} value={rec.lineStop[f.key]} onChange={(v) => set(`lineStop.${f.key}`, v)} />
          ))}
        </div>
      </Panel>

      <Panel title="Line Stop UBC (menit)" icon={AlertTriangle}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr 1fr 1fr", gap: 10 }}>
          <NumField label="Process" value={rec.lineStopUBC.process} onChange={(v) => set("lineStopUBC.process", v)} />
          <NumField label="Body Short" value={rec.lineStopUBC.bodyShort} onChange={(v) => set("lineStopUBC.bodyShort", v)} />
          <NumField label="Body Full" value={rec.lineStopUBC.bodyFull} onChange={(v) => set("lineStopUBC.bodyFull", v)} />
          <NumField label="Equipment (barcode error)" value={rec.lineStopUBC.equipmentBarcode} onChange={(v) => set("lineStopUBC.equipmentBarcode", v)} />
          <NumField label="Equipment (LC)" value={rec.lineStopUBC.equipmentLC} onChange={(v) => set("lineStopUBC.equipmentLC", v)} />
          <NumField label="Berencana" value={rec.lineStopUBC.berencana} onChange={(v) => set("lineStopUBC.berencana", v)} />
        </div>
      </Panel>

      <Panel title="Jadwal & OT (CCR)" icon={Gauge}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
          <NumField label="Shift" value={rec.scheduleCCR.shift} onChange={(v) => set("scheduleCCR.shift", v)} />
          <NumField label="Plan OT CCR (jam)" value={rec.scheduleCCR.planOtCcr} onChange={(v) => set("scheduleCCR.planOtCcr", v)} />
          <NumField label="Jam Kerja OT (jam)" value={rec.scheduleCCR.jamKerjaOt} onChange={(v) => set("scheduleCCR.jamKerjaOt", v)} />
        </div>
      </Panel>

      <Panel title="Jam Kerja, Lembur & Efisiensi" icon={Gauge}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr 1fr 1fr", gap: 10 }}>
          <NumField label="WT Planning (jam)" value={rec.workingHour.planning} onChange={(v) => set("workingHour.planning", v)} />
          <NumField label="WT Actual (jam)" value={rec.workingHour.actual} onChange={(v) => set("workingHour.actual", v)} />
          <NumField label="OT Planning (jam)" value={rec.ot.planning} onChange={(v) => set("ot.planning", v)} />
          <NumField label="OT Actual (jam)" value={rec.ot.actual} onChange={(v) => set("ot.actual", v)} />
          <NumField label="MH" value={rec.efficiency.mh} onChange={(v) => set("efficiency.mh", v)} />
          <NumField label="C3 MH" value={rec.efficiency.c3mh} onChange={(v) => set("efficiency.c3mh", v)} />
        </div>
      </Panel>

      <Panel title="Defect Quality (Q/G) — Total" icon={AlertTriangle}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 10 }}>
          {DEFECT_FIELDS.map((d) => (
            <NumField key={d} label={d} value={rec.defect[d]} onChange={(v) => set(`defect.${d}`, v)} />
          ))}
        </div>
      </Panel>

      <Panel title="Defect Detail per Proses & Model" icon={AlertTriangle}>
        {DEFECT_AREAS.map((area) => {
          const areaData = rec.defectDetail[area.key] || {};
          return (
            <div key={area.key} style={{ marginBottom: 16, paddingBottom: 16, borderBottom: "1px solid #2A323A" }}>
              <div style={{ fontSize: 12, color: "#F5A623", marginBottom: 8, fontFamily: "'Oswald', sans-serif", textTransform: "uppercase", letterSpacing: "0.05em" }}>
                {area.label}
              </div>
              {area.perModel.map((f) => {
                const vals = areaData[f.key] || {};
                const total = sum(vals);
                return (
                  <div key={f.key} style={{ marginBottom: 10 }}>
                    <div style={{ fontSize: 11, color: "#8B95A1", marginBottom: 4 }}>
                      {f.label} <span style={{ color: "#6B7480" }}>(total: {total})</span>
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 10 }}>
                      {MODELS.map((mo) => (
                        <NumField
                          key={mo.key}
                          label={mo.label}
                          value={vals[mo.key] ?? 0}
                          onChange={(v) => set(`defectDetail.${area.key}.${f.key}.${mo.key}`, v)}
                        />
                      ))}
                    </div>
                  </div>
                );
              })}
              {area.flat.length > 0 && (
                <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 10 }}>
                  {area.flat.map((f) => (
                    <NumField
                      key={f.key}
                      label={f.label}
                      value={areaData[f.key] ?? 0}
                      onChange={(v) => set(`defectDetail.${area.key}.${f.key}`, v)}
                    />
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </Panel>

      <Panel title="Stok Material (KG)" icon={Droplets}>
        {MATERIALS.map((m) => {
          const mat = rec.material[m.key] || { stockAwal: 0, bon: 0, stokAkhir: 0, ratio: {} };
          const ratio = mat.ratio || {};
          const pemakaian = num(mat.stockAwal) + num(mat.bon) - num(mat.stokAkhir);
          const pemakaianUnit = totalActual > 0 ? pemakaian / totalActual : 0;
          const overStd = m.std && pemakaianUnit > m.std;
          return (
            <div key={m.key} style={{ marginBottom: 14, paddingBottom: 14, borderBottom: "1px solid #2A323A" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 8 }}>
                <div style={{ fontSize: 12, color: "#F5A623", fontFamily: "'Oswald', sans-serif", textTransform: "uppercase", letterSpacing: "0.05em" }}>{m.label}</div>
                <div style={{ fontSize: 10, color: "#6B7480", fontFamily: "'IBM Plex Mono', monospace" }}>STD: {m.std} kg/unit</div>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 10, marginBottom: 10 }}>
                <NumField label="Stok Awal" value={mat.stockAwal} onChange={(v) => set(`material.${m.key}.stockAwal`, v)} />
                <NumField label="Bon" value={mat.bon} onChange={(v) => set(`material.${m.key}.bon`, v)} />
                <NumField label="Stok Akhir" value={mat.stokAkhir} onChange={(v) => set(`material.${m.key}.stokAkhir`, v)} />
                <div>
                  <div style={{ fontSize: 11, color: "#8B95A1", marginBottom: 4 }}>Pemakaian (auto)</div>
                  <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: "#E7EAEE", padding: "7px 0" }}>{fmt(pemakaian, 2)} kg</div>
                </div>
              </div>
              <div style={{ fontSize: 10, color: "#6B7480", marginBottom: 6, textTransform: "uppercase", letterSpacing: "0.04em", fontFamily: "'Inter', sans-serif" }}>Rasio pemakaian per model</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 10 }}>
                {MODELS.map((mo) => (
                  <NumField
                    key={mo.key}
                    label={mo.label}
                    value={ratio[mo.key] ?? 0}
                    onChange={(v) => set(`material.${m.key}.ratio.${mo.key}`, v)}
                  />
                ))}
              </div>
              <div style={{ marginTop: 8, fontSize: 11, fontFamily: "'IBM Plex Mono', monospace", color: overStd ? "#E8543E" : "#4CAF7D" }}>
                Pemakaian/Unit: {fmt(pemakaianUnit, 4)} kg/unit {overStd ? "— di atas STD" : "— sesuai/di bawah STD"}
              </div>
            </div>
          );
        })}
      </Panel>

      <button
        onClick={handleSave}
        style={{
          display: "flex", alignItems: "center", gap: 8, background: savedFlash ? "#4CAF7D" : "#F5A623",
          color: "#14181C", border: "none", borderRadius: 3, padding: "10px 22px",
          fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 13, textTransform: "uppercase",
          letterSpacing: "0.04em", cursor: "pointer", marginTop: 4, marginBottom: 30, transition: "background 0.2s",
        }}
      >
        <Save size={15} /> {savedFlash ? "Tersimpan" : "Simpan Data Harian"}
      </button>
    </div>
  );
}

// ---------- History table ----------

function HistoryTable({ records, onDelete }) {
  const [month, setMonth] = useState("ALL");
  const [confirmId, setConfirmId] = useState(null);

  const sorted = useMemo(
    () => Object.values(records).sort((a, b) => b.date.localeCompare(a.date)),
    [records]
  );
  const months = useMemo(() => {
    const s = new Set(sorted.map((r) => r.date.slice(0, 7)));
    return Array.from(s).sort();
  }, [sorted]);
  const filtered = month === "ALL" ? sorted : sorted.filter((r) => r.date.slice(0, 7) === month);

  if (sorted.length === 0) {
    return (
      <div style={{ textAlign: "center", padding: "80px 20px", color: "#6B7480" }}>
        <History size={40} color="#3A4148" style={{ marginBottom: 12 }} />
        <div style={{ fontFamily: "'Oswald', sans-serif", fontSize: 16, color: "#8B95A1" }}>Belum ada riwayat</div>
      </div>
    );
}

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 12 }}>
        <select
          value={month}
          onChange={(e) => setMonth(e.target.value)}
          style={{
            background: "#1D2329", border: "1px solid #2A323A", color: "#E7EAEE",
            borderRadius: 3, padding: "6px 10px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12,
          }}
        >
          <option value="ALL">Semua Periode</option>
          {months.map((m) => <option key={m} value={m}>{m}</option>)}
        </select>
      </div>

      <div style={{ overflowX: "auto", border: "1px solid #2A323A", borderRadius: 4 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }}>
          <thead>
            <tr style={{ background: "#232A31" }}>
              {["Tanggal", "Planning", "Actual", "Reject", "OK Ratio", "Line Stop", "WT Actual", "", ].map((h, i) => (
                <th key={i} style={{ padding: "9px 12px", textAlign: i === 0 ? "left" : "right", color: "#8B95A1", fontWeight: 500, fontFamily: "'Oswald', sans-serif", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.04em" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.map((r) => {
              const actual = sum(r.unit.actual);
              const reject = sum(r.unit.reject);
              const ok = actual > 0 ? ((actual - reject) / actual) * 100 : 0;
              return (
                <tr key={r.date} style={{ borderTop: "1px solid #2A323A" }}>
                  <td style={{ padding: "9px 12px", color: "#E7EAEE" }}>{r.date}</td>
                  <td style={{ padding: "9px 12px", textAlign: "right", color: "#8B95A1" }}>{fmt(r.unit.planning)}</td>
                  <td style={{ padding: "9px 12px", textAlign: "right", color: "#F5A623" }}>{fmt(actual)}</td>
                  <td style={{ padding: "9px 12px", textAlign: "right", color: "#E8543E" }}>{fmt(reject)}</td>
                  <td style={{ padding: "9px 12px", textAlign: "right", color: ok >= 95 ? "#4CAF7D" : "#E8543E" }}>{fmt(ok, 1)}%</td>
                  <td style={{ padding: "9px 12px", textAlign: "right", color: "#8B95A1" }}>{fmt(sum(r.lineStop))}</td>
                  <td style={{ padding: "9px 12px", textAlign: "right", color: "#8B95A1" }}>{fmt(r.workingHour.actual)}</td>
                  <td style={{ padding: "6px 10px", textAlign: "right" }}>
                    {confirmId === r.date ? (
                      <span style={{ display: "inline-flex", gap: 6, alignItems: "center" }}>
                        <button onClick={() => { onDelete(r.date); setConfirmId(null); }} style={{ background: "#E8543E", color: "#fff", border: "none", borderRadius: 3, padding: "4px 8px", fontSize: 11, cursor: "pointer" }}>Hapus</button>
                        <button onClick={() => setConfirmId(null)} style={{ background: "transparent", color: "#8B95A1", border: "1px solid #2A323A", borderRadius: 3, padding: "4px 8px", fontSize: 11, cursor: "pointer" }}><X size={11} /></button>
                      </span>
                    ) : (
                      <button onClick={() => setConfirmId(r.date)} style={{ background: "transparent", color: "#6B7480", border: "none", cursor: "pointer", padding: 4 }}>
                        <Trash2 size={14} />
                      </button>
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ---------- Export to Excel ----------

function exportToExcel(records) {
  const rows = Object.values(records).sort((a, b) => a.date.localeCompare(b.date));
  if (rows.length === 0) return;

  const sheetRows = rows.map((r) => {
    const actual = sum(r.unit.actual);
    const reject = sum(r.unit.reject);
    const ok = actual > 0 ? ((actual - reject) / actual) * 100 : 0;
    const row = {
      Tanggal: r.date,
      "Planning (unit)": num(r.unit.planning),
    };
    MODELS.forEach((m) => (row[`Actual ${m.label}`] = num(r.unit.actual[m.key])));
    row["Total Actual"] = actual;
    MODELS.forEach((m) => (row[`Reject ${m.label}`] = num(r.unit.reject[m.key])));
    row["Total Reject"] = reject;
    row["OK Ratio (%)"] = Number(ok.toFixed(2));
    LINESTOP_FIELDS.forEach((f) => (row[`Line Stop Sealer - ${f.label}`] = num(r.lineStop[f.key])));
    row["Total Line Stop Sealer"] = sum(r.lineStop);
    const ubc = r.lineStopUBC || {};
    row["Line Stop UBC - Process"] = num(ubc.process);
    row["Line Stop UBC - Body Short"] = num(ubc.bodyShort);
    row["Line Stop UBC - Body Full"] = num(ubc.bodyFull);
    row["Line Stop UBC - Equipment (barcode)"] = num(ubc.equipmentBarcode);
    row["Line Stop UBC - Equipment (LC)"] = num(ubc.equipmentLC);
    row["Line Stop UBC - Berencana"] = num(ubc.berencana);
    row["Total Line Stop UBC"] = sum(ubc);
    const ccr = r.scheduleCCR || {};
    row["Shift"] = num(ccr.shift);
    row["Plan OT CCR"] = num(ccr.planOtCcr);
    row["Jam Kerja OT"] = num(ccr.jamKerjaOt);
    row["WT Planning"] = num(r.workingHour.planning);
    row["WT Actual"] = num(r.workingHour.actual);
    row["OT Planning"] = num(r.ot.planning);
    row["OT Actual"] = num(r.ot.actual);
    row["MH"] = num(r.efficiency.mh);
    row["MH/U"] = actual > 0 ? Number((num(r.efficiency.mh) / actual).toFixed(3)) : 0;
    row["C3 MH"] = num(r.efficiency.c3mh);
    DEFECT_FIELDS.forEach((d) => (row[`Defect ${d}`] = num(r.defect[d])));
    DEFECT_AREAS.forEach((area) => {
      const areaData = (r.defectDetail && r.defectDetail[area.key]) || {};
      area.perModel.forEach((f) => {
        const vals = areaData[f.key] || {};
        MODELS.forEach((mo) => {
          row[`${area.label} - ${f.label} - ${mo.label}`] = num(vals[mo.key]);
        });
        row[`${area.label} - ${f.label} - Total`] = sum(vals);
      });
      area.flat.forEach((f) => {
        row[`${area.label} - ${f.label}`] = num(areaData[f.key]);
      });
    });
    MATERIALS.forEach((m) => {
      const mat = r.material[m.key] || { stockAwal: 0, bon: 0, stokAkhir: 0, ratio: {} };
      const pemakaian = num(mat.stockAwal) + num(mat.bon) - num(mat.stokAkhir);
      row[`${m.label} - Stok Awal`] = num(mat.stockAwal);
      row[`${m.label} - Bon`] = num(mat.bon);
      row[`${m.label} - Stok Akhir`] = num(mat.stokAkhir);
      row[`${m.label} - Pemakaian`] = Number(pemakaian.toFixed(2));
      row[`${m.label} - Pemakaian/Unit`] = actual > 0 ? Number((pemakaian / actual).toFixed(4)) : 0;
      row[`${m.label} - STD`] = m.std;
      MODELS.forEach((mo) => {
        row[`${m.label} - Rasio ${mo.label}`] = num((mat.ratio || {})[mo.key]);
      });
    });
    return row;
  });

  const ws = XLSX.utils.json_to_sheet(sheetRows);
  ws["!cols"] = Object.keys(sheetRows[0]).map((k) => ({ wch: Math.max(12, k.length + 2) }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Data Base Sealer");
  const stamp = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `Data_Base_Sealer_Export_${stamp}.xlsx`);
}

// ---------- App shell ----------

export default function App() {
  const { records, loaded, error, saveRecord, deleteRecord } = useRecords();
  const [tab, setTab] = useState("dashboard");

  const tabs = [
    { key: "dashboard", label: "Dashboard", icon: Gauge },
    { key: "input", label: "Input Data", icon: Plus },
    { key: "history", label: "Riwayat", icon: History },
  ];

  return (
    <div style={{ minHeight: "100vh", background: "#14181C", color: "#E7EAEE", fontFamily: "'Inter', sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Oswald:wght@400;500;600;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
        input[type="date"]::-webkit-calendar-picker-indicator { filter: invert(0.7); }
        ::-webkit-scrollbar { height: 8px; width: 8px; }
        ::-webkit-scrollbar-thumb { background: #2A323A; border-radius: 4px; }
      `}</style>

      <div style={{ borderBottom: "1px solid #2A323A" }}>
        <div style={{ maxWidth: 1100, margin: "0 auto", padding: "18px 20px 0" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16 }}>
            <Factory size={22} color="#F5A623" />
            <div>
              <div style={{ fontFamily: "'Oswald', sans-serif", fontSize: 18, fontWeight: 600, letterSpacing: "0.02em" }}>DATA BASE SEALER — RED SHIFT</div>
              <div style={{ fontSize: 11, color: "#6B7480" }}>Monitoring produksi harian · 2025</div>
            </div>
            <div style={{ flex: 1 }} />
            <button
              onClick={() => exportToExcel(records)}
              disabled={Object.keys(records).length === 0}
              style={{
                display: "flex", alignItems: "center", gap: 7, background: "transparent",
                border: "1px solid #2A323A", color: Object.keys(records).length === 0 ? "#4A525A" : "#E7EAEE",
                borderRadius: 3, padding: "8px 14px", fontFamily: "'Oswald', sans-serif", fontSize: 12,
                textTransform: "uppercase", letterSpacing: "0.04em",
                cursor: Object.keys(records).length === 0 ? "not-allowed" : "pointer", marginBottom: 16,
              }}
            >
              <Download size={14} /> Export ke Excel
            </button>
          </div>
          <div style={{ display: "flex", gap: 4 }}>
            {tabs.map((t) => {
              const active = tab === t.key;
              const Icon = t.icon;
              return (
                <button
                  key={t.key}
                  onClick={() => setTab(t.key)}
                  style={{
                    display: "flex", alignItems: "center", gap: 6,
                    background: "transparent", border: "none", cursor: "pointer",
                    padding: "9px 16px", fontFamily: "'Oswald', sans-serif", fontSize: 13,
                    textTransform: "uppercase", letterSpacing: "0.04em",
                    color: active ? "#F5A623" : "#8B95A1",
                    borderBottom: active ? "2px solid #F5A623" : "2px solid transparent",
                  }}
                >
                  <Icon size={14} /> {t.label}
                </button>
              );
            })}
          </div>
        </div>
      </div>
      <HazardBar />

      <div style={{ maxWidth: 1100, margin: "0 auto", padding: "22px 20px 60px" }}>
        {!loaded ? (
          <div style={{ padding: 60, textAlign: "center", color: "#6B7480", fontFamily: "'IBM Plex Mono', monospace" }}>Memuat data…</div>
        ) : (
          <>
            {error && (
              <div style={{ background: "#E8543E22", border: "1px solid #E8543E", color: "#E8543E", padding: "8px 12px", borderRadius: 3, fontSize: 12, marginBottom: 14 }}>
                {error}
              </div>
            )}
            {tab === "dashboard" && <Dashboard records={records} />}
            {tab === "input" && <InputForm records={records} onSave={saveRecord} />}
            {tab === "history" && <HistoryTable records={records} onDelete={deleteRecord} />}
          </>
        )}
      </div>
    </div>
  );
  }
    
  }
