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
