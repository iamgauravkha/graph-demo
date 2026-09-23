import React, { useMemo, useState } from "react";
import {
  ComposedChart,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  ResponsiveContainer,
  Label,
} from "recharts";
import * as XLSX from "xlsx";

/* ================================================================== */
/*  Colours taken from the Kohmer datasheet                            */
/* ================================================================== */
const C = {
  rpm: "#1b9a94",
  rpmCalc: "#6fc3bd",
  power: "#f0707f",
  eff: "#6f8f55",
  fei: "#1f5a2e",
  sound: "#5a2ca0",
  range: "#e5141f",
  grid: "#e3e5e8",
  ink: "#1d1d1f",
  navy: "#1f3a6e",
  specBg: "#e4e7ea",
  frame: "#9fb7b3",
};
const FONT = 'Carlito, Calibri, "Segoe UI", Arial, sans-serif';

/* ================================================================== */
/*  Hard-coded data: KESO-315 (from Keso-315_Sample_Datasheet.xlsx)    */
/*  row = [airflow m3/h, static Pa, power kW, total eff %, FEI, dBA]   */
/* ================================================================== */
const RAW = {
  2880: [
    [3387.6, 0, 0.4, 8, 0.565, 94.73],
    [3096, 241, 0.469, 49.4, 1.291, null],
    [2869.2, 442, 0.558, 66.5, 1.571, 90.91],
    [2530.8, 734, 0.685, 77.1, 1.745, null],
    [2203.2, 923, 0.713, 80.1, 1.819, 88.57],
    [1832.4, 1062, 0.694, 78.2, 1.824, null],
    [1522.8, 1118, 0.645, 73.4, 1.779, 90.59],
    [993.6, 1108, 0.508, 59.9, 1.643, null],
    [396, 1082, 0.385, 30.9, 1.278, null],
    [0, 1141, 0.348, 0, 0.831, null],
  ],
  3000: [
    [3528.75, 0, 0.4521, 8, 0.565, 95.62],
    [3225, 261.5, 0.5301, 49.4, 1.291, null],
    [2988.75, 479.6, 0.6307, 66.5, 1.571, 91.79],
    [2636.25, 796.44, 0.7742, 77.1, 1.745, null],
    [2295, 1001.52, 0.8059, 80.1, 1.819, 89.45],
    [1908.75, 1152.34, 0.7844, 78.2, 1.824, null],
    [1586.25, 1213.11, 0.729, 73.4, 1.779, 91.48],
    [1035, 1202.26, 0.5742, 59.9, 1.643, null],
    [412.5, 1174.05, 0.4352, 30.9, 1.278, null],
    [0, 1238.06, 0.3933, 0, 0.831, null],
  ],
  3200: [
    [3764, 0, 0.5487, 8, 0.565, 97.02],
    [3440, 297.53, 0.6433, 49.4, 1.291, null],
    [3188, 545.68, 0.7654, 66.5, 1.571, 93.2],
    [2812, 906.17, 0.9396, 77.1, 1.745, null],
    [2448, 1139.51, 0.9781, 80.1, 1.819, 90.85],
    [2036, 1311.11, 0.952, 78.2, 1.824, null],
    [1692, 1380.25, 0.8848, 73.4, 1.779, 92.88],
    [1104, 1367.9, 0.6968, 59.9, 1.643, null],
    [440, 1335.8, 0.5281, 30.9, 1.278, null],
    [0, 1408.64, 0.4774, 0, 0.831, null],
  ],
  3450: [
    [4058.06, 0, 0.6876, 8, 0.565, 98.65],
    [3708.75, 345.84, 0.8062, 49.4, 1.291, null],
    [3437.06, 634.27, 0.9592, 66.5, 1.571, 94.83],
    [3031.69, 1053.29, 1.1775, 77.1, 1.745, null],
    [2639.25, 1324.51, 1.2257, 80.1, 1.819, 92.49],
    [2195.06, 1523.97, 1.193, 78.2, 1.824, null],
    [1824.19, 1604.33, 1.1088, 73.4, 1.779, 94.51],
    [1190.25, 1589.98, 0.8733, 59.9, 1.643, null],
    [474.38, 1552.67, 0.6618, 30.9, 1.278, null],
    [0, 1637.34, 0.5982, 0, 0.831, null],
  ],
};
const SAMPLE_CURVES = Object.entries(RAW).map(([speed, rows]) => ({
  speed: Number(speed),
  points: rows.map(([q, p, w, eff, fei, sound], idx) => ({
    idx,
    q,
    p,
    w,
    eff,
    fei,
    sound,
  })),
}));

/* Datasheet text: layout and sample values from the Reso-630 PDF page */
const DEFAULT_SHEET = {
  company: "KUMARAN",
  brand: "KOHMER",
  tagline: "Moving Air Adding Value",
  modelPrefix: "Plug Fan Model: KDPF - ",
  specs: [
    { head: true, label: "Description", value: "Value", unit: "" },
    { label: "Reference Density :", value: "1.204", unit: "kg/m³" },
    { label: "Medium Temperature :", value: "20", unit: "°C" },
    { label: "Maximum Fan Weight :", value: "161", unit: "kg" },
    { head: true, label: "Feed data" },
    { label: "Frequency :", value: "50", unit: "Hz" },
    { head: true, label: "Rated motor data" },
    {
      label: "Voltage- Phases- Frequency :",
      value: "415 V / 3 Phase / 50 Hz",
      unit: "",
    },
    { label: "Frame Size :", value: "132 S", unit: "" },
    { label: "Power (Pₙ):", value: "7.5", unit: "kW" },
    { label: "Speed (Nₙ):", value: "1455", unit: "min⁻¹" },
    { label: "Current (Iₙ):", value: "15.3", unit: "A" },
    { head: true, label: "Operational limits" },
    { label: "Maximum Fan Speed (nMax.):", value: "1740", unit: "min⁻¹" },
    { label: "Minimum Fan Speed (nMin.):", value: "756", unit: "min⁻¹" },
  ],
  dims1: [
    ["A", "795"],
    ["B", "437.5"],
    ["C", "50"],
    ["D", "625"],
    ["E", "760"],
    ["F", "398"],
    ["G", "811"],
    ["H", "715"],
    ["I", "674"],
  ],
  dims2: [
    ["J", "-"],
    ["K", "12"],
    ["L", "610"],
    ["M", "640"],
    ["N", "12"],
    ["O", "635"],
    ["P", "80"],
    ["n x J", "5 x 112"],
  ],
  soundRows: [
    {
      q: "15401",
      p: "0",
      n: "1440",
      w: "2.275",
      b: ["92", "98", "99", "95", "93", "89", "94", "92"],
      overall: "100",
    },
    {
      q: "12917",
      p: "497",
      n: "1440",
      w: "3.194",
      b: ["88", "93", "93", "90", "89", "86", "85", "80"],
      overall: "94",
    },
    {
      q: "9468",
      p: "1006",
      n: "1440",
      w: "3.889",
      b: ["91", "95", "93", "87", "86", "82", "80", "75"],
      overall: "91",
    },
    {
      q: "3960",
      p: "1124",
      n: "1440",
      w: "2.896",
      b: ["94", "97", "96", "91", "89", "84", "80", "74"],
      overall: "94",
    },
  ],
  footnote:
    "Performance certified is for installation type A:Free Inlet- Free Outlet. Performance ratings do not include the effects of Appurtenances (accessories) Power rating kW is fan input power. The Sound, Power Level ratings are shown in decibels,referred to 10⁻¹² watts, calculated per ANSI / AMCA Standard 210-16. Values shown are for inlet Lw and LwA sound power levels for Installation type A: Free Inlet- Free Outlet. Ratings include the effects of duct and correction.\nThe A-weighted sound ratings shown have been calculated per ANSI/AMCA Standard 210-16.",
  docCode: "KEBPL-KO-Aug 2026",
  website: "www.kohmer.in",
  page: "5 of 27",
  images: {
    companyLogo: "",
    brandLogo: "",
    product: "",
    drawing: "",
    badges: "",
  },
};

const DEFAULT_OPTS = {
  log: true,
  minEff: 50,
  fill: false,
  from: 2880,
  to: 3450,
  step: 100,
};
const BANDS = ["63", "125", "250", "500", "1000", "2000", "4000", "8000"];

/* ================================================================== */
/*  Parsing helpers (Excel upload + paste)                             */
/* ================================================================== */
const toNum = (v) => {
  if (v === null || v === undefined || String(v).trim() === "") return null;
  const n = parseFloat(String(v).replace(/,/g, ""));
  return Number.isFinite(n) ? n : null;
};

function findCols(header) {
  const h = header.map((c) => String(c ?? "").toLowerCase());
  const f = (t) => h.findIndex(t);
  return {
    model: f((x) => /model/.test(x)),
    speed: f((x) => /speed|rpm/.test(x)),
    q: f((x) => /air|flow/.test(x)),
    p: f((x) => /static|pressure/.test(x)),
    w: f((x) => /(power|kw)/.test(x) && !/sound|db/.test(x)),
    eff: f((x) => /eff/.test(x) && !/fei/.test(x)),
    fei: f((x) => /fei/.test(x)),
    sound: f((x) => /sound|db/.test(x)),
  };
}

/* Reads blocks laid out like the Excel file: header row with "Model",
   first data row carries Model + Speed, blank row ends a block. */
function parseSheetRows(rows) {
  const models = {};
  let cols = null,
    cur = null,
    model = null;
  for (const row of rows) {
    if (row.some((c) => /^model$/i.test(String(c ?? "").trim()))) {
      cols = findCols(row);
      cur = null;
      continue;
    }
    if (!cols) continue;
    const get = (k) => (cols[k] >= 0 ? toNum(row[cols[k]]) : null);
    const q = get("q"),
      p = get("p");
    if (q === null && p === null) {
      cur = null;
      continue;
    }
    const m = cols.model >= 0 ? String(row[cols.model] ?? "").trim() : "";
    const s = get("speed");
    if (m) model = m;
    if (s !== null || !cur) {
      const key = model || "Fan";
      models[key] = models[key] || { model: key, curves: [] };
      cur = { speed: s ?? 0, points: [] };
      models[key].curves.push(cur);
    }
    cur.points.push({
      idx: cur.points.length,
      q,
      p,
      w: get("w"),
      eff: get("eff"),
      fei: get("fei"),
      sound: get("sound"),
    });
  }
  return Object.values(models).filter((m) => m.curves.length);
}

function parsePasted(text) {
  const rows = text
    .replace(/\r/g, "")
    .split("\n")
    .map((l) => l.split("\t"));
  const hasHeader = rows.some((r) => r.some((c) => /^model$/i.test(c.trim())));
  if (!hasHeader) {
    const width = Math.max(...rows.map((r) => r.length));
    const header = [
      "Model",
      "Speed",
      "AirFlow",
      "Static",
      "Power",
      "TotalEff",
      "FEI",
      "Sound",
    ];
    rows.unshift(width >= 8 ? header : header.slice(1));
  }
  return parseSheetRows(rows);
}

/* editable table rows <-> curves */
const curvesToRows = (curves) =>
  curves.flatMap((c) =>
    c.points.map((pt, i) => ({
      speed: i === 0 ? String(c.speed) : "",
      ...Object.fromEntries(
        ["q", "p", "w", "eff", "fei", "sound"].map((k) => [
          k,
          pt[k] == null ? "" : String(pt[k]),
        ]),
      ),
    })),
  );

function rowsToCurves(rows) {
  const curves = [];
  let cur = null;
  rows.forEach((r) => {
    const s = toNum(r.speed),
      q = toNum(r.q),
      p = toNum(r.p);
    if (q === null && p === null) {
      if (s === null) cur = null;
      return;
    }
    if (s !== null || !cur) {
      cur = { speed: s ?? 0, points: [] };
      curves.push(cur);
    }
    cur.points.push({
      idx: cur.points.length,
      q,
      p,
      w: toNum(r.w),
      eff: toNum(r.eff),
      fei: toNum(r.fei),
      sound: toNum(r.sound),
    });
  });
  return curves;
}

/* Fan laws: Q ~ n, P ~ n², W ~ n³, Lw + 50·log10(n2/n1) */
function scaleCurve(ref, n) {
  const r = n / ref.speed;
  return {
    speed: n,
    calc: true,
    points: ref.points.map((pt) => ({
      ...pt,
      q: pt.q == null ? null : pt.q * r,
      p: pt.p == null ? null : pt.p * r * r,
      w: pt.w == null ? null : pt.w * r ** 3,
      sound: pt.sound == null ? null : pt.sound + 50 * Math.log10(r),
    })),
  };
}

function buildCurves(measured, opts) {
  const base = measured.filter((c) => c.speed > 0 && c.points.length);
  if (!base.length) return [];
  const out = base.map((c) => ({ ...c, calc: false }));
  if (opts.fill && opts.step > 0 && opts.to >= opts.from) {
    const speeds = [];
    for (let n = opts.from; n < opts.to && speeds.length < 30; n += opts.step)
      speeds.push(Math.round(n));
    speeds.push(Math.round(opts.to));
    speeds
      .filter((n) => n > 0 && !out.some((c) => c.speed === n))
      .forEach((n) => {
        const ref = base.reduce((a, c) =>
          Math.abs(c.speed - n) < Math.abs(a.speed - n) ? c : a,
        );
        out.push(scaleCurve(ref, n));
      });
  }
  return out.sort((a, b) => a.speed - b.speed);
}

/* ================================================================== */
/*  Axis helpers                                                       */
/* ================================================================== */
const STEPS = [1, 1.5, 2, 3, 4, 5, 7];
function niceLogDomain(min, max) {
  const cand = [];
  for (
    let e = Math.floor(Math.log10(min)) - 1;
    e <= Math.ceil(Math.log10(max)) + 1;
    e++
  )
    STEPS.forEach((s) => cand.push(+(s * 10 ** e).toPrecision(3)));
  cand.sort((a, b) => a - b);
  const lo = [...cand].reverse().find((v) => v <= min) ?? min;
  const hi = cand.find((v) => v >= max) ?? max;
  return { domain: [lo, hi], ticks: cand.filter((v) => v >= lo && v <= hi) };
}
const fmt = (v, d = 1) =>
  v == null || !Number.isFinite(v) ? "–" : Number(v).toFixed(d);
const trim = (v, d = 3) => (v == null ? "" : String(+Number(v).toFixed(d)));

function AxisTick({ x, y, payload, orient }) {
  const v = payload.value;
  const decade =
    v > 0 && Math.abs(Math.log10(v) - Math.round(Math.log10(v))) < 1e-9;
  const text =
    v >= 1000 ? v.toLocaleString("en-US") : String(+Number(v).toPrecision(4));
  return orient === "x" ? (
    <text
      x={x}
      y={y + 12}
      textAnchor="middle"
      fontSize={10}
      fontWeight={decade ? 700 : 400}
      fill={C.ink}
    >
      {text}
    </text>
  ) : (
    <text
      x={x - 4}
      y={y + 3}
      textAnchor="end"
      fontSize={10}
      fontWeight={decade ? 700 : 400}
      fill={C.ink}
    >
      {text}
    </text>
  );
}

function Tag({
  x,
  y,
  text,
  color,
  anchor = "middle",
  bold = false,
  border = 1,
}) {
  const w = text.length * 5.6 + 8,
    h = 14;
  const x0 = anchor === "middle" ? x - w / 2 : anchor === "end" ? x - w : x;
  return (
    <g pointerEvents="none">
      <rect
        x={x0}
        y={y - h / 2}
        width={w}
        height={h}
        fill="#fff"
        stroke={color}
        strokeWidth={border}
      />
      <text
        x={x0 + w / 2}
        y={y + 3.5}
        textAnchor="middle"
        fontSize={10}
        fill={color}
        fontWeight={bold ? 700 : 600}
      >
        {text}
      </text>
    </g>
  );
}

/* ================================================================== */
/*  Performance chart                                                  */
/* ================================================================== */
const LEGEND = [
  {
    key: "rpm",
    label: "RPM",
    icon: (
      <rect
        x="1"
        y="2"
        width="26"
        height="6"
        fill="#fff"
        stroke={C.rpm}
        strokeWidth="1.5"
      />
    ),
  },
  {
    key: "power",
    label: "Power",
    icon: (
      <rect
        x="1"
        y="2"
        width="26"
        height="6"
        fill="#fff"
        stroke={C.power}
        strokeWidth="1.5"
      />
    ),
  },
  {
    key: "eff",
    label: "Total Efficiency",
    icon: (
      <line
        x1="1"
        y1="5"
        x2="27"
        y2="5"
        stroke={C.eff}
        strokeWidth="4"
        strokeDasharray="1.5 1.5"
      />
    ),
  },
  {
    key: "fei",
    label: "FEI",
    icon: <rect x="1" y="3" width="26" height="4" fill={C.fei} />,
  },
  {
    key: "sound",
    label: "Sound",
    icon: <rect x="1" y="3" width="26" height="4" fill={C.sound} />,
  },
];

function PerformanceChart({ curves, opts, height = 660 }) {
  const [show, setShow] = useState({
    rpm: true,
    power: true,
    eff: true,
    fei: true,
    sound: true,
  });
  const [hover, setHover] = useState(null);
  const log = opts.log;

  const chart = useMemo(() => {
    const okP = (pt) =>
      pt &&
      pt.q != null &&
      pt.p != null &&
      (log ? pt.q > 0 && pt.p > 0 : pt.q >= 0 && pt.p >= 0);
    const okW = (pt) =>
      pt &&
      pt.q != null &&
      pt.w != null &&
      (log ? pt.q > 0 && pt.w > 0 : pt.q >= 0 && pt.w >= 0);
    const byQ = (a, b) => a.q - b.q;
    const tagged = (c, pt) => ({ ...pt, speed: c.speed, calc: c.calc });

    const rpmLines = curves.map((c) => ({
      ...c,
      data: c.points
        .filter(okP)
        .map((pt) => tagged(c, pt))
        .sort(byQ),
    }));
    const powerLines = curves.map((c) => ({
      ...c,
      data: c.points
        .filter(okW)
        .map((pt) => tagged(c, pt))
        .sort(byQ),
    }));

    /* iso-efficiency lines: same row index across all speeds */
    const maxLen = Math.max(0, ...curves.map((c) => c.points.length));
    const isoLines = [];
    for (let i = 0; i < maxLen; i++) {
      const data = curves
        .map((c) => (okP(c.points[i]) ? tagged(c, c.points[i]) : null))
        .filter(Boolean)
        .sort(byQ);
      const ref = data.find((d) => !d.calc) || data[0];
      if (data.length && ref?.eff > 0)
        isoLines.push({ idx: i, eff: ref.eff, fei: ref.fei, data });
    }

    /* red selection-range lines: where efficiency drops to opts.minEff either side of peak */
    const left = [],
      right = [];
    const cross = (a, b, t) => {
      const f = (a.eff - t) / (a.eff - b.eff);
      return { q: a.q + (b.q - a.q) * f, p: a.p + (b.p - a.p) * f };
    };
    curves.forEach((c) => {
      const pts = c.points
        .filter((pt) => pt.eff != null && pt.q != null && pt.p != null)
        .sort(byQ);
      if (!pts.length) return;
      let bi = 0;
      pts.forEach((pt, i) => {
        if (pt.eff > pts[bi].eff) bi = i;
      });
      const t = opts.minEff;
      if (pts[bi].eff < t) return;
      for (let j = bi - 1; j >= 0; j--)
        if (pts[j].eff < t) {
          const x = cross(pts[j + 1], pts[j], t);
          if (okP(x)) left.push(x);
          break;
        }
      for (let j = bi + 1; j < pts.length; j++)
        if (pts[j].eff < t) {
          const x = cross(pts[j - 1], pts[j], t);
          if (okP(x)) right.push(x);
          break;
        }
    });
    const rangeLines = [left, right]
      .filter((l) => l.length > 1)
      .map((l) => l.sort(byQ));

    const allP = rpmLines.flatMap((l) => l.data);
    const allW = powerLines.flatMap((l) => l.data);
    const qs = [...allP, ...allW].map((d) => d.q).filter((v) => v > 0);
    const ps = allP.map((d) => d.p).filter((v) => v > 0);
    const ws = allW.map((d) => d.w).filter((v) => v > 0);
    if (!qs.length || !ps.length) return null;

    const [minQ, maxQ, minP, maxP] = [
      Math.min(...qs),
      Math.max(...qs),
      Math.min(...ps),
      Math.max(...ps),
    ];
    const minW = ws.length ? Math.min(...ws) : 0.1,
      maxW = ws.length ? Math.max(...ws) : 1;

    let x, p, w;
    if (log) {
      x = niceLogDomain(minQ / 1.05, maxQ * 1.3);
      p = niceLogDomain(minP / 1.6, maxP * 1.5);
      w = { domain: [minW / 1.4, maxW * 3.2] };
    } else {
      x = { domain: [0, Math.ceil((maxQ * 1.2) / 500) * 500] };
      p = { domain: [0, Math.ceil((maxP * 1.3) / 100) * 100] };
      w = { domain: [0, maxW * 1.6] };
    }

    /* extend each power curve horizontally to the right edge, like the datasheet */
    powerLines.forEach((l) => {
      const last = l.data[l.data.length - 1];
      if (last) l.data = [...l.data, { ...last, q: x.domain[1], ext: true }];
    });

    return { rpmLines, powerLines, isoLines, rangeLines, x, p, w };
  }, [curves, log, opts.minEff]);

  if (!chart)
    return (
      <p className="text-center py-20 text-sm">
        No plottable points. Open “Edit data” and add airflow and pressure
        values.
      </p>
    );

  const hoverCircle = (props) => (
    <circle
      cx={props.cx}
      cy={props.cy}
      r={6}
      fill="transparent"
      style={{ cursor: "crosshair" }}
      onMouseEnter={() =>
        setHover({ ...props.payload, cx: props.cx, cy: props.cy })
      }
    />
  );

  const rpmDot = (line) => (props) => {
    const { cx, cy, payload, index } = props;
    if (cx == null) return <g key={index} />;
    const last = index === line.data.length - 1;
    return (
      <g key={`r${line.speed}-${index}`}>
        {show.sound && payload.sound != null && (
          <Tag
            x={cx - 5}
            y={cy + 10}
            anchor="end"
            text={`${Math.round(payload.sound)} dB(A)`}
            color={C.sound}
          />
        )}
        {show.rpm && last && (
          <Tag
            x={cx + 3}
            y={cy + 12}
            anchor="start"
            text={`${Math.round(line.speed)} RPM${line.calc ? "*" : ""}`}
            color={line.calc ? C.rpmCalc : C.rpm}
          />
        )}
        {hoverCircle(props)}
      </g>
    );
  };

  const powerDot = (line) => (props) => {
    const { cx, cy, payload, index } = props;
    if (cx == null) return <g key={index} />;
    if (payload.ext)
      return (
        <text
          key={`w${line.speed}`}
          x={cx + 5}
          y={cy + 3}
          fontSize={9.5}
          fill={C.power}
        >
          {trim(payload.w)}
        </text>
      );
    return <g key={`w${line.speed}-${index}`}>{hoverCircle(props)}</g>;
  };

  const isoDot =
    (line) =>
    ({ cx, cy, index }) => {
      if (index !== line.data.length - 1 || cx == null)
        return <g key={index} />;
      return (
        <g key={`i${line.idx}`}>
          {show.eff && (
            <Tag
              x={cx}
              y={cy - 13}
              text={`${trim(line.eff, 1)}%`}
              color={C.eff}
            />
          )}
          {show.fei && line.fei != null && (
            <Tag
              x={cx}
              y={cy - 30}
              text={trim(line.fei, 2)}
              color={C.fei}
              bold
              border={1.6}
            />
          )}
        </g>
      );
    };

  const scale = log ? "log" : "linear";

  return (
    <div>
      <div
        className="flex flex-wrap justify-center gap-4 pt-1"
        style={{ fontSize: 11 }}
      >
        {LEGEND.map((l) => (
          <button
            key={l.key}
            onClick={() => setShow((s) => ({ ...s, [l.key]: !s[l.key] }))}
            aria-pressed={show[l.key]}
            className="flex items-center gap-1"
            style={{ opacity: show[l.key] ? 1 : 0.35 }}
          >
            <svg width="28" height="10">
              {l.icon}
            </svg>
            {l.label}
          </button>
        ))}
      </div>

      <div
        className="relative"
        style={{ height }}
        onMouseLeave={() => setHover(null)}
      >
        <ResponsiveContainer>
          <ComposedChart margin={{ top: 30, right: 10, bottom: 22, left: 16 }}>
            <CartesianGrid stroke={C.grid} />
            <XAxis
              type="number"
              dataKey="q"
              scale={scale}
              domain={chart.x.domain}
              ticks={chart.x.ticks}
              tick={<AxisTick orient="x" />}
              axisLine={{ stroke: "#bbb" }}
              tickLine={false}
            >
              <Label
                value="Air Flow (m³/hr)"
                position="bottom"
                offset={6}
                style={{ fontSize: 10, fontWeight: 700, fill: C.ink }}
              />
            </XAxis>
            <YAxis
              yAxisId="p"
              type="number"
              scale={scale}
              domain={chart.p.domain}
              ticks={chart.p.ticks}
              tick={<AxisTick orient="y" />}
              axisLine={{ stroke: "#bbb" }}
              tickLine={false}
              width={48}
            >
              <Label
                value="Static Pressure (Pa)"
                angle={-90}
                position="insideLeft"
                offset={-8}
                style={{
                  fontSize: 10,
                  fontWeight: 700,
                  fill: C.ink,
                  textAnchor: "middle",
                }}
              />
            </YAxis>
            <YAxis
              yAxisId="w"
              orientation="right"
              type="number"
              scale={scale}
              domain={chart.w.domain}
              tick={false}
              axisLine={false}
              tickLine={false}
              width={52}
            >
              {show.power && (
                <Label
                  value="Power (kW)"
                  angle={90}
                  position="insideRight"
                  style={{
                    fontSize: 10,
                    fontWeight: 700,
                    fill: C.power,
                    textAnchor: "middle",
                  }}
                />
              )}
            </YAxis>

            {chart.rangeLines.map((l, i) => (
              <Line
                key={`range${i}`}
                yAxisId="p"
                data={l}
                dataKey="p"
                type="linear"
                stroke={C.range}
                strokeWidth={2}
                dot={false}
                activeDot={false}
                isAnimationActive={false}
              />
            ))}
            {show.power &&
              chart.powerLines.map((l) => (
                <Line
                  key={`w${l.speed}`}
                  yAxisId="w"
                  data={l.data}
                  dataKey="w"
                  type="monotone"
                  stroke={C.power}
                  strokeWidth={1.2}
                  strokeOpacity={l.calc ? 0.6 : 1}
                  dot={powerDot(l)}
                  activeDot={false}
                  isAnimationActive={false}
                />
              ))}
            {chart.rpmLines.map((l) => (
              <Line
                key={`r${l.speed}`}
                yAxisId="p"
                data={l.data}
                dataKey="p"
                type="monotone"
                stroke={show.rpm ? (l.calc ? C.rpmCalc : C.rpm) : "transparent"}
                strokeWidth={2}
                dot={rpmDot(l)}
                activeDot={false}
                isAnimationActive={false}
              />
            ))}
            {/* efficiency lines last so their labels sit on top */}
            {chart.isoLines.map((l) => (
              <Line
                key={`e${l.idx}`}
                yAxisId="p"
                data={l.data}
                dataKey="p"
                type="linear"
                stroke={show.eff ? C.eff : "transparent"}
                strokeDasharray="3 3"
                strokeWidth={0.9}
                dot={isoDot(l)}
                activeDot={false}
                isAnimationActive={false}
              />
            ))}
          </ComposedChart>
        </ResponsiveContainer>

        {hover && (
          <div
            className="no-print absolute bg-white border shadow px-2 py-1 pointer-events-none"
            style={{
              left: Math.min(hover.cx + 12, 640),
              top: hover.cy + 12,
              fontSize: 11,
              lineHeight: 1.4,
              borderColor: C.rpm,
            }}
          >
            <b>
              {Math.round(hover.speed)} RPM{hover.calc ? " (fan law)" : ""}
            </b>
            <br />
            Air flow {fmt(hover.q)} m³/h
            <br />
            Static pressure {fmt(hover.p)} Pa
            <br />
            Power {fmt(hover.w, 3)} kW
            <br />
            Total eff. {fmt(hover.eff)}% · FEI {fmt(hover.fei, 3)}
            {hover.sound != null && (
              <>
                <br />
                Sound {fmt(hover.sound, 1)} dB(A)
              </>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

/* ================================================================== */
/*  Datasheet page                                                     */
/* ================================================================== */
const cell = {
  border: "1px solid #555",
  padding: "1px 4px",
  textAlign: "center",
};

function ImgOr({ src, label, style }) {
  return src ? (
    <img src={src} alt={label} style={{ objectFit: "contain", ...style }} />
  ) : (
    <div
      className="flex items-center justify-center text-center"
      style={{
        border: "1px dashed #aaa",
        color: "#888",
        fontSize: 11,
        ...style,
      }}
    >
      {label}
    </div>
  );
}

function DimTable({ dims }) {
  return (
    <table
      style={{
        borderCollapse: "collapse",
        width: "100%",
        fontSize: 9,
        marginTop: 4,
      }}
    >
      <tbody>
        <tr>
          {dims.map(([k], i) => (
            <td key={i} style={{ ...cell, fontWeight: 700 }}>
              {k}
            </td>
          ))}
        </tr>
        <tr>
          {dims.map(([, v], i) => (
            <td key={i} style={cell}>
              {v}
            </td>
          ))}
        </tr>
      </tbody>
    </table>
  );
}

function Datasheet({ model, sheet, curves, opts }) {
  const im = sheet.images;
  return (
    <div
      className="datasheet-page bg-white mx-auto"
      style={{
        width: 920,
        padding: "18px 26px",
        fontFamily: FONT,
        color: C.ink,
      }}
    >
      {/* Header */}
      <div
        className="flex items-end justify-between"
        style={{ borderBottom: "1.5px solid #333", paddingBottom: 4 }}
      >
        {im.companyLogo ? (
          <img
            src={im.companyLogo}
            alt={sheet.company}
            style={{ height: 42 }}
          />
        ) : (
          <div className="flex items-end gap-1">
            {/* <div
              style={{
                width: 30,
                height: 34,
                border: `3px solid ${C.navy}`,
                borderRight: "none",
                color: C.navy,
                fontWeight: 800,
                fontSize: 22,
                lineHeight: "28px",
                paddingLeft: 3,
              }}
            >
              K
            </div>
            <span
              style={{
                color: C.navy,
                fontWeight: 800,
                fontSize: 26,
                letterSpacing: 0.5,
              }}
            >
              {sheet.company}
            </span> */}
          </div>
        )}
        {im.brandLogo ? (
          <img src={im.brandLogo} alt={sheet.brand} style={{ height: 42 }} />
        ) : (
          <div className="text-right">
            {/* <div className="flex items-center gap-1 justify-end">
              <span
                style={{
                  background: C.range,
                  color: "#fff",
                  borderRadius: "50% 0 0 50%",
                  padding: "0 6px",
                  fontSize: 14,
                }}
              >
                ✻
              </span>
              <span
                style={{
                  color: C.navy,
                  fontWeight: 800,
                  fontSize: 22,
                  borderBottom: `2px solid ${C.range}`,
                }}
              >
                {sheet.brand}
              </span>
            </div>
            <div style={{ color: C.range, fontSize: 9 }}>{sheet.tagline}</div> */}
          </div>
        )}
      </div>

      {/* Title + specs + drawing */}
      <div className="flex gap-4" style={{ marginTop: 6 }}>
        <div style={{ width: 440 }}>
          <div style={{ fontSize: 17, margin: "0 0 6px 10px" }}>
            {sheet.modelPrefix}
            {model}
          </div>
          <div
            style={{
              background: C.specBg,
              padding: "6px 12px",
              fontSize: 12.5,
            }}
          >
            {sheet.specs.map((s, i) => (
              <div
                key={i}
                className="flex"
                style={{
                  fontWeight: s.head ? 700 : 400,
                  marginTop: s.head && i > 0 ? 8 : 0,
                  lineHeight: 1.45,
                }}
              >
                <span style={{ flex: "0 0 215px" }}>{s.label}</span>
                <span style={{ flex: 1 }}>{s.value}</span>
                <span style={{ width: 50, textAlign: "right" }}>{s.unit}</span>
              </div>
            ))}
          </div>
        </div>
        <div style={{ flex: 1 }}>
          <div className="flex justify-end">
            <ImgOr
              src={im.product}
              label="Product photo"
              style={{ width: 150, height: 110 }}
            />
          </div>
          <ImgOr
            src={im.drawing}
            label="Dimensional drawing"
            style={{ width: "100%", height: 170, marginTop: 4 }}
          />
          <DimTable dims={sheet.dims1} />
          <DimTable dims={sheet.dims2} />
        </div>
      </div>

      {/* Chart */}
      <div
        style={{
          border: `1px solid ${C.frame}`,
          marginTop: 8,
          padding: "2px 4px",
        }}
      >
        <PerformanceChart curves={curves} opts={opts} />
      </div>

      {/* Sound table */}
      <table
        style={{
          borderCollapse: "collapse",
          width: "100%",
          fontSize: 10.5,
          marginTop: 8,
        }}
      >
        <tbody>
          <tr>
            <td
              rowSpan={sheet.soundRows.length + 2}
              style={{ ...cell, width: 70 }}
            >
              {model}
            </td>
            <td rowSpan={2} style={cell}>
              Flow
              <br />
              (CMH)
            </td>
            <td rowSpan={2} style={cell}>
              Static Pressure
              <br />
              (Pa)
            </td>
            <td rowSpan={2} style={cell}>
              Speed
              <br />
              (RPM)
            </td>
            <td rowSpan={2} style={cell}>
              Absorbed Power
              <br />
              (BkW) (kW)
            </td>
            <td colSpan={8} style={cell}>
              Sound Power Levels (Lwo) (dB)
            </td>
            <td rowSpan={2} style={cell}>
              Overall (LwoA)
              <br />
              (dBA)
            </td>
          </tr>
          <tr>
            {BANDS.map((b) => (
              <td key={b} style={cell}>
                {b}
              </td>
            ))}
          </tr>
          {sheet.soundRows.map((r, i) => (
            <tr key={i}>
              <td style={cell}>{r.q}</td>
              <td style={cell}>{r.p}</td>
              <td style={cell}>{r.n}</td>
              <td style={cell}>{r.w}</td>
              {r.b.map((v, j) => (
                <td key={j} style={cell}>
                  {v}
                </td>
              ))}
              <td style={cell}>{r.overall}</td>
            </tr>
          ))}
        </tbody>
      </table>

      {/* Footnote */}
      {/* <div className="flex gap-4 items-start" style={{ marginTop: 10 }}>
        <ImgOr
          src={im.badges}
          label="Certification badges"
          style={{ width: 100, height: 70, flex: "0 0 100px" }}
        />
        <p
          style={{
            fontSize: 9.5,
            lineHeight: 1.35,
            whiteSpace: "pre-line",
            margin: 0,
          }}
        >
          {sheet.footnote}
        </p>
      </div> */}

      {/* <div
        className="flex justify-between"
        style={{ fontSize: 13, marginTop: 14 }}
      >
        <span>{sheet.docCode}</span>
        <a
          href={`https://${sheet.website.replace(/^https?:\/\//, "")}`}
          style={{ color: "#1a56c4", textDecoration: "underline" }}
        >
          {sheet.website}
        </a>
        <span>{sheet.page}</span>
      </div> */}
    </div>
  );
}

/* ================================================================== */
/*  Data editor modal                                                  */
/* ================================================================== */
const CURVE_COLS = [
  { key: "speed", label: "Speed (RPM)" },
  { key: "q", label: "AirFlowRate m³/hr" },
  { key: "p", label: "StaticPressure Pa" },
  { key: "w", label: "AbsorbedPower kW" },
  { key: "eff", label: "TotalEff %" },
  { key: "fei", label: "FEI" },
  { key: "sound", label: "Sound Power dBA" },
];
const inputCls = "w-full border border-gray-300 rounded px-1 py-0.5 text-sm";

function Field({ label, children }) {
  return (
    <label className="flex flex-col gap-1 text-sm">
      <span className="text-gray-600">{label}</span>
      {children}
    </label>
  );
}

function EditorModal({ initial, onApply, onClose }) {
  const [tab, setTab] = useState("curve");
  const [model, setModel] = useState(initial.model);
  const [rows, setRows] = useState(() => curvesToRows(initial.curves));
  const [sheet, setSheet] = useState(() =>
    JSON.parse(JSON.stringify(initial.sheet)),
  );
  const [opts, setOpts] = useState({ ...initial.opts });
  const [paste, setPaste] = useState("");
  const [msg, setMsg] = useState("");

  const setRow = (i, k, v) =>
    setRows((rs) => rs.map((r, j) => (j === i ? { ...r, [k]: v } : r)));
  const emptyRow = () => Object.fromEntries(CURVE_COLS.map((c) => [c.key, ""]));
  const upd = (patch) => setSheet((s) => ({ ...s, ...patch }));

  const loadParsed = (parsed, source) => {
    if (!parsed.length) {
      setMsg(
        `No curve data found in ${source}. Include the header row (Model, Speed, AirFlowRate, StaticPressure…) or paste the 8 columns from Model to Sound.`,
      );
      return;
    }
    const m = parsed[0];
    setModel(m.model);
    setRows(curvesToRows(m.curves));
    const speeds = m.curves.map((c) => c.speed);
    setOpts((o) => ({
      ...o,
      from: Math.min(...speeds),
      to: Math.max(...speeds),
    }));
    setMsg(
      `Loaded ${m.model}: ${m.curves.length} speeds, ${m.curves.reduce((s, c) => s + c.points.length, 0)} rows from ${source}.` +
        (parsed.length > 1
          ? ` The file also has ${parsed
              .slice(1)
              .map((x) => x.model)
              .join(", ")}; only the first model is used.`
          : ""),
    );
  };

  const onFile = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const wb = XLSX.read(await file.arrayBuffer());
      const ws = wb.Sheets[wb.SheetNames[0]];
      loadParsed(
        parseSheetRows(
          XLSX.utils.sheet_to_json(ws, {
            header: 1,
            defval: "",
            blankrows: true,
          }),
        ),
        file.name,
      );
    } catch (err) {
      setMsg(`Could not read ${file.name}: ${err.message}`);
    }
    e.target.value = "";
  };

  const onImage = (key) => (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const r = new FileReader();
    r.onload = () => upd({ images: { ...sheet.images, [key]: r.result } });
    r.readAsDataURL(file);
  };

  const apply = () => {
    const curves = rowsToCurves(rows);
    if (!curves.length) {
      setMsg(
        "Add at least one row with a speed, airflow and static pressure before applying.",
      );
      setTab("curve");
      return;
    }
    onApply({ model, curves, sheet, opts });
  };

  const TABS = [
    ["curve", "Curve data"],
    ["sheet", "Datasheet details"],
    ["sound", "Sound table"],
    ["images", "Images"],
  ];

  return (
    <div
      className="no-print fixed inset-0 z-50 flex items-center justify-center p-4"
      style={{ background: "rgba(20,30,40,.45)", fontFamily: FONT }}
    >
      <div
        className="bg-white rounded-lg shadow-xl flex flex-col w-full"
        style={{ maxWidth: 1040, maxHeight: "92vh" }}
      >
        <div className="flex items-center justify-between px-5 py-3 border-b">
          <h2 className="text-lg font-semibold">Edit datasheet data</h2>
          <button
            onClick={onClose}
            className="text-2xl leading-none px-2"
            aria-label="Close"
          >
            ×
          </button>
        </div>

        <div className="flex gap-1 px-5 pt-3 border-b">
          {TABS.map(([k, l]) => (
            <button
              key={k}
              onClick={() => setTab(k)}
              className="px-3 py-2 text-sm rounded-t"
              style={{
                background: tab === k ? C.rpm : "transparent",
                color: tab === k ? "#fff" : C.ink,
              }}
            >
              {l}
            </button>
          ))}
        </div>

        <div className="overflow-auto px-5 py-4 flex-1">
          {msg && (
            <p
              className="text-sm mb-3 px-3 py-2 rounded"
              style={{ background: "#eef7f6" }}
            >
              {msg}
            </p>
          )}

          {tab === "curve" && (
            <div className="flex flex-col gap-4">
              <div className="flex flex-wrap gap-4 items-end">
                <Field label="Model">
                  <input
                    className={inputCls}
                    value={model}
                    onChange={(e) => setModel(e.target.value)}
                  />
                </Field>
                <label
                  className="px-3 py-1.5 rounded border text-sm cursor-pointer"
                  style={{ borderColor: C.rpm, color: C.rpm }}
                >
                  Upload Excel (.xlsx)
                  <input
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    className="hidden"
                    onChange={onFile}
                  />
                </label>
              </div>

              <div>
                <Field label="Or paste straight from Excel (with or without the header rows), then press Load pasted data">
                  <textarea
                    className={inputCls}
                    rows={3}
                    value={paste}
                    onChange={(e) => setPaste(e.target.value)}
                    placeholder={
                      "Model\tSpeed\tAirFlowRate m3\\hr\tStaticPressure Pa\tAbsorbedPower kW\tTotalEff\tFEI\tSound Power dBA\nKESO-315\t2880\t3387.6\t0\t0.4\t8\t0.565\t94.7"
                    }
                  />
                </Field>
                <button
                  className="mt-2 px-3 py-1.5 rounded text-sm text-white"
                  style={{ background: C.rpm }}
                  onClick={() => loadParsed(parsePasted(paste), "pasted text")}
                >
                  Load pasted data
                </button>
              </div>

              <div
                className="flex flex-wrap gap-4 items-end p-3 rounded"
                style={{ background: "#f5f6f7" }}
              >
                <Field label="Selection range (red lines) at efficiency %">
                  <input
                    type="number"
                    className={inputCls}
                    style={{ width: 90 }}
                    value={opts.minEff}
                    onChange={(e) =>
                      setOpts({ ...opts, minEff: Number(e.target.value) })
                    }
                  />
                </Field>
                <label className="flex items-center gap-2 text-sm">
                  <input
                    type="checkbox"
                    checked={opts.fill}
                    onChange={(e) =>
                      setOpts({ ...opts, fill: e.target.checked })
                    }
                  />
                  Add fan-law speeds between
                </label>
                <Field label="From RPM">
                  <input
                    type="number"
                    className={inputCls}
                    style={{ width: 90 }}
                    value={opts.from}
                    onChange={(e) =>
                      setOpts({ ...opts, from: Number(e.target.value) })
                    }
                  />
                </Field>
                <Field label="To RPM">
                  <input
                    type="number"
                    className={inputCls}
                    style={{ width: 90 }}
                    value={opts.to}
                    onChange={(e) =>
                      setOpts({ ...opts, to: Number(e.target.value) })
                    }
                  />
                </Field>
                <Field label="Step RPM">
                  <input
                    type="number"
                    className={inputCls}
                    style={{ width: 90 }}
                    value={opts.step}
                    onChange={(e) =>
                      setOpts({ ...opts, step: Number(e.target.value) })
                    }
                  />
                </Field>
              </div>

              <p className="text-sm text-gray-600">
                Put the speed on the first row of each block and leave it blank
                on the rows below, the same as your Excel sheet. A new speed
                value starts a new curve.
              </p>
              <div
                className="overflow-auto border rounded"
                style={{ maxHeight: 380 }}
              >
                <table
                  className="w-full text-sm"
                  style={{ borderCollapse: "collapse" }}
                >
                  <thead className="sticky top-0 bg-gray-100">
                    <tr>
                      <th className="px-1 py-1 text-left w-8">#</th>
                      {CURVE_COLS.map((c) => (
                        <th
                          key={c.key}
                          className="px-1 py-1 text-left font-semibold"
                        >
                          {c.label}
                        </th>
                      ))}
                      <th />
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map((r, i) => (
                      <tr
                        key={i}
                        style={{
                          borderTop: r.speed
                            ? `2px solid ${C.rpm}`
                            : "1px solid #eee",
                        }}
                      >
                        <td className="px-1 text-gray-500">{i + 1}</td>
                        {CURVE_COLS.map((c) => (
                          <td key={c.key} className="px-1 py-0.5">
                            <input
                              className={inputCls}
                              value={r[c.key]}
                              onChange={(e) => setRow(i, c.key, e.target.value)}
                              style={
                                c.key === "speed" && r.speed
                                  ? { fontWeight: 700 }
                                  : undefined
                              }
                            />
                          </td>
                        ))}
                        <td className="px-1">
                          <button
                            className="text-gray-500 px-1"
                            aria-label={`Delete row ${i + 1}`}
                            onClick={() =>
                              setRows((rs) => rs.filter((_, j) => j !== i))
                            }
                          >
                            ×
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="flex gap-2 text-sm">
                <button
                  className="px-3 py-1.5 rounded border"
                  onClick={() => setRows((rs) => [...rs, emptyRow()])}
                >
                  Add row
                </button>
                <button
                  className="px-3 py-1.5 rounded border"
                  onClick={() =>
                    setRows((rs) => [
                      ...rs,
                      ...Array.from({ length: 10 }, emptyRow),
                    ])
                  }
                >
                  Add speed block (10 rows)
                </button>
                <button
                  className="px-3 py-1.5 rounded border"
                  onClick={() => setRows([])}
                >
                  Clear table
                </button>
              </div>
            </div>
          )}

          {tab === "sheet" && (
            <div className="flex flex-col gap-4">
              <div
                className="grid gap-3"
                style={{ gridTemplateColumns: "repeat(3, minmax(0,1fr))" }}
              >
                <Field label="Company (left logo text)">
                  <input
                    className={inputCls}
                    value={sheet.company}
                    onChange={(e) => upd({ company: e.target.value })}
                  />
                </Field>
                <Field label="Brand (right logo text)">
                  <input
                    className={inputCls}
                    value={sheet.brand}
                    onChange={(e) => upd({ brand: e.target.value })}
                  />
                </Field>
                <Field label="Brand tagline">
                  <input
                    className={inputCls}
                    value={sheet.tagline}
                    onChange={(e) => upd({ tagline: e.target.value })}
                  />
                </Field>
                <Field label="Title before model">
                  <input
                    className={inputCls}
                    value={sheet.modelPrefix}
                    onChange={(e) => upd({ modelPrefix: e.target.value })}
                  />
                </Field>
                <Field label="Document code">
                  <input
                    className={inputCls}
                    value={sheet.docCode}
                    onChange={(e) => upd({ docCode: e.target.value })}
                  />
                </Field>
                <Field label="Website">
                  <input
                    className={inputCls}
                    value={sheet.website}
                    onChange={(e) => upd({ website: e.target.value })}
                  />
                </Field>
                <Field label="Page number">
                  <input
                    className={inputCls}
                    value={sheet.page}
                    onChange={(e) => upd({ page: e.target.value })}
                  />
                </Field>
              </div>

              <h3 className="font-semibold">Specification table</h3>
              <table className="text-sm w-full">
                <thead>
                  <tr className="text-left text-gray-600">
                    <th className="w-24">Heading row</th>
                    <th>Label</th>
                    <th>Value</th>
                    <th className="w-28">Unit</th>
                    <th />
                  </tr>
                </thead>
                <tbody>
                  {sheet.specs.map((s, i) => {
                    const set = (k, v) =>
                      upd({
                        specs: sheet.specs.map((x, j) =>
                          j === i ? { ...x, [k]: v } : x,
                        ),
                      });
                    return (
                      <tr key={i}>
                        <td>
                          <input
                            type="checkbox"
                            checked={!!s.head}
                            onChange={(e) => set("head", e.target.checked)}
                          />
                        </td>
                        <td className="pr-1">
                          <input
                            className={inputCls}
                            value={s.label || ""}
                            onChange={(e) => set("label", e.target.value)}
                            style={s.head ? { fontWeight: 700 } : undefined}
                          />
                        </td>
                        <td className="pr-1">
                          <input
                            className={inputCls}
                            value={s.value || ""}
                            onChange={(e) => set("value", e.target.value)}
                          />
                        </td>
                        <td className="pr-1">
                          <input
                            className={inputCls}
                            value={s.unit || ""}
                            onChange={(e) => set("unit", e.target.value)}
                          />
                        </td>
                        <td>
                          <button
                            className="px-1 text-gray-500"
                            aria-label="Delete spec row"
                            onClick={() =>
                              upd({
                                specs: sheet.specs.filter((_, j) => j !== i),
                              })
                            }
                          >
                            ×
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
              <button
                className="self-start px-3 py-1.5 rounded border text-sm"
                onClick={() =>
                  upd({
                    specs: [...sheet.specs, { label: "", value: "", unit: "" }],
                  })
                }
              >
                Add spec row
              </button>

              <h3 className="font-semibold">Dimensions (mm)</h3>
              {["dims1", "dims2"].map((key) => (
                <div key={key} className="flex flex-wrap gap-2">
                  {sheet[key].map(([k, v], i) => {
                    const set = (idx, val) =>
                      upd({
                        [key]: sheet[key].map((d, j) =>
                          j === i ? (idx === 0 ? [val, d[1]] : [d[0], val]) : d,
                        ),
                      });
                    return (
                      <div
                        key={i}
                        className="flex flex-col gap-1"
                        style={{ width: 78 }}
                      >
                        <input
                          className={inputCls}
                          value={k}
                          onChange={(e) => set(0, e.target.value)}
                          style={{ fontWeight: 700 }}
                          aria-label="Dimension letter"
                        />
                        <input
                          className={inputCls}
                          value={v}
                          onChange={(e) => set(1, e.target.value)}
                          aria-label={`Dimension ${k}`}
                        />
                      </div>
                    );
                  })}
                </div>
              ))}

              <Field label="Footnote">
                <textarea
                  className={inputCls}
                  rows={5}
                  value={sheet.footnote}
                  onChange={(e) => upd({ footnote: e.target.value })}
                />
              </Field>
            </div>
          )}

          {tab === "sound" && (
            <div className="flex flex-col gap-3">
              <div className="overflow-auto">
                <table
                  className="text-sm"
                  style={{ borderCollapse: "collapse" }}
                >
                  <thead>
                    <tr className="bg-gray-100 text-left">
                      {[
                        "Flow (CMH)",
                        "Static (Pa)",
                        "Speed (RPM)",
                        "Power (kW)",
                        ...BANDS.map((b) => `${b} Hz`),
                        "Overall dBA",
                      ].map((h) => (
                        <th
                          key={h}
                          className="px-1 py-1 font-semibold whitespace-nowrap"
                        >
                          {h}
                        </th>
                      ))}
                      <th />
                    </tr>
                  </thead>
                  <tbody>
                    {sheet.soundRows.map((r, i) => {
                      const set = (patch) =>
                        upd({
                          soundRows: sheet.soundRows.map((x, j) =>
                            j === i ? { ...x, ...patch } : x,
                          ),
                        });
                      return (
                        <tr key={i}>
                          {["q", "p", "n", "w"].map((k) => (
                            <td key={k} className="px-0.5">
                              <input
                                className={inputCls}
                                style={{ width: 72 }}
                                value={r[k]}
                                onChange={(e) => set({ [k]: e.target.value })}
                              />
                            </td>
                          ))}
                          {r.b.map((v, j) => (
                            <td key={j} className="px-0.5">
                              <input
                                className={inputCls}
                                style={{ width: 46 }}
                                value={v}
                                onChange={(e) =>
                                  set({
                                    b: r.b.map((x, m) =>
                                      m === j ? e.target.value : x,
                                    ),
                                  })
                                }
                              />
                            </td>
                          ))}
                          <td className="px-0.5">
                            <input
                              className={inputCls}
                              style={{ width: 60 }}
                              value={r.overall}
                              onChange={(e) => set({ overall: e.target.value })}
                            />
                          </td>
                          <td>
                            <button
                              className="px-1 text-gray-500"
                              aria-label="Delete sound row"
                              onClick={() =>
                                upd({
                                  soundRows: sheet.soundRows.filter(
                                    (_, j) => j !== i,
                                  ),
                                })
                              }
                            >
                              ×
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              <button
                className="self-start px-3 py-1.5 rounded border text-sm"
                onClick={() =>
                  upd({
                    soundRows: [
                      ...sheet.soundRows,
                      {
                        q: "",
                        p: "",
                        n: "",
                        w: "",
                        b: BANDS.map(() => ""),
                        overall: "",
                      },
                    ],
                  })
                }
              >
                Add row
              </button>
            </div>
          )}

          {tab === "images" && (
            <div
              className="grid gap-4"
              style={{
                gridTemplateColumns: "repeat(auto-fill, minmax(220px,1fr))",
              }}
            >
              {[
                ["companyLogo", "Company logo (left)"],
                ["brandLogo", "Brand logo (right)"],
                ["product", "Product photo"],
                ["drawing", "Dimensional drawing"],
                ["badges", "Certification badges"],
              ].map(([k, l]) => (
                <div key={k} className="border rounded p-3 flex flex-col gap-2">
                  <span className="text-sm font-semibold">{l}</span>
                  <ImgOr
                    src={sheet.images[k]}
                    label="No image"
                    style={{ width: "100%", height: 110 }}
                  />
                  <div className="flex gap-2 text-sm">
                    <label className="px-2 py-1 rounded border cursor-pointer">
                      Upload
                      <input
                        type="file"
                        accept="image/*"
                        className="hidden"
                        onChange={onImage(k)}
                      />
                    </label>
                    {sheet.images[k] && (
                      <button
                        className="px-2 py-1 rounded border"
                        onClick={() =>
                          upd({ images: { ...sheet.images, [k]: "" } })
                        }
                      >
                        Remove
                      </button>
                    )}
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>

        <div className="flex justify-end gap-2 px-5 py-3 border-t">
          <button
            className="px-4 py-2 rounded border text-sm"
            onClick={onClose}
          >
            Cancel
          </button>
          <button
            className="px-4 py-2 rounded text-sm text-white"
            style={{ background: C.rpm }}
            onClick={apply}
          >
            Apply to datasheet
          </button>
        </div>
      </div>
    </div>
  );
}

/* Minimal copy of the utility classes used here, so the component also works in projects without Tailwind */
const BASE_CSS = `
.fds *, .fds *::before, .fds *::after { box-sizing: border-box; }
.fds button { background: transparent; border: 0 solid #d1d5db; font: inherit; color: inherit; cursor: pointer; }
.fds input, .fds textarea, .fds select { font: inherit; color: inherit; background: #fff; }
.fds h2, .fds h3, .fds p { margin: 0; }
.fds .flex{display:flex}.fds .grid{display:grid}.fds .hidden{display:none}
.fds .flex-col{flex-direction:column}.fds .flex-wrap{flex-wrap:wrap}.fds .flex-1{flex:1 1 0%}
.fds .items-center{align-items:center}.fds .items-end{align-items:flex-end}.fds .items-start{align-items:flex-start}
.fds .justify-between{justify-content:space-between}.fds .justify-center{justify-content:center}.fds .justify-end{justify-content:flex-end}
.fds .self-start{align-self:flex-start}
.fds .gap-1{gap:.25rem}.fds .gap-2{gap:.5rem}.fds .gap-3{gap:.75rem}.fds .gap-4{gap:1rem}
.fds .relative{position:relative}.fds .absolute{position:absolute}.fds .fixed{position:fixed}.fds .sticky{position:sticky}
.fds .inset-0{inset:0}.fds .top-0{top:0}.fds .z-50{z-index:50}
.fds .w-full{width:100%}.fds .w-8{width:2rem}.fds .w-24{width:6rem}.fds .w-28{width:7rem}.fds .min-h-screen{min-height:100vh}
.fds .mx-auto{margin-left:auto;margin-right:auto}.fds .mb-3{margin-bottom:.75rem}.fds .mt-2{margin-top:.5rem}
.fds .p-3{padding:.75rem}.fds .p-4{padding:1rem}.fds .pt-1{padding-top:.25rem}.fds .pt-3{padding-top:.75rem}.fds .pr-1{padding-right:.25rem}
.fds .px-0\\.5{padding-left:.125rem;padding-right:.125rem}.fds .px-1{padding-left:.25rem;padding-right:.25rem}.fds .px-2{padding-left:.5rem;padding-right:.5rem}
.fds .px-3{padding-left:.75rem;padding-right:.75rem}.fds .px-4{padding-left:1rem;padding-right:1rem}.fds .px-5{padding-left:1.25rem;padding-right:1.25rem}
.fds .py-0\\.5{padding-top:.125rem;padding-bottom:.125rem}.fds .py-1{padding-top:.25rem;padding-bottom:.25rem}.fds .py-1\\.5{padding-top:.375rem;padding-bottom:.375rem}
.fds .py-2{padding-top:.5rem;padding-bottom:.5rem}.fds .py-3{padding-top:.75rem;padding-bottom:.75rem}.fds .py-4{padding-top:1rem;padding-bottom:1rem}.fds .py-20{padding-top:5rem;padding-bottom:5rem}
.fds .border{border-width:1px;border-style:solid;border-color:#e5e7eb}.fds .border-b{border-bottom:1px solid #e5e7eb}.fds .border-t{border-top:1px solid #e5e7eb}
.fds .border-gray-300{border-color:#d1d5db}
.fds .rounded{border-radius:.25rem}.fds .rounded-lg{border-radius:.5rem}.fds .rounded-t{border-top-left-radius:.25rem;border-top-right-radius:.25rem}
.fds .shadow{box-shadow:0 1px 3px rgba(0,0,0,.15)}.fds .shadow-lg{box-shadow:0 8px 24px rgba(0,0,0,.15)}.fds .shadow-xl{box-shadow:0 16px 40px rgba(0,0,0,.25)}
.fds .bg-white{background:#fff}.fds .bg-gray-100{background:#f3f4f6}
.fds .overflow-auto{overflow:auto}.fds .overflow-hidden{overflow:hidden}.fds .overflow-x-auto{overflow-x:auto}
.fds .cursor-pointer{cursor:pointer}.fds .pointer-events-none{pointer-events:none}
.fds .text-sm{font-size:.875rem;line-height:1.25rem}.fds .text-lg{font-size:1.125rem}.fds .text-2xl{font-size:1.5rem}
.fds .font-semibold{font-weight:600}.fds .leading-none{line-height:1}.fds .whitespace-nowrap{white-space:nowrap}
.fds .text-left{text-align:left}.fds .text-center{text-align:center}.fds .text-right{text-align:right}
.fds .text-white{color:#fff}.fds .text-gray-500{color:#6b7280}.fds .text-gray-600{color:#4b5563}
.fds button:focus-visible, .fds input:focus-visible, .fds textarea:focus-visible { outline: 2px solid #1b9a94; outline-offset: 1px; }
@media print {
  .fds .no-print { display: none !important; }
  .fds { background: #fff !important; padding: 0 !important; }
  .fds .datasheet-page { box-shadow: none !important; }
  @page { size: A4; margin: 6mm; }
}
`;

/* ================================================================== */
/*  Main                                                               */
/* ================================================================== */
export default function FanDatasheet() {
  const [data, setData] = useState({
    model: "KESO-315",
    curves: SAMPLE_CURVES,
    sheet: DEFAULT_SHEET,
    opts: DEFAULT_OPTS,
  });
  const [editing, setEditing] = useState(false);
  const curves = useMemo(
    () => buildCurves(data.curves, data.opts),
    [data.curves, data.opts],
  );
  const setOpt = (patch) =>
    setData((d) => ({ ...d, opts: { ...d.opts, ...patch } }));

  return (
    <div
      className="fds min-h-screen py-4 w-full"
      style={{ background: "#d9dde1", fontFamily: FONT }}
    >
      <style>{BASE_CSS}</style>

      <div className="no-print flex flex-wrap items-center justify-center gap-2 mb-3 text-sm">
        <button
          className="px-4 py-2 rounded text-white"
          style={{ background: C.rpm }}
          onClick={() => setEditing(true)}
        >
          Edit data
        </button>
        <div className="flex rounded border overflow-hidden bg-white">
          {["Log", "Linear"].map((m) => {
            const on = (m === "Log") === data.opts.log;
            return (
              <button
                key={m}
                className="px-3 py-2"
                onClick={() => setOpt({ log: m === "Log" })}
                style={{
                  background: on ? C.navy : "#fff",
                  color: on ? "#fff" : C.ink,
                }}
              >
                {m}
              </button>
            );
          })}
        </div>
        <button
          className="px-4 py-2 rounded border bg-white"
          onClick={() => window.print()}
        >
          Print / Save PDF
        </button>
      </div>

      <div className="overflow-x-auto">
        <div
          className="datasheet-page shadow-lg mx-auto"
          style={{ width: 920 }}
        >
          <Datasheet
            model={data.model}
            sheet={data.sheet}
            curves={curves}
            opts={data.opts}
          />
        </div>
      </div>

      {editing && (
        <EditorModal
          initial={data}
          onClose={() => setEditing(false)}
          onApply={(next) => {
            setData(next);
            setEditing(false);
          }}
        />
      )}
    </div>
  );
}
