// SVG chart renderer for ReflexGrid: column/bar (clustered, stacked, 100%), line, area, pie,
// doughnut, scatter and radar charts, drawn from the engine's resolved chart data.
import React from "react";

const FONT = "Calibri, Carlito, 'Segoe UI', system-ui, Arial, sans-serif";

function niceTicks(lo, hi, count) {
  if (!isFinite(lo) || !isFinite(hi)) return [0, 1];
  if (lo === hi) {
    if (lo === 0) return [0, 1];
    lo = Math.min(0, lo);
    hi = Math.max(0, hi);
    if (lo === hi) hi = lo + 1;
  }
  const span = hi - lo;
  const raw = span / Math.max(count, 1);
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const norm = raw / mag;
  const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10) * mag;
  const start = Math.floor(lo / step) * step;
  const end = Math.ceil(hi / step) * step;
  const out = [];
  for (let v = start; v <= end + step * 1e-9; v += step) out.push(Math.abs(v) < step * 1e-9 ? 0 : v);
  return out;
}

function fmtTick(v, percent) {
  if (percent) return Math.round(v * 100) + "%";
  const a = Math.abs(v);
  if (a >= 1e9) return (v / 1e9).toFixed(a % 1e9 === 0 ? 0 : 1) + "B";
  if (a >= 1e6) return (v / 1e6).toFixed(a % 1e6 === 0 ? 0 : 1) + "M";
  if (a >= 1e4) return Math.round(v).toLocaleString("en-US");
  if (Number.isInteger(v)) return v.toLocaleString("en-US");
  return String(Math.round(v * 1000) / 1000);
}

function seriesColor(chart, i) {
  const s = chart.series[i];
  if (s && s.color) return s.color;
  const p = chart.palette || ["#156082", "#E97132", "#196B24", "#0F9ED5", "#A02B93", "#4EA72E"];
  const base = p[i % p.length];
  if (i < p.length) return base;
  // Darker shades for series past the six accents.
  const k = Math.floor(i / p.length);
  const n = parseInt(base.slice(1), 16);
  const f = Math.max(0.35, 1 - 0.25 * k);
  const r = Math.round(((n >> 16) & 255) * f), g = Math.round(((n >> 8) & 255) * f), b = Math.round((n & 255) * f);
  return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
}

function Legend({ chart, x, y, w, h, position, items }) {
  if (position === "none" || !items.length) return null;
  const vertical = position === "left" || position === "right";
  const out = [];
  let cx = x, cy = y;
  const fs = 11;
  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    const tw = Math.min(140, it.label.length * 6.2 + 18);
    if (!vertical && cx + tw > x + w && cx > x) {
      cx = x;
      cy += 16;
    }
    out.push(
      <g key={i} transform={`translate(${cx},${cy})`}>
        <rect x={0} y={-8} width={9} height={9} fill={it.color} />
        <text x={13} y={0} fontSize={fs} fill="#404040" fontFamily={FONT}>
          {it.label.length > 22 ? it.label.slice(0, 21) + "…" : it.label}
        </text>
      </g>
    );
    if (vertical) cy += 16;
    else cx += tw + 8;
  }
  return <g>{out}</g>;
}

function legendBox(position, W, H, n) {
  // Space reserved for the legend: [top, right, bottom, left].
  if (position === "none" || n === 0) return [0, 0, 0, 0];
  if (position === "right") return [0, 120, 0, 0];
  if (position === "left") return [0, 0, 0, 120];
  if (position === "top") return [22, 0, 0, 0];
  return [0, 0, 24, 0];
}

function PieChart({ chart, W, H, top, doughnut }) {
  const s = chart.series[0] || { values: [] };
  const vals = s.values.map((v) => (typeof v === "number" && v > 0 ? v : 0));
  const total = vals.reduce((a, b) => a + b, 0);
  const items = chart.categories.map((c, i) => ({ label: String(c), color: seriesColor(chart, i) }));
  const lb = legendBox(chart.legend, W, H, items.length);
  const cx = (W - lb[1] + lb[3]) / 2;
  const cy = top + (H - top - lb[2] + lb[0]) / 2;
  const r = Math.max(10, Math.min(W - lb[1] - lb[3], H - top - lb[0] - lb[2]) / 2 - 12);
  const inner = doughnut ? r * 0.5 : 0;
  let a0 = -Math.PI / 2;
  const slices = [];
  vals.forEach((v, i) => {
    if (total <= 0 || v <= 0) return;
    const a1 = a0 + (v / total) * Math.PI * 2;
    const large = a1 - a0 > Math.PI ? 1 : 0;
    const p = (ang, rad) => [cx + rad * Math.cos(ang), cy + rad * Math.sin(ang)];
    const [x0, y0] = p(a0, r), [x1, y1] = p(a1, r);
    let d;
    if (vals.filter((x) => x > 0).length === 1) {
      d = doughnut
        ? `M ${cx - r} ${cy} A ${r} ${r} 0 1 1 ${cx + r} ${cy} A ${r} ${r} 0 1 1 ${cx - r} ${cy} M ${cx - inner} ${cy} A ${inner} ${inner} 0 1 0 ${cx + inner} ${cy} A ${inner} ${inner} 0 1 0 ${cx - inner} ${cy}`
        : `M ${cx - r} ${cy} A ${r} ${r} 0 1 1 ${cx + r} ${cy} A ${r} ${r} 0 1 1 ${cx - r} ${cy}`;
    } else if (doughnut) {
      const [xi0, yi0] = p(a0, inner), [xi1, yi1] = p(a1, inner);
      d = `M ${x0} ${y0} A ${r} ${r} 0 ${large} 1 ${x1} ${y1} L ${xi1} ${yi1} A ${inner} ${inner} 0 ${large} 0 ${xi0} ${yi0} Z`;
    } else {
      d = `M ${cx} ${cy} L ${x0} ${y0} A ${r} ${r} 0 ${large} 1 ${x1} ${y1} Z`;
    }
    const mid = (a0 + a1) / 2;
    const [lx, ly] = p(mid, doughnut ? (r + inner) / 2 : r * 0.65);
    slices.push(
      <g key={i}>
        <path d={d} fill={seriesColor(chart, i)} stroke="#FFFFFF" strokeWidth={1.5} fillRule="evenodd" />
        {chart.dataLabels && (
          <text x={lx} y={ly + 4} textAnchor="middle" fontSize={11} fill="#FFFFFF" fontFamily={FONT}>
            {Math.round((v / total) * 100) + "%"}
          </text>
        )}
      </g>
    );
    a0 = a1;
  });
  const lx = chart.legend === "right" ? W - lb[1] + 8 : chart.legend === "left" ? 8 : 12;
  const ly = chart.legend === "top" ? top + 12 : chart.legend === "bottom" ? H - 10 : top + 20;
  return (
    <g>
      {slices}
      <Legend chart={chart} x={lx} y={ly} w={chart.legend === "right" || chart.legend === "left" ? 110 : W - 24} h={H} position={chart.legend} items={items} />
    </g>
  );
}

function RadarChart({ chart, W, H, top }) {
  const items = chart.series.map((s, i) => ({ label: s.name, color: seriesColor(chart, i) }));
  const lb = legendBox(chart.legend, W, H, items.length);
  const cx = (W - lb[1] + lb[3]) / 2;
  const cy = top + (H - top - lb[2] + lb[0]) / 2;
  const r = Math.max(10, Math.min(W - lb[1] - lb[3], H - top - lb[0] - lb[2]) / 2 - 24);
  const n = Math.max(chart.categories.length, ...chart.series.map((s) => s.values.length), 3);
  let hi = -Infinity;
  chart.series.forEach((s) => s.values.forEach((v) => { if (typeof v === "number") hi = Math.max(hi, v); }));
  const ticks = niceTicks(0, hi > 0 ? hi : 1, 4);
  const max = ticks[ticks.length - 1] || 1;
  const pt = (i, v) => {
    const a = -Math.PI / 2 + (i / n) * Math.PI * 2;
    const rr = (Math.max(v, 0) / max) * r;
    return [cx + rr * Math.cos(a), cy + rr * Math.sin(a)];
  };
  const rings = ticks.slice(1).map((t, k) => (
    <polygon key={"r" + k} points={Array.from({ length: n }, (_, i) => pt(i, t).join(",")).join(" ")} fill="none" stroke="#D9D9D9" />
  ));
  const spokes = Array.from({ length: n }, (_, i) => {
    const [x, y] = pt(i, max);
    const [lx, ly] = pt(i, max * 1.12);
    return (
      <g key={"s" + i}>
        <line x1={cx} y1={cy} x2={x} y2={y} stroke="#D9D9D9" />
        <text x={lx} y={ly + 4} textAnchor="middle" fontSize={10} fill="#595959" fontFamily={FONT}>
          {String(chart.categories[i] ?? i + 1)}
        </text>
      </g>
    );
  });
  const polys = chart.series.map((s, k) => (
    <polygon key={"p" + k} points={s.values.map((v, i) => pt(i, typeof v === "number" ? v : 0).join(",")).join(" ")} fill="none" stroke={seriesColor(chart, k)} strokeWidth={2} />
  ));
  const lx = chart.legend === "right" ? W - lb[1] + 8 : 12;
  const ly = chart.legend === "top" ? top + 12 : chart.legend === "bottom" ? H - 10 : top + 20;
  return (
    <g>
      {rings}
      {spokes}
      {polys}
      <Legend chart={chart} x={lx} y={ly} w={W - 24} h={H} position={chart.legend} items={items} />
    </g>
  );
}

function XYChart({ chart, W, H, top }) {
  const kind = chart.kind;
  const horizontal = kind.startsWith("bar");
  const stacked = kind.endsWith("Stacked") || kind === "lineStacked" || kind === "areaStacked";
  const pct = kind.endsWith("Stacked100");
  const isStacked = stacked || pct;
  const scatter = kind.startsWith("scatter");
  const isLine = kind.startsWith("line");
  const isArea = kind.startsWith("area");
  const isBar = kind.startsWith("column") || kind.startsWith("bar");
  const series = chart.series;
  const nCat = Math.max(chart.categories.length, ...series.map((s) => s.values.length), 1);
  const items = series.map((s, i) => ({ label: s.name, color: seriesColor(chart, i) }));
  const lb = legendBox(chart.legend, W, H, series.length > 1 || chart.legend !== "none" ? items.length : 0);
  // Value range.
  let lo = Infinity, hi = -Infinity;
  if (isStacked) {
    for (let i = 0; i < nCat; i++) {
      let pos = 0, neg = 0, tot = 0;
      series.forEach((s) => {
        const v = typeof s.values[i] === "number" ? s.values[i] : 0;
        if (v >= 0) pos += v; else neg += v;
        tot += Math.abs(v);
      });
      if (pct) { hi = Math.max(hi, tot > 0 ? pos / tot : 0); lo = Math.min(lo, tot > 0 ? neg / tot : 0); }
      else { hi = Math.max(hi, pos); lo = Math.min(lo, neg); }
    }
  } else {
    series.forEach((s) => s.values.forEach((v) => { if (typeof v === "number") { lo = Math.min(lo, v); hi = Math.max(hi, v); } }));
  }
  if (!isFinite(lo)) { lo = 0; hi = 1; }
  if (isBar || isArea || isStacked) { lo = Math.min(lo, 0); hi = Math.max(hi, 0); }
  const ticks = niceTicks(lo, hi, 5);
  const vmin = ticks[0], vmax = ticks[ticks.length - 1];
  // X range for scatter.
  let xlo = Infinity, xhi = -Infinity;
  if (scatter) {
    chart.categories.forEach((c, i) => { const x = typeof c === "number" ? c : i + 1; xlo = Math.min(xlo, x); xhi = Math.max(xhi, x); });
    if (!isFinite(xlo)) { xlo = 0; xhi = nCat; }
  }
  const xticks = scatter ? niceTicks(xlo, xhi, 6) : null;
  const tickLabelW = Math.max(...ticks.map((t) => fmtTick(t, pct).length)) * 6.5 + 10;
  const left = 8 + lb[3] + (horizontal ? Math.min(110, Math.max(...chart.categories.map((c) => String(c).length), 2) * 6.2 + 10) : tickLabelW) + (chart.yTitle ? 16 : 0);
  const right = W - 12 - lb[1];
  const plotTop = top + 8 + lb[0];
  const bottom = H - 8 - lb[2] - (horizontal ? 16 : 18) - (chart.xTitle ? 16 : 0);
  const pw = Math.max(right - left, 10), ph = Math.max(bottom - plotTop, 10);
  const vscale = (v) => (horizontal ? left + ((v - vmin) / (vmax - vmin || 1)) * pw : bottom - ((v - vmin) / (vmax - vmin || 1)) * ph);
  const els = [];
  // Gridlines and value axis labels.
  ticks.forEach((t, k) => {
    const p = vscale(t);
    if (horizontal) {
      if (chart.gridlines) els.push(<line key={"g" + k} x1={p} y1={plotTop} x2={p} y2={bottom} stroke="#D9D9D9" />);
      els.push(<text key={"t" + k} x={p} y={bottom + 14} textAnchor="middle" fontSize={10} fill="#595959" fontFamily={FONT}>{fmtTick(t, pct)}</text>);
    } else {
      if (chart.gridlines) els.push(<line key={"g" + k} x1={left} y1={p} x2={right} y2={p} stroke="#D9D9D9" />);
      els.push(<text key={"t" + k} x={left - 6} y={p + 4} textAnchor="end" fontSize={10} fill="#595959" fontFamily={FONT}>{fmtTick(t, pct)}</text>);
    }
  });
  // Category axis.
  const band = (horizontal ? ph : pw) / nCat;
  const catPos = (i) => (horizontal ? plotTop + band * (i + 0.5) : left + band * (i + 0.5));
  const xscale = (x) => left + ((x - xticks[0]) / ((xticks[xticks.length - 1] - xticks[0]) || 1)) * pw;
  if (scatter) {
    xticks.forEach((t, k) => {
      const p = xscale(t);
      if (chart.gridlines) els.push(<line key={"xg" + k} x1={p} y1={plotTop} x2={p} y2={bottom} stroke="#EDEDED" />);
      els.push(<text key={"xt" + k} x={p} y={bottom + 14} textAnchor="middle" fontSize={10} fill="#595959" fontFamily={FONT}>{fmtTick(t, false)}</text>);
    });
  } else {
    const every = Math.max(1, Math.ceil(nCat / Math.max(1, (horizontal ? ph : pw) / 40)));
    for (let i = 0; i < nCat; i++) {
      if (i % every) continue;
      const label = String(chart.categories[i] ?? i + 1);
      const short = label.length > 14 ? label.slice(0, 13) + "…" : label;
      if (horizontal) els.push(<text key={"c" + i} x={left - 6} y={catPos(i) + 4} textAnchor="end" fontSize={10} fill="#595959" fontFamily={FONT}>{short}</text>);
      else els.push(<text key={"c" + i} x={catPos(i)} y={bottom + 14} textAnchor="middle" fontSize={10} fill="#595959" fontFamily={FONT}>{short}</text>);
    }
  }
  const zero = vscale(Math.min(Math.max(0, vmin), vmax));
  els.push(horizontal ? <line key="axis" x1={zero} y1={plotTop} x2={zero} y2={bottom} stroke="#BFBFBF" /> : <line key="axis" x1={left} y1={zero} x2={right} y2={zero} stroke="#BFBFBF" />);
  // Data.
  if (isBar) {
    const groupW = band * 0.7;
    const barW = isStacked ? groupW : groupW / Math.max(series.length, 1);
    const base = [], baseNeg = [];
    for (let i = 0; i < nCat; i++) {
      let tot = 0;
      series.forEach((s) => { const v = s.values[i]; if (typeof v === "number") tot += Math.abs(v); });
      series.forEach((s, k) => {
        let v = s.values[i];
        if (typeof v !== "number") return;
        if (pct) v = tot > 0 ? v / tot : 0;
        let a, b;
        if (isStacked) {
          const p0 = v >= 0 ? base[i] || 0 : baseNeg[i] || 0;
          a = p0; b = p0 + v;
          if (v >= 0) base[i] = b; else baseNeg[i] = b;
        } else { a = 0; b = v; }
        const pa = vscale(Math.min(Math.max(a, vmin), vmax)), pb = vscale(Math.min(Math.max(b, vmin), vmax));
        const off = catPos(i) - groupW / 2 + (isStacked ? 0 : k * barW);
        const rect = horizontal
          ? { x: Math.min(pa, pb), y: off, width: Math.abs(pb - pa), height: Math.max(barW - 1, 1) }
          : { x: off, y: Math.min(pa, pb), width: Math.max(barW - 1, 1), height: Math.abs(pb - pa) };
        els.push(<rect key={`b${k}_${i}`} {...rect} fill={seriesColor(chart, k)} />);
        if (chart.dataLabels) {
          const lx = horizontal ? rect.x + rect.width + 3 : rect.x + rect.width / 2;
          const ly = horizontal ? rect.y + rect.height / 2 + 4 : rect.y - 3;
          els.push(<text key={`l${k}_${i}`} x={lx} y={ly} textAnchor={horizontal ? "start" : "middle"} fontSize={9} fill="#404040" fontFamily={FONT}>{fmtTick(s.values[i], false)}</text>);
        }
      });
    }
  } else {
    const acc = [];
    series.forEach((s, k) => {
      const pts = [];
      for (let i = 0; i < nCat; i++) {
        let v = s.values[i];
        if (typeof v !== "number") { pts.push(null); continue; }
        if (isStacked) { acc[i] = (acc[i] || 0) + v; v = acc[i]; }
        const x = scatter ? xscale(typeof chart.categories[i] === "number" ? chart.categories[i] : i + 1) : catPos(i);
        pts.push([x, vscale(v), s.values[i]]);
      }
      const color = seriesColor(chart, k);
      const valid = pts.filter(Boolean);
      if (isArea && valid.length) {
        const d = "M " + valid.map((p) => `${p[0]} ${p[1]}`).join(" L ") + ` L ${valid[valid.length - 1][0]} ${zero} L ${valid[0][0]} ${zero} Z`;
        els.push(<path key={"a" + k} d={d} fill={color} fillOpacity={isStacked ? 0.85 : 0.6} stroke={color} />);
      } else if (!scatter || kind === "scatterLines") {
        let d = "";
        let pen = false;
        pts.forEach((p) => { if (!p) { pen = false; return; } d += (pen ? " L " : " M ") + `${p[0]} ${p[1]}`; pen = true; });
        els.push(<path key={"l" + k} d={d} fill="none" stroke={color} strokeWidth={2.25} strokeLinejoin="round" />);
      }
      if (scatter || kind === "lineMarkers") {
        valid.forEach((p, i) => els.push(<circle key={`m${k}_${i}`} cx={p[0]} cy={p[1]} r={3.5} fill={color} stroke="#FFFFFF" strokeWidth={0.75} />));
      }
      if (chart.dataLabels) valid.forEach((p, i) => els.push(<text key={`dl${k}_${i}`} x={p[0]} y={p[1] - 6} textAnchor="middle" fontSize={9} fill="#404040" fontFamily={FONT}>{fmtTick(p[2], false)}</text>));
    });
  }
  if (chart.xTitle) els.push(<text key="xt" x={left + pw / 2} y={H - 8 - lb[2]} textAnchor="middle" fontSize={11} fill="#404040" fontFamily={FONT}>{chart.xTitle}</text>);
  if (chart.yTitle) els.push(<text key="yt" x={0} y={0} transform={`translate(${12 + lb[3]},${plotTop + ph / 2}) rotate(-90)`} textAnchor="middle" fontSize={11} fill="#404040" fontFamily={FONT}>{chart.yTitle}</text>);
  const showLegend = chart.legend !== "none";
  const lx = chart.legend === "right" ? W - lb[1] + 4 : chart.legend === "left" ? 8 : left;
  const ly = chart.legend === "top" ? top + 14 : chart.legend === "bottom" ? H - 10 : plotTop + 10;
  return (
    <g>
      {els}
      {showLegend && <Legend chart={chart} x={lx} y={ly} w={chart.legend === "right" || chart.legend === "left" ? 110 : pw} h={H} position={chart.legend} items={items} />}
    </g>
  );
}

export function ChartSvg({ chart, width, height }) {
  const W = Math.max(width, 60), H = Math.max(height, 40);
  const top = chart.title ? 30 : 8;
  let body;
  if (!chart.series || !chart.series.length) {
    body = <text x={W / 2} y={H / 2} textAnchor="middle" fontSize={12} fill="#7F7F7F" fontFamily={FONT}>No data</text>;
  } else if (chart.kind === "pie" || chart.kind === "doughnut") {
    body = <PieChart chart={chart} W={W} H={H} top={top} doughnut={chart.kind === "doughnut"} />;
  } else if (chart.kind === "radar") {
    body = <RadarChart chart={chart} W={W} H={H} top={top} />;
  } else {
    body = <XYChart chart={chart} W={W} H={H} top={top} />;
  }
  return (
    <svg width={W} height={H} viewBox={`0 0 ${W} ${H}`} style={{ display: "block", background: "#FFFFFF" }}>
      {chart.title && (
        <text x={W / 2} y={21} textAnchor="middle" fontSize={14} fill="#404040" fontFamily={FONT}>
          {chart.title}
        </text>
      )}
      {body}
    </svg>
  );
}

// Sparkline inside a cell.
export function Sparkline({ spark, width, height }) {
  const vals = spark.values || [];
  const nums = vals.filter((v) => typeof v === "number");
  if (!nums.length) return null;
  const pad = 3;
  const w = Math.max(width - pad * 2, 4), h = Math.max(height - pad * 2, 4);
  const lo = Math.min(...nums, spark.kind === "line" ? Infinity : 0), hi = Math.max(...nums, spark.kind === "line" ? -Infinity : 0);
  const y = (v) => pad + h - ((v - lo) / ((hi - lo) || 1)) * h;
  const n = vals.length;
  const color = spark.color || "#156082";
  if (spark.kind === "column" || spark.kind === "stacked") {
    const bw = w / n;
    return (
      <svg width={width} height={height} style={{ position: "absolute", left: 0, top: 0 }}>
        {vals.map((v, i) => {
          if (typeof v !== "number") return null;
          if (spark.kind === "stacked") {
            const mid = pad + h / 2;
            return <rect key={i} x={pad + i * bw + bw * 0.1} y={v >= 0 ? pad : mid} width={bw * 0.8} height={h / 2} fill={v >= 0 ? color : "#C00000"} />;
          }
          const y0 = y(Math.max(0, lo)), y1 = y(v);
          return <rect key={i} x={pad + i * bw + bw * 0.1} y={Math.min(y0, y1)} width={bw * 0.8} height={Math.max(Math.abs(y1 - y0), 1)} fill={v >= 0 ? color : "#C00000"} />;
        })}
      </svg>
    );
  }
  let d = "";
  let pen = false;
  vals.forEach((v, i) => {
    if (typeof v !== "number") { pen = false; return; }
    const x = pad + (n > 1 ? (i / (n - 1)) * w : w / 2);
    d += (pen ? " L " : " M ") + x + " " + y(v);
    pen = true;
  });
  return (
    <svg width={width} height={height} style={{ position: "absolute", left: 0, top: 0 }}>
      <path d={d} fill="none" stroke={color} strokeWidth={1.25} />
    </svg>
  );
}

// Conditional-format icons (our own drawings).
export function CfIcon({ set, index, size }) {
  const s = size || 14;
  const n = parseInt(set, 10) || 3;
  const name = set.replace(/^\d/, "");
  const trafficColors = n === 4 ? ["#C00000", "#FFC000", "#00B050", "#262626"] : ["#E04B4B", "#F2C12E", "#3DA35D"];
  let el;
  const levelColor = (i) => {
    const pal3 = ["#E04B4B", "#F2C12E", "#3DA35D"];
    if (n === 3) return pal3[i];
    const pal5 = ["#E04B4B", "#F08A3C", "#F2C12E", "#9BC53D", "#3DA35D"];
    if (n === 4) return ["#E04B4B", "#F2C12E", "#9BC53D", "#3DA35D"][i];
    return pal5[i];
  };
  if (name.startsWith("Arrows")) {
    const gray = name.includes("Gray");
    const dirs = n === 3 ? [180, 90, 0] : n === 4 ? [180, 135, 45, 0] : [180, 135, 90, 45, 0];
    const rot = dirs[index] ?? 90;
    el = (
      <g transform={`rotate(${rot - 90} 7 7)`}>
        <path d="M7 1 L12.5 7 L9 7 L9 13 L5 13 L5 7 L1.5 7 Z" transform="rotate(90 7 7)" fill={gray ? "#7F7F7F" : levelColor(index)} />
      </g>
    );
  } else if (name.startsWith("TrafficLights") || name === "Signs" || name === "RedToBlack") {
    const c = name === "RedToBlack" ? ["#262626", "#7F7F7F", "#D9A0A0", "#E04B4B"][index] : n === 4 ? ["#262626", "#E04B4B", "#F2C12E", "#3DA35D"][index] : trafficColors[index];
    el = name === "Signs"
      ? (index === 0 ? <polygon points="7,1 13,7 7,13 1,7" fill={c} /> : index === 1 ? <polygon points="7,1.5 13,12.5 1,12.5" fill={c} /> : <circle cx={7} cy={7} r={6} fill={c} />)
      : <circle cx={7} cy={7} r={6} fill={c} stroke="#00000022" />;
  } else if (name.startsWith("Symbols")) {
    const c = ["#E04B4B", "#F2C12E", "#3DA35D"][index];
    const glyph = index === 0 ? "M4 4 L10 10 M10 4 L4 10" : index === 1 ? "M7 3 L7 8.5 M7 10.5 L7 11" : "M3.5 7.5 L6 10 L10.5 4";
    el = name === "Symbols"
      ? (<g><circle cx={7} cy={7} r={6.2} fill={c} /><path d={glyph} stroke="#FFFFFF" strokeWidth={2} fill="none" strokeLinecap="round" /></g>)
      : <path d={glyph} stroke={c} strokeWidth={2.4} fill="none" strokeLinecap="round" />;
  } else if (name === "Flags") {
    const c = ["#E04B4B", "#F2C12E", "#3DA35D"][index];
    el = <g><path d="M3 1.5 L3 13" stroke="#404040" strokeWidth={1.2} /><path d="M3.5 2 L12 4.5 L3.5 7.5 Z" fill={c} /></g>;
  } else if (name === "Stars") {
    const fill = index === 2 ? "#F2C12E" : index === 1 ? "url(#half)" : "#FFFFFF";
    el = (
      <g>
        <defs><linearGradient id="half"><stop offset="50%" stopColor="#F2C12E" /><stop offset="50%" stopColor="#FFFFFF" /></linearGradient></defs>
        <polygon points="7,1 8.8,5.2 13.3,5.5 9.8,8.4 10.9,12.8 7,10.4 3.1,12.8 4.2,8.4 0.7,5.5 5.2,5.2" fill={fill} stroke="#C9A227" strokeWidth={0.8} />
      </g>
    );
  } else if (name === "Triangles") {
    el = index === 2 ? <polygon points="7,2 12.5,11 1.5,11" fill="#3DA35D" /> : index === 1 ? <rect x={2} y={5.5} width={10} height={3} fill="#F2C12E" /> : <polygon points="1.5,3 12.5,3 7,12" fill="#E04B4B" />;
  } else if (name === "Rating") {
    const bars = [];
    for (let i = 0; i < 4; i++) bars.push(<rect key={i} x={1 + i * 3.2} y={10 - i * 2.5} width={2.4} height={3 + i * 2.5} fill={i < index + (n === 5 ? 0 : 1) ? "#3A6EA5" : "#D0D7E2"} />);
    el = <g>{bars}</g>;
  } else if (name === "Quarters") {
    const frac = index / 4;
    const a = frac * Math.PI * 2 - Math.PI / 2;
    const x = 7 + 6 * Math.cos(a), y = 7 + 6 * Math.sin(a);
    el = (
      <g>
        <circle cx={7} cy={7} r={6} fill="#FFFFFF" stroke="#404040" />
        {index === 4 ? <circle cx={7} cy={7} r={6} fill="#404040" /> : index > 0 ? <path d={`M7 7 L7 1 A6 6 0 ${frac > 0.5 ? 1 : 0} 1 ${x} ${y} Z`} fill="#404040" /> : null}
      </g>
    );
  } else if (name === "Boxes") {
    const cells = [];
    for (let i = 0; i < 4; i++) cells.push(<rect key={i} x={1 + (i % 2) * 6.5} y={1 + Math.floor(i / 2) * 6.5} width={5.5} height={5.5} fill={i < index ? "#3A6EA5" : "#D0D7E2"} />);
    el = <g>{cells}</g>;
  } else {
    el = <circle cx={7} cy={7} r={6} fill={levelColor(Math.min(index, n - 1))} />;
  }
  return (
    <svg width={s} height={s} viewBox="0 0 14 14" style={{ flex: "none" }}>
      {el}
    </svg>
  );
}
