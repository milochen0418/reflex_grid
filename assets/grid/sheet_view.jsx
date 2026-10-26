// ReflexGrid sheet view: name box + formula bar + a virtualized grid (canvas for gridlines,
// fills and borders; DOM for text, selection, editor and objects). Talks to the Python engine
// through one event channel: onEvent({type, seq, ...}).
import React, { useState, useEffect, useRef, useMemo, useCallback, useLayoutEffect } from "react";
import { getBackendURL, getToken } from "$/utils/state";
import env from "$/env.json";
import { ChartSvg, Sparkline, CfIcon } from "./sheet_charts.jsx";

const MAXR = 1048576;
const MAXC = 16384;
const ACCENT = "#156082";
const SEL_FILL = "rgba(21, 96, 130, 0.12)";
const GRID = "#DADCE0";
const HEAD_BG = "#F3F4F6";
const HEAD_SEL = "#D7E3EA";
const REF_COLORS = ["#2F6FDB", "#D9480F", "#7A3FB8", "#1E8E4E", "#C2255C", "#B07A00", "#0B7285", "#5C940D"];
const IS_MAC = typeof navigator !== "undefined" && /Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent || "");

// ---------------------------------------------------------------- helpers

export function colName(c) {
  let s = "";
  c += 1;
  while (c > 0) {
    const m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

function a1(r, c) {
  return colName(c) + (r + 1);
}

function rangeA1(g) {
  const [r0, c0, r1, c1] = g;
  if (c0 === 0 && c1 === MAXC - 1 && !(r0 === 0 && r1 === MAXR - 1)) return `${r0 + 1}:${r1 + 1}`;
  if (r0 === 0 && r1 === MAXR - 1 && !(c0 === 0 && c1 === MAXC - 1)) return `${colName(c0)}:${colName(c1)}`;
  if (r0 === r1 && c0 === c1) return a1(r0, c0);
  return `${a1(r0, c0)}:${a1(r1, c1)}`;
}

function norm(a, b) {
  return [Math.min(a[0], b[0]), Math.min(a[1], b[1]), Math.max(a[0], b[0]), Math.max(a[1], b[1])];
}

function contains(g, r, c) {
  return g[0] <= r && r <= g[2] && g[1] <= c && c <= g[3];
}

function intersects(a, b) {
  return a[0] <= b[2] && b[0] <= a[2] && a[1] <= b[3] && b[1] <= a[3];
}

function key(r, c) {
  return r * MAXC + c;
}

function fontFamily(name) {
  const n = (name || "Calibri").replace(/'/g, "");
  if (/^calibri/i.test(n)) return "Calibri, Carlito, 'Segoe UI', system-ui, Arial, sans-serif";
  if (/^aptos/i.test(n)) return "Aptos, Carlito, 'Segoe UI', system-ui, Arial, sans-serif";
  if (/^cambria/i.test(n)) return "Cambria, Caladea, Georgia, serif";
  if (/times/i.test(n)) return "'Times New Roman', Tinos, Times, serif";
  if (/courier|consolas|mono/i.test(n)) return `'${n}', Consolas, 'Courier New', monospace`;
  return `'${n}', Carlito, 'Segoe UI', system-ui, Arial, sans-serif`;
}

// Browsers draw Calibri/Carlito digits at their exact width (7.4px at 11pt); Excel snaps them to
// 7px at 100%. Numbers are judged against the column width with Excel's metrics.
const EXCEL_DIGIT_FIT = 0.94;

class Lines {
  constructor(custom, def, max) {
    this.def = def;
    this.max = max;
    const keys = Object.keys(custom || {})
      .map(Number)
      .filter((k) => k >= 0 && k < max)
      .sort((a, b) => a - b);
    this.keys = keys;
    this.pos = new Float64Array(keys.length);
    this.size = new Float64Array(keys.length);
    let pos = 0;
    let prev = 0;
    keys.forEach((k, i) => {
      pos += (k - prev) * def;
      this.pos[i] = pos;
      const s = Number(custom[k]) || 0;
      this.size[i] = s;
      pos += s;
      prev = k + 1;
    });
  }
  idx(k) {
    let lo = 0,
      hi = this.keys.length;
    while (lo < hi) {
      const m = (lo + hi) >> 1;
      if (this.keys[m] < k) lo = m + 1;
      else hi = m;
    }
    return lo;
  }
  start(k) {
    const i = this.idx(k);
    if (i < this.keys.length && this.keys[i] === k) return this.pos[i];
    if (i === 0) return k * this.def;
    const l = this.keys[i - 1];
    return this.pos[i - 1] + this.size[i - 1] + (k - l - 1) * this.def;
  }
  sizeOf(k) {
    const i = this.idx(k);
    return i < this.keys.length && this.keys[i] === k ? this.size[i] : this.def;
  }
  at(p) {
    if (p <= 0) return 0;
    let lo = 0,
      hi = this.keys.length;
    while (lo < hi) {
      const m = (lo + hi) >> 1;
      if (this.pos[m] <= p) lo = m + 1;
      else hi = m;
    }
    let baseLine = 0,
      basePos = 0;
    if (lo > 0) {
      const i = lo - 1;
      const l = this.keys[i],
        ps = this.pos[i],
        s = this.size[i];
      if (p < ps + s) return Math.min(l, this.max - 1);
      baseLine = l + 1;
      basePos = ps + s;
    }
    if (this.def <= 0) return Math.min(baseLine, this.max - 1);
    return Math.min(baseLine + Math.floor((p - basePos) / this.def), this.max - 1);
  }
}

function fontProps(st, z, scale) {
  const fs = (st.fs || 11) * (96 / 72) * z * (scale || 1);
  return { fontStyle: st.i ? "italic" : "normal", fontWeight: st.b ? 700 : 400, fontSize: fs + "px", fontFamily: fontFamily(st.ff) };
}

let measureCtx = null;
function textWidth(text, font) {
  if (!measureCtx) {
    const c = document.createElement("canvas");
    measureCtx = c.getContext("2d");
  }
  measureCtx.font = font;
  return measureCtx.measureText(text).width;
}

function borderSpec(spec) {
  // [style, color] -> {w, dash, double}
  const st = spec[0];
  const color = spec[1] || "#000000";
  switch (st) {
    case "medium":
      return { w: 2, color };
    case "thick":
      return { w: 3, color };
    case "double":
      return { w: 1, color, double: true };
    case "dashed":
      return { w: 1, color, dash: [3, 1] };
    case "dotted":
      return { w: 1, color, dash: [1, 1] };
    case "hair":
      return { w: 1, color, dash: [1, 1], alpha: 0.6 };
    case "mediumDashed":
      return { w: 2, color, dash: [6, 2] };
    case "dashDot":
      return { w: 1, color, dash: [5, 2, 1, 2] };
    case "mediumDashDot":
    case "slantDashDot":
      return { w: 2, color, dash: [6, 2, 2, 2] };
    case "dashDotDot":
      return { w: 1, color, dash: [5, 2, 1, 2, 1, 2] };
    case "mediumDashDotDot":
      return { w: 2, color, dash: [6, 2, 2, 2, 2, 2] };
    default:
      return { w: 1, color };
  }
}

// Draws one border along an edge: horizontal (y fixed) or vertical (x fixed). `at` is the pixel
// of the gridline the border sits on.
function drawEdge(ctx, spec, horizontal, at, from, to) {
  const b = borderSpec(spec);
  ctx.fillStyle = b.color;
  ctx.globalAlpha = b.alpha || 1;
  const len = to - from;
  if (len <= 0) return;
  const lines = b.double ? [at - 1, at + 1] : [at - Math.floor((b.w - 1) / 2)];
  for (const p of lines) {
    if (b.dash) {
      let x = from;
      let k = 0;
      while (x < to) {
        const seg = b.dash[k % b.dash.length];
        if (k % 2 === 0) {
          const s = Math.min(seg, to - x);
          if (horizontal) ctx.fillRect(x, p, s, b.w);
          else ctx.fillRect(p, x, b.w, s);
        }
        x += seg;
        k++;
      }
    } else if (horizontal) ctx.fillRect(from, p, len, b.double ? 1 : b.w);
    else ctx.fillRect(p, from, b.double ? 1 : b.w, len);
  }
  ctx.globalAlpha = 1;
}

// How far browsers scroll a box (they stop a little past 15 million px).
const SCROLL_CAP = 15_000_000;

// References in formula text, for colour highlighting while editing: cells and ranges, whole columns
// (C:C) and whole rows (4:4).
const REF_RE = /(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)?!?(?:\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}(?![A-Za-z0-9_(])|\$?\d{1,7}:\$?\d{1,7}(?![0-9]))/g;
function formulaRefs(text) {
  if (!text || text[0] !== "=") return [];
  const out = [];
  let m;
  REF_RE.lastIndex = 0;
  // Skip text inside string literals.
  const masked = text.replace(/"[^"]*"?/g, (s) => " ".repeat(s.length));
  while ((m = REF_RE.exec(masked))) {
    const s = m[0];
    const before = masked[m.index - 1];
    if (before && /[A-Za-z0-9_.$]/.test(before)) continue;
    const after = masked[m.index + s.length];
    if (after === "(") continue;
    if (s.includes("!")) continue; // other sheets are not drawn
    const parts = s.split(":");
    const col = (t) => {
      let c = 0;
      for (const ch of t.toUpperCase()) c = c * 26 + (ch.charCodeAt(0) - 64);
      return c - 1;
    };
    const p = (t) => {
      const mm = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(t);
      return mm ? [parseInt(mm[2], 10) - 1, col(mm[1])] : null;
    };
    const cs = parts.map((t) => /^\$?([A-Za-z]{1,3})$/.exec(t));
    const rs = parts.map((t) => /^\$?(\d+)$/.exec(t));
    let range = null;
    if (parts.length === 2 && cs[0] && cs[1]) {
      const c0 = col(cs[0][1]), c1 = col(cs[1][1]);
      if (c0 < MAXC && c1 < MAXC) range = [0, Math.min(c0, c1), MAXR - 1, Math.max(c0, c1)];
    } else if (parts.length === 2 && rs[0] && rs[1]) {
      const r0 = parseInt(rs[0][1], 10) - 1, r1 = parseInt(rs[1][1], 10) - 1;
      if (r0 >= 0 && r1 >= 0 && r0 < MAXR && r1 < MAXR) range = [Math.min(r0, r1), 0, Math.max(r0, r1), MAXC - 1];
    } else {
      const a = p(parts[0]);
      const b = parts[1] ? p(parts[1]) : a;
      if (a && b && a[1] < MAXC && b[1] < MAXC) range = norm(a, b);
    }
    if (!range) continue;
    out.push({ start: m.index, end: m.index + s.length, range });
  }
  return out;
}

// Whether a reference can be inserted at `pos` (point mode): after an operator, "(" or ",".
function canPoint(text, pos) {
  if (!text || text[0] !== "=") return false;
  const before = text.slice(0, pos).replace(/\s+$/, "");
  const ch = before[before.length - 1];
  return ch !== undefined && "=(,+-*/^&<>:;%{".includes(ch);
}

function toggleAbs(ref) {
  const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(ref);
  if (!m) return ref;
  const [, a, col, b, row] = m;
  // A1 -> $A$1 -> A$1 -> $A1 -> A1
  if (!a && !b) return `$${col}$${row}`;
  if (a && b) return `${col}$${row}`;
  if (!a && b) return `$${col}${row}`;
  return `${col}${row}`;
}

// ---------------------------------------------------------------- component

export function SheetView(props) {
  const { geometry, data, objects, ui, functions, names, onEvent, focusKey, editRequest, peers, rev, changes, me } = props;
  const g = geometry || {};
  const z = (g.zoom || 100) / 100;
  const rows = useMemo(() => new Lines(g.rows, g.defaultRowHeight || 20, MAXR), [g.rows, g.defaultRowHeight]);
  const cols = useMemo(() => new Lines(g.cols, g.defaultColWidth || 64, MAXC), [g.cols, g.defaultColWidth]);
  const fr = (g.freeze && g.freeze[0]) || 0;
  const fc = (g.freeze && g.freeze[1]) || 0;
  const headings = g.headings !== false;
  const sheetIndex = ui && ui.sheet !== undefined ? ui.sheet : 0;
  // Sheets are followed by name: someone else adding or removing a sheet moves the index.
  const sheetName = ui && ui.sheetName !== undefined ? ui.sheetName : ui && ui.sheets && ui.sheets[sheetIndex] ? ui.sheets[sheetIndex].name : "";

  const scrollerRef = useRef(null);
  const viewRef = useRef(null);
  const canvasRef = useRef(null);
  const kbdRef = useRef(null);
  const barRef = useRef(null);
  const [view, setView] = useState({ w: 800, h: 500 });
  const [scroll, setScroll] = useState({ top: 0, left: 0 });
  const [sel, setSelState] = useState({ active: [0, 0], anchor: [0, 0], ranges: [[0, 0, 0, 0]] });
  const selRef = useRef(sel);
  const [ed, setEdState] = useState(null);
  const edRef = useRef(null);
  const [pending, setPending] = useState({});
  const [menu, setMenu] = useState(null);
  const [drag, setDrag] = useState(null); // visual state of resize/fill/move drags
  const dragRef = useRef(null);
  const [objSel, setObjSel] = useState(null);
  const [objDrag, setObjDrag] = useState(null);
  const [hover, setHover] = useState(null); // [r, c] of the cell under the pointer, when it has a note, comment or link
  const [listOpen, setListOpen] = useState(false);
  const [ac, setAc] = useState(null); // function autocomplete {items, index, start}
  const pointRef = useRef(null); // formula point mode {start, end, anchor, cur}
  const seqRef = useRef(0);
  // What this tab last copied in ReflexGrid: the text it put on the system clipboard and
  // the server's token for it. On paste, a system clipboard still equal to this is our own
  // copy, so only the token is sent, not the cells (a big block would not fit in one event).
  const lastClipRef = useRef(null);
  const pendingCopyRef = useRef(null);
  const uiRef = useRef(ui);
  uiRef.current = ui;
  const lastWindowRef = useRef(null);
  const sheetRef = useRef(sheetName);
  const sheetIdxRef = useRef(sheetIndex);
  const onEventRef = useRef(onEvent);
  onEventRef.current = onEvent;
  const composingRef = useRef(false);
  const extentRef = useRef({ r: 100, c: 26 });
  const autoScrollRef = useRef(null);

  const [fontTick, setFontTick] = useState(0);
  useEffect(() => {
    if (!document.fonts) return;
    let alive = true;
    Promise.all(["400 15px Carlito", "700 15px Carlito", "italic 400 15px Carlito"].map((f) => document.fonts.load(f).catch(() => null)))
      .then(() => document.fonts.ready)
      .then(() => alive && setFontTick((t) => t + 1));
    return () => { alive = false; };
  }, []);

  const setSel = useCallback((s) => {
    selRef.current = s;
    setSelState(s);
  }, []);
  const setEd = useCallback((e) => {
    edRef.current = e;
    setEdState(e);
  }, []);

  // Where the caret goes once React has put new editor text in place (synchronously after the
  // render, so fast typing can't land before it).
  const caretRef = useRef(null);
  useLayoutEffect(() => {
    const c = caretRef.current;
    if (!c || !ed) return;
    caretRef.current = null;
    const el = ed.bar ? barRef.current : kbdRef.current;
    if (!el) return;
    if (document.activeElement !== el) el.focus({ preventScroll: true });
    const p = Math.min(c.pos, el.value.length);
    try {
      el.setSelectionRange(p, p);
    } catch (_) {}
  }, [ed]);

  const send = useCallback(
    (type, payload) => {
      const seq = ++seqRef.current;
      if (onEvent) onEvent({ type, seq, ...(payload || {}) });
      return seq;
    },
    [onEvent]
  );

  // ------------------------------------------------------------ data index

  const cellMap = useMemo(() => {
    const m = new Map();
    if (data && data.cells && data.sheet === sheetIndex) for (const e of data.cells) m.set(key(e.r, e.c), e);
    return m;
  }, [data, sheetIndex]);
  const styles = (data && data.styles) || {};
  const merges = (data && data.sheet === sheetIndex && data.merges) || [];
  const mergeAt = useCallback(
    (r, c) => {
      for (const m of merges) if (contains(m, r, c)) return m;
      return null;
    },
    [merges]
  );
  const noteMap = useMemo(() => {
    const m = new Map();
    if (data && data.notes && data.sheet === sheetIndex) for (const n of data.notes) m.set(key(n.r, n.c), n);
    return m;
  }, [data, sheetIndex]);
  const linkMap = useMemo(() => {
    const m = new Map();
    if (data && data.links && data.sheet === sheetIndex) for (const l of data.links) m.set(key(l.r, l.c), l);
    return m;
  }, [data, sheetIndex]);
  const sparkMap = useMemo(() => {
    const m = new Map();
    if (data && data.sparklines && data.sheet === sheetIndex) for (const s of data.sparklines) m.set(key(s.r, s.c), s);
    return m;
  }, [data, sheetIndex]);
  const rowStyles = (data && data.rowStyles) || {};
  const colStyles = (data && data.colStyles) || {};

  // ------------------------------------------------------------ geometry

  const maxRowVisible = Math.max(sel.active[0], (ui && ui.selection && ui.selection.active && ui.selection.active[0]) || 0);
  const digits = String(Math.min(MAXR, Math.max(rows.at(rows.start(fr) + (scroll.top + view.h) / z) + 1, 100))).length;
  const HH = headings ? Math.round(20 * z) : 0;
  const RW = headings ? Math.round(Math.max(28, 9 + digits * 7.2) * z) : 0;
  const FH = rows.start(fr) * z;
  const FW = cols.start(fc) * z;
  const mainX = RW + FW;
  const mainY = HH + FH;
  const mainW = Math.max(view.w - mainX, 0);
  const mainH = Math.max(view.h - mainY, 0);

  const ry = useCallback((r) => (r < fr ? HH + rows.start(r) * z : mainY + (rows.start(r) - rows.start(fr)) * z - scroll.top), [fr, HH, rows, z, mainY, scroll.top]);
  const cx = useCallback((c) => (c < fc ? RW + cols.start(c) * z : mainX + (cols.start(c) - cols.start(fc)) * z - scroll.left), [fc, RW, cols, z, mainX, scroll.left]);
  const rh = useCallback((r) => rows.sizeOf(r) * z, [rows, z]);
  const cw = useCallback((c) => cols.sizeOf(c) * z, [cols, z]);

  const mainR0 = rows.at(rows.start(fr) + scroll.top / z);
  const mainR1 = rows.at(rows.start(fr) + (scroll.top + mainH) / z);
  const mainC0 = cols.at(cols.start(fc) + scroll.left / z);
  const mainC1 = cols.at(cols.start(fc) + (scroll.left + mainW) / z);

  // Scrollable extent grows with the used range, the selection and how far the user went.
  const ext = extentRef.current;
  ext.r = Math.max(ext.r, (g.usedRows || 0) + 60, mainR1 + 40, maxRowVisible + 40);
  ext.c = Math.max(ext.c, (g.usedCols || 0) + 12, mainC1 + 8, sel.active[1] + 8);
  ext.r = Math.min(ext.r, MAXR);
  ext.c = Math.min(ext.c, MAXC);
  // Browsers can't scroll much past 15 million px (750,000 default rows at 100%). A sheet taller
  // than that scrolls its first half one to one and the rest of its rows squeezed into the other
  // half, so every row up to 1,048,576 can be reached; scroll.top is in sheet pixels either way.
  const extH = mainY + (rows.start(ext.r) - rows.start(fr)) * z + 2;
  const fullH = mainY + (rows.start(MAXR) - rows.start(fr)) * z + 2;
  const squeezed = fullH > SCROLL_CAP && extH - view.h > SCROLL_CAP / 2;
  const totalH = squeezed ? SCROLL_CAP : Math.min(extH, SCROLL_CAP);
  const totalW = Math.min(mainX + (cols.start(ext.c) - cols.start(fc)) * z + 2, SCROLL_CAP);
  const vmap = useMemo(() => {
    if (!squeezed) return { toSheet: (s) => s, toScroll: (t) => t };
    const half = SCROLL_CAP / 2;
    const dom = Math.max(SCROLL_CAP - view.h - half, 1);
    const sheet = Math.max(fullH - view.h - half, 1);
    return {
      toSheet: (s) => (s <= half ? s : half + ((s - half) * sheet) / dom),
      toScroll: (t) => (t <= half ? t : half + ((t - half) * dom) / sheet),
    };
  }, [squeezed, fullH, view.h]);
  const vmapRef = useRef(vmap);
  vmapRef.current = vmap;

  const hit = useCallback(
    (x, y) => {
      // Screen (view-relative) point -> [row, col]; -1 for headers.
      let r, c;
      if (y < HH) r = -1;
      else if (y < mainY) r = rows.at((y - HH) / z);
      else r = rows.at(rows.start(fr) + (y - mainY + scroll.top) / z);
      if (x < RW) c = -1;
      else if (x < mainX) c = cols.at((x - RW) / z);
      else c = cols.at(cols.start(fc) + (x - mainX + scroll.left) / z);
      return [r, c];
    },
    [HH, mainY, rows, z, fr, scroll.top, RW, mainX, cols, fc, scroll.left]
  );

  const panes = useMemo(() => {
    const p = [{ key: "main", x: mainX, y: mainY, w: mainW, h: mainH, r0: Math.max(mainR0, fr), r1: mainR1, c0: Math.max(mainC0, fc), c1: mainC1 }];
    if (fr) p.push({ key: "top", x: mainX, y: HH, w: mainW, h: FH, r0: 0, r1: fr - 1, c0: Math.max(mainC0, fc), c1: mainC1 });
    if (fc) p.push({ key: "left", x: RW, y: mainY, w: FW, h: mainH, r0: Math.max(mainR0, fr), r1: mainR1, c0: 0, c1: fc - 1 });
    if (fr && fc) p.push({ key: "corner", x: RW, y: HH, w: FW, h: FH, r0: 0, r1: fr - 1, c0: 0, c1: fc - 1 });
    return p;
  }, [mainX, mainY, mainW, mainH, mainR0, mainR1, mainC0, mainC1, fr, fc, HH, FH, RW, FW]);

  // Rectangle of a range on screen (unclipped).
  const rectOf = useCallback(
    (g0) => {
      const [r0, c0, r1, c1] = g0;
      const x0 = cx(c0);
      const y0 = ry(r0);
      const x1 = cx(c1) + cw(c1);
      const y1 = ry(r1) + rh(r1);
      return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
    },
    [cx, ry, cw, rh]
  );

  // Clip a range to what a pane shows (full rows/columns get clipped to the visible lines).
  const clipToPane = useCallback((g0, p) => {
    const r0 = Math.max(g0[0], p.r0),
      r1 = Math.min(g0[2], p.r1),
      c0 = Math.max(g0[1], p.c0),
      c1 = Math.min(g0[3], p.c1);
    if (r0 > r1 || c0 > c1) return null;
    return [r0, c0, r1, c1];
  }, []);

  // A range's box in pane coordinates, cut a few pixels outside the pane: browsers can't lay out a box
  // millions of pixels tall (a whole column, a selection down to row 1,048,576), and what lies outside
  // the pane doesn't show anyway.
  const paneBox = (rc, p, grow) => {
    const x0 = rc.x - p.x - 1, y0 = rc.y - p.y - 1;
    const l = Math.max(x0, -8), t = Math.max(y0, -8);
    const r = Math.min(x0 + rc.w + grow, p.w + 8), b = Math.min(y0 + rc.h + grow, p.h + 8);
    return { left: l, top: t, width: Math.max(r - l, 0), height: Math.max(b - t, 0) };
  };

  // ------------------------------------------------------------ sizing, scrolling

  useLayoutEffect(() => {
    const el = scrollerRef.current;
    if (!el) return;
    const ro = new ResizeObserver(() => setView({ w: el.clientWidth, h: el.clientHeight }));
    ro.observe(el);
    setView({ w: el.clientWidth, h: el.clientHeight });
    return () => ro.disconnect();
  }, []);

  // Zooming keeps the same top-left cell in view (scroll offsets scale with the zoom).
  const zoomRef = useRef(z);
  useLayoutEffect(() => {
    const prev = zoomRef.current;
    zoomRef.current = z;
    const el = scrollerRef.current;
    if (!el || prev === z) return;
    const k = z / prev;
    const top = scroll.top * k, left = el.scrollLeft * k;
    requestAnimationFrame(() => {
      el.scrollTop = vmapRef.current.toScroll(top);
      el.scrollLeft = left;
    });
  }, [z]);

  const onScroll = useCallback(() => {
    const el = scrollerRef.current;
    if (!el) return;
    setScroll({ top: vmapRef.current.toSheet(el.scrollTop), left: el.scrollLeft });
    setMenu(null);
  }, []);

  // Ask the engine for the cells around the visible area.
  useEffect(() => {
    if (!onEvent) return;
    const want = [Math.max(0, mainR0 - 40), Math.max(0, mainC0 - 6), Math.min(MAXR - 1, mainR1 + 40), Math.min(MAXC - 1, mainC1 + 6)];
    const have = data && data.window && data.sheet === sheetIndex ? data.window : null;
    const visible = [mainR0, mainC0, mainR1, mainC1];
    const covered = have && have[0] <= visible[0] && have[1] <= visible[1] && have[2] >= visible[2] && have[3] >= visible[3];
    if (covered) return;
    const last = lastWindowRef.current;
    if (last && last.sheet === sheetIndex && last.w.join() === want.join() && Date.now() - last.t < 1500) return;
    const t = setTimeout(() => {
      lastWindowRef.current = { sheet: sheetIndex, w: want, t: Date.now() };
      onEvent({ type: "viewport", seq: seqRef.current, window: want });
    }, 40);
    return () => clearTimeout(t);
  }, [mainR0, mainR1, mainC0, mainC1, data, sheetIndex, onEvent]);

  const ensureVisible = useCallback(
    (r, c) => {
      const el = scrollerRef.current;
      if (!el) return;
      const now = vmapRef.current.toSheet(el.scrollTop);
      let top = now,
        left = el.scrollLeft;
      if (r >= fr) {
        const y0 = (rows.start(r) - rows.start(fr)) * z;
        const y1 = y0 + rows.sizeOf(r) * z;
        if (y0 < top) top = y0;
        else if (y1 > top + mainH) top = Math.max(0, y1 - mainH);
      }
      if (c >= fc) {
        const x0 = (cols.start(c) - cols.start(fc)) * z;
        const x1 = x0 + cols.sizeOf(c) * z;
        if (x0 < left) left = x0;
        else if (x1 > left + mainW) left = Math.max(0, x1 - mainW);
      }
      const ext = extentRef.current;
      if (r + 20 > ext.r) ext.r = Math.min(MAXR, r + 60);
      if (c + 8 > ext.c) ext.c = Math.min(MAXC, c + 12);
      if (top !== now || left !== el.scrollLeft) {
        // The spacer may need to grow first (and the squeezed scrolling start): once more if so.
        const go = (again) => {
          el.scrollTop = vmapRef.current.toScroll(top);
          el.scrollLeft = left;
          if (again && Math.abs(vmapRef.current.toSheet(el.scrollTop) - top) > 4 * z) requestAnimationFrame(() => go(false));
        };
        requestAnimationFrame(() => go(true));
      }
    },
    [fr, fc, rows, cols, z, mainH, mainW]
  );

  // A keyboard move from the scrolled rows (columns) into the frozen ones puts the scrolled area
  // back at its top (left), as Excel does: otherwise rows stay hidden behind the frozen ones.
  const backToPanes = useCallback(
    (from, to) => {
      const el = scrollerRef.current;
      if (!el) return;
      if (to[0] < fr && from[0] >= fr) el.scrollTop = 0;
      if (to[1] < fc && from[1] >= fc) el.scrollLeft = 0;
    },
    [fr, fc]
  );

  // ------------------------------------------------------------ server state sync

  // A chart or picture this tab just inserted is selected, as in Excel: Delete then removes it,
  // not the data it was made from.
  const selObjRef = useRef(null);
  useEffect(() => {
    const so = ui && ui.selectObject;
    if (!so || selObjRef.current === so.n) return;
    selObjRef.current = so.n;
    setObjSel(so.id);
  }, [ui]);

  const reopenedRef = useRef(0);
  // The next selection from the server extends the current one (Ctrl+Shift+Arrow): reveal its moving
  // end, not the anchor, as Excel does.
  const revealEndRef = useRef(false);
  useEffect(() => {
    if (!ui || !ui.selection) return;
    // Someone renaming this sheet keeps its place and drops its old name: not another sheet.
    const prev = sheetRef.current;
    const renamed = prev !== sheetName && sheetIdxRef.current === ui.sheet && !(ui.sheets || []).some((x) => x.name === prev);
    const sheetChanged = prev !== sheetName && !renamed;
    sheetRef.current = sheetName;
    sheetIdxRef.current = ui.sheet;
    if (sheetChanged) {
      const el = scrollerRef.current;
      if (el) {
        el.scrollTop = 0;
        el.scrollLeft = 0;
      }
      extentRef.current = { r: 100, c: 26 };
      setObjSel(null);
      if (edRef.current) setEd(null);
    }
    // Passive redraws (someone else's change, a ping, a scroll) keep this tab's selection: it
    // may have moved since it last told the server.
    if (((ui.seq || 0) >= seqRef.current && !ui.passive) || sheetChanged) {
      const s = ui.selection;
      const cur = selRef.current;
      const same = cur.active[0] === s.active[0] && cur.active[1] === s.active[1] && JSON.stringify(cur.ranges) === JSON.stringify(s.ranges);
      if (!same && !dragRef.current) {
        setSel({ active: s.active, anchor: s.anchor || s.active, ranges: s.ranges });
        const g0 = s.ranges[s.ranges.length - 1];
        const an = s.anchor || s.active;
        if (revealEndRef.current && g0) ensureVisible(an[0] === g0[2] ? g0[0] : g0[2], an[1] === g0[3] ? g0[1] : g0[3]);
        else ensureVisible(s.active[0], s.active[1]);
      }
      revealEndRef.current = false;
    }
    setPending((p) => {
      const n = {};
      let changed = false;
      for (const k in p) {
        if (p[k].seq > (ui.seq || 0)) n[k] = p[k];
        else changed = true;
      }
      return changed ? n : p;
    });
    // A failed commit reopens the editor with what was typed, where the server kept the selection.
    // Once: the error stays in the passive redraws that follow it (which don't move the selection,
    // and may reach us in the same update) until the next event.
    const err = ui.lastError;
    if (err && err.seq === seqRef.current && err.kind === "commit" && !edRef.current && reopenedRef.current !== err.seq) {
      reopenedRef.current = err.seq;
      const s = ui.selection;
      setSel({ active: s.active, anchor: s.anchor || s.active, ranges: s.ranges });
      setEd({ r: err.r, c: err.c, text: err.text, mode: "edit", bar: false });
      setTimeout(() => kbdRef.current && kbdRef.current.focus(), 0);
    }
    if (ui.clipboardSet && ui.clipboardSet.seq === seqRef.current) {
      const set = ui.clipboardSet;
      // Text here means the server read the range (off-screen or big): put it on the clipboard.
      if ("text" in set && navigator.clipboard) navigator.clipboard.writeText(set.text).catch(() => {});
      const pc = pendingCopyRef.current;
      const text = "text" in set ? set.text : pc ? pc.text : null;
      const cut = pc ? pc.cut : false;
      // Only a copy (not a cut, which is consumed when pasted) gets the paste-by-token fast path.
      lastClipRef.current = !cut && text != null ? { text, token: set.token } : null;
    }
  }, [ui]);

  useEffect(() => {
    if (kbdRef.current && !edRef.current) kbdRef.current.focus({ preventScroll: true });
  }, [focusKey]);

  // ------------------------------------------------------------ other people on the book

  // The book changed after what this tab drew: ask for its window again, one request at a time.
  const syncRef = useRef(0);
  useEffect(() => {
    if (!ui || rev === undefined || rev === null || !onEventRef.current) return;
    const drawn = ui.bookRev || 0;
    if (rev <= drawn || syncRef.current > drawn) return;
    syncRef.current = rev;
    onEventRef.current({ type: "sync", seq: seqRef.current });
  }, [rev, ui]);

  // Still here: keeps this tab among the people on the book (and its drawing fresh).
  useEffect(() => {
    const t = setInterval(() => onEventRef.current && onEventRef.current({ type: "ping", seq: seqRef.current }), 5000);
    return () => clearInterval(t);
  }, []);

  // Closing or reloading the tab: the others stop seeing it at once (otherwise it would take
  // the server a few missed pings to notice).
  // A beacon, because a websocket message doesn't get out of a page that is going away. It goes
  // to the backend the way Reflex's own events do (behind a proxy, on the page's host).
  const peerRef = useRef(me);
  peerRef.current = me;
  useEffect(() => {
    const bye = () => {
      const peer = peerRef.current;
      if (!peer || !navigator.sendBeacon) return;
      try {
        const url = new URL("/_reflex_grid/bye", getBackendURL(env.PING));
        navigator.sendBeacon(url.toString(), JSON.stringify({ token: getToken(), peer }));
      } catch (_) {}
    };
    window.addEventListener("pagehide", bye);
    return () => window.removeEventListener("pagehide", bye);
  }, []);

  // Tell the others which cell this tab is typing in.
  const editingSent = useRef("");
  useEffect(() => {
    const cur = ed ? `${ed.r},${ed.c}` : "";
    if (editingSent.current === cur || !onEventRef.current) return;
    editingSent.current = cur;
    onEventRef.current({ type: "presence", seq: seqRef.current, editing: ed ? [ed.r, ed.c] : null });
  }, [ed]);

  // Everyone else on this sheet, drawn with their selection in their colour.
  const others = useMemo(() => (peers || []).filter((q) => q && q.id !== me && q.sheet === sheetName && q.active), [peers, me, sheetName]);

  // A name tag shows for a moment after someone moves (and while they type), so it doesn't sit
  // on the cells above theirs; otherwise a small flag marks their cell.
  const TAG_MS = 3000;
  const peerMoved = useRef({});
  const [, setTagTick] = useState(0);
  useEffect(() => {
    const now = Date.now();
    const seen = peerMoved.current;
    let moved = false;
    for (const q of peers || []) {
      if (!q || !q.active) continue;
      const sig = `${q.sheet}|${q.active[0]},${q.active[1]}|${q.editing ? 1 : 0}`;
      if (!seen[q.id] || seen[q.id].sig !== sig) {
        seen[q.id] = { sig, at: now };
        moved = true;
      }
    }
    if (!moved) return;
    const t = setTimeout(() => setTagTick((n) => n + 1), TAG_MS + 50);
    return () => clearTimeout(t);
  }, [peers]);

  // Cells someone else just changed light up in their colour for a moment.
  const [flashes, setFlashes] = useState([]);
  const seenChangeRef = useRef(null);
  useEffect(() => {
    if (!ui || ui.bookRev === undefined) return;
    const list = changes || [];
    const top = list.length ? list[list.length - 1].rev : 0;
    if (seenChangeRef.current === null) {
      // What happened before this tab drew the book isn't news.
      seenChangeRef.current = Math.max(top, ui.bookRev || 0);
      return;
    }
    const fresh = list.filter((c) => c.rev > seenChangeRef.current && c.author !== me);
    seenChangeRef.current = Math.max(seenChangeRef.current, top);
    const now = Date.now();
    const add = [];
    for (const c of fresh) for (const area of c.areas || []) add.push({ id: `${c.rev}:${add.length}`, sheet: area[0], g: area[1], color: c.color || "#5F6368", until: now + 1600 });
    if (!add.length) return;
    setFlashes((f) => [...f.filter((x) => x.until > now), ...add].slice(-40));
    setTimeout(() => setFlashes((f) => f.filter((x) => x.until > Date.now())), 1700);
  }, [changes, me, ui]);

  // Keys pressed while nothing has focus (after a dialog closed, a click on the page background)
  // belong to the grid, as in Excel.
  useEffect(() => {
    const onKey = (e) => {
      const a = document.activeElement;
      if (a && a !== document.body && a !== document.documentElement) return;
      if (document.querySelector('[role="dialog"]')) return;
      const k = kbdRef.current;
      if (!k) return;
      k.focus({ preventScroll: true });
      if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey && !edRef.current) {
        e.preventDefault();
        const [r, c] = selRef.current.active;
        caretRef.current = { pos: 1 };
        setEd({ r, c, text: e.key, mode: "enter", bar: false });
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [setEd]);

  // The ribbon or a dialog asked to start editing with some text (Insert Function).
  const editReqRef = useRef(null);
  useEffect(() => {
    if (!editRequest || !editRequest.seq || editReqRef.current === editRequest.seq) return;
    editReqRef.current = editRequest.seq;
    beginEdit(editRequest.text || "", "edit");
    setTimeout(() => updateAc(editRequest.text || "", (editRequest.text || "").length), 10);
  }, [editRequest]);

  // Tell the server which object is selected (for the Chart Design tab).
  const objSelSent = useRef(null);
  useEffect(() => {
    if (objSelSent.current === objSel) return;
    objSelSent.current = objSel;
    const isChart = objSel !== null && ((objects && objects.charts) || []).some((c) => c.id === objSel);
    if (onEvent) onEvent({ type: "objectSelect", seq: seqRef.current, id: objSel, kind: objSel === null ? "" : isChart ? "chart" : "image" });
  }, [objSel]);

  // ------------------------------------------------------------ selection helpers

  const isHiddenRow = useCallback((r) => rows.sizeOf(r) === 0, [rows]);
  const isHiddenCol = useCallback((c) => cols.sizeOf(c) === 0, [cols]);

  const step = useCallback(
    (r, c, dr, dc) => {
      const m = mergeAt(r, c);
      if (m) {
        if (dr > 0) r = m[2];
        if (dr < 0) r = m[0];
        if (dc > 0) c = m[3];
        if (dc < 0) c = m[1];
      }
      for (let i = 0; i < Math.abs(dr); i++) {
        let n = r + Math.sign(dr);
        let guard = 0;
        while (n >= 0 && n < MAXR && isHiddenRow(n) && guard++ < 100000) n += Math.sign(dr);
        if (n < 0 || n >= MAXR) break;
        r = n;
      }
      for (let i = 0; i < Math.abs(dc); i++) {
        let n = c + Math.sign(dc);
        let guard = 0;
        while (n >= 0 && n < MAXC && isHiddenCol(n) && guard++ < 20000) n += Math.sign(dc);
        if (n < 0 || n >= MAXC) break;
        c = n;
      }
      const m2 = mergeAt(r, c);
      return m2 ? [m2[0], m2[1]] : [r, c];
    },
    [mergeAt, isHiddenRow, isHiddenCol]
  );

  const expandMerges = useCallback(
    (g0) => {
      let cur = g0;
      for (let k = 0; k < 20; k++) {
        let grown = cur;
        for (const m of merges) if (intersects(m, grown)) grown = [Math.min(grown[0], m[0]), Math.min(grown[1], m[1]), Math.max(grown[2], m[2]), Math.max(grown[3], m[3])];
        if (grown.join() === cur.join()) break;
        cur = grown;
      }
      return cur;
    },
    [merges]
  );

  const selectCell = useCallback(
    (r, c, opts) => {
      const m = mergeAt(r, c);
      const s = m ? { active: [m[0], m[1]], anchor: [m[0], m[1]], ranges: [m] } : { active: [r, c], anchor: [r, c], ranges: [[r, c, r, c]] };
      setSel(s);
      if (!opts || !opts.noScroll) ensureVisible(r, c);
      return s;
    },
    [mergeAt, setSel, ensureVisible]
  );

  const syncSel = useCallback(
    (s) => {
      s = s || selRef.current;
      send("select", { sel: { active: s.active, anchor: s.anchor, ranges: s.ranges } });
    },
    [send]
  );

  const selTimer = useRef(null);
  const syncSelSoon = useCallback(() => {
    if (selTimer.current) clearTimeout(selTimer.current);
    selTimer.current = setTimeout(() => {
      selTimer.current = null;
      syncSel();
    }, 60);
  }, [syncSel]);

  // ------------------------------------------------------------ editing

  const cellText = useCallback(
    (r, c) => {
      const e = cellMap.get(key(r, c));
      if (e && e.in !== undefined) return e.in;
      if (e && e.t) return e.t;
      if (ui && ui.selection && ui.selection.active && ui.selection.active[0] === r && ui.selection.active[1] === c && ui.sheet === sheetIndex) return ui.formulaBar || "";
      return "";
    },
    [ui, cellMap, sheetIndex]
  );

  const beginEdit = useCallback(
    (text, mode, bar) => {
      const [r, c] = selRef.current.active;
      if (!bar) caretRef.current = { pos: (text || "").length };
      setEd({ r, c, text, mode, bar: !!bar });
      setListOpen(false);
      pointRef.current = null;
    },
    [setEd]
  );

  const moveAfter = useCallback(
    (move) => {
      const s = selRef.current;
      const single = s.ranges.length === 1 && (s.ranges[0][0] === s.ranges[0][2] && s.ranges[0][1] === s.ranges[0][3] || (mergeAt(s.active[0], s.active[1]) || []).join() === s.ranges[0].join());
      const fwd = move === "down" || move === "right";
      const byRow = move === "right" || move === "left";
      if (move === "none") return;
      if (single) {
        const [dr, dc] = byRow ? [0, fwd ? 1 : -1] : [fwd ? 1 : -1, 0];
        const [r, c] = step(s.active[0], s.active[1], dr, dc);
        selectCell(r, c);
        return;
      }
      // Cycle inside the selection.
      const g0 = s.ranges[s.ranges.length - 1];
      let [r, c] = s.active;
      const d = fwd ? 1 : -1;
      if (byRow) {
        c += d;
        if (c > g0[3]) { c = g0[1]; r += 1; }
        else if (c < g0[1]) { c = g0[3]; r -= 1; }
        if (r > g0[2]) r = g0[0];
        else if (r < g0[0]) r = g0[2];
      } else {
        r += d;
        if (r > g0[2]) { r = g0[0]; c += 1; }
        else if (r < g0[0]) { r = g0[2]; c -= 1; }
        if (c > g0[3]) c = g0[1];
        else if (c < g0[1]) c = g0[3];
      }
      setSel({ ...s, active: [r, c] });
      ensureVisible(r, c);
    },
    [step, selectCell, setSel, ensureVisible, mergeAt]
  );

  const commit = useCallback(
    (move, mode) => {
      const e = edRef.current;
      if (!e) return;
      setEd(null);
      setAc(null);
      pointRef.current = null;
      const seq = send("commit", { r: e.r, c: e.c, text: e.text, move: move || "down", mode: mode || "normal" });
      if (!(mode === "fill" || mode === "array")) {
        const show = e.text && e.text[0] !== "=" ? e.text : null;
        if (show !== null) setPending((p) => ({ ...p, [key(e.r, e.c)]: { text: show, seq } }));
      }
      moveAfter(mode === "fill" ? "none" : move || "down");
      setTimeout(() => kbdRef.current && kbdRef.current.focus({ preventScroll: true }), 0);
    },
    [send, setEd, moveAfter]
  );

  const cancelEdit = useCallback(() => {
    setEd(null);
    setAc(null);
    pointRef.current = null;
    setTimeout(() => kbdRef.current && kbdRef.current.focus({ preventScroll: true }), 0);
  }, [setEd]);

  // Function autocomplete and argument hints.
  const fnIndex = useMemo(() => (functions || []).map((f) => ({ ...f, upper: f.name.toUpperCase() })), [functions]);
  const updateAc = useCallback(
    (text, caret) => {
      if (!text || text[0] !== "=") {
        setAc(null);
        return;
      }
      const before = text.slice(0, caret);
      const m = /([A-Za-z_][A-Za-z0-9_.]*)$/.exec(before);
      if (!m || (before[m.index - 1] && /[A-Za-z0-9_$!:"']/.test(before[m.index - 1]))) {
        setAc(null);
        return;
      }
      const prefix = m[1].toUpperCase();
      const items = fnIndex.filter((f) => f.upper.startsWith(prefix)).slice(0, 12);
      if (!items.length) {
        setAc(null);
        return;
      }
      setAc({ items, index: 0, start: m.index, end: caret });
    },
    [fnIndex]
  );

  const acceptAc = useCallback(
    (item) => {
      const e = edRef.current;
      const a = ac;
      if (!e || !a) return;
      const fn = item || a.items[a.index];
      const text = e.text.slice(0, a.start) + fn.name + "(" + e.text.slice(a.end);
      const caret = a.start + fn.name.length + 1;
      caretRef.current = { pos: caret };
      setEd({ ...e, text });
      setAc(null);
    },
    [ac, setEd]
  );

  const argHint = useMemo(() => {
    if (!ed || !ed.text || ed.text[0] !== "=") return null;
    const el = ed.bar ? barRef.current : kbdRef.current;
    const caret = el && el.selectionStart != null ? el.selectionStart : ed.text.length;
    const before = ed.text.slice(0, caret);
    let depth = 0;
    let argi = 0;
    let inStr = false;
    for (let i = before.length - 1; i >= 0; i--) {
      const ch = before[i];
      if (ch === '"') inStr = !inStr;
      if (inStr) continue;
      if (ch === ")") depth++;
      else if (ch === "(") {
        if (depth === 0) {
          const m = /([A-Za-z_][A-Za-z0-9_.]*)$/.exec(before.slice(0, i));
          if (!m) return null;
          const f = fnIndex.find((x) => x.upper === m[1].toUpperCase());
          return f ? { fn: f, arg: argi } : null;
        }
        depth--;
      } else if (ch === "," && depth === 0) argi++;
    }
    return null;
  }, [ed, fnIndex]);

  // Point mode: insert or replace a reference at the caret.
  const pointTo = useCallback(
    (g0) => {
      const e = edRef.current;
      if (!e) return false;
      const el = e.bar ? barRef.current : kbdRef.current;
      let pm = pointRef.current;
      const ref = rangeA1(g0);
      if (!pm) {
        const caret = el && el.selectionStart != null ? el.selectionStart : e.text.length;
        if (!canPoint(e.text, caret)) return false;
        pm = { start: caret, end: caret };
      }
      const text = e.text.slice(0, pm.start) + ref + e.text.slice(pm.end);
      pointRef.current = { ...pm, end: pm.start + ref.length, range: g0 };
      caretRef.current = { pos: pm.start + ref.length };
      setEd({ ...e, text, mode: "enter" });
      return true;
    },
    [setEd]
  );

  // ------------------------------------------------------------ clipboard

  const selectionTsv = useCallback(() => {
    const s = selRef.current;
    if (s.ranges.length !== 1) return null;
    const [r0, c0, r1, c1] = s.ranges[0];
    const have = data && data.window;
    if (!have || r0 < have[0] || c0 < have[1] || r1 > have[2] || c1 > have[3] || (r1 - r0 + 1) * (c1 - c0 + 1) > 200000) return null;
    const lines = [];
    for (let r = r0; r <= r1; r++) {
      if (isHiddenRow(r)) continue;
      const parts = [];
      for (let c = c0; c <= c1; c++) {
        const e = cellMap.get(key(r, c));
        let t = e && e.t ? e.t : "";
        if (/[\t\n"]/.test(t)) t = '"' + t.replace(/"/g, '""') + '"';
        parts.push(t);
      }
      lines.push(parts.join("\t"));
    }
    return lines.join("\n") + (lines.length ? "\n" : "");
  }, [data, cellMap, isHiddenRow]);

  const onCopyEvt = useCallback(
    (e, cut) => {
      if (edRef.current) return;
      const text = selectionTsv();
      if (text !== null && e.clipboardData) {
        e.clipboardData.setData("text/plain", text);
        e.preventDefault();
      }
      pendingCopyRef.current = { text, cut: !!cut };
      send("copy", { cut: !!cut, text });
    },
    [selectionTsv, send]
  );

  // Paste what is on the system clipboard: our own copy by its token (the big block stays on the
  // server), anything else as its text.
  const sendPaste = useCallback(
    (text) => {
      const lc = lastClipRef.current;
      // Only while the server still holds that copy (its marquee shows): after Esc, a pasted cut or
      // a macro's CutCopyMode = False, the text is pasted as before.
      const held = uiRef.current && uiRef.current.clipboard;
      if (lc && lc.token && held && text != null && text === lc.text) send("paste", { token: lc.token });
      else send("paste", { text });
    },
    [send]
  );

  const onPasteEvt = useCallback(
    (e) => {
      if (edRef.current) return;
      const cd = e.clipboardData;
      if (!cd) return;
      e.preventDefault();
      // Only the text goes to the server (it pastes TSV, not an app's much bigger HTML).
      sendPaste(cd.getData("text/plain"));
    },
    [sendPaste]
  );

  // ------------------------------------------------------------ keyboard

  const cmd = useCallback((id, params) => send("command", { id, params: params || {} }), [send]);
  const dialog = useCallback((name, params) => send("dialog", { name, params: params || {} }), [send]);

  const onKeyDown = useCallback(
    (e) => {
      const mod = e.ctrlKey || e.metaKey;
      const editing = edRef.current;
      if (e.nativeEvent.isComposing || e.keyCode === 229) return;
      if (editing) {
        if (ac) {
          if (e.key === "ArrowDown" || e.key === "ArrowUp") {
            e.preventDefault();
            const n = ac.items.length;
            setAc({ ...ac, index: (ac.index + (e.key === "ArrowDown" ? 1 : n - 1)) % n });
            return;
          }
          if (e.key === "Tab") {
            e.preventDefault();
            acceptAc();
            return;
          }
          if (e.key === "Escape") {
            e.preventDefault();
            setAc(null);
            return;
          }
        }
        if (e.key === "Enter") {
          if (e.altKey) {
            // Line break inside the cell.
            e.preventDefault();
            const el = e.target;
            const p = el.selectionStart;
            const text = editing.text.slice(0, p) + "\n" + editing.text.slice(el.selectionEnd);
            caretRef.current = { pos: p + 1 };
            setEd({ ...editing, text, mode: "edit" });
            return;
          }
          e.preventDefault();
          if (mod && e.shiftKey) commit("none", "array");
          else if (mod) commit("none", "fill");
          else commit(e.shiftKey ? "up" : "down");
          return;
        }
        if (e.key === "Tab") {
          e.preventDefault();
          commit(e.shiftKey ? "left" : "right");
          return;
        }
        if (e.key === "Escape") {
          e.preventDefault();
          cancelEdit();
          return;
        }
        if (e.key === "F2") {
          e.preventDefault();
          setEd({ ...editing, mode: editing.mode === "enter" ? "edit" : "enter" });
          return;
        }
        if (e.key === "F4" && editing.text[0] === "=") {
          e.preventDefault();
          const el = e.target;
          const caret = el.selectionStart;
          const refs = formulaRefs(editing.text);
          const hitRef = refs.find((x) => x.start <= caret && caret <= x.end);
          if (hitRef) {
            const part = editing.text.slice(hitRef.start, hitRef.end);
            const toggled = part.split(":").map(toggleAbs).join(":");
            const text = editing.text.slice(0, hitRef.start) + toggled + editing.text.slice(hitRef.end);
            caretRef.current = { pos: hitRef.start + toggled.length };
            setEd({ ...editing, text });
          }
          return;
        }
        if (editing.mode === "enter" && !editing.bar && e.key.startsWith("Arrow")) {
          const dr = e.key === "ArrowDown" ? 1 : e.key === "ArrowUp" ? -1 : 0;
          const dc = e.key === "ArrowRight" ? 1 : e.key === "ArrowLeft" ? -1 : 0;
          const el = e.target;
          const caret = el.selectionStart;
          if (editing.text[0] === "=" && (pointRef.current || canPoint(editing.text, caret))) {
            e.preventDefault();
            const pm = pointRef.current;
            const base = pm && pm.range ? (e.shiftKey ? pm.cur || [pm.range[2], pm.range[3]] : [pm.range[0], pm.range[1]]) : [editing.r, editing.c];
            const anchor = pm && pm.anchor && e.shiftKey ? pm.anchor : null;
            const [r, c] = step(base[0], base[1], dr, dc);
            const g0 = anchor ? norm(anchor, [r, c]) : [r, c, r, c];
            pointTo(g0);
            pointRef.current = { ...pointRef.current, anchor: anchor || [r, c], cur: [r, c] };
            ensureVisible(r, c);
            return;
          }
          e.preventDefault();
          commit(dr > 0 ? "down" : dr < 0 ? "up" : dc > 0 ? "right" : "left");
          return;
        }
        return;
      }
      // ---- not editing
      const s = selRef.current;
      const k = e.key;
      if (objSel !== null && (k === "Delete" || k === "Backspace")) {
        e.preventDefault();
        cmd("object.delete", { id: objSel });
        setObjSel(null);
        return;
      }
      if (k.startsWith("Arrow")) {
        e.preventDefault();
        const dr = k === "ArrowDown" ? 1 : k === "ArrowUp" ? -1 : 0;
        const dc = k === "ArrowRight" ? 1 : k === "ArrowLeft" ? -1 : 0;
        if (e.altKey && k === "ArrowDown") {
          if (ui && ui.dropdown) setListOpen(true);
          return;
        }
        if (mod) {
          revealEndRef.current = e.shiftKey;
          cmd("selection.move", { dr, dc, jump: true, extend: e.shiftKey });
          return;
        }
        if (e.shiftKey) {
          const g0 = s.ranges[s.ranges.length - 1];
          const anchor = s.active;
          let cr = g0[2] !== anchor[0] || g0[0] === g0[2] ? g0[2] : g0[0];
          let cc = g0[3] !== anchor[1] || g0[1] === g0[3] ? g0[3] : g0[1];
          if (anchor[0] === g0[2] && g0[0] !== g0[2]) cr = g0[0];
          if (anchor[1] === g0[3] && g0[1] !== g0[3]) cc = g0[1];
          const [nr, nc] = step(cr, cc, dr, dc);
          const ng = expandMerges(norm(anchor, [nr, nc]));
          const ns = { active: anchor, anchor, ranges: [...s.ranges.slice(0, -1), ng] };
          setSel(ns);
          // Shift+Arrow on whole columns (or rows) keeps the view on the rows (columns) it showed:
          // the extension's far end is the sheet's edge, not somewhere to scroll to.
          const wholeCols = g0[0] === 0 && g0[2] === MAXR - 1 && dr === 0;
          const wholeRows = g0[1] === 0 && g0[3] === MAXC - 1 && dc === 0;
          ensureVisible(wholeCols ? anchor[0] : nr, wholeRows ? anchor[1] : nc);
          syncSelSoon();
          return;
        }
        const [r, c] = step(s.active[0], s.active[1], dr, dc);
        selectCell(r, c);
        backToPanes(s.active, [r, c]);
        syncSelSoon();
        return;
      }
      if (k === "Enter" || k === "Tab") {
        e.preventDefault();
        // Enter while cells are marked as copied pastes them into the selection and ends copy mode,
        // as in Excel.
        if (k === "Enter" && !e.shiftKey && !mod && !e.altKey && ui && ui.clipboard) {
          lastClipRef.current = null;
          send("paste", {});
          cmd("edit.clearClipboard");
          return;
        }
        const move = k === "Enter" ? (e.shiftKey ? "up" : "down") : e.shiftKey ? "left" : "right";
        moveAfter(move);
        syncSelSoon();
        return;
      }
      if (k === "Home" || k === "End" || k === "PageDown" || k === "PageUp") {
        e.preventDefault();
        if (k === "Home" && mod) {
          selectCell(fr, fc);
          syncSelSoon();
        } else if (k === "Home") {
          selectCell(s.active[0], 0);
          syncSelSoon();
        } else if (k === "End" && mod) {
          cmd("edit.goToSpecial", { kind: "lastCell" });
        } else {
          const visRows = Math.max(1, Math.floor(mainH / ((g.defaultRowHeight || 20) * z)) - 1);
          const visCols = Math.max(1, Math.floor(mainW / ((g.defaultColWidth || 64) * z)) - 1);
          const d = k === "PageDown" ? 1 : -1;
          if (e.altKey) {
            const [r, c] = step(s.active[0], s.active[1], 0, d * visCols);
            selectCell(r, c);
            backToPanes(s.active, [r, c]);
          } else {
            const [r, c] = step(s.active[0], s.active[1], d * visRows, 0);
            selectCell(r, c);
            const el = scrollerRef.current;
            // Into the frozen rows: the scrolled rows go back to the top, so none stay hidden
            // behind the frozen ones (row 10, Page Down, Page Up, Page Up used to end on row 1 with
            // rows 2-9 out of sight).
            if (el && r < fr && s.active[0] >= fr) el.scrollTop = 0;
            else if (el) el.scrollTop = Math.max(0, vmapRef.current.toScroll(vmapRef.current.toSheet(el.scrollTop) + d * visRows * (g.defaultRowHeight || 20) * z));
          }
          syncSelSoon();
        }
        return;
      }
      if (k === "F2") {
        e.preventDefault();
        if (ui && ui.protected && ui.ribbon && ui.ribbon.locked) return;
        beginEdit(cellText(s.active[0], s.active[1]), "edit");
        return;
      }
      if (k === "Delete") {
        e.preventDefault();
        cmd("edit.clearContents");
        return;
      }
      if (k === "Backspace") {
        e.preventDefault();
        beginEdit("", "edit");
        return;
      }
      if (k === "Escape") {
        e.preventDefault();
        setMenu(null);
        setListOpen(false);
        if (ui && (ui.clipboard || ui.formatPainter)) {
          lastClipRef.current = null;
          cmd("edit.clearClipboard");
          if (ui.formatPainter) cmd("edit.cancelFormatPainter");
        }
        return;
      }
      if (k === "F9") {
        e.preventDefault();
        cmd("formulas.calculateNow");
        return;
      }
      if (k === "F5") {
        e.preventDefault();
        dialog("goTo");
        return;
      }
      if (k === "F11" && e.shiftKey) {
        e.preventDefault();
        cmd("home.insertSheet");
        return;
      }
      if (k === "F11" && e.altKey) {
        e.preventDefault();
        dialog("vba");
        return;
      }
      if (k === "F8" && e.altKey) {
        e.preventDefault();
        dialog("macros");
        return;
      }
      if (k === "F2" && e.shiftKey) {
        e.preventDefault();
        dialog("note");
        return;
      }
      if (e.altKey && (k === "=" || e.code === "Equal")) {
        e.preventDefault();
        cmd("formulas.autoSum");
        return;
      }
      if (mod) {
        const lower = k.toLowerCase();
        const shortcuts = {
          z: ["edit.undo"],
          y: ["edit.redo"],
          b: ["home.bold"],
          i: ["home.italic"],
          u: ["home.underline"],
          5: ["home.strikethrough"],
          d: ["edit.fillDown"],
          r: ["edit.fillRight"],
          "`": ["formulas.showFormulas"],
        };
        if (lower === "z" && e.shiftKey) {
          e.preventDefault();
          cmd("edit.redo");
          return;
        }
        if (lower === "c" || lower === "x" || lower === "v") {
          if (lower === "v" && e.altKey) {
            e.preventDefault();
            dialog("pasteSpecial");
          }
          return; // native clipboard events
        }
        if (lower === "a") {
          e.preventDefault();
          cmd("edit.selectAll");
          return;
        }
        if (k === " " && !e.shiftKey) {
          e.preventDefault();
          cmd("selection.column");
          return;
        }
        if (k === " " && e.shiftKey) {
          e.preventDefault();
          cmd("edit.selectAll", { sheet: true });
          return;
        }
        if (lower === "l" && e.shiftKey) {
          e.preventDefault();
          cmd("data.filter");
          return;
        }
        if (lower === "t" && !e.shiftKey) {
          e.preventDefault();
          dialog("insertTable");
          return;
        }
        if (lower === "1") {
          e.preventDefault();
          dialog("formatCells");
          return;
        }
        if (lower === "f") {
          e.preventDefault();
          dialog("find");
          return;
        }
        if (lower === "h") {
          e.preventDefault();
          dialog("replace");
          return;
        }
        if (lower === "g") {
          e.preventDefault();
          dialog("goTo");
          return;
        }
        if (lower === "k") {
          e.preventDefault();
          dialog("hyperlink");
          return;
        }
        if (lower === "s") {
          e.preventDefault();
          cmd("ui.save");
          return;
        }
        if (lower === "p") {
          e.preventDefault();
          cmd("ui.print");
          return;
        }
        if (lower === "o") {
          e.preventDefault();
          dialog("open");
          return;
        }
        if (k === "PageDown" || k === "PageUp") {
          e.preventDefault();
          cmd(k === "PageDown" ? "sheet.next" : "sheet.previous");
          return;
        }
        if (k === ";" || k === ":") {
          e.preventDefault();
          const now = new Date();
          const text = e.shiftKey || k === ":"
            ? `${now.getHours() % 12 || 12}:${String(now.getMinutes()).padStart(2, "0")} ${now.getHours() < 12 ? "AM" : "PM"}`
            : `${now.getMonth() + 1}/${now.getDate()}/${now.getFullYear()}`;
          beginEdit(text, "enter");
          return;
        }
        if (k === "-" ) {
          e.preventDefault();
          const g0 = s.ranges[s.ranges.length - 1];
          if (g0[1] === 0 && g0[3] === MAXC - 1) cmd("home.deleteRows");
          else if (g0[0] === 0 && g0[2] === MAXR - 1) cmd("home.deleteColumns");
          else dialog("deleteCells");
          return;
        }
        if (k === "+" || (k === "=" && e.shiftKey)) {
          e.preventDefault();
          const g0 = s.ranges[s.ranges.length - 1];
          if (g0[1] === 0 && g0[3] === MAXC - 1) cmd("home.insertRows");
          else if (g0[0] === 0 && g0[2] === MAXR - 1) cmd("home.insertColumns");
          else dialog("insertCells");
          return;
        }
        if (e.shiftKey) {
          const fmt = { "$": "currency", "%": "percent", "#": "shortDate", "@": "time", "!": "number", "~": "general" }[k];
          if (fmt) {
            e.preventDefault();
            if (fmt === "percent") cmd("home.percent");
            else if (fmt === "number") cmd("home.numberFormat", { code: "#,##0.00" });
            else cmd("home.numberFormat", { format: fmt });
            return;
          }
          if (k === "(") { e.preventDefault(); cmd("home.unhideRows"); return; }
        }
        if (k === "9") { e.preventDefault(); cmd("home.hideRows"); return; }
        if (k === "0") { e.preventDefault(); cmd("home.hideColumns"); return; }
        const sc = shortcuts[lower];
        if (sc) {
          e.preventDefault();
          cmd(sc[0], sc[1]);
        }
        return;
      }
      if (k === " " && e.shiftKey) {
        e.preventDefault();
        cmd("selection.row");
        return;
      }
      // Printable keys fall through: the textarea receives the character and editing starts
      // in onInput.
    },
    [ac, acceptAc, commit, cancelEdit, setEd, step, pointTo, ensureVisible, backToPanes, objSel, cmd, dialog, ui, expandMerges, setSel, syncSelSoon, selectCell, moveAfter, fr, fc, mainH, mainW, g.defaultRowHeight, g.defaultColWidth, z, beginEdit, cellText]
  );

  const onKbdInput = useCallback(
    (e) => {
      const v = e.target.value;
      const editing = edRef.current;
      if (!editing) {
        if (v === "") return;
        if (ui && ui.protected && ui.ribbon && ui.ribbon.locked) {
          e.target.value = "";
          return;
        }
        const [r, c] = selRef.current.active;
        setEd({ r, c, text: v, mode: "enter", bar: false });
        setListOpen(false);
        pointRef.current = null;
        updateAc(v, e.target.selectionStart);
        return;
      }
      if (pointRef.current && !composingRef.current) pointRef.current = null;
      setEd({ ...editing, text: v });
      updateAc(v, e.target.selectionStart);
    },
    [setEd, updateAc, ui]
  );

  // ------------------------------------------------------------ mouse

  const localPoint = useCallback((ev) => {
    const el = viewRef.current;
    const b = el.getBoundingClientRect();
    return [ev.clientX - b.left, ev.clientY - b.top];
  }, []);

  const headerBoundary = useCallback(
    (x, y) => {
      // Column boundary under a column header, or row boundary under a row header.
      if (y < HH && x >= RW) {
        const [, c] = hit(x, y);
        const left = cx(c);
        const right = left + cw(c);
        if (Math.abs(x - right) <= 4) return { axis: "cols", index: c };
        if (Math.abs(x - left) <= 4 && c > 0) {
          let p = c - 1;
          while (p > 0 && cols.sizeOf(p) === 0) p--;
          return { axis: "cols", index: p };
        }
      }
      if (x < RW && y >= HH) {
        const [r] = hit(x, y);
        const top = ry(r);
        const bottom = top + rh(r);
        if (Math.abs(y - bottom) <= 3) return { axis: "rows", index: r };
        if (Math.abs(y - top) <= 3 && r > 0) {
          let p = r - 1;
          while (p > 0 && rows.sizeOf(p) === 0) p--;
          return { axis: "rows", index: p };
        }
      }
      return null;
    },
    [HH, RW, hit, cx, cw, ry, rh, cols, rows]
  );

  const selRect = useMemo(() => {
    const g0 = sel.ranges[sel.ranges.length - 1];
    if (!g0) return null;
    return rectOf([g0[0], g0[1], Math.min(g0[2], ext.r), Math.min(g0[3], ext.c)]);
  }, [sel, rectOf, ext.r, ext.c]);

  const stopAutoScroll = () => {
    if (autoScrollRef.current) {
      clearInterval(autoScrollRef.current);
      autoScrollRef.current = null;
    }
  };

  const onMouseDown = useCallback(
    (ev) => {
      if (ev.button === 2) return;
      setMenu(null);
      const [x, y] = localPoint(ev);
      const mod = ev.ctrlKey || ev.metaKey;
      const editing = edRef.current;
      ev.preventDefault();
      // Resize?
      const hb = headerBoundary(x, y);
      if (hb) {
        if (ev.detail === 2) {
          const s = selRef.current;
          const full = s.ranges.filter((g0) => (hb.axis === "cols" ? g0[0] === 0 && g0[2] === MAXR - 1 && g0[1] <= hb.index && hb.index <= g0[3] : g0[1] === 0 && g0[3] === MAXC - 1 && g0[0] <= hb.index && hb.index <= g0[2]));
          const idx = full.length ? full.flatMap((g0) => (hb.axis === "cols" ? range(g0[1], g0[3]) : range(g0[0], g0[2]))) : [hb.index];
          send("resize", { axis: hb.axis, indices: idx, autofit: true });
          return;
        }
        const startSize = hb.axis === "cols" ? cols.sizeOf(hb.index) : rows.sizeOf(hb.index);
        dragRef.current = { kind: "resize", axis: hb.axis, index: hb.index, start: hb.axis === "cols" ? ev.clientX : ev.clientY, startSize, size: startSize };
        setDrag({ ...dragRef.current });
        return;
      }
      const [r, c] = hit(x, y);
      // Formula point mode: clicking cells inserts references.
      if (editing && editing.text && editing.text[0] === "=" && r >= 0 && c >= 0) {
        if (pointTo([r, c, r, c])) {
          pointRef.current = { ...pointRef.current, anchor: [r, c], cur: [r, c] };
          dragRef.current = { kind: "point", anchor: [r, c] };
          return;
        }
      }
      // ... and a column or row header inserts the whole column (C:C) or row (4:4), as in Excel;
      // dragging across headers extends it (C:E).
      if (editing && editing.text && editing.text[0] === "=" && (r < 0) !== (c < 0)) {
        const axis = r < 0 ? "cols" : "rows";
        const g0 = axis === "cols" ? [0, c, MAXR - 1, c] : [r, 0, r, MAXC - 1];
        if (pointTo(g0)) {
          pointRef.current = { ...pointRef.current, anchor: [g0[0], g0[1]], cur: [g0[0], g0[1]] };
          dragRef.current = { kind: "point", axis, anchor: axis === "cols" ? c : r };
          return;
        }
      }
      if (editing) commit("none");
      setObjSel(null);
      kbdRef.current && kbdRef.current.focus({ preventScroll: true });
      const s = selRef.current;
      // Fill handle / move border on the current range.
      if (selRect && r >= 0 && c >= 0 && s.ranges.length === 1) {
        const hx = selRect.x + selRect.w, hy = selRect.y + selRect.h;
        if (Math.abs(x - hx) <= 5 && Math.abs(y - hy) <= 5) {
          if (ev.detail === 2) {
            // Double-clicking the handle fills down as far as the data next to it goes.
            dragRef.current = null;
            setDrag(null);
            cmd("edit.autoFillDown", { range: rangeA1(s.ranges[0]) });
            return;
          }
          dragRef.current = { kind: "fill", source: s.ranges[0], target: s.ranges[0], alt: mod };
          setDrag({ ...dragRef.current });
          return;
        }
        const g0 = s.ranges[0];
        const nearEdge =
          (Math.abs(x - selRect.x) <= 3 || Math.abs(x - hx) <= 3) && y >= selRect.y && y <= hy ||
          (Math.abs(y - selRect.y) <= 3 || Math.abs(y - hy) <= 3) && x >= selRect.x && x <= hx;
        if (nearEdge && !(g0[1] === 0 && g0[3] === MAXC - 1) && !(g0[0] === 0 && g0[2] === MAXR - 1)) {
          dragRef.current = { kind: "move", source: g0, grab: [r - g0[0], c - g0[1]], to: [g0[0], g0[1]], copy: mod };
          setDrag({ ...dragRef.current });
          return;
        }
      }
      // Links follow on Ctrl/Cmd+click.
      if (mod && r >= 0 && c >= 0) {
        const e = cellMap.get(key(r, c));
        if (e && e.link) {
          send("link", { r, c });
          return;
        }
      }
      let ns;
      if (r < 0 && c < 0) {
        ns = { active: [Math.max(mainR0, 0), 0], anchor: [0, 0], ranges: [[0, 0, MAXR - 1, MAXC - 1]] };
        ns.active = [0, 0];
        dragRef.current = null;
      } else if (r < 0) {
        const g0 = ev.shiftKey ? [0, Math.min(s.anchor[1], c), MAXR - 1, Math.max(s.anchor[1], c)] : [0, c, MAXR - 1, c];
        const act = ev.shiftKey ? s.active : [Math.max(fr, mainR0), c];
        ns = { active: act, anchor: ev.shiftKey ? s.anchor : [0, c], ranges: mod ? [...s.ranges, g0] : [g0] };
        dragRef.current = { kind: "colsel", anchor: ev.shiftKey ? s.anchor[1] : c, add: mod };
      } else if (c < 0) {
        const g0 = ev.shiftKey ? [Math.min(s.anchor[0], r), 0, Math.max(s.anchor[0], r), MAXC - 1] : [r, 0, r, MAXC - 1];
        const act = ev.shiftKey ? s.active : [r, Math.max(fc, mainC0)];
        ns = { active: act, anchor: ev.shiftKey ? s.anchor : [r, 0], ranges: mod ? [...s.ranges, g0] : [g0] };
        dragRef.current = { kind: "rowsel", anchor: ev.shiftKey ? s.anchor[0] : r, add: mod };
      } else if (ev.shiftKey) {
        const g0 = expandMerges(norm(s.active, [r, c]));
        ns = { active: s.active, anchor: s.active, ranges: [...s.ranges.slice(0, -1), g0] };
        dragRef.current = { kind: "cells", anchor: s.active, add: false, keep: s.ranges.slice(0, -1) };
      } else {
        const m = mergeAt(r, c);
        const g0 = m || [r, c, r, c];
        const act = m ? [m[0], m[1]] : [r, c];
        ns = { active: act, anchor: act, ranges: mod ? [...s.ranges, g0] : [g0] };
        dragRef.current = { kind: "cells", anchor: act, add: mod, keep: mod ? s.ranges : [] };
        if (ev.detail === 2 && !mod) {
          // Double-click edits the cell.
          setSel(ns);
          dragRef.current = null;
          if (ui && ui.protected && ui.ribbon && ui.ribbon.locked) return;
          setTimeout(() => beginEdit(cellText(act[0], act[1]), "edit"), 0);
          return;
        }
      }
      setSel(ns);
      if (!dragRef.current) syncSel(ns);
    },
    [localPoint, headerBoundary, hit, pointTo, commit, selRect, cellMap, send, cmd, mergeAt, expandMerges, setSel, syncSel, cols, rows, fr, fc, mainR0, mainC0, ui, beginEdit, cellText]
  );

  useEffect(() => {
    const move = (ev) => {
      const d = dragRef.current;
      if (!d) return;
      const el = viewRef.current;
      if (!el) return;
      const [x, y] = localPoint(ev);
      if (d.kind === "resize") {
        const delta = (d.axis === "cols" ? ev.clientX : ev.clientY) - d.start;
        d.size = Math.max(0, Math.round(d.startSize + delta / z));
        setDrag({ ...d });
        return;
      }
      // Auto-scroll near the edges.
      const sc = scrollerRef.current;
      const dx = x > view.w - 10 ? 20 : x < mainX && x > RW && fc === 0 ? -20 : x < RW ? -20 : 0;
      const dy = y > view.h - 10 ? 20 : y < HH ? -20 : 0;
      stopAutoScroll();
      if ((dx || dy) && sc) {
        autoScrollRef.current = setInterval(() => {
          sc.scrollTop = Math.max(0, sc.scrollTop + dy);
          sc.scrollLeft = Math.max(0, sc.scrollLeft + dx);
        }, 30);
      }
      let [r, c] = hit(Math.max(x, RW + 1), Math.max(y, HH + 1));
      r = Math.max(0, r);
      c = Math.max(0, c);
      if (d.kind === "point" && d.axis) {
        const g0 = d.axis === "cols" ? [0, Math.min(d.anchor, c), MAXR - 1, Math.max(d.anchor, c)] : [Math.min(d.anchor, r), 0, Math.max(d.anchor, r), MAXC - 1];
        pointTo(g0);
        return;
      }
      if (d.kind === "point") {
        const g0 = norm(d.anchor, [r, c]);
        pointTo(g0);
        pointRef.current = { ...pointRef.current, anchor: d.anchor, cur: [r, c] };
        return;
      }
      if (d.kind === "cells") {
        const g0 = expandMerges(norm(d.anchor, [r, c]));
        const s = selRef.current;
        setSel({ active: s.active, anchor: d.anchor, ranges: [...(d.keep || []), g0] });
        return;
      }
      if (d.kind === "colsel") {
        const [, cc] = hit(Math.max(x, RW + 1), HH + 1);
        const g0 = [0, Math.min(d.anchor, cc), MAXR - 1, Math.max(d.anchor, cc)];
        const s = selRef.current;
        setSel({ ...s, ranges: [...s.ranges.slice(0, -1), g0] });
        return;
      }
      if (d.kind === "rowsel") {
        const [rr] = hit(RW + 1, Math.max(y, HH + 1));
        const g0 = [Math.min(d.anchor, rr), 0, Math.max(d.anchor, rr), MAXC - 1];
        const s = selRef.current;
        setSel({ ...s, ranges: [...s.ranges.slice(0, -1), g0] });
        return;
      }
      if (d.kind === "fill") {
        const src = d.source;
        let t = src;
        const downBy = r - src[2], upBy = src[0] - r, rightBy = c - src[3], leftBy = src[1] - c;
        const vert = Math.max(downBy, upBy) >= Math.max(rightBy, leftBy);
        if (vert && downBy > 0) t = [src[0], src[1], r, src[3]];
        else if (vert && upBy > 0) t = [r, src[1], src[2], src[3]];
        else if (!vert && rightBy > 0) t = [src[0], src[1], src[2], c];
        else if (!vert && leftBy > 0) t = [src[0], c, src[2], src[3]];
        else if (r < src[2] && r >= src[0] && vert) t = [src[0], src[1], r, src[3]];
        else if (c < src[3] && c >= src[1]) t = [src[0], src[1], src[2], c];
        d.target = t;
        d.alt = ev.ctrlKey || ev.metaKey;
        setDrag({ ...d });
        return;
      }
      if (d.kind === "move") {
        const src = d.source;
        const tr = Math.max(0, Math.min(r - d.grab[0], MAXR - (src[2] - src[0]) - 1));
        const tc = Math.max(0, Math.min(c - d.grab[1], MAXC - (src[3] - src[1]) - 1));
        d.to = [tr, tc];
        d.copy = ev.ctrlKey || ev.metaKey;
        setDrag({ ...d });
      }
    };
    const up = () => {
      const d = dragRef.current;
      stopAutoScroll();
      if (!d) return;
      dragRef.current = null;
      setDrag(null);
      if (d.kind === "resize") {
        const s = selRef.current;
        const full = s.ranges.filter((g0) => (d.axis === "cols" ? g0[0] === 0 && g0[2] === MAXR - 1 && g0[1] <= d.index && d.index <= g0[3] : g0[1] === 0 && g0[3] === MAXC - 1 && g0[0] <= d.index && d.index <= g0[2]));
        const idx = full.length ? full.flatMap((g0) => (d.axis === "cols" ? range(g0[1], g0[3]) : range(g0[0], g0[2]))) : [d.index];
        if (d.size !== d.startSize) send("resize", { axis: d.axis, indices: idx, size: d.size });
        return;
      }
      if (d.kind === "point") return;
      if (d.kind === "fill") {
        if (d.target.join() !== d.source.join()) {
          send("fill", { source: rangeA1(d.source), target: rangeA1(d.target), mode: d.alt ? "alternate" : "series" });
          setSel({ ...selRef.current, ranges: [d.target] });
        }
        return;
      }
      if (d.kind === "move") {
        if (d.to[0] !== d.source[0] || d.to[1] !== d.source[1]) {
          send("move", { from: rangeA1(d.source), to: a1(d.to[0], d.to[1]), copy: !!d.copy });
          const h = d.source[2] - d.source[0], w = d.source[3] - d.source[1];
          setSel({ active: d.to, anchor: d.to, ranges: [[d.to[0], d.to[1], d.to[0] + h, d.to[1] + w]] });
        }
        return;
      }
      syncSel();
      if (ui && ui.formatPainter) cmd("edit.applyFormatPainter", { range: rangeA1(selRef.current.ranges[selRef.current.ranges.length - 1]) });
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
    return () => {
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
    };
  }, [localPoint, hit, z, view, mainX, RW, HH, fc, expandMerges, setSel, pointTo, send, syncSel, ui, cmd]);

  const onContextMenu = useCallback(
    (ev) => {
      ev.preventDefault();
      if (edRef.current) return;
      const [x, y] = localPoint(ev);
      const [r, c] = hit(x, y);
      const s = selRef.current;
      const inside = s.ranges.some((g0) => contains(g0, Math.max(r, 0), Math.max(c, 0)) || (r < 0 && g0[0] === 0 && g0[2] === MAXR - 1 && g0[1] <= c && c <= g0[3]) || (c < 0 && g0[1] === 0 && g0[3] === MAXC - 1 && g0[0] <= r && r <= g0[2]));
      if (!inside) {
        if (r < 0 && c >= 0) setSel({ active: [Math.max(fr, mainR0), c], anchor: [0, c], ranges: [[0, c, MAXR - 1, c]] });
        else if (c < 0 && r >= 0) setSel({ active: [r, Math.max(fc, mainC0)], anchor: [r, 0], ranges: [[r, 0, r, MAXC - 1]] });
        else if (r >= 0 && c >= 0) selectCell(r, c, { noScroll: true });
        syncSel();
      }
      const kind = r < 0 && c >= 0 ? "col" : c < 0 && r >= 0 ? "row" : "cell";
      const b = viewRef.current.getBoundingClientRect();
      // The keyboard goes to the grid, as after a left click (Escape closes the menu).
      kbdRef.current && kbdRef.current.focus({ preventScroll: true });
      setMenu({ x: ev.clientX - b.left, y: ev.clientY - b.top, kind, r, c });
    },
    [localPoint, hit, setSel, selectCell, syncSel, fr, fc, mainR0, mainC0]
  );

  const menuItems = useMemo(() => {
    if (!menu) return [];
    const common = [
      { label: "Cut", id: "ui.cut", sc: "Ctrl+X" },
      { label: "Copy", id: "ui.copy", sc: "Ctrl+C" },
      { label: "Paste", id: "ui.paste", sc: "Ctrl+V" },
      { label: "Paste Special…", dialog: "pasteSpecial" },
      { divider: true },
    ];
    if (menu.kind === "col")
      return [...common, { label: "Insert", id: "home.insertColumns" }, { label: "Delete", id: "home.deleteColumns" }, { label: "Clear Contents", id: "edit.clearContents" }, { divider: true }, { label: "Format Cells…", dialog: "formatCells" }, { label: "Column Width…", dialog: "columnWidth" }, { label: "AutoFit Column Width", id: "home.autofitColumnWidth" }, { label: "Hide", id: "home.hideColumns" }, { label: "Unhide", id: "home.unhideColumns" }];
    if (menu.kind === "row")
      return [...common, { label: "Insert", id: "home.insertRows" }, { label: "Delete", id: "home.deleteRows" }, { label: "Clear Contents", id: "edit.clearContents" }, { divider: true }, { label: "Format Cells…", dialog: "formatCells" }, { label: "Row Height…", dialog: "rowHeight" }, { label: "AutoFit Row Height", id: "home.autofitRowHeight" }, { label: "Hide", id: "home.hideRows" }, { label: "Unhide", id: "home.unhideRows" }];
    const note = noteMap.get(key(sel.active[0], sel.active[1]));
    return [
      ...common,
      { label: "Insert…", dialog: "insertCells" },
      { label: "Delete…", dialog: "deleteCells" },
      { label: "Clear Contents", id: "edit.clearContents", sc: "Delete" },
      { divider: true },
      { label: "Sort A to Z", id: "data.sortAscending" },
      { label: "Sort Z to A", id: "data.sortDescending" },
      { label: "Filter", id: "data.filter" },
      { divider: true },
      ...(note && note.threaded
        ? [{ label: "Reply to Comment", dialog: "note", params: { threaded: true } }, { label: "Delete Comment", id: "review.deleteComment" }]
        : note
          ? [{ label: "Edit Note", dialog: "note" }, { label: "Delete Note", id: "review.deleteComment" }]
          : [{ label: "New Comment", dialog: "note", params: { threaded: true } }, { label: "New Note", dialog: "note" }]),
      { label: "Format Cells…", dialog: "formatCells", sc: "Ctrl+1" },
      ui && ui.dropdown ? { label: "Pick From Drop-down List", id: "ui.pickList" } : null,
      { label: "Define Name…", dialog: "defineName" },
      { label: "Link…", dialog: "hyperlink", sc: "Ctrl+K" },
    ].filter(Boolean);
  }, [menu, noteMap, sel, ui]);

  const runMenu = useCallback(
    (it) => {
      setMenu(null);
      kbdRef.current && kbdRef.current.focus({ preventScroll: true });
      if (it.dialog) return dialog(it.dialog, it.params);
      if (it.id === "ui.copy" || it.id === "ui.cut") {
        const text = selectionTsv();
        if (text !== null && navigator.clipboard) navigator.clipboard.writeText(text).catch(() => {});
        pendingCopyRef.current = { text, cut: it.id === "ui.cut" };
        send("copy", { cut: it.id === "ui.cut", text });
        return;
      }
      if (it.id === "ui.paste") {
        if (navigator.clipboard && navigator.clipboard.readText) {
          navigator.clipboard.readText().then((text) => sendPaste(text)).catch(() => sendPaste(null));
        } else sendPaste(null);
        return;
      }
      if (it.id === "ui.pickList") {
        setListOpen(true);
        return;
      }
      cmd(it.id, it.params);
    },
    [dialog, cmd, send, selectionTsv, sendPaste]
  );

  // ------------------------------------------------------------ canvas: gridlines, fills, borders

  const styleOf = useCallback((sid) => (sid ? styles[String(sid)] || null : null), [styles]);

  // Overflowing text spans (row -> list of [c0, c1] covered by text from one cell).
  const layout = useMemo(() => {
    const out = [];
    const occupied = (r, c) => {
      const e = cellMap.get(key(r, c));
      return !!(e && (e.t || e.sp)) || pending[key(r, c)] !== undefined || mergeAt(r, c) !== null;
    };
    for (const p of panes) {
      if (p.w <= 0 || p.h <= 0) continue;
      for (let r = p.r0; r <= p.r1; r++) {
        if (rows.sizeOf(r) === 0) continue;
        for (let c = p.c0; c <= p.c1; c++) {
          const e = cellMap.get(key(r, c));
          const pend = pending[key(r, c)];
          if (!e && !pend) continue;
          const m = mergeAt(r, c);
          if (m && (m[0] !== r || m[1] !== c)) continue;
          if (m && (m[0] < p.r0 || m[1] < p.c0) && !(m[0] === r && m[1] === c)) continue;
          out.push({ pane: p.key, r, c, e: e || { r, c }, pend, m });
        }
      }
      // Merge anchors scrolled out of the pane still draw their merged area.
      for (const m of merges) {
        if (m[2] < p.r0 || m[0] > p.r1 || m[3] < p.c0 || m[1] > p.c1) continue;
        if (m[0] >= p.r0 && m[1] >= p.c0) continue;
        const e = cellMap.get(key(m[0], m[1]));
        out.push({ pane: p.key, r: m[0], c: m[1], e: e || { r: m[0], c: m[1] }, pend: pending[key(m[0], m[1])], m });
      }
    }
    // Overflow extents.
    for (const item of out) {
      const { e, m, r, c, pend } = item;
      const st = styleOf(e.s) || {};
      const text = pend ? pend.text : e.t;
      item.text = text;
      if (text && !m && !st.wr && st.ha === "centerContinuous") {
        // Center Across Selection: the text (a number too) centres over the empty cells to its
        // right that are also Center Across Selection, without a merge.
        let c1 = c;
        while (c1 + 1 < MAXC && c1 + 1 <= (c < fc ? fc - 1 : MAXC - 1) && !occupied(r, c1 + 1)) {
          const n = cellMap.get(key(r, c1 + 1));
          const s = n ? n.s : rowStyles[String(r)] !== undefined ? rowStyles[String(r)] : colStyles[String(c1 + 1)];
          if ((styleOf(s) || {}).ha !== "centerContinuous") break;
          c1++;
        }
        if (c1 !== c) item.span = [c, c1];
        continue;
      }
      if (!text || m || st.wr || st.sh || (e.n && !pend) || e.x && e.x.icon) continue;
      const fs = (st.fs || 11) * (96 / 72) * z;
      const font = `${st.i ? "italic " : ""}${st.b ? "700 " : "400 "}${fs}px ${fontFamily(st.ff)}`;
      const width = textWidth(text, font) + 6 * z;
      const w0 = cw(c);
      if (width <= w0) continue;
      const align = st.ha || (e.a === "r" ? "right" : e.a === "c" ? "center" : "left");
      let c0 = c, c1 = c, acc = w0;
      if (align === "left" || align === "general" || align === "fill" || align === "justify" || align === "distributed") {
        while (acc < width && c1 + 1 < MAXC && !occupied(r, c1 + 1) && c1 + 1 <= (c < fc ? fc - 1 : MAXC - 1)) { c1++; acc += cw(c1); }
      } else if (align === "right") {
        while (acc < width && c0 - 1 >= (c < fc ? 0 : fc) && !occupied(r, c0 - 1)) { c0--; acc += cw(c0); }
      } else {
        let grow = true;
        while (acc < width && grow) {
          grow = false;
          if (c1 + 1 < MAXC && !occupied(r, c1 + 1)) { c1++; acc += cw(c1); grow = true; }
          if (acc < width && c0 - 1 >= (c < fc ? 0 : fc) && !occupied(r, c0 - 1)) { c0--; acc += cw(c0); grow = true; }
        }
      }
      if (c0 !== c || c1 !== c) item.span = [c0, c1];
    }
    return out;
  }, [panes, cellMap, pending, mergeAt, merges, rows, styleOf, z, cw, fc, fontTick]); // row and column styles come with cellMap's data

  useLayoutEffect(() => {
    const cv = canvasRef.current;
    if (!cv) return;
    const dpr = window.devicePixelRatio || 1;
    const W = view.w, H = view.h;
    if (cv.width !== Math.round(W * dpr) || cv.height !== Math.round(H * dpr)) {
      cv.width = Math.round(W * dpr);
      cv.height = Math.round(H * dpr);
      cv.style.width = W + "px";
      cv.style.height = H + "px";
    }
    const ctx = cv.getContext("2d");
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.fillStyle = "#FFFFFF";
    ctx.fillRect(0, 0, W, H);
    for (const p of panes) {
      if (p.w <= 0 || p.h <= 0) continue;
      ctx.save();
      ctx.beginPath();
      ctx.rect(p.x, p.y, p.w, p.h);
      ctx.clip();
      // Row and column styles (fills of whole rows/columns).
      for (let c = p.c0; c <= p.c1; c++) {
        const st = styleOf(colStyles[String(c)]);
        if (st && st.bg) {
          ctx.fillStyle = st.bg;
          ctx.fillRect(cx(c) - 1, p.y, cw(c) + 1, p.h);
        }
      }
      for (let r = p.r0; r <= p.r1; r++) {
        const st = styleOf(rowStyles[String(r)]);
        if (st && st.bg) {
          ctx.fillStyle = st.bg;
          ctx.fillRect(p.x, ry(r) - 1, p.w, rh(r) + 1);
        }
      }
      // Gridlines.
      if (g.gridlines !== false) {
        ctx.fillStyle = GRID;
        for (let r = p.r0; r <= p.r1 + 1; r++) {
          const y = Math.round(ry(r) + rh(r)) - 1;
          if (rh(r) > 0) ctx.fillRect(p.x, y, p.w, 1);
        }
        for (let c = p.c0; c <= p.c1 + 1; c++) {
          const x = Math.round(cx(c) + cw(c)) - 1;
          if (cw(c) > 0) ctx.fillRect(x, p.y, 1, p.h);
        }
      }
      // Overflowing text and merged areas hide the gridlines they cover.
      for (const it of layout) {
        if (it.pane !== p.key) continue;
        if (it.span) {
          const x0 = cx(it.span[0]), x1 = cx(it.span[1]) + cw(it.span[1]);
          ctx.fillStyle = "#FFFFFF";
          ctx.fillRect(x0, ry(it.r), x1 - x0 - 1, rh(it.r) - 1);
        }
        if (it.m) {
          const rc = rectOf(it.m);
          ctx.fillStyle = "#FFFFFF";
          ctx.fillRect(rc.x, rc.y, rc.w - 1, rc.h - 1);
        }
      }
      // Fills (cell style, table style, conditional formats).
      for (const it of layout) {
        if (it.pane !== p.key) continue;
        const e = it.e;
        const st = styleOf(e.s) || {};
        const bg = (e.x && e.x.bg) || st.bg || (e.tx && e.tx.bg);
        if (!bg) continue;
        const rc = it.m ? rectOf(it.m) : { x: cx(it.c), y: ry(it.r), w: cw(it.c), h: rh(it.r) };
        ctx.fillStyle = bg;
        ctx.fillRect(rc.x - 1, rc.y - 1, rc.w + 1, rc.h + 1);
      }
      // Data bars.
      for (const it of layout) {
        if (it.pane !== p.key || !it.e.x || !it.e.x.bar) continue;
        const [frac, color] = it.e.x.bar;
        const x = cx(it.c) + 2, y = ry(it.r) + 2, w = (cw(it.c) - 5) * frac, h = rh(it.r) - 5;
        if (w <= 0 || h <= 0) continue;
        const grd = ctx.createLinearGradient(x, 0, x + w, 0);
        grd.addColorStop(0, color);
        grd.addColorStop(1, "#FFFFFF");
        ctx.fillStyle = grd;
        ctx.fillRect(x, y, w, h);
        ctx.fillStyle = color;
        ctx.fillRect(x, y, w, 1);
        ctx.fillRect(x, y + h - 1, w, 1);
        ctx.fillRect(x, y, 1, h);
      }
      // Borders.
      const sidesOf = (e) => {
        const st = styleOf(e.s) || {};
        const tx = e.tx || {};
        const x = e.x || {};
        return { bt: x.bt || st.bt || tx.bt, bb: x.bb || st.bb || tx.bb, bl: x.bl || st.bl || tx.bl, br: x.br || st.br || tx.br };
      };
      for (const it of layout) {
        if (it.pane !== p.key) continue;
        const e = it.e;
        const st = styleOf(e.s) || {};
        if (it.m) {
          // A merged area's edges are drawn from the cells along them, where Excel keeps them (its
          // right border on the right column's cells, its bottom one on the bottom row's).
          const m = it.m;
          const at = (r, c) => cellMap.get(key(r, c)) || (r === m[0] && c === m[1] ? e : null);
          for (let c = m[1]; c <= m[3]; c++) {
            const x0 = Math.round(cx(c)) - 1, x1 = Math.round(cx(c) + cw(c));
            const t = at(m[0], c), b = at(m[2], c);
            const ts = t && sidesOf(t).bt, bs = b && sidesOf(b).bb;
            if (ts) drawEdge(ctx, ts, true, Math.round(ry(m[0])) - 1, x0, x1);
            if (bs) drawEdge(ctx, bs, true, Math.round(ry(m[2]) + rh(m[2])) - 1, x0, x1);
          }
          for (let r = m[0]; r <= m[2]; r++) {
            const y0 = Math.round(ry(r)) - 1, y1 = Math.round(ry(r) + rh(r));
            const l = at(r, m[1]), rr = at(r, m[3]);
            const ls = l && sidesOf(l).bl, rs = rr && sidesOf(rr).br;
            if (ls) drawEdge(ctx, ls, false, Math.round(cx(m[1])) - 1, y0, y1);
            if (rs) drawEdge(ctx, rs, false, Math.round(cx(m[3]) + cw(m[3])) - 1, y0, y1);
          }
          if (!st.dd && !st.du) continue;
        }
        const sides = it.m ? {} : sidesOf(e);
        if (!sides.bt && !sides.bb && !sides.bl && !sides.br && !st.dd && !st.du) continue;
        const rc = it.m ? rectOf(it.m) : { x: cx(it.c), y: ry(it.r), w: cw(it.c), h: rh(it.r) };
        const L = Math.round(rc.x) - 1, T = Math.round(rc.y) - 1, R = Math.round(rc.x + rc.w) - 1, B = Math.round(rc.y + rc.h) - 1;
        if (sides.bt) drawEdge(ctx, sides.bt, true, T, L, R + 1);
        if (sides.bb) drawEdge(ctx, sides.bb, true, B, L, R + 1);
        if (sides.bl) drawEdge(ctx, sides.bl, false, L, T, B + 1);
        if (sides.br) drawEdge(ctx, sides.br, false, R, T, B + 1);
        if (st.dd || st.du) {
          ctx.save();
          if (st.dd) {
            ctx.strokeStyle = st.dd[1] || "#000";
            ctx.beginPath();
            ctx.moveTo(L, T);
            ctx.lineTo(R, B);
            ctx.stroke();
          }
          if (st.du) {
            ctx.strokeStyle = st.du[1] || "#000";
            ctx.beginPath();
            ctx.moveTo(L, B);
            ctx.lineTo(R, T);
            ctx.stroke();
          }
          ctx.restore();
        }
      }
      ctx.restore();
    }
    // Pane separators for frozen panes.
    ctx.fillStyle = "#9AA4AE";
    if (fr) ctx.fillRect(RW, mainY - 1, view.w - RW, 1);
    if (fc) ctx.fillRect(mainX - 1, HH, 1, view.h - HH);
  });

  // ------------------------------------------------------------ render pieces

  const cellsLayer = (p) => {
    const out = [];
    for (const it of layout) {
      if (it.pane !== p.key) continue;
      const { e, r, c, m } = it;
      const st = styleOf(e.s) || {};
      const x = e.x || {};
      const tx = e.tx || {};
      let rc = m ? rectOf(m) : { x: cx(c), y: ry(r), w: cw(c), h: rh(r) };
      if (rc.w <= 0 || rc.h <= 0) continue;
      const spark = sparkMap.get(key(r, c));
      let text = it.text;
      if (x.hide) text = "";
      const hasIcon = x.icon;
      if (!text && !hasIcon && !spark && !e.link) continue;
      const fsPt = st.fs || 11;
      const fs = fsPt * (96 / 72) * z * (st.va ? 0.75 : 1);
      const bold = st.b || x.b || tx.b;
      const italic = st.i || x.i;
      const family = fontFamily(st.ff);
      const font = `${italic ? "italic " : ""}${bold ? "700 " : "400 "}${fs}px ${family}`;
      let halign = st.ha || (e.a === "r" ? "right" : e.a === "c" ? "center" : "left");
      if (halign === "general") halign = "left";
      const wrap = !!st.wr || halign === "justify" || halign === "distributed";
      if (it.span) {
        const x0 = cx(it.span[0]), x1 = cx(it.span[1]) + cw(it.span[1]);
        rc = { ...rc, x: x0, w: x1 - x0 };
      }
      const pad = 3 * z;
      let shown = text;
      // With Wrap Text off a multi-line value shows on one line, as in Excel (the breaks stay in the
      // value; the editor shows them).
      if (!wrap && shown && shown.includes("\n")) shown = shown.replace(/\n/g, "");
      // Shrink to Fit: a smaller font instead of spilling over or showing ###.
      let fsShown = fs;
      if (st.sh && !wrap && text && !it.span) {
        const avail = rc.w - 2 * pad - 1;
        const tw = textWidth(text, font);
        if (tw > avail && avail > 0) fsShown = Math.max(1, (fs * avail) / tw);
      }
      if (e.n && text && !it.pend && !wrap && fsShown === fs) {
        const avail = rc.w - 3;
        if (textWidth(text, font) * EXCEL_DIGIT_FIT > avail) {
          const hw = textWidth("#", font) || 7;
          shown = "#".repeat(Math.max(1, Math.floor(avail / hw)));
        }
      }
      const color = x.c || e.fc || st.c || tx.c || "#000000";
      const deco = [st.u || x.u ? "underline" : "", st.st || x.st ? "line-through" : ""].filter(Boolean).join(" ");
      const valign = st.vt === "top" ? "flex-start" : st.vt === "center" || st.vt === "distributed" || st.vt === "justify" ? "center" : "flex-end";
      const justify = halign === "right" ? "flex-end" : halign === "center" || halign === "centerContinuous" ? "center" : "flex-start";
      const indent = (st.ind || 0) * 9 * z;
      const style = {
        position: "absolute",
        left: rc.x - p.x,
        top: rc.y - p.y,
        width: rc.w - 1,
        height: rc.h - 1,
        display: "flex",
        alignItems: valign,
        justifyContent: justify,
        padding: `0 ${pad + (halign === "right" ? indent : 0)}px 0 ${pad + (halign !== "right" ? indent : 0)}px`,
        boxSizing: "border-box",
        overflow: "hidden",
        fontStyle: italic ? "italic" : "normal",
        fontWeight: bold ? 700 : 400,
        fontSize: fsShown + "px",
        fontFamily: family,
        color,
        textDecoration: deco || undefined,
        textDecorationStyle: st.u === "double" || st.u === "doubleAccounting" ? "double" : undefined,
        whiteSpace: wrap ? "pre-wrap" : "pre",
        wordBreak: wrap ? "break-word" : undefined,
        lineHeight: 1.15,
        cursor: e.link ? "pointer" : undefined,
        pointerEvents: "none",
      };
      let inner;
      if (e.fill && !wrap && shown === text) {
        const [, pos] = e.fill;
        inner = (
          <>
            <span>{text.slice(0, pos)}</span>
            <span style={{ flex: 1 }} />
            <span>{text.slice(pos)}</span>
          </>
        );
        style.justifyContent = "stretch";
      } else if (wrap) {
        inner = <div style={{ width: "100%", textAlign: halign === "right" ? "right" : halign === "center" ? "center" : halign === "justify" ? "justify" : "left" }}>{shown}</div>;
      } else if (st.rot) {
        const deg = st.rot === 255 ? 0 : st.rot > 90 ? st.rot - 90 : -st.rot;
        inner = st.rot === 255 ? <span style={{ writingMode: "vertical-rl", textOrientation: "upright" }}>{shown}</span> : <span style={{ transform: `rotate(${deg}deg)`, display: "inline-block" }}>{shown}</span>;
      } else {
        inner = st.va ? <span style={{ position: "relative", top: st.va === "superscript" ? -fs * 0.4 : fs * 0.2 }}>{shown}</span> : shown;
      }
      out.push(
        <div key={`${p.key}:${r}:${c}`} className="rg-cell" data-cell={a1(r, c)} style={style}>
          {hasIcon && <span style={{ marginRight: "auto", display: "flex", alignItems: "center" }}><CfIcon set={x.icon[0]} index={x.icon[1]} size={Math.min(16 * z, rc.h - 2)} /></span>}
          {spark ? null : inner}
          {spark && (
            <div className="rg-spark" data-spark={a1(r, c)} style={{ position: "absolute", inset: 0 }}>
              <Sparkline spark={spark} width={rc.w - 1} height={rc.h - 1} />
            </div>
          )}
        </div>
      );
    }
    // Sparklines in empty cells.
    for (const s of (data && data.sparklines) || []) {
      if (s.r < p.r0 || s.r > p.r1 || s.c < p.c0 || s.c > p.c1 || cellMap.get(key(s.r, s.c))) continue;
      out.push(
        <div key={`sp:${p.key}:${s.r}:${s.c}`} className="rg-spark" data-spark={a1(s.r, s.c)} style={{ position: "absolute", left: cx(s.c) - p.x, top: ry(s.r) - p.y, width: cw(s.c) - 1, height: rh(s.r) - 1, pointerEvents: "none" }}>
          <Sparkline spark={s} width={cw(s.c) - 1} height={rh(s.r) - 1} />
        </div>
      );
    }
    return out;
  };

  const decorLayer = (p) => {
    const out = [];
    // Note indicators.
    for (const n of (data && data.notes) || []) {
      if (n.r < p.r0 || n.r > p.r1 || n.c < p.c0 || n.c > p.c1) continue;
      const m = mergeAt(n.r, n.c);
      const rc = m ? rectOf(m) : { x: cx(n.c), y: ry(n.r), w: cw(n.c), h: rh(n.r) };
      const color = n.threaded ? "#7A3FB8" : "#D7263D";
      out.push(
        <div
          key={`n:${p.key}:${n.r}:${n.c}`}
          className={n.threaded ? "rg-comment-mark" : "rg-note-mark"}
          style={{ position: "absolute", left: rc.x + rc.w - p.x - 8, top: rc.y - p.y, width: 8, height: 8 }}
        >
          <div style={{ width: 0, height: 0, borderTop: `6px solid ${color}`, borderLeft: "6px solid transparent", marginLeft: 1 }} />
        </div>
      );
    }
    // Filter buttons.
    for (const b of (data && data.buttons) || []) {
      if (b.r < p.r0 || b.r > p.r1 || b.c < p.c0 || b.c > p.c1) continue;
      const size = Math.min(16 * z, rh(b.r) - 2);
      out.push(
        <div
          key={`f:${p.key}:${b.r}:${b.c}`}
          className="rg-filter-btn"
          data-filter={a1(b.r, b.c)}
          title={b.active ? "Filtered" : "Filter"}
          onMouseDown={(ev) => {
            ev.stopPropagation();
            ev.preventDefault();
            send("filter", { r: b.r, c: b.c, table: b.table || null });
          }}
          style={{ position: "absolute", left: cx(b.c) + cw(b.c) - p.x - size - 2, top: ry(b.r) + rh(b.r) - p.y - size - 2, width: size, height: size, border: "1px solid #A6ADB4", background: "#F7F8F9", borderRadius: 2, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", pointerEvents: "auto", boxSizing: "border-box" }}
        >
          {b.active ? (
            <svg width={10} height={10} viewBox="0 0 10 10"><path d="M1 1 L9 1 L6 5 L6 9 L4 8 L4 5 Z" fill={ACCENT} /></svg>
          ) : (
            <svg width={8} height={8} viewBox="0 0 8 8"><path d="M1 2 L7 2 L4 6 Z" fill="#404040" /></svg>
          )}
        </div>
      );
    }
    return out;
  };

  const peerLayer = (p) => {
    const out = [];
    for (const f of flashes) {
      if (f.sheet !== sheetName) continue;
      const cg = clipToPane(f.g, p);
      if (!cg) continue;
      const rc = rectOf(cg);
      out.push(<div key={`fl${f.id}`} className="rg-flash" style={{ position: "absolute", left: rc.x - p.x, top: rc.y - p.y, width: rc.w - 1, height: rc.h - 1, background: f.color, pointerEvents: "none" }} />);
    }
    // Ranges first, then each person's active cell; people on the same cell share one row of tags.
    for (const q of others) {
      (q.ranges || []).forEach((g0, i) => {
        if (g0[0] === g0[2] && g0[1] === g0[3]) return;
        const cg = clipToPane(g0, p);
        if (!cg) return;
        const rc = rectOf(cg);
        out.push(<div key={`pr${q.id}:${i}`} className="rg-peer-range" style={{ position: "absolute", ...paneBox(rc, p, 1), border: `1px solid ${q.color}`, background: q.color + "1F", boxSizing: "border-box", pointerEvents: "none" }} />);
      });
    }
    const groups = new Map();
    for (const q of others) {
      const am = mergeAt(q.active[0], q.active[1]);
      const ag = am || [q.active[0], q.active[1], q.active[0], q.active[1]];
      const k = ag.join(",");
      if (!groups.has(k)) groups.set(k, { ag, people: [] });
      groups.get(k).people.push(q);
    }
    const now = Date.now();
    for (const [k, grp] of groups) {
      if (!clipToPane(grp.ag, p)) continue;
      const rc = rectOf(grp.ag);
      const tags = [];
      grp.people.forEach((q, i) => {
        const editing = !!q.editing;
        const inset = Math.min(i * 2, Math.floor(Math.min(rc.w, rc.h) / 2) - 2);
        out.push(<div key={`pa${q.id}`} className="rg-peer" data-peer={q.name} data-peer-cell={a1(q.active[0], q.active[1])} data-editing={editing ? "1" : "0"} style={{ position: "absolute", left: rc.x - p.x - 1 + inset, top: rc.y - p.y - 1 + inset, width: rc.w + 1 - 2 * inset, height: rc.h + 1 - 2 * inset, border: `2px ${editing ? "dashed" : "solid"} ${q.color}`, boxSizing: "border-box", pointerEvents: "none" }} />);
        const moved = peerMoved.current[q.id];
        if (editing || !moved || now - moved.at < TAG_MS) tags.push(q);
        else out.push(<div key={`pf${q.id}`} className="rg-peer-flag" title={q.name} style={{ position: "absolute", left: rc.x - p.x + rc.w - 7 - i * 7, top: rc.y - p.y - 1, width: 6, height: 6, background: q.color, pointerEvents: "none" }} />);
      });
      if (!tags.length) continue;
      const above = rc.y - p.y - 17 >= 0;
      out.push(
        <div key={`pt${k}`} style={{ position: "absolute", left: rc.x - p.x - 1, top: above ? rc.y - p.y - 17 : rc.y - p.y + rc.h, display: "flex", gap: 2, pointerEvents: "none", zIndex: 4 }}>
          {tags.map((q) => (
            <div key={q.id} className="rg-peer-tag" data-peer={q.name} style={{ background: q.color, color: "#FFFFFF", font: "600 11px 'Segoe UI', system-ui, sans-serif", lineHeight: "16px", height: 16, padding: "0 5px", borderRadius: above ? "3px 3px 3px 0" : "0 3px 3px 3px", whiteSpace: "nowrap" }}>
              {q.name}{q.editing ? " ✎" : ""}
            </div>
          ))}
        </div>
      );
    }
    return out;
  };

  const selectionLayer = (p) => {
    const out = [];
    const s = sel;
    const multi = s.ranges.length > 1 || !(s.ranges[0][0] === s.ranges[0][2] && s.ranges[0][1] === s.ranges[0][3]);
    s.ranges.forEach((g0, i) => {
      const cg = clipToPane(g0, p);
      if (!cg) return;
      const rc = rectOf(cg);
      const isCur = i === s.ranges.length - 1;
      if (multi) out.push(<div key={`sf${i}`} style={{ position: "absolute", left: rc.x - p.x, top: rc.y - p.y, width: rc.w, height: rc.h, background: SEL_FILL, pointerEvents: "none" }} />);
      if (isCur && s.ranges.length === 1) {
        const full = rectOf(g0);
        out.push(<div key={`sb${i}`} className="rg-selection" style={{ position: "absolute", ...paneBox(full, p, 1), border: `2px solid ${ACCENT}`, boxSizing: "border-box", pointerEvents: "none" }} />);
      }
    });
    // Active cell (white, with its own border inside a multi-cell selection).
    const am = mergeAt(s.active[0], s.active[1]);
    const ag = am || [s.active[0], s.active[1], s.active[0], s.active[1]];
    if (clipToPane(ag, p)) {
      const rc = rectOf(ag);
      if (multi) out.push(<div key="act" style={{ position: "absolute", left: rc.x - p.x, top: rc.y - p.y, width: rc.w - 1, height: rc.h - 1, background: "rgba(255,255,255,0.0)", outline: s.ranges.length > 1 ? `1px solid ${ACCENT}` : "none", pointerEvents: "none" }} />);
      if (!multi) out.push(<div key="actb" className="rg-active" style={{ position: "absolute", left: rc.x - p.x - 1, top: rc.y - p.y - 1, width: rc.w + 1, height: rc.h + 1, border: `2px solid ${ACCENT}`, boxSizing: "border-box", pointerEvents: "none" }} />);
    }
    // Fill handle.
    if (s.ranges.length === 1 && !ed) {
      const g0 = s.ranges[0];
      const corner = [g0[2] >= MAXR - 1 ? g0[0] : g0[2], g0[3] >= MAXC - 1 ? g0[1] : g0[3]];
      if (g0[2] < MAXR - 1 || g0[3] < MAXC - 1) {
        const rr = g0[2] >= MAXR - 1 ? Math.min(g0[2], p.r1) : g0[2];
        const ccc = g0[3] >= MAXC - 1 ? Math.min(g0[3], p.c1) : g0[3];
        if (contains([p.r0, p.c0, p.r1, p.c1], rr, ccc)) {
          const hx = cx(ccc) + cw(ccc) - p.x, hy = ry(rr) + rh(rr) - p.y;
          out.push(<div key="fh" className="rg-fill-handle" style={{ position: "absolute", left: hx - 4, top: hy - 4, width: 7, height: 7, background: ACCENT, border: "1px solid #FFFFFF", boxSizing: "border-box", cursor: "crosshair", pointerEvents: "none" }} />);
        }
      }
    }
    // Copy marquee.
    if (ui && ui.clipboard && ui.clipboard.sheet === sheetIndex) {
      const cg = clipToPane(ui.clipboard.range, p);
      if (cg) {
        const bx = paneBox(rectOf(ui.clipboard.range), p, 2);
        out.push(
          <svg key="marquee" style={{ position: "absolute", ...bx, pointerEvents: "none", overflow: "visible" }}>
            <rect x={1} y={1} width={Math.max(bx.width - 3, 1)} height={Math.max(bx.height - 3, 1)} fill="none" stroke={ACCENT} strokeWidth={2} strokeDasharray="5 3" className="rg-marquee" />
          </svg>
        );
      }
    }
    // Spill range border.
    if (ui && ui.spill && ui.sheet === sheetIndex) {
      const sg = ui.spill.slice(2);
      if (clipToPane(sg, p)) {
        const rc = rectOf(sg);
        out.push(<div key="spill" style={{ position: "absolute", ...paneBox(rc, p, 1), border: "1px dashed #2F6FDB", boxSizing: "border-box", pointerEvents: "none" }} />);
      }
    }
    // Formula references while editing.
    if (ed && ed.text && ed.text[0] === "=") {
      formulaRefs(ed.text).forEach((ref, i) => {
        const cg = clipToPane(ref.range, p);
        if (!cg) return;
        const rc = rectOf(ref.range);
        const color = REF_COLORS[i % REF_COLORS.length];
        out.push(<div key={`ref${i}`} style={{ position: "absolute", ...paneBox(rc, p, 1), border: `2px solid ${color}`, background: color + "14", boxSizing: "border-box", pointerEvents: "none" }} />);
      });
    }
    // Drag previews.
    if (drag && (drag.kind === "fill" || drag.kind === "move")) {
      const g0 = drag.kind === "fill" ? drag.target : [drag.to[0], drag.to[1], drag.to[0] + drag.source[2] - drag.source[0], drag.to[1] + drag.source[3] - drag.source[1]];
      if (clipToPane(g0, p)) {
        const rc = rectOf(g0);
        out.push(<div key="dragp" style={{ position: "absolute", ...paneBox(rc, p, 1), border: `2px dashed ${drag.kind === "fill" ? "#5F6B76" : ACCENT}`, boxSizing: "border-box", pointerEvents: "none" }} />);
      }
    }
    return out;
  };

  // Objects (charts, pictures) in sheet coordinates.
  const sheetToScreen = useCallback(
    (x, y) => {
      const sx = x < cols.start(fc) ? RW + x * z : mainX + (x - cols.start(fc)) * z - scroll.left;
      const sy = y < rows.start(fr) ? HH + y * z : mainY + (y - rows.start(fr)) * z - scroll.top;
      return [sx, sy];
    },
    [cols, rows, fc, fr, RW, mainX, HH, mainY, z, scroll.left, scroll.top]
  );

  const objectLayer = () => {
    const out = [];
    const items = [
      ...(((objects && objects.charts) || []).map((c) => ({ kind: "chart", id: c.id, x: c.x, y: c.y, width: c.width, height: c.height, chart: c }))),
      ...(((objects && objects.images) || []).map((im) => ({ kind: "image", id: im.id, x: im.x, y: im.y, width: im.width, height: im.height, src: im.src, alt: im.alt }))),
    ];
    for (const o of items) {
      const live = objDrag && objDrag.id === o.id ? objDrag : null;
      const ox = live ? live.x : o.x, oy = live ? live.y : o.y, ow = live ? live.width : o.width, oh = live ? live.height : o.height;
      const [sx, sy] = sheetToScreen(ox, oy);
      const w = ow * z, h = oh * z;
      if (sx > view.w || sy > view.h || sx + w < RW || sy + h < HH) continue;
      const selected = objSel === o.id;
      const startDrag = (ev, mode) => {
        ev.stopPropagation();
        ev.preventDefault();
        if (edRef.current) commit("none");
        setObjSel(o.id);
        kbdRef.current && kbdRef.current.focus({ preventScroll: true });
        const start = { mx: ev.clientX, my: ev.clientY, x: o.x, y: o.y, width: o.width, height: o.height };
        let last = null;
        const moveFn = (e2) => {
          const dx = (e2.clientX - start.mx) / z, dy = (e2.clientY - start.my) / z;
          let n = { id: o.id, x: start.x, y: start.y, width: start.width, height: start.height };
          if (mode === "move") n = { ...n, x: Math.max(0, start.x + dx), y: Math.max(0, start.y + dy) };
          else {
            // A handle ("nw", "e"…) resizes; a move never does, though "move" has an "e" too.
            if (mode.includes("e")) n.width = Math.max(20, start.width + dx);
            if (mode.includes("s")) n.height = Math.max(20, start.height + dy);
            if (mode.includes("w")) { n.width = Math.max(20, start.width - dx); n.x = start.x + (start.width - n.width); }
            if (mode.includes("n")) { n.height = Math.max(20, start.height - dy); n.y = start.y + (start.height - n.height); }
          }
          last = n;
          setObjDrag(n);
        };
        const upFn = (e2) => {
          window.removeEventListener("mousemove", moveFn);
          window.removeEventListener("mouseup", upFn);
          setObjDrag(null);
          // Sent here, not from a state updater: React may call those twice (Strict Mode), which
          // would send the move twice and leave an empty undo step.
          if (last && (last.x !== o.x || last.y !== o.y || last.width !== o.width || last.height !== o.height)) {
            send("object", { id: o.id, x: last.x, y: last.y, width: last.width, height: last.height });
          }
        };
        window.addEventListener("mousemove", moveFn);
        window.addEventListener("mouseup", upFn);
      };
      const handles = selected
        ? ["nw", "n", "ne", "e", "se", "s", "sw", "w"].map((hd) => {
            const hx = hd.includes("w") ? -4 : hd.includes("e") ? w - 4 : w / 2 - 4;
            const hy = hd.includes("n") ? -4 : hd.includes("s") ? h - 4 : h / 2 - 4;
            return <div key={hd} onMouseDown={(ev) => startDrag(ev, hd)} style={{ position: "absolute", left: hx, top: hy, width: 8, height: 8, background: "#FFFFFF", border: `1px solid ${ACCENT}`, borderRadius: 4, cursor: `${hd}-resize` }} />;
          })
        : null;
      out.push(
        <div
          key={`${o.kind}${o.id}`}
          className={o.kind === "chart" ? "rg-chart" : "rg-image"}
          data-object={o.id}
          onMouseDown={(ev) => startDrag(ev, "move")}
          style={{ position: "absolute", left: sx, top: sy, width: w, height: h, border: selected ? `1px solid ${ACCENT}` : "1px solid #D0D4D9", background: "#FFFFFF", boxSizing: "border-box", cursor: "move", boxShadow: selected ? "0 1px 6px rgba(0,0,0,0.18)" : "none" }}
        >
          {o.kind === "chart" ? <ChartSvg chart={o.chart} width={w - 2} height={h - 2} /> : <img src={o.src} alt={o.alt || ""} draggable={false} style={{ width: "100%", height: "100%", objectFit: "fill", display: "block" }} />}
          {handles}
        </div>
      );
    }
    return out;
  };

  // ------------------------------------------------------------ headers

  const headerLayer = () => {
    if (!headings) return null;
    const out = [];
    const s = sel;
    const colSel = (c) => s.ranges.some((g0) => g0[1] <= c && c <= g0[3]);
    const colFull = (c) => s.ranges.some((g0) => g0[0] === 0 && g0[2] === MAXR - 1 && g0[1] <= c && c <= g0[3]);
    const rowSel = (r) => s.ranges.some((g0) => g0[0] <= r && r <= g0[2]);
    const rowFull = (r) => s.ranges.some((g0) => g0[1] === 0 && g0[3] === MAXC - 1 && g0[0] <= r && r <= g0[2]);
    const hfontSize = `${12 * z}px`;
    const hfontFamily = "'Segoe UI', system-ui, Arial, sans-serif";
    const colRanges = [];
    if (fc) colRanges.push([0, fc - 1]);
    colRanges.push([Math.max(mainC0, fc), mainC1]);
    for (const [c0, c1] of colRanges) {
      for (let c = c0; c <= c1; c++) {
        const w = cw(c);
        if (w <= 0) continue;
        const x = cx(c);
        if (c >= fc && x + w < mainX) continue;
        const on = colSel(c), full = colFull(c);
        out.push(
          <div key={`ch${c}`} className="rg-colhead" style={{ position: "absolute", left: x, top: 0, width: w - 1, height: HH - 1, background: full ? ACCENT : on ? HEAD_SEL : HEAD_BG, color: full ? "#FFFFFF" : on ? ACCENT : "#3C4043", fontSize: hfontSize, fontFamily: hfontFamily, fontWeight: on ? 600 : 400, display: "flex", alignItems: "center", justifyContent: "center", borderRight: `1px solid ${GRID}`, borderBottom: on && !full ? `2px solid ${ACCENT}` : `1px solid ${GRID}`, boxSizing: "content-box", userSelect: "none", clipPath: c >= fc ? `inset(0 0 0 ${Math.max(0, mainX - x)}px)` : undefined }}
          >
            {colName(c)}
          </div>
        );
      }
    }
    const rowRanges = [];
    if (fr) rowRanges.push([0, fr - 1]);
    rowRanges.push([Math.max(mainR0, fr), mainR1]);
    for (const [r0, r1] of rowRanges) {
      for (let r = r0; r <= r1; r++) {
        const h = rh(r);
        if (h <= 0) continue;
        const y = ry(r);
        if (r >= fr && y + h < mainY) continue;
        const on = rowSel(r), full = rowFull(r);
        out.push(
          <div key={`rh${r}`} className="rg-rowhead" style={{ position: "absolute", left: 0, top: y, width: RW - 1, height: h - 1, background: full ? ACCENT : on ? HEAD_SEL : HEAD_BG, color: full ? "#FFFFFF" : on ? ACCENT : "#3C4043", fontSize: hfontSize, fontFamily: hfontFamily, fontWeight: on ? 600 : 400, display: "flex", alignItems: "center", justifyContent: "center", borderBottom: `1px solid ${GRID}`, borderRight: on && !full ? `2px solid ${ACCENT}` : `1px solid ${GRID}`, boxSizing: "content-box", userSelect: "none", clipPath: r >= fr ? `inset(${Math.max(0, mainY - y)}px 0 0 0)` : undefined }}
          >
            {r + 1}
          </div>
        );
      }
    }
    out.push(
      <div key="corner" className="rg-corner" style={{ position: "absolute", left: 0, top: 0, width: RW - 1, height: HH - 1, background: HEAD_BG, borderRight: `1px solid ${GRID}`, borderBottom: `1px solid ${GRID}` }}>
        <svg width={RW - 1} height={HH - 1}><path d={`M ${RW - 5} ${HH - 14} L ${RW - 5} ${HH - 5} L ${RW - 14} ${HH - 5} Z`} fill="#B6BEC6" /></svg>
      </div>
    );
    if (drag && drag.kind === "resize") {
      if (drag.axis === "cols") {
        const x = cx(drag.index) + drag.size * z;
        out.push(<div key="rz" style={{ position: "absolute", left: x - 1, top: 0, width: 1, height: view.h, background: ACCENT }} />);
        out.push(<div key="rzt" style={{ position: "absolute", left: x + 6, top: HH + 4, background: "#FFFFE1", border: "1px solid #A0A0A0", padding: "1px 6px", font: "12px system-ui", zIndex: 5 }}>{`Width: ${drag.size} px`}</div>);
      } else {
        const y = ry(drag.index) + drag.size * z;
        out.push(<div key="rz" style={{ position: "absolute", left: 0, top: y - 1, width: view.w, height: 1, background: ACCENT }} />);
        out.push(<div key="rzt" style={{ position: "absolute", left: RW + 6, top: y + 6, background: "#FFFFE1", border: "1px solid #A0A0A0", padding: "1px 6px", font: "12px system-ui", zIndex: 5 }}>{`Height: ${drag.size} px`}</div>);
      }
    }
    return out;
  };

  // ------------------------------------------------------------ editor overlay

  const activeRect = useMemo(() => {
    const [r, c] = ed ? [ed.r, ed.c] : sel.active;
    const m = mergeAt(r, c);
    return rectOf(m || [r, c, r, c]);
  }, [ed, sel, mergeAt, rectOf]);

  const editorStyle = useMemo(() => {
    const [r, c] = ed ? [ed.r, ed.c] : sel.active;
    const e = cellMap.get(key(r, c));
    const st = (e && styleOf(e.s)) || {};
    const fs = (st.fs || 11) * (96 / 72) * z;
    const font = `${st.i ? "italic " : ""}${st.b ? "700 " : "400 "}${fs}px ${fontFamily(st.ff)}`;
    const fp = fontProps(st, z);
    if (!ed || ed.bar) {
      return { position: "absolute", left: activeRect.x, top: activeRect.y, width: 1, height: 1, opacity: 0, padding: 0, border: 0, resize: "none", overflow: "hidden", ...fp, lineHeight: "normal", zIndex: 1, outline: "none", whiteSpace: "pre" };
    }
    const lines = ed.text.split("\n");
    let maxW = 0;
    for (const l of lines) maxW = Math.max(maxW, textWidth(l, font));
    const w = Math.min(Math.max(activeRect.w, maxW + 12 * z), Math.max(view.w - activeRect.x - 4, activeRect.w));
    const lineH = fs * 1.2;
    const h = Math.max(activeRect.h, lines.length * lineH + 6);
    return {
      position: "absolute",
      left: activeRect.x - 1,
      top: activeRect.y - 1,
      width: w + 1,
      height: h + 1,
      ...fp,
      lineHeight: `${lineH}px`,
      padding: `${Math.max(1, (activeRect.h - lineH) / 2 - 1)}px ${3 * z}px 0 ${3 * z}px`,
      border: `2px solid ${ACCENT}`,
      boxSizing: "border-box",
      background: "#FFFFFF",
      color: (st && st.c) || "#000000",
      resize: "none",
      outline: "none",
      overflow: "hidden",
      zIndex: 6,
      whiteSpace: "pre",
      boxShadow: "0 1px 4px rgba(0,0,0,0.15)",
    };
  }, [ed, sel, cellMap, styleOf, z, activeRect, view.w]);

  // ------------------------------------------------------------ render

  const barValue = ed
    ? ed.text
    : ui && ui.formulaBar !== undefined && ui.selection && sel.active[0] === ui.selection.active[0] && sel.active[1] === ui.selection.active[1]
    ? ui.formulaBar
    : (() => {
        const e = cellMap.get(key(sel.active[0], sel.active[1])) || {};
        return e.in !== undefined ? e.in : e.t || "";
      })();
  const nameBoxValue = ui && ui.nameBox && ui.selection && sel.active[0] === ui.selection.active[0] && sel.active[1] === ui.selection.active[1] && JSON.stringify(sel.ranges) === JSON.stringify(ui.selection.ranges) ? ui.nameBox : (() => {
    const g0 = sel.ranges[sel.ranges.length - 1];
    if (dragRef.current && dragRef.current.kind === "cells" && g0) return `${g0[2] - g0[0] + 1}R x ${g0[3] - g0[1] + 1}C`;
    return a1(sel.active[0], sel.active[1]);
  })();
  const [nameEdit, setNameEdit] = useState(null);

  const dropdownBtn = ui && ui.dropdown && !ed && ui.sheet === sheetIndex && ui.selection && sel.active[0] === ui.selection.active[0] && sel.active[1] === ui.selection.active[1];
  const inputMsg = ui && ui.inputMessage && !ed && ui.selection && sel.active[0] === ui.selection.active[0] && sel.active[1] === ui.selection.active[1] ? ui.inputMessage : null;
  const hovered = menu ? null : hover; // a context menu hides what the pointer shows
  const visibleNotes = ((data && data.notes) || []).filter((n) => n.visible || (hovered && hovered[0] === n.r && hovered[1] === n.c));
  const hoverLink = hovered ? linkMap.get(key(hovered[0], hovered[1])) : null;

  return (
    <div className="rg-sheetview" style={{ display: "flex", flexDirection: "column", width: "100%", height: "100%", minHeight: 0, fontFamily: "'Segoe UI', system-ui, -apple-system, Arial, sans-serif", color: "#1F2937", colorScheme: "light" }}>
      {/* Name box + formula bar */}
      <div className="rg-fbar" style={{ display: "flex", alignItems: "stretch", gap: 0, borderBottom: "1px solid #D5D9DE", background: "#FFFFFF", minHeight: 28 }}>
        <div style={{ position: "relative", width: 110, borderRight: "1px solid #D5D9DE", display: "flex" }}>
          <input
            id="rg-namebox"
            value={nameEdit !== null ? nameEdit : nameBoxValue}
            onFocus={(e) => {
              setNameEdit(nameBoxValue);
              // Only while it still has the focus: select() would take it back from the grid after
              // a quick Enter, and what is typed next would land in the name box.
              const el = e.target;
              setTimeout(() => { if (document.activeElement === el) el.select(); }, 0);
            }}
            onChange={(e) => setNameEdit(e.target.value)}
            onBlur={() => setNameEdit(null)}
            onKeyDown={(e) => {
              if (e.key === "Enter") {
                e.preventDefault();
                const t = (nameEdit || "").trim();
                setNameEdit(null);
                if (t) send("namebox", { text: t });
                kbdRef.current && kbdRef.current.focus();
              } else if (e.key === "Escape") {
                setNameEdit(null);
                kbdRef.current && kbdRef.current.focus();
              }
            }}
            list="rg-names"
            spellCheck={false}
            style={{ flex: 1, minWidth: 0, border: 0, outline: "none", padding: "0 6px", font: "13px 'Segoe UI', system-ui, sans-serif", background: "transparent", color: "#1F2937" }}
          />
          <datalist id="rg-names">{(names || []).map((n) => <option key={n} value={n} />)}</datalist>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 2, padding: "0 6px", borderRight: "1px solid #D5D9DE", color: ed ? "#3C4043" : "#B0B6BC" }}>
          <button title="Cancel" disabled={!ed} onMouseDown={(e) => { e.preventDefault(); cancelEdit(); }} style={btnStyle}>✕</button>
          <button title="Enter" disabled={!ed} onMouseDown={(e) => { e.preventDefault(); commit("none"); }} style={btnStyle}>✓</button>
          <button title="Insert Function" onMouseDown={(e) => { e.preventDefault(); dialog("insertFunction"); }} style={{ ...btnStyle, color: "#3C4043", fontStyle: "italic", fontFamily: "Georgia, serif" }}>fx</button>
        </div>
        <textarea
          id="rg-formula-bar"
          ref={barRef}
          value={barValue}
          rows={Math.min(Math.max(1, (barValue || "").split("\n").length), 6)}
          onFocus={() => {
            if (!edRef.current) {
              if (ui && ui.protected && ui.ribbon && ui.ribbon.locked) return;
              const [r, c] = selRef.current.active;
              setEd({ r, c, text: barValue || "", mode: "edit", bar: true });
            } else if (!edRef.current.bar) setEd({ ...edRef.current, bar: true, mode: "edit" });
          }}
          onChange={(e) => {
            const v = e.target.value;
            if (!edRef.current) {
              const [r, c] = selRef.current.active;
              setEd({ r, c, text: v, mode: "edit", bar: true });
            } else setEd({ ...edRef.current, text: v, bar: true });
            updateAc(v, e.target.selectionStart);
          }}
          onKeyDown={onKeyDown}
          spellCheck={false}
          style={{ flex: 1, minWidth: 0, border: 0, outline: "none", resize: "none", padding: "5px 8px", font: "13px Consolas, 'Cascadia Mono', Menlo, monospace", lineHeight: "18px", background: "#FFFFFF", color: "#111827" }}
        />
      </div>
      {/* Grid */}
      <div style={{ position: "relative", flex: 1, minHeight: 0 }}>
        <div ref={scrollerRef} className="rg-scroller" onScroll={onScroll} style={{ position: "absolute", inset: 0, overflow: "scroll", background: "#FFFFFF" }}>
          <div style={{ width: totalW, height: totalH, position: "relative" }}>
            <div
              ref={viewRef}
              className="rg-view"
              onMouseDown={onMouseDown}
              onContextMenu={onContextMenu}
              onMouseMove={(ev) => {
                if (dragRef.current) return;
                const [x, y] = localPoint(ev);
                const hb = headerBoundary(x, y);
                const el = viewRef.current;
                if (!el) return;
                let cursor = "cell";
                if (hb) cursor = hb.axis === "cols" ? "col-resize" : "row-resize";
                else if (y < HH || x < RW) cursor = y < HH && x < RW ? "default" : y < HH ? "s-resize" : "e-resize";
                else if (selRect) {
                  const hx = selRect.x + selRect.w, hy = selRect.y + selRect.h;
                  if (Math.abs(x - hx) <= 5 && Math.abs(y - hy) <= 5) cursor = "crosshair";
                  else if (((Math.abs(x - selRect.x) <= 3 || Math.abs(x - hx) <= 3) && y >= selRect.y && y <= hy) || ((Math.abs(y - selRect.y) <= 3 || Math.abs(y - hy) <= 3) && x >= selRect.x && x <= hx)) cursor = "move";
                }
                if (el.style.cursor !== cursor) el.style.cursor = cursor;
                // A cell's note, comment or link shows while the pointer is on the cell, as in Excel.
                let cell = null;
                if (!hb && x >= RW && y >= HH && !(ev.target.closest && ev.target.closest("[data-object]"))) {
                  let [r, c] = hit(x, y);
                  const m = mergeAt(r, c);
                  if (m) [r, c] = [m[0], m[1]];
                  if (noteMap.has(key(r, c)) || linkMap.has(key(r, c))) cell = [r, c];
                }
                setHover((h) => (h === cell || (h && cell && h[0] === cell[0] && h[1] === cell[1]) ? h : cell));
              }}
              onMouseLeave={() => setHover(null)}
              style={{ position: "sticky", top: 0, left: 0, width: view.w, height: view.h, overflow: "hidden", userSelect: "none" }}
            >
              <canvas ref={canvasRef} style={{ position: "absolute", left: 0, top: 0 }} />
              {panes.map((p) =>
                p.w > 0 && p.h > 0 ? (
                  <div key={p.key} className={`rg-pane rg-pane-${p.key}`} style={{ position: "absolute", left: p.x, top: p.y, width: p.w, height: p.h, overflow: "hidden", pointerEvents: "none" }}>
                    {cellsLayer(p)}
                    {peerLayer(p)}
                    {selectionLayer(p)}
                    <div style={{ position: "absolute", inset: 0, pointerEvents: "none" }}>{decorLayer(p)}</div>
                  </div>
                ) : null
              )}
              {headerLayer()}
              <div style={{ position: "absolute", left: RW, top: HH, right: 0, bottom: 0, overflow: "hidden", pointerEvents: "none" }}>
                <div style={{ position: "absolute", left: -RW, top: -HH, width: view.w, height: view.h, pointerEvents: "none" }}>
                  <div style={{ pointerEvents: "auto" }}>{objectLayer()}</div>
                </div>
              </div>
              {visibleNotes.map((n) => {
                const rc = rectOf([n.r, n.c, n.r, n.c]);
                return (
                  <div key={`note${n.r}:${n.c}`} className="rg-note" style={{ position: "absolute", left: rc.x + rc.w + 8, top: rc.y, maxWidth: 220, minWidth: 120, background: n.threaded ? "#FFFFFF" : "#FFFFE1", border: "1px solid #A0A0A0", boxShadow: "1px 1px 4px rgba(0,0,0,0.2)", padding: "4px 6px", font: "12px 'Segoe UI', system-ui, sans-serif", whiteSpace: "pre-wrap", zIndex: 8, pointerEvents: "none" }}>
                    <div style={{ fontWeight: 600, marginBottom: 2 }}>{n.author}{n.resolved ? " (resolved)" : ""}:</div>
                    <div>{n.text}</div>
                    {(n.replies || []).map((rp, i) => (
                      <div key={i} style={{ marginTop: 4, borderTop: "1px solid #E0E0E0", paddingTop: 3 }}><b>{rp[0]}:</b> {rp[1]}</div>
                    ))}
                  </div>
                );
              })}
              {hoverLink && (() => {
                const rc = rectOf(mergeAt(hoverLink.r, hoverLink.c) || [hoverLink.r, hoverLink.c, hoverLink.r, hoverLink.c]);
                return (
                  <div className="rg-link-tip" style={{ position: "absolute", left: rc.x, top: rc.y + rc.h + 4, maxWidth: 320, background: "#FFFFFF", border: "1px solid #A0A0A0", boxShadow: "1px 1px 4px rgba(0,0,0,0.2)", padding: "4px 6px", font: "12px 'Segoe UI', system-ui, sans-serif", color: "#1F2937", whiteSpace: "pre-wrap", overflowWrap: "anywhere", zIndex: 8, pointerEvents: "none" }}>
                    <div>{hoverLink.tooltip || hoverLink.target}</div>
                    <div style={{ color: "#6B7280", marginTop: 2 }}>{IS_MAC ? "⌘+click" : "Ctrl+click"} to follow the link.</div>
                  </div>
                );
              })()}
              {inputMsg && !(dropdownBtn && listOpen) && (
                // Hidden while the cell's list is open, which takes the same place.
                <div className="rg-input-msg" style={{ position: "absolute", left: activeRect.x + 12, top: activeRect.y + activeRect.h + 6, maxWidth: 240, background: "#FFFFE1", border: "1px solid #A0A0A0", padding: "4px 6px", font: "12px 'Segoe UI', system-ui, sans-serif", zIndex: 7, pointerEvents: "none" }}>
                  {inputMsg.title && <div style={{ fontWeight: 600 }}>{inputMsg.title}</div>}
                  <div style={{ whiteSpace: "pre-wrap" }}>{inputMsg.message}</div>
                </div>
              )}
              {dropdownBtn && (
                <div
                  className="rg-list-btn"
                  onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); setListOpen((v) => !v); }}
                  style={{ position: "absolute", left: activeRect.x + activeRect.w + 1, top: activeRect.y + activeRect.h - 17, width: 16, height: 16, background: "#F3F4F6", border: "1px solid #A6ADB4", boxSizing: "border-box", display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", zIndex: 7 }}
                >
                  <svg width={8} height={8} viewBox="0 0 8 8"><path d="M1 2 L7 2 L4 6 Z" fill="#404040" /></svg>
                </div>
              )}
              {dropdownBtn && listOpen && (
                <div className="rg-list" style={{ position: "absolute", left: activeRect.x, top: activeRect.y + activeRect.h + 1, minWidth: Math.max(activeRect.w + 16, 120), maxHeight: 200, overflowY: "auto", background: "#FFFFFF", border: "1px solid #A6ADB4", boxShadow: "0 2px 8px rgba(0,0,0,0.18)", zIndex: 9, font: "13px 'Segoe UI', system-ui, sans-serif" }}>
                  {ui.dropdown.map((item, i) => (
                    <div
                      key={i}
                      className="rg-list-item"
                      onMouseDown={(e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        setListOpen(false);
                        const [r, c] = selRef.current.active;
                        const seq = send("commit", { r, c, text: item, move: "none", mode: "normal" });
                        setPending((p) => ({ ...p, [key(r, c)]: { text: item, seq } }));
                        kbdRef.current && kbdRef.current.focus();
                      }}
                      onMouseEnter={(e) => (e.currentTarget.style.background = HEAD_SEL)}
                      onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
                      style={{ padding: "3px 8px", cursor: "pointer", whiteSpace: "nowrap" }}
                    >
                      {item}
                    </div>
                  ))}
                </div>
              )}
              <textarea
                ref={kbdRef}
                id="rg-editor"
                className={ed && !ed.bar ? "rg-editor rg-editing" : "rg-editor"}
                value={ed && !ed.bar ? ed.text : ""}
                onChange={onKbdInput}
                onKeyDown={onKeyDown}
                onCompositionStart={() => {
                  composingRef.current = true;
                  if (!edRef.current) {
                    const [r, c] = selRef.current.active;
                    setEd({ r, c, text: "", mode: "enter", bar: false });
                  }
                }}
                onCompositionEnd={(e) => {
                  composingRef.current = false;
                  const editing = edRef.current;
                  if (editing) setEd({ ...editing, text: e.target.value });
                }}
                onCopy={(e) => onCopyEvt(e, false)}
                onCut={(e) => onCopyEvt(e, true)}
                onPaste={onPasteEvt}
                onMouseDown={(e) => e.stopPropagation()}
                onKeyUp={(e) => { if (edRef.current && edRef.current.text[0] === "=") updateAc(edRef.current.text, e.target.selectionStart); }}
                spellCheck={false}
                autoComplete="off"
                autoCorrect="off"
                autoCapitalize="off"
                style={editorStyle}
              />
              {ac && ed && (
                <div className="rg-ac" style={{ position: "absolute", left: (ed.bar ? RW + 4 : activeRect.x), top: ed.bar ? 2 : activeRect.y + Math.max(activeRect.h, 20) + 2, background: "#FFFFFF", border: "1px solid #A6ADB4", boxShadow: "0 2px 8px rgba(0,0,0,0.18)", zIndex: 10, minWidth: 220, font: "12px 'Segoe UI', system-ui, sans-serif" }}>
                  {ac.items.map((f, i) => (
                    <div key={f.name} onMouseDown={(e) => { e.preventDefault(); acceptAc(f); }} style={{ padding: "3px 8px", background: i === ac.index ? HEAD_SEL : "transparent", cursor: "pointer" }}>
                      <b>{f.name}</b>
                      {i === ac.index && <div style={{ color: "#5F6B76", fontSize: 11, maxWidth: 320, whiteSpace: "normal" }}>{f.description}</div>}
                    </div>
                  ))}
                </div>
              )}
              {!ac && argHint && ed && (
                <div className="rg-arghint" style={{ position: "absolute", left: ed.bar ? RW + 4 : activeRect.x, top: ed.bar ? 2 : activeRect.y + Math.max(activeRect.h, 20) + 2, background: "#FFFFE1", border: "1px solid #A0A0A0", padding: "2px 6px", zIndex: 10, font: "12px 'Segoe UI', system-ui, sans-serif", pointerEvents: "none", whiteSpace: "nowrap" }}>
                  {renderSignature(argHint.fn.signature, argHint.arg)}
                </div>
              )}
              {menu && (
                <div className="rg-menu" onMouseDown={(e) => e.stopPropagation()} style={{ position: "absolute", left: Math.min(menu.x, view.w - 230), top: Math.min(menu.y, Math.max(0, view.h - menuItems.length * 26 - 10)), width: 220, background: "#FFFFFF", border: "1px solid #C4C9CE", boxShadow: "0 4px 14px rgba(0,0,0,0.2)", padding: "4px 0", zIndex: 20, font: "13px 'Segoe UI', system-ui, sans-serif" }}>
                  {menuItems.map((it, i) =>
                    it.divider ? (
                      <div key={i} style={{ height: 1, background: "#E3E6E9", margin: "4px 0" }} />
                    ) : (
                      <div
                        key={i}
                        className="rg-menu-item"
                        onMouseDown={(e) => { e.preventDefault(); runMenu(it); }}
                        onMouseEnter={(e) => (e.currentTarget.style.background = HEAD_SEL)}
                        onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
                        style={{ padding: "4px 14px", display: "flex", justifyContent: "space-between", cursor: "default" }}
                      >
                        <span>{it.label}</span>
                        {it.sc && <span style={{ color: "#8A9199", fontSize: 12 }}>{it.sc}</span>}
                      </div>
                    )
                  )}
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

const btnStyle = { border: 0, background: "transparent", width: 22, height: 22, padding: 0, cursor: "pointer", font: "14px system-ui", color: "inherit" };

function range(a, b) {
  const out = [];
  for (let i = a; i <= b && out.length < 16384; i++) out.push(i);
  return out;
}

function renderSignature(sig, argIndex) {
  const open = sig.indexOf("(");
  if (open < 0) return sig;
  const name = sig.slice(0, open);
  const inner = sig.slice(open + 1, sig.lastIndexOf(")"));
  const parts = [];
  let depth = 0, cur = "";
  for (const ch of inner) {
    if (ch === "[" ) depth++;
    if (ch === "]") depth--;
    if (ch === "," && depth <= 0) { parts.push(cur.trim()); cur = ""; continue; }
    cur += ch;
  }
  if (cur.trim()) parts.push(cur.trim());
  const idx = Math.min(argIndex, parts.length - 1);
  return (
    <span>
      {name}(
      {parts.map((p, i) => (
        <span key={i}>
          {i > 0 ? ", " : ""}
          {i === idx ? <b>{p}</b> : p}
        </span>
      ))}
      )
    </span>
  );
}
