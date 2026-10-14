// ReflexGrid's form designer: a UserForm's design (the server's snapshot of it) to edit as in the VBA
// editor. Click to select (the form itself when nothing is under the pointer), drag to move, drag a
// handle to resize, Delete to remove, the arrow keys to nudge (Shift: resize); with a toolbox tool
// chosen, click (or drag a rectangle) to place a new control — in a Frame or on a MultiPage's page
// when it is dropped there. Positions snap to a 6-point grid (hold Alt not to). Double-click opens
// the control's event procedure. Every change goes to the server, which changes the form in place.
//
// <FormDesigner form={snapshot} selected="TextBox1" tool="CommandButton" pictures={{}} onEvent={...} />
import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Controls, PX, TITLE_H, fontCss, px } from "./userform.jsx";

const GRID = 6;
const HANDLE = 7;
const TAB_H = 15;
// Default sizes of new controls, in points (as the VBA editor's toolbox draws them).
const NEW_SIZE = {
  Label: [72, 18], TextBox: [72, 18], ComboBox: [72, 18], ListBox: [72, 72], CheckBox: [108, 18],
  OptionButton: [108, 18], ToggleButton: [24, 24], Frame: [144, 72], CommandButton: [72, 24],
  TabStrip: [144, 108], MultiPage: [144, 108], ScrollBar: [12.75, 63.75], SpinButton: [12.75, 25.5], Image: [72, 72],
};

function snap(v, on) {
  return on ? Math.round(v / GRID) * GRID : Math.round(v * 4) / 4;
}

// Every control with its rectangle in the form's inside (points), deepest last; containers say
// where their own controls start.
function layout(controls) {
  const out = [];
  const visit = (list, ox, oy, parent) => {
    for (const c of list || []) {
      const r = { c, x: ox + (c.left || 0), y: oy + (c.top || 0), w: c.width || 0, h: c.height || 0, parent };
      out.push(r);
      if (c.kind === "Frame") visit(c.children, r.x, r.y, c.name);
      else if (c.kind === "MultiPage") {
        const pages = c.children || [];
        const page = pages[Math.max(0, Math.min(pages.length - 1, c.value || 0))];
        if (page) visit(page.children, r.x, r.y + TAB_H, c.name);
      }
    }
  };
  visit(controls, 0, 0, "");
  return out;
}

function inside(r, x, y) {
  return x >= r.x && x <= r.x + r.w && y >= r.y && y <= r.y + r.h;
}

export function FormDesigner(props) {
  const { form, selected, tool, pictures, onEvent } = props;
  const client = useRef(null);
  // What the pointer is doing; a ref too, as the button may come up before React redraws.
  const [drag, setDragState] = useState(null);
  const dragRef = useRef(null);
  const setDrag = (d) => {
    dragRef.current = d;
    setDragState(d);
  };
  const send = useCallback((e) => onEvent && onEvent(e), [onEvent]);
  const rects = useMemo(() => (form ? layout(form.controls) : []), [form]);
  const byName = useMemo(() => Object.fromEntries(rects.map((r) => [r.c.name, r])), [rects]);

  // A move or resize sent to the server shows where it goes until the server's answer comes.
  const [pending, setPending] = useState(null);
  useEffect(() => {
    setPending(null);
  }, [form]);

  if (!form) return null;
  const w = px(form.insideWidth);
  const h = px(form.insideHeight);

  const ctx = {
    send: () => {},
    edit: () => {},
    focused: () => {},
    focusNext: () => {},
    hasDefault: false,
    hasCancel: false,
    inert: true,
    off: false,
    design: false,
    box: (c) => ({ position: "absolute", left: px(c.left), top: px(c.top), width: px(c.width), height: px(c.height) }),
    pic: (p) => (!p ? "" : (pictures && pictures[p]) || (p.startsWith("data:") ? p : "")),
  };

  const point = (e) => {
    const b = client.current.getBoundingClientRect();
    return { x: (e.clientX - b.left) / PX, y: (e.clientY - b.top) / PX };
  };

  const hit = (x, y) => {
    for (let i = rects.length - 1; i >= 0; i--) if (inside(rects[i], x, y)) return rects[i];
    return null;
  };

  const containerAt = (x, y) => {
    for (let i = rects.length - 1; i >= 0; i--) {
      const r = rects[i];
      if ((r.c.kind === "Frame" || r.c.kind === "MultiPage") && inside(r, x, y)) return r;
    }
    return null;
  };

  const onPointerDown = (e) => {
    if (e.button !== 0) return;
    e.preventDefault();
    client.current.focus({ preventScroll: true });
    const p = point(e);
    if (tool) {
      setDrag({ mode: "add", x0: p.x, y0: p.y, x: p.x, y: p.y, snap: !e.altKey });
      e.currentTarget.setPointerCapture(e.pointerId);
      return;
    }
    // A MultiPage's tab: its page comes up.
    const tab = e.target.closest && e.target.closest("[data-tab]");
    const r = hit(p.x, p.y);
    if (tab && r && r.c.kind === "MultiPage") {
      const i = parseInt(tab.getAttribute("data-tab"), 10);
      if (i !== r.c.value) send({ event: "page", control: r.c.name, value: i });
      if (selected !== r.c.name) send({ event: "select", control: r.c.name });
      return;
    }
    if (!r) {
      if (selected) send({ event: "select", control: "" });
      return;
    }
    if (selected !== r.c.name) send({ event: "select", control: r.c.name });
    setDrag({ mode: "move", name: r.c.name, x0: p.x, y0: p.y, dx: 0, dy: 0, snap: !e.altKey });
    e.currentTarget.setPointerCapture(e.pointerId);
  };

  const onHandleDown = (e, name, edge) => {
    e.preventDefault();
    e.stopPropagation();
    const p = point(e);
    setDrag({ mode: "size", name, edge, x0: p.x, y0: p.y, dx: 0, dy: 0, snap: !e.altKey });
    client.current.setPointerCapture(e.pointerId);
  };

  const onPointerMove = (e) => {
    const d = dragRef.current;
    if (!d) return;
    const p = point(e);
    if (d.mode === "add") setDrag({ ...d, x: p.x, y: p.y, snap: !e.altKey });
    else setDrag({ ...d, dx: p.x - d.x0, dy: p.y - d.y0, snap: !e.altKey });
  };

  // The rectangle a move or resize makes of a control (or the form, for "").
  const moved = (d) => {
    const r = d.name ? byName[d.name] : { c: { left: 0, top: 0, width: form.insideWidth, height: form.insideHeight }, x: 0, y: 0 };
    if (!r) return null;
    let left = r.c.left, top = r.c.top, width = r.c.width, height = r.c.height;
    if (d.mode === "move") {
      left = snap(left + d.dx, d.snap);
      top = snap(top + d.dy, d.snap);
    } else {
      const e = d.edge;
      if (e.includes("e")) width = Math.max(3, snap(left + width + d.dx, d.snap) - left);
      if (e.includes("s")) height = Math.max(3, snap(top + height + d.dy, d.snap) - top);
      if (e.includes("w")) {
        const nl = Math.min(left + width - 3, snap(left + d.dx, d.snap));
        width = width + (left - nl);
        left = nl;
      }
      if (e.includes("n")) {
        const nt = Math.min(top + height - 3, snap(top + d.dy, d.snap));
        height = height + (top - nt);
        top = nt;
      }
    }
    return { left: Math.max(0, left), top: Math.max(0, top), width, height };
  };

  const onPointerUp = (e) => {
    const d = dragRef.current;
    if (!d) return;
    setDrag(null);
    if (d.mode === "add") {
      const [dw, dh] = NEW_SIZE[tool] || [72, 18];
      const x0 = Math.min(d.x0, d.x), y0 = Math.min(d.y0, d.y);
      const big = Math.abs(d.x - d.x0) > 4 || Math.abs(d.y - d.y0) > 4;
      const host = containerAt(d.x0, d.y0);
      const ox = host ? host.x : 0;
      const oy = host ? host.y + (host.c.kind === "MultiPage" ? TAB_H : 0) : 0;
      send({
        event: "add",
        kind: tool,
        parent: host ? host.c.name : "",
        left: Math.max(0, snap(x0 - ox, d.snap)),
        top: Math.max(0, snap(y0 - oy, d.snap)),
        width: big ? Math.max(6, snap(Math.abs(d.x - d.x0), d.snap)) : dw,
        height: big ? Math.max(6, snap(Math.abs(d.y - d.y0), d.snap)) : dh,
      });
      return;
    }
    if (Math.abs(d.dx) < 0.5 && Math.abs(d.dy) < 0.5) return;
    const g = moved(d);
    if (g) {
      setPending({ name: d.name || "", g });
      send({ event: "geometry", control: d.name || "", ...g });
    }
  };

  const onKeyDown = (e) => {
    if (!selected) return;
    if (e.key === "Delete" || e.key === "Backspace") {
      e.preventDefault();
      send({ event: "delete", control: selected });
      return;
    }
    const step = e.altKey ? 0.75 : GRID;
    const d = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] }[e.key];
    if (!d) return;
    e.preventDefault();
    const r = byName[selected];
    if (!r) return;
    const c = r.c;
    if (e.shiftKey) send({ event: "geometry", control: selected, width: Math.max(3, c.width + d[0]), height: Math.max(3, c.height + d[1]) });
    else send({ event: "geometry", control: selected, left: Math.max(0, c.left + d[0]), top: Math.max(0, c.top + d[1]) });
  };

  const onDoubleClick = (e) => {
    const p = point(e);
    const r = hit(p.x, p.y);
    send({ event: "code", control: r ? r.c.name : "" });
  };

  // What is drawn as selected (following the pointer while it drags).
  let sel = null;
  if (selected && byName[selected]) {
    const r = byName[selected];
    const g = drag && drag.mode !== "add" && drag.name === selected ? moved(drag) : pending && pending.name === selected ? pending.g : null;
    const pr = r.parent ? byName[r.parent] : null;
    const ox = pr ? pr.x : 0;
    const oy = pr ? pr.y + (pr.c.kind === "MultiPage" ? TAB_H : 0) : 0;
    sel = g ? { x: ox + g.left, y: oy + g.top, w: g.width, h: g.height } : { x: r.x, y: r.y, w: r.w, h: r.h };
  }
  const formDrag = drag && drag.mode === "size" && !drag.name ? moved(drag) : pending && !pending.name ? pending.g : null;
  const shownW = formDrag ? px(formDrag.width) : w;
  const shownH = formDrag ? px(formDrag.height) : h;

  const handles = (box, name) =>
    ["nw", "n", "ne", "e", "se", "s", "sw", "w"].map((edge) => {
      const cx = edge.includes("w") ? 0 : edge.includes("e") ? box.w : box.w / 2;
      const cy = edge.includes("n") ? 0 : edge.includes("s") ? box.h : box.h / 2;
      const cursor = { n: "ns-resize", s: "ns-resize", e: "ew-resize", w: "ew-resize", ne: "nesw-resize", sw: "nesw-resize", nw: "nwse-resize", se: "nwse-resize" }[edge];
      return (
        <div
          key={edge}
          data-handle={edge}
          onPointerDown={(e) => onHandleDown(e, name, edge)}
          style={{ position: "absolute", left: px(box.x + cx) - HANDLE / 2, top: px(box.y + cy) - HANDLE / 2, width: HANDLE, height: HANDLE, background: "#FFFFFF", border: "1px solid #0B5CAD", cursor, boxSizing: "border-box", zIndex: 3 }}
        />
      );
    });

  let addBox = null;
  if (drag && drag.mode === "add") {
    addBox = { x: Math.min(drag.x0, drag.x), y: Math.min(drag.y0, drag.y), w: Math.abs(drag.x - drag.x0), h: Math.abs(drag.y - drag.y0) };
  }

  return (
    <div className="rg-form-designer" data-form={form.name} style={{ display: "inline-block", padding: 16, colorScheme: "light", color: "#000000", userSelect: "none" }}>
      <div style={{ width: shownW + 2, border: selected ? "1px solid #9A9A9A" : "1px solid #0B5CAD", background: "#FFFFFF", boxShadow: "0 2px 10px rgba(0,0,0,.15)" }}>
        <div
          onPointerDown={(e) => {
            e.preventDefault();
            if (selected) send({ event: "select", control: "" });
          }}
          onDoubleClick={() => send({ event: "code", control: "" })}
          style={{ height: TITLE_H, display: "flex", alignItems: "center", padding: "0 8px", fontFamily: "'Segoe UI', Tahoma, Arial, sans-serif", fontSize: 12, color: "#000000", background: "#FFFFFF", cursor: "default" }}
        >
          <span style={{ flex: 1, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{form.caption}</span>
          <span style={{ color: "#555", fontSize: 13 }}>✕</span>
        </div>
        <div
          ref={client}
          tabIndex={0}
          className="rg-form-designer-client"
          data-tool={tool || ""}
          onPointerDown={onPointerDown}
          onPointerMove={onPointerMove}
          onPointerUp={onPointerUp}
          onKeyDown={onKeyDown}
          onDoubleClick={onDoubleClick}
          style={{
            position: "relative",
            width: shownW,
            height: shownH,
            background: form.backColor,
            backgroundImage: `radial-gradient(circle, rgba(0,0,0,.28) 0.8px, transparent 0.9px)`,
            backgroundSize: `${px(GRID)}px ${px(GRID)}px`,
            backgroundPosition: `-${px(GRID) / 2}px -${px(GRID) / 2}px`,
            color: form.foreColor,
            overflow: "visible",
            outline: "none",
            cursor: tool ? "crosshair" : "default",
            ...fontCss(form.font),
          }}
        >
          <div style={{ position: "absolute", inset: 0, overflow: "hidden" }}>
            <Controls list={form.controls || []} ctx={ctx} values={{}} />
          </div>
          {sel ? <div style={{ position: "absolute", left: px(sel.x) - 1, top: px(sel.y) - 1, width: px(sel.w) + 2, height: px(sel.h) + 2, border: "1px dashed #0B5CAD", pointerEvents: "none", zIndex: 2 }} /> : null}
          {sel ? handles(sel, selected) : null}
          {!selected ? (
            <div
              data-handle="form-se"
              onPointerDown={(e) => onHandleDown(e, "", "se")}
              style={{ position: "absolute", left: shownW - HANDLE / 2, top: shownH - HANDLE / 2, width: HANDLE, height: HANDLE, background: "#FFFFFF", border: "1px solid #0B5CAD", cursor: "nwse-resize", zIndex: 3 }}
            />
          ) : null}
          {addBox ? <div style={{ position: "absolute", left: px(addBox.x), top: px(addBox.y), width: px(addBox.w), height: px(addBox.h), border: "1px dashed #333", pointerEvents: "none" }} /> : null}
        </div>
      </div>
    </div>
  );
}
