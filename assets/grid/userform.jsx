// ReflexGrid's UserForms: the forms a macro shows, drawn from the server's snapshots
// (gridcore.vba.msforms.UserForm.snapshot) and sending back what the person does.
//
// Sizes are in points (1 pt = 4/3 px). What the person types stays here until the next event,
// which carries it ("values"), unless the form's code handles the control's Change or KeyPress:
// then each change goes at once, as VBA runs those events key by key. Clicks always go at once.
// A modal form (the one the macro waits on) has a backdrop that keeps the page out of reach.
//
// <UserForms forms={[...]} waiting={id} blocked={bool} design={false} onEvent={(e) => ...} />
// With design={true} the form is a still picture of its design (the VBA editor's form view);
// double-clicking a control asks for its event procedure.
import React, { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";

export const PX = 4 / 3;
const FONT_FALLBACK = ", Tahoma, Verdana, 'Segoe UI', Arial, sans-serif";
export const TITLE_H = 26;

export function px(pt) {
  return Math.round((pt || 0) * PX * 100) / 100;
}

export function fontCss(f) {
  if (!f) return {};
  const deco = [f.underline ? "underline" : "", f.strike ? "line-through" : ""].filter(Boolean).join(" ");
  return {
    fontFamily: `'${(f.name || "Tahoma").replace(/'/g, "")}'${FONT_FALLBACK}`,
    fontSize: `${px(f.size || 8.25)}px`,
    fontWeight: f.bold ? 700 : 400,
    fontStyle: f.italic ? "italic" : "normal",
    textDecoration: deco || "none",
  };
}

const ALIGN = { 1: "left", 2: "center", 3: "right" };

// SpecialEffect / BorderStyle as CSS borders.
function borderCss(c, fallbackSunken) {
  const eff = c.specialEffect;
  if (c.borderStyle === 1) return { border: `1px solid ${c.borderColor || "#646464"}` };
  if (eff === 2 || (eff === undefined && fallbackSunken)) return { border: "1px solid #7A7A7A", boxShadow: "inset 1px 1px 0 #D4D4D4" };
  if (eff === 1) return { border: "2px outset #E3E3E3" };
  if (eff === 3) return { border: "2px groove #E3E3E3" };
  if (eff === 6) return { border: "2px ridge #E3E3E3" };
  return { border: "none" };
}

function has(c, ev) {
  return Array.isArray(c.events) && c.events.indexOf(ev) >= 0;
}

function colWidths(spec, count, total) {
  // "72 pt;0 pt;1 in" → px widths (blank: share what is left).
  const parts = (spec || "").split(";");
  const out = [];
  let used = 0;
  let blanks = 0;
  for (let i = 0; i < count; i++) {
    const s = (parts[i] || "").trim().toLowerCase();
    let w = null;
    const m = s.match(/^([0-9.]+)\s*(pt|in|cm|mm)?$/);
    if (m) {
      const n = parseFloat(m[1]);
      const unit = m[2] || "pt";
      w = px(unit === "in" ? n * 72 : unit === "cm" ? (n * 72) / 2.54 : unit === "mm" ? (n * 72) / 25.4 : n);
    }
    if (w === null) blanks++;
    else used += w;
    out.push(w);
  }
  const share = blanks ? Math.max(24, (total - used) / blanks) : 0;
  return out.map((w) => (w === null ? share : w));
}

// ------------------------------------------------------------------ controls

function Caption({ c, style }) {
  return (
    <span style={{ whiteSpace: c.wordWrap ? "pre-wrap" : "pre", overflow: "hidden", ...style }}>
      {c.caption}
    </span>
  );
}

function TextBox({ c, value, ctx }) {
  const ref = useRef(null);
  const keyed = has(c, "keypress");
  const multi = !!c.multiLine;
  const onKeyDown = (e) => {
    if (e.key === "Enter" && !(multi && (c.enterKeyBehavior || e.shiftKey || e.ctrlKey))) {
      if (ctx.hasDefault) {
        e.preventDefault();
        ctx.send({ event: "enter" });
        return;
      }
      if (!multi) {
        e.preventDefault();
        ctx.focusNext(e.currentTarget);
        return;
      }
    }
    if (e.key === "Escape" && ctx.hasCancel) {
      e.preventDefault();
      ctx.send({ event: "escape" });
      return;
    }
    if (keyed && e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      // KeyPress: the form's code sees the character first (and may drop it).
      e.preventDefault();
      const el = e.currentTarget;
      ctx.send({ event: "keypress", control: c.name, key: e.key, selStart: el.selectionStart, selLength: el.selectionEnd - el.selectionStart });
      return;
    }
    if (has(c, "keydown") && e.key.length > 1) ctx.send({ event: "keydown", control: c.name, keyCode: e.keyCode, shift: (e.shiftKey ? 1 : 0) | (e.ctrlKey ? 2 : 0) | (e.altKey ? 4 : 0) });
  };
  useEffect(() => {
    // The code moved the caret (SelStart) or typed for the person (KeyPress).
    const el = ref.current;
    if (el && document.activeElement === el && typeof c.selStart === "number" && keyed) {
      try {
        el.setSelectionRange(c.selStart, c.selStart + (c.selLength || 0));
      } catch (err) {}
    }
  }, [c.text, c.selStart]);
  const style = {
    ...ctx.box(c),
    ...borderCss(c, true),
    background: c.opaque === false ? "transparent" : c.backColor,
    color: c.foreColor,
    ...fontCss(c.font),
    textAlign: ALIGN[c.textAlign] || "left",
    padding: "1px 2px",
    resize: "none",
    outline: "none",
    boxSizing: "border-box",
    whiteSpace: c.wordWrap && multi ? "pre-wrap" : "pre",
  };
  const common = {
    ref,
    value: value == null ? "" : value,
    readOnly: !!c.locked || ctx.design,
    disabled: !c.enabled || ctx.off,
    maxLength: c.maxLength > 0 ? c.maxLength : undefined,
    onChange: (e) => ctx.edit(c, e.target.value, has(c, "change")),
    onKeyDown,
    onFocus: () => ctx.focused(c),
    onDoubleClick: () => has(c, "dblclick") && ctx.send({ event: "dblclick", control: c.name }),
    title: c.tip || undefined,
    "data-control": c.name,
    tabIndex: c.tabStop ? 0 : -1,
    spellCheck: false,
    style,
  };
  if (multi) return <textarea {...common} />;
  return <input type={c.passwordChar ? "password" : "text"} {...common} />;
}

function ListRows({ c, rows, selected, onPick, height }) {
  const n = Math.max(1, c.columnCount || 1);
  const widths = colWidths(c.columnWidths, n, px(c.width) - 4);
  return (
    <div style={{ overflowY: "auto", overflowX: "hidden", height, background: c.backColor, color: c.foreColor }}>
      {rows.map((row, i) => (
        <div
          key={i}
          data-row={i}
          onMouseDown={(e) => {
            e.preventDefault();
            onPick(i, e);
          }}
          style={{
            display: "flex",
            whiteSpace: "nowrap",
            cursor: "default",
            padding: "0 2px",
            background: selected.indexOf(i) >= 0 ? "#0078D7" : "transparent",
            color: selected.indexOf(i) >= 0 ? "#FFFFFF" : c.foreColor,
          }}
        >
          {widths.map((w, j) =>
            w > 0 ? (
              <span key={j} style={{ width: w, minWidth: w, overflow: "hidden", textOverflow: "ellipsis" }}>
                {row[j] == null ? "" : row[j]}
                {" "}
              </span>
            ) : null
          )}
        </div>
      ))}
    </div>
  );
}

function ListBox({ c, value, ctx }) {
  const multi = c.multiSelect || 0;
  const selected = value != null ? value : c.selected || [];
  const pick = (i, e) => {
    if (!c.enabled || c.locked || ctx.inert) return;
    ctx.focused(c);
    let next;
    if (!multi) next = i;
    else if (multi === 1 || e.ctrlKey || e.metaKey) next = selected.indexOf(i) >= 0 ? selected.filter((x) => x !== i) : selected.concat([i]);
    else next = [i];
    ctx.send({ event: "click", control: c.name, value: next });
  };
  const onKeyDown = (e) => {
    const n = (c.list || []).length;
    if (e.key === "Enter" && ctx.hasDefault) {
      e.preventDefault();
      ctx.send({ event: "enter" });
      return;
    }
    if (e.key === "Escape" && ctx.hasCancel) {
      e.preventDefault();
      ctx.send({ event: "escape" });
      return;
    }
    if (multi || !n || !c.enabled || c.locked || ctx.inert) return;
    const cur = typeof c.listIndex === "number" ? c.listIndex : -1;
    let next = null;
    if (e.key === "ArrowDown") next = Math.min(n - 1, cur + 1);
    else if (e.key === "ArrowUp") next = Math.max(0, cur - 1);
    else if (e.key === "Home") next = 0;
    else if (e.key === "End") next = n - 1;
    if (next !== null) {
      e.preventDefault();
      if (next !== cur) ctx.send({ event: "click", control: c.name, value: next });
    }
  };
  return (
    <div
      data-control={c.name}
      tabIndex={c.tabStop ? 0 : -1}
      onFocus={() => ctx.focused(c)}
      onKeyDown={onKeyDown}
      onDoubleClick={() => has(c, "dblclick") && ctx.send({ event: "dblclick", control: c.name })}
      title={c.tip || undefined}
      style={{ ...ctx.box(c), ...borderCss(c, true), boxSizing: "border-box", ...fontCss(c.font), background: c.backColor, opacity: c.enabled ? 1 : 0.6, outline: "none" }}
    >
      <ListRows c={c} rows={c.list || []} selected={Array.isArray(selected) ? selected : [selected]} onPick={pick} height="100%" />
    </div>
  );
}

function ComboBox({ c, value, ctx }) {
  const [open, setOpen] = useState(false);
  const self = useRef(null);
  useEffect(() => {
    if (!open) return undefined;
    const away = (e) => {
      if (self.current && !self.current.contains(e.target)) setOpen(false);
    };
    document.addEventListener("mousedown", away, true);
    return () => document.removeEventListener("mousedown", away, true);
  }, [open]);
  const text = value != null ? value : c.text;
  const listOnly = c.style === 2;
  const pick = (i) => {
    setOpen(false);
    const col = (c.textColumn > 0 ? c.textColumn : 1) - 1;
    const row = (c.list || [])[i] || [];
    ctx.send({ event: "click", control: c.name, value: row[col] == null ? "" : String(row[col]) });
  };
  const rows = c.list || [];
  const box = ctx.box(c);
  return (
    <div ref={self} data-control={c.name} style={{ ...box, boxSizing: "border-box", ...borderCss(c, true), background: c.backColor, display: "flex", ...fontCss(c.font) }} title={c.tip || undefined}>
      {listOnly ? (
        <div
          tabIndex={c.tabStop ? 0 : -1}
          onFocus={() => ctx.focused(c)}
          onMouseDown={(e) => {
            if (!c.enabled || c.locked || ctx.inert) return;
            e.preventDefault();
            ctx.focused(c);
            setOpen(!open);
          }}
          onKeyDown={(e) => {
            if (e.key === "ArrowDown" || e.key === "F4") setOpen(true);
            if (e.key === "Escape") setOpen(false);
          }}
          style={{ flex: 1, padding: "1px 2px", whiteSpace: "pre", overflow: "hidden", color: c.foreColor, cursor: "default", outline: "none" }}
        >
          {text || " "}
        </div>
      ) : (
        <input
          value={text == null ? "" : text}
          readOnly={!!c.locked || ctx.design}
          disabled={!c.enabled || ctx.off}
          maxLength={c.maxLength > 0 ? c.maxLength : undefined}
          onChange={(e) => ctx.edit(c, e.target.value, has(c, "change"))}
          onFocus={() => ctx.focused(c)}
          onKeyDown={(e) => {
            if (e.key === "ArrowDown" && e.altKey) setOpen(true);
            else if (e.key === "Enter" && ctx.hasDefault) {
              e.preventDefault();
              ctx.send({ event: "enter" });
            } else if (e.key === "Escape") {
              if (open) setOpen(false);
              else if (ctx.hasCancel) ctx.send({ event: "escape" });
            }
          }}
          data-control-input={c.name}
          tabIndex={c.tabStop ? 0 : -1}
          spellCheck={false}
          style={{ flex: 1, minWidth: 0, border: "none", outline: "none", background: "transparent", color: c.foreColor, padding: "1px 2px", ...fontCss(c.font) }}
        />
      )}
      <div
        onMouseDown={(e) => {
          e.preventDefault();
          if (!c.enabled || c.locked || ctx.inert) return;
          ctx.focused(c);
          if (!open && has(c, "dropbuttonclick")) ctx.send({ event: "dropbuttonclick", control: c.name });
          setOpen(!open);
        }}
        style={{ width: 15, borderLeft: "1px solid #ACACAC", display: "flex", alignItems: "center", justifyContent: "center", background: "#F0F0F0", cursor: "default", fontSize: 8, color: "#333" }}
      >
        ▼
      </div>
      {open && (
        <div
          style={{ position: "absolute", left: -1, top: "100%", width: "calc(100% + 2px)", maxHeight: px(12) * Math.max(1, Math.min(c.listRows || 8, rows.length || 1)) + 4, background: "#FFFFFF", border: "1px solid #646464", zIndex: 5, boxShadow: "0 2px 6px rgba(0,0,0,.25)", overflow: "hidden" }}
          onMouseLeave={() => {}}
        >
          <ListRows c={c} rows={rows} selected={c.listIndex >= 0 ? [c.listIndex] : []} onPick={pick} height="100%" />
        </div>
      )}
    </div>
  );
}

function Toggle({ c, value, ctx }) {
  const v = value !== undefined ? value : c.value;
  const radio = c.kind === "OptionButton";
  const click = (e) => {
    e.preventDefault();
    if (!c.enabled || c.locked || ctx.inert) return;
    ctx.focused(c);
    ctx.send({ event: "click", control: c.name });
  };
  if (c.kind === "ToggleButton") {
    return (
      <button
        type="button"
        data-control={c.name}
        tabIndex={c.tabStop ? 0 : -1}
        disabled={!c.enabled || ctx.off}
        onMouseDown={(e) => e.preventDefault()}
        onClick={click}
        onFocus={() => ctx.focused(c)}
        title={c.tip || undefined}
        style={{ ...ctx.box(c), ...fontCss(c.font), color: c.foreColor, background: v ? "#CCE4F7" : c.backColor, border: v ? "1px solid #005499" : "1px solid #ADADAD", boxShadow: v ? "inset 1px 1px 2px rgba(0,0,0,.25)" : "none", padding: 0, boxSizing: "border-box" }}
      >
        <Caption c={c} />
      </button>
    );
  }
  const mark = (
    <span
      style={{
        width: px(9.75),
        height: px(9.75),
        minWidth: px(9.75),
        border: "1px solid #333333",
        borderRadius: radio ? "50%" : 0,
        background: "#FFFFFF",
        display: "inline-flex",
        alignItems: "center",
        justifyContent: "center",
        fontSize: px(8),
        lineHeight: 1,
        color: "#000000",
        boxSizing: "border-box",
      }}
    >
      {radio ? v ? <span style={{ width: "50%", height: "50%", borderRadius: "50%", background: "#000" }} /> : null : v === null ? <span style={{ width: "60%", height: "60%", background: "#555" }} /> : v ? "✓" : ""}
    </span>
  );
  const right = c.alignment !== 0;
  return (
    <div
      role={radio ? "radio" : "checkbox"}
      aria-checked={v === null ? "mixed" : !!v}
      data-control={c.name}
      tabIndex={c.tabStop ? 0 : -1}
      onMouseDown={(e) => e.preventDefault()}
      onClick={click}
      onKeyDown={(e) => {
        if (e.key === " ") click(e);
        else if (e.key === "Enter" && ctx.hasDefault) ctx.send({ event: "enter" });
        else if (e.key === "Escape" && ctx.hasCancel) ctx.send({ event: "escape" });
      }}
      onFocus={() => ctx.focused(c)}
      title={c.tip || undefined}
      style={{
        ...ctx.box(c),
        display: "flex",
        flexDirection: right ? "row" : "row-reverse",
        alignItems: "center",
        gap: 4,
        ...fontCss(c.font),
        color: c.enabled ? c.foreColor : "#8C8C8C",
        background: c.opaque === false ? "transparent" : c.backColor,
        cursor: "default",
        outline: "none",
        userSelect: "none",
        boxSizing: "border-box",
        overflow: "hidden",
      }}
    >
      {mark}
      <Caption c={c} style={{ flex: 1, textAlign: right ? "left" : "right" }} />
    </div>
  );
}

function CommandButton({ c, ctx }) {
  return (
    <button
      type="button"
      data-control={c.name}
      tabIndex={c.tabStop ? 0 : -1}
      disabled={!c.enabled || ctx.off}
      onClick={() => {
        ctx.focused(c, true);
        ctx.send({ event: "click", control: c.name });
      }}
      onDoubleClick={() => has(c, "dblclick") && ctx.send({ event: "dblclick", control: c.name })}
      onFocus={() => ctx.focused(c)}
      title={c.tip || undefined}
      style={{
        ...ctx.box(c),
        ...fontCss(c.font),
        color: c.enabled ? c.foreColor : "#8C8C8C",
        background: c.backColor,
        border: c.default ? "2px solid #0078D7" : "1px solid #ADADAD",
        borderRadius: 0,
        padding: 0,
        boxSizing: "border-box",
        cursor: "default",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        overflow: "hidden",
      }}
    >
      {c.picture ? <img src={ctx.pic(c.picture)} alt="" style={{ maxWidth: "100%", maxHeight: "100%" }} /> : null}
      <Caption c={c} />
    </button>
  );
}

function Label({ c, ctx }) {
  return (
    <div
      data-control={c.name}
      onClick={() => has(c, "click") && ctx.send({ event: "click", control: c.name })}
      onDoubleClick={() => has(c, "dblclick") && ctx.send({ event: "dblclick", control: c.name })}
      title={c.tip || undefined}
      style={{
        ...ctx.box(c),
        ...fontCss(c.font),
        ...borderCss(c, false),
        color: c.enabled ? c.foreColor : "#8C8C8C",
        background: c.opaque === false ? "transparent" : c.backColor,
        textAlign: ALIGN[c.textAlign] || "left",
        whiteSpace: c.wordWrap ? "pre-wrap" : "pre",
        overflow: "hidden",
        boxSizing: "border-box",
        cursor: "default",
        userSelect: "none",
      }}
    >
      {c.picture ? <img src={ctx.pic(c.picture)} alt="" style={{ maxWidth: "100%", maxHeight: "100%" }} /> : null}
      {c.caption}
    </div>
  );
}

function ImageCtl({ c, ctx }) {
  const fit = { 0: "none", 1: "fill", 3: "contain" }[c.pictureSizeMode || 0] || "none";
  return (
    <div
      data-control={c.name}
      onClick={() => has(c, "click") && ctx.send({ event: "click", control: c.name })}
      onDoubleClick={() => has(c, "dblclick") && ctx.send({ event: "dblclick", control: c.name })}
      title={c.tip || undefined}
      style={{ ...ctx.box(c), ...borderCss(c, false), background: c.opaque === false ? "transparent" : c.backColor, overflow: "hidden", boxSizing: "border-box" }}
    >
      {c.picture ? <img src={ctx.pic(c.picture)} alt="" style={{ width: "100%", height: "100%", objectFit: fit }} /> : null}
    </div>
  );
}

function ScrollBar({ c, value, ctx }) {
  const v = value != null ? value : c.value;
  const vertical = c.orientation === 0 || (c.orientation !== 1 && c.height > c.width);
  const lo = Math.min(c.min, c.max);
  const hi = Math.max(c.min, c.max);
  return (
    <input
      type="range"
      data-control={c.name}
      min={lo}
      max={hi}
      step={c.smallChange || 1}
      value={v}
      disabled={!c.enabled || ctx.off}
      onFocus={() => ctx.focused(c)}
      onChange={(e) => ctx.edit(c, parseInt(e.target.value, 10), false)}
      onMouseUp={(e) => ctx.send({ event: "scroll", control: c.name, value: parseInt(e.currentTarget.value, 10) })}
      onKeyUp={(e) => ctx.send({ event: "scroll", control: c.name, value: parseInt(e.currentTarget.value, 10) })}
      title={c.tip || undefined}
      style={{
        ...ctx.box(c),
        margin: 0,
        writingMode: vertical ? "vertical-lr" : "horizontal-tb",
        direction: vertical && c.max >= c.min ? "ltr" : undefined,
        accentColor: "#6B6B6B",
      }}
    />
  );
}

function SpinButton({ c, ctx }) {
  const vertical = c.orientation === 0 || (c.orientation !== 1 && c.height >= c.width);
  const btn = (up, label) => (
    <div
      onMouseDown={(e) => {
        e.preventDefault();
        if (!c.enabled || ctx.inert) return;
        ctx.focused(c);
        ctx.send({ event: "spin", control: c.name, up });
      }}
      style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", border: "1px solid #ADADAD", background: "#E1E1E1", fontSize: 8, color: c.enabled ? "#222" : "#999", cursor: "default", userSelect: "none" }}
    >
      {label}
    </div>
  );
  return (
    <div data-control={c.name} title={c.tip || undefined} style={{ ...ctx.box(c), display: "flex", flexDirection: vertical ? "column" : "row-reverse", boxSizing: "border-box" }}>
      {btn(true, vertical ? "▲" : "▶")}
      {btn(false, vertical ? "▼" : "◀")}
    </div>
  );
}

function Tabs({ c, captions, value, ctx }) {
  return (
    <div style={{ display: "flex", height: px(15), borderBottom: "1px solid #ADADAD", ...fontCss(c.font) }}>
      {captions.map((t, i) => (
        <div
          key={i}
          data-tab={i}
          onMouseDown={(e) => {
            e.preventDefault();
            if (!c.enabled || ctx.inert || i === value) return;
            ctx.send({ event: "click", control: c.name, value: i });
          }}
          style={{
            padding: "0 8px",
            display: "flex",
            alignItems: "center",
            border: "1px solid #ADADAD",
            borderBottom: i === value ? "1px solid " + (c.backColor || "#F0F0F0") : "1px solid #ADADAD",
            marginBottom: -1,
            marginRight: -1,
            background: i === value ? c.backColor || "#F0F0F0" : "#E8E8E8",
            color: c.foreColor,
            cursor: "default",
            whiteSpace: "nowrap",
            userSelect: "none",
            position: "relative",
            top: i === value ? 0 : 2,
          }}
        >
          {t}
        </div>
      ))}
    </div>
  );
}

function Container({ c, ctx, values }) {
  if (c.kind === "MultiPage") {
    const pages = c.children || [];
    const v = Math.max(0, Math.min(pages.length - 1, c.value || 0));
    const page = pages[v];
    return (
      <div data-control={c.name} style={{ ...ctx.box(c), boxSizing: "border-box" }} title={c.tip || undefined}>
        <Tabs c={c} captions={pages.map((p) => p.caption)} value={v} ctx={ctx} />
        <div style={{ position: "absolute", left: 0, top: px(15), right: 0, bottom: 0, border: "1px solid #ADADAD", borderTop: "none", background: (page && page.backColor) || c.backColor, overflow: "hidden" }}>
          {page ? <Controls list={page.children || []} ctx={ctx} values={values} /> : null}
        </div>
      </div>
    );
  }
  if (c.kind === "TabStrip") {
    return (
      <div data-control={c.name} style={{ ...ctx.box(c), boxSizing: "border-box", background: c.backColor }} title={c.tip || undefined}>
        <Tabs c={c} captions={c.tabs || []} value={c.value} ctx={ctx} />
        <div style={{ position: "absolute", left: 0, top: px(15), right: 0, bottom: 0, border: "1px solid #ADADAD", borderTop: "none" }} />
      </div>
    );
  }
  // A Frame: its caption on its border, its controls inside.
  return (
    <div
      data-control={c.name}
      onClick={(e) => e.target === e.currentTarget && has(c, "click") && ctx.send({ event: "click", control: c.name })}
      title={c.tip || undefined}
      style={{ ...ctx.box(c), boxSizing: "border-box", ...borderCss({ ...c, specialEffect: c.specialEffect === undefined ? 3 : c.specialEffect }, false), background: c.opaque === false ? "transparent" : c.backColor }}
    >
      {c.caption ? (
        <span style={{ position: "absolute", left: 6, top: -px(c.font ? c.font.size : 8.25) * 0.7, padding: "0 2px", background: c.backColor, color: c.foreColor, ...fontCss(c.font), whiteSpace: "pre", lineHeight: 1 }}>{c.caption}</span>
      ) : null}
      <div style={{ position: "absolute", inset: 0, overflow: "hidden" }}>
        <Controls list={c.children || []} ctx={ctx} values={values} />
      </div>
    </div>
  );
}

function Control({ c, ctx, values }) {
  if (!c.visible && !ctx.design) return null;
  if (ctx.design) {
    // The designer's picture: double-clicking a control opens its event procedure.
    return (
      <div
        style={{ display: "contents" }}
        onDoubleClick={(e) => {
          e.stopPropagation();
          ctx.send({ event: "dblclick", control: c.name });
        }}
      >
        <ControlBody c={c} ctx={ctx} values={values} />
      </div>
    );
  }
  return <ControlBody c={c} ctx={ctx} values={values} />;
}

function ControlBody({ c, ctx, values }) {
  const value = values[c.name];
  switch (c.kind) {
    case "TextBox":
      return <TextBox c={c} value={value !== undefined ? value : c.text} ctx={ctx} />;
    case "ComboBox":
      return <ComboBox c={c} value={value} ctx={ctx} />;
    case "ListBox":
      return <ListBox c={c} value={value} ctx={ctx} />;
    case "CheckBox":
    case "OptionButton":
    case "ToggleButton":
      return <Toggle c={c} value={value} ctx={ctx} />;
    case "CommandButton":
      return <CommandButton c={c} ctx={ctx} />;
    case "Label":
      return <Label c={c} ctx={ctx} />;
    case "Image":
      return <ImageCtl c={c} ctx={ctx} />;
    case "ScrollBar":
      return <ScrollBar c={c} value={value} ctx={ctx} />;
    case "SpinButton":
      return <SpinButton c={c} ctx={ctx} />;
    case "Frame":
    case "MultiPage":
    case "TabStrip":
      return <Container c={c} ctx={ctx} values={values} />;
    default:
      return (
        <div data-control={c.name} style={{ ...ctx.box(c), border: "1px dashed #999", color: "#666", fontSize: 11, overflow: "hidden" }} title={`${c.kind} (not shown in ReflexGrid)`}>
          {c.kind}
        </div>
      );
  }
}

export function Controls({ list, ctx, values }) {
  return (
    <>
      {list.map((c) => (
        <Control key={c.name} c={c} ctx={ctx} values={values} />
      ))}
    </>
  );
}

export function walk(list, f) {
  for (const c of list || []) {
    f(c);
    walk(c.children, f);
  }
}

// ------------------------------------------------------------------ a form

function FormWindow({ form, active, inert, design, onEvent, z, pictures }) {
  const [pos, setPos] = useState(null);
  const [values, setValues] = useState({});
  const pending = useRef({});
  const seq = useRef(0);
  const sentSeq = useRef(0);
  const focusName = useRef(form.active || "");
  const queued = useRef(false);
  const inflight = useRef(false);
  const rootRef = useRef(null);
  const w = px(form.insideWidth);
  const h = px(form.insideHeight);

  // Where the form starts: centred (StartUpPosition 1, 2, 3) or where its Left/Top say.
  useLayoutEffect(() => {
    if (pos !== null) return;
    const vw = window.innerWidth || 1200;
    const vh = window.innerHeight || 800;
    if (design) setPos({ x: 0, y: 0 });
    else if (form.startUp === 0) setPos({ x: Math.max(0, Math.min(px(form.left), vw - w - 12)), y: Math.max(0, Math.min(px(form.top), vh - h - TITLE_H - 12)) });
    else setPos({ x: Math.max(8, (vw - w) / 2), y: Math.max(8, (vh - h - TITLE_H) / 2) });
  }, []);

  // A new snapshot: what was sent is in it now; keep what was typed since.
  useEffect(() => {
    inflight.current = false;
    const keep = {};
    let changed = false;
    for (const [k, v] of Object.entries(pending.current)) {
      if (v.seq > sentSeq.current) keep[k] = v;
      else changed = true;
    }
    pending.current = keep;
    if (changed) setValues(Object.fromEntries(Object.entries(keep).map(([k, v]) => [k, v.value])));
    if (queued.current) {
      queued.current = false;
      flushChanges();
    }
  }, [form]);

  // The code moved the focus (SetFocus) or the form just opened (its first control has it).
  useEffect(() => {
    if (design || inert || !form.active || pos === null) return;
    const root = rootRef.current;
    if (!root) return;
    const el = root.querySelector(`[data-control="${CSS.escape(form.active)}"] input, [data-control-input="${CSS.escape(form.active)}"], [data-control="${CSS.escape(form.active)}"]`);
    if (el && document.activeElement !== el && !(root.contains(document.activeElement) && focusName.current === form.active)) {
      focusName.current = form.active;
      try {
        el.focus({ preventScroll: true });
      } catch (err) {}
    }
  }, [form.active, inert, pos === null]);

  const send = useCallback(
    (ev) => {
      if (design) {
        if (ev.event === "dblclick") onEvent({ form: form.id, event: { event: "design", control: ev.control || "" } });
        return;
      }
      const vals = {};
      for (const [k, v] of Object.entries(pending.current)) vals[k] = v.value;
      sentSeq.current = seq.current;
      inflight.current = true;
      onEvent({ form: form.id, event: { ...ev, values: vals, focus: focusName.current } });
    },
    [form.id, design, onEvent]
  );

  const flushChanges = () => {
    // Controls whose Change the code handles: their latest text, as one event each.
    for (const [k, v] of Object.entries(pending.current)) {
      if (v.seq > sentSeq.current && v.live) {
        send({ event: "change", control: k, value: v.value });
        return;
      }
    }
  };

  const edit = (c, value, live) => {
    seq.current += 1;
    pending.current = { ...pending.current, [c.name]: { value, seq: seq.current, live } };
    setValues((old) => ({ ...old, [c.name]: value }));
    if (live) {
      if (inflight.current) queued.current = true;
      else send({ event: "change", control: c.name, value });
    }
  };

  const controlsByName = useMemo(() => {
    const m = {};
    walk(form.controls, (c) => (m[c.name] = c));
    return m;
  }, [form]);

  const focused = (c, quiet) => {
    if (focusName.current === c.name) return;
    const old = controlsByName[focusName.current];
    focusName.current = c.name;
    if (quiet) return;
    const watch = (x) => x && (has(x, "enter") || has(x, "exit") || has(x, "beforeupdate") || has(x, "afterupdate"));
    if (watch(c) || watch(old)) send({ event: "focus", control: c.name });
  };

  const focusNext = (el) => {
    const root = rootRef.current;
    if (!root) return;
    const all = Array.from(root.querySelectorAll('[tabindex="0"], input:not([tabindex="-1"]), textarea:not([tabindex="-1"]), button:not([tabindex="-1"])')).filter((x) => !x.disabled);
    const i = all.indexOf(el);
    if (i >= 0 && all.length) all[(i + 1) % all.length].focus();
  };

  let hasDefault = false;
  let hasCancel = false;
  walk(form.controls, (c) => {
    if (c.kind === "CommandButton" && c.visible && c.enabled) {
      if (c.default) hasDefault = true;
      if (c.cancel) hasCancel = true;
    }
  });

  const ctx = {
    send,
    edit,
    focused,
    focusNext,
    hasDefault,
    hasCancel,
    inert: inert || design,
    // Disabled for real only when the form can't be used (in the designer's picture controls
    // stay live enough to be double-clicked).
    off: inert && !design,
    design,
    box: (c) => ({ position: "absolute", left: px(c.left), top: px(c.top), width: px(c.width), height: px(c.height) }),
    pic: (p) => (!p ? "" : (pictures && pictures[p]) || (p.startsWith("data:") ? p : "")),
  };

  // Dragging by the title bar.
  const drag = useRef(null);
  const onTitleDown = (e) => {
    if (design || e.button !== 0) return;
    drag.current = { x: e.clientX - pos.x, y: e.clientY - pos.y };
    e.currentTarget.setPointerCapture(e.pointerId);
  };
  const onTitleMove = (e) => {
    if (!drag.current) return;
    setPos({ x: Math.max(-w + 60, e.clientX - drag.current.x), y: Math.max(0, e.clientY - drag.current.y) });
  };
  const onTitleUp = () => {
    drag.current = null;
  };

  if (pos === null) return null;
  return (
    <div
      ref={rootRef}
      className="rg-userform"
      data-form={form.name}
      data-form-id={form.id}
      role="dialog"
      aria-label={form.caption}
      onKeyDown={(e) => {
        if (design) return;
        if (e.key === "Escape" && hasCancel && !e.defaultPrevented && e.target.tagName !== "INPUT" && e.target.tagName !== "TEXTAREA") send({ event: "escape" });
        e.stopPropagation();
      }}
      style={{
        position: design ? "relative" : "fixed",
        left: design ? 0 : pos.x,
        top: design ? 0 : pos.y,
        zIndex: z,
        width: w + 2,
        background: "#FFFFFF",
        border: active ? "1px solid #0078D7" : "1px solid #9A9A9A",
        boxShadow: design ? "none" : "0 8px 28px rgba(0,0,0,.28)",
        colorScheme: "light",
        color: "#000000",
        userSelect: "none",
        pointerEvents: "auto",
      }}
    >
      <div
        onPointerDown={onTitleDown}
        onPointerMove={onTitleMove}
        onPointerUp={onTitleUp}
        onDoubleClick={() => design && onEvent({ form: form.id, event: { event: "design", control: "" } })}
        style={{ height: TITLE_H, display: "flex", alignItems: "center", padding: "0 0 0 8px", background: "#FFFFFF", cursor: design ? "default" : "move", fontFamily: "'Segoe UI', Tahoma, Arial, sans-serif", fontSize: 12, color: active ? "#000000" : "#6B6B6B" }}
      >
        <span style={{ flex: 1, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{form.caption}</span>
        <button
          type="button"
          aria-label="Close"
          data-form-close={form.name}
          disabled={inert && !design}
          onPointerDown={(e) => e.stopPropagation()}
          onClick={() => !design && send({ event: "close" })}
          style={{ width: 40, height: TITLE_H, border: "none", background: "transparent", fontSize: 14, color: "#222", cursor: "default" }}
          onMouseEnter={(e) => !design && (e.currentTarget.style.background = "#E81123")}
          onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
        >
          ✕
        </button>
      </div>
      <div
        onClick={(e) => e.target === e.currentTarget && has(form, "click") && send({ event: "click" })}
        onDoubleClick={(e) => e.target === e.currentTarget && (design || has(form, "dblclick")) && send({ event: "dblclick" })}
        style={{ position: "relative", width: w, height: h, background: form.backColor, color: form.foreColor, overflow: "hidden", ...fontCss(form.font), opacity: inert && !design ? 0.85 : 1, pointerEvents: inert && !design ? "none" : "auto" }}
      >
        {form.picture ? <img src={ctx.pic(form.picture)} alt="" style={{ position: "absolute", left: 0, top: 0, pointerEvents: "none" }} /> : null}
        <Controls list={form.controls || []} ctx={ctx} values={values} />
      </div>
    </div>
  );
}

export function UserForms(props) {
  const { forms, waiting, blocked, design, onEvent, pictures } = props;
  const list = Array.isArray(forms) ? forms : [];
  const send = useCallback((e) => onEvent && onEvent(e), [onEvent]);
  if (!list.length) return null;
  if (design) {
    return (
      <div style={{ display: "inline-block", padding: 12 }}>
        {list.map((f) => (
          <FormWindow key={f.id} form={f} active={true} inert={true} design={true} onEvent={send} z={1} pictures={pictures} />
        ))}
      </div>
    );
  }
  // The form the macro waits on is on top; while it is up the other forms can't be used.
  const top = list.find((f) => f.id === waiting) || null;
  const modalUp = !!top || list.some((f) => f.modal);
  return (
    <>
      {modalUp ? <div className="rg-userform-backdrop" style={{ position: "fixed", inset: 0, zIndex: 40, background: "rgba(0,0,0,0.06)", pointerEvents: "auto" }} /> : null}
      {list.map((f, i) => {
        const isTop = top ? f.id === top.id : f.modal || !modalUp;
        return <FormWindow key={f.id} form={f} active={isTop && !blocked} inert={blocked || !isTop} design={false} onEvent={send} z={isTop ? 42 + i : 39} pictures={pictures} />;
      })}
    </>
  );
}
