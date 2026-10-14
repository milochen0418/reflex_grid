// ReflexGrid's VBA code editor: a textarea over a highlighted copy of its text, with line numbers.
// Keys: Tab / Shift+Tab indent, Enter keeps the indent, Ctrl/Cmd+S applies the code, F5 runs the
// Sub the cursor is in. The code goes to the server on onChange (debounced) and onSave.
import React, { useEffect, useMemo, useRef, useState, useCallback } from "react";

const KEYWORDS = new Set(
  (
    "AddressOf Alias And As Attribute Boolean ByRef Byte ByVal Call Case Const Currency Date Decimal Declare " +
    "DefBool DefByte DefCur DefDate DefDbl DefInt DefLng DefLngLng DefLngPtr DefObj DefSng DefStr DefVar Dim Do " +
    "Double Each Else ElseIf Empty End Enum Eqv Erase Error Event Exit Explicit False For Friend Function Get " +
    "GoSub GoTo If Imp Implements In Integer Is Let Lib Like Long LongLong LongPtr Loop LSet Me Mod New Next Not " +
    "Nothing Null Object On Option Optional Or ParamArray Preserve Private Property PtrSafe Public RaiseEvent " +
    "ReDim Resume Return RSet Select Set Single Static Step Stop String Sub Then To True Type TypeOf Until " +
    "Variant WEnd While With WithEvents Xor Base Compare Binary Text Module Debug Print"
  )
    .split(" ")
    .map((w) => w.toLowerCase())
);

const COLORS = { keyword: "#00008B", comment: "#008000", string: "#A31515", number: "#098658" };

function escapeHtml(s) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// One line of VBA as HTML spans (comments, strings, keywords, numbers).
function highlightLine(line) {
  let out = "";
  let i = 0;
  const n = line.length;
  while (i < n) {
    const ch = line[i];
    if (ch === "'") {
      out += `<span style="color:${COLORS.comment}">${escapeHtml(line.slice(i))}</span>`;
      break;
    }
    if (ch === '"') {
      let j = i + 1;
      while (j < n) {
        if (line[j] === '"') {
          if (line[j + 1] === '"') {
            j += 2;
            continue;
          }
          break;
        }
        j++;
      }
      out += `<span style="color:${COLORS.string}">${escapeHtml(line.slice(i, j + 1))}</span>`;
      i = j + 1;
      continue;
    }
    if (/[A-Za-z_]/.test(ch)) {
      let j = i + 1;
      while (j < n && /[A-Za-z0-9_]/.test(line[j])) j++;
      const word = line.slice(i, j);
      const prev = i > 0 ? line[i - 1] : "";
      if (word.toLowerCase() === "rem" && line.slice(0, i).trim() === "") {
        out += `<span style="color:${COLORS.comment}">${escapeHtml(line.slice(i))}</span>`;
        break;
      }
      if (KEYWORDS.has(word.toLowerCase()) && prev !== ".") {
        out += `<span style="color:${COLORS.keyword}">${escapeHtml(word)}</span>`;
      } else {
        out += escapeHtml(word);
      }
      i = j;
      continue;
    }
    if (/[0-9]/.test(ch)) {
      let j = i + 1;
      while (j < n && /[0-9.eE#&!@%^]/.test(line[j])) j++;
      out += `<span style="color:${COLORS.number}">${escapeHtml(line.slice(i, j))}</span>`;
      i = j;
      continue;
    }
    out += escapeHtml(ch);
    i++;
  }
  return out;
}

const PROC_START = /^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)/i;

// The procedure the cursor is in: its name, or "".
function procedureAt(text, line) {
  const lines = text.split("\n");
  for (let i = Math.min(line, lines.length - 1); i >= 0; i--) {
    const m = PROC_START.exec(lines[i]);
    if (m) return m[2];
    if (i < line && /^\s*End\s+(Sub|Function|Property)\b/i.test(lines[i])) return "";
  }
  return "";
}

export function VbaEditor(props) {
  const { value, docKey, errorLine, readOnly, onChange, onSave, onRun, fontSize } = props;
  const [text, setText] = useState(value || "");
  const area = useRef(null);
  const pre = useRef(null);
  const gutter = useRef(null);
  const timer = useRef(null);
  const sent = useRef(value || "");

  // Another module (or a reload after someone else's change) replaces the text.
  useEffect(() => {
    setText(value || "");
    sent.current = value || "";
    if (area.current) {
      area.current.scrollTop = 0;
      area.current.setSelectionRange(0, 0);
    }
  }, [docKey]);

  const flush = useCallback(
    (t) => {
      if (timer.current) {
        clearTimeout(timer.current);
        timer.current = null;
      }
      if (t !== sent.current) {
        sent.current = t;
        if (onChange) onChange(t);
      }
    },
    [onChange]
  );

  useEffect(() => () => timer.current && clearTimeout(timer.current), []);

  const change = (t) => {
    setText(t);
    if (timer.current) clearTimeout(timer.current);
    timer.current = setTimeout(() => flush(t), 400);
  };

  const html = useMemo(() => {
    const lines = text.split("\n");
    return lines
      .map((l, i) => {
        const body = highlightLine(l) || " ";
        if (errorLine && i + 1 === errorLine) return `<span style="background:#FDE2E1;display:inline-block;min-width:100%">${body}</span>`;
        return body;
      })
      .join("\n") + "\n";
  }, [text, errorLine]);

  const lineCount = useMemo(() => text.split("\n").length, [text]);

  const syncScroll = () => {
    const a = area.current;
    if (!a) return;
    if (pre.current) {
      pre.current.scrollTop = a.scrollTop;
      pre.current.scrollLeft = a.scrollLeft;
    }
    if (gutter.current) gutter.current.scrollTop = a.scrollTop;
  };

  // Scroll an error line into view.
  useEffect(() => {
    const a = area.current;
    if (!a || !errorLine) return;
    const lh = (fontSize || 13) * 1.5;
    const top = (errorLine - 1) * lh;
    if (top < a.scrollTop || top > a.scrollTop + a.clientHeight - lh) a.scrollTop = Math.max(0, top - a.clientHeight / 3);
    syncScroll();
  }, [errorLine, docKey]);

  const replaceRange = (a, start, end, insert, selStart, selEnd) => {
    const t = a.value.slice(0, start) + insert + a.value.slice(end);
    change(t);
    requestAnimationFrame(() => {
      a.value = t;
      a.setSelectionRange(selStart, selEnd);
    });
  };

  const onKeyDown = (e) => {
    const a = area.current;
    if (!a) return;
    const mod = e.ctrlKey || e.metaKey;
    if (mod && (e.key === "s" || e.key === "S")) {
      e.preventDefault();
      flush(a.value);
      if (onSave) onSave(a.value);
      return;
    }
    if (e.key === "F5") {
      e.preventDefault();
      flush(a.value);
      const line = a.value.slice(0, a.selectionStart).split("\n").length - 1;
      if (onRun) onRun(a.value, procedureAt(a.value, line));
      return;
    }
    if (readOnly) return;
    if (e.key === "Tab") {
      e.preventDefault();
      const v = a.value;
      const s = a.selectionStart;
      const en = a.selectionEnd;
      const lineStart = v.lastIndexOf("\n", s - 1) + 1;
      if (s === en && !e.shiftKey) {
        replaceRange(a, s, en, "    ", s + 4, s + 4);
        return;
      }
      const lineEnd = v.indexOf("\n", en - (en > s && v[en - 1] === "\n" ? 1 : 0));
      const block = v.slice(lineStart, lineEnd === -1 ? v.length : lineEnd);
      const lines = block.split("\n");
      const out = lines.map((l) => (e.shiftKey ? l.replace(/^( {1,4}|\t)/, "") : "    " + l)).join("\n");
      replaceRange(a, lineStart, lineStart + block.length, out, lineStart, lineStart + out.length);
      return;
    }
    if (e.key === "Enter" && !e.shiftKey && !mod) {
      e.preventDefault();
      const v = a.value;
      const s = a.selectionStart;
      const lineStart = v.lastIndexOf("\n", s - 1) + 1;
      const indent = /^[ \t]*/.exec(v.slice(lineStart, s))[0];
      const ins = "\n" + indent;
      replaceRange(a, s, a.selectionEnd, ins, s + ins.length, s + ins.length);
      return;
    }
  };

  const fs = fontSize || 13;
  const lh = fs * 1.5;
  const mono = "'SF Mono', Menlo, Consolas, 'Liberation Mono', 'Courier New', monospace";
  const common = {
    margin: 0,
    padding: "6px 8px",
    fontFamily: mono,
    fontSize: fs + "px",
    lineHeight: lh + "px",
    whiteSpace: "pre",
    tabSize: 4,
    border: 0,
    boxSizing: "border-box",
  };
  const numbers = [];
  for (let i = 1; i <= lineCount; i++) numbers.push(i);

  return (
    <div style={{ display: "flex", width: "100%", height: "100%", background: "#FFFFFF", border: "1px solid #C9CED4", borderRadius: 4, overflow: "hidden" }}>
      <div
        ref={gutter}
        aria-hidden="true"
        style={{ ...common, padding: "6px 6px 6px 4px", width: 46, flexShrink: 0, overflow: "hidden", textAlign: "right", color: "#8A9199", background: "#F3F4F6", borderRight: "1px solid #E1E4E8", userSelect: "none" }}
      >
        {numbers.map((n) => (
          <div key={n} style={{ height: lh, color: n === errorLine ? "#C5221F" : undefined, fontWeight: n === errorLine ? 700 : 400 }}>
            {n}
          </div>
        ))}
        <div style={{ height: 40 }} />
      </div>
      <div style={{ position: "relative", flex: 1, minWidth: 0 }}>
        <pre
          ref={pre}
          aria-hidden="true"
          style={{ ...common, position: "absolute", inset: 0, overflow: "hidden", color: "#1F2937", pointerEvents: "none" }}
          dangerouslySetInnerHTML={{ __html: html }}
        />
        <textarea
          ref={area}
          id="vba-code"
          value={text}
          readOnly={!!readOnly}
          spellCheck={false}
          autoCapitalize="off"
          autoComplete="off"
          autoCorrect="off"
          wrap="off"
          onChange={(e) => change(e.target.value)}
          onBlur={(e) => flush(e.target.value)}
          onScroll={syncScroll}
          onKeyDown={onKeyDown}
          style={{ ...common, position: "absolute", inset: 0, width: "100%", height: "100%", resize: "none", outline: "none", overflow: "auto", background: "transparent", color: "transparent", caretColor: "#111827" }}
        />
      </div>
    </div>
  );
}
