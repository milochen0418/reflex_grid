// ReflexGrid's VBA code editor: a textarea over a highlighted copy of its text, with line numbers.
// Keys: Tab / Shift+Tab indent, Enter keeps the indent, Ctrl/Cmd+S applies the code, F5 runs the
// Sub the cursor is in. The code goes to the server on onChange (debounced) and onSave. A new
// gotoKey puts the caret at the end of line gotoLine (an event procedure the form view opened).
// With the AI (a hub address): its bar and the change it writes (vba_ai.jsx), and the grey
// suggestion at the cursor, which Tab takes and Esc dismisses.
import React, { useEffect, useMemo, useRef, useState, useCallback } from "react";
import { AiBar, DraftPane } from "./vba_ai.jsx";

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
const GHOST_WAIT_MS = 500; // a pause in the typing before a suggestion is asked for
const GHOST_GIVE_UP_MS = 30000;

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
  const { value, docKey, errorLine, readOnly, onChange, onSave, onRun, fontSize, gotoLine, gotoKey, module, hub, draft, ghost, suggest, fixable, words, onAi } = props;
  const [text, setText] = useState(value || "");
  const [shownGhost, setShownGhost] = useState(null); // {at, text}
  const area = useRef(null);
  const pre = useRef(null);
  const gutter = useRef(null);
  const timer = useRef(null);
  const sent = useRef(value || "");
  const bar = useRef(null);
  const ghostTimer = useRef(null);
  const ghostReq = useRef(0); // the suggestion last asked for
  const ghostSince = useRef(0); // when it was asked, while the answer is awaited
  const ghostAt = useRef(null); // {text, at}: where it was asked for
  const canAi = !!hub && !readOnly;
  // The change for the module on screen.
  const shown = draft && draft.status && draft.docKey === docKey ? draft : null;
  const covering = shown && ["writing", "checking", "done"].includes(shown.status);

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

  useEffect(() => () => ghostTimer.current && clearTimeout(ghostTimer.current), []);

  // Module changed under a change the AI was writing: that change is over.
  useEffect(() => {
    if (draft && draft.status && draft.docKey !== docKey && onAi) onAi({ type: "close" });
  }, [docKey, draft && draft.seq]);

  const askGhost = () => {
    const a = area.current;
    if (!canAi || !suggest || draft && draft.status || !a || document.activeElement !== a) return;
    const at = a.selectionStart;
    const v = a.value;
    if (at !== a.selectionEnd || !v.trim()) return;
    if (ghostSince.current && Date.now() - ghostSince.current < GHOST_GIVE_UP_MS) return; // one at a time
    const end = v.indexOf("\n", at);
    // At the end of a line, or before what closes it; not in a comment.
    if (!/^[\s)\]"',]*$/.test(v.slice(at, end < 0 ? v.length : end))) return;
    if (v.slice(v.lastIndexOf("\n", at - 1) + 1, at).trimStart().startsWith("'")) return;
    ghostReq.current += 1;
    ghostSince.current = Date.now();
    ghostAt.current = { text: v, at };
    onAi({ type: "suggest", req: ghostReq.current, prefix: v.slice(0, at), suffix: v.slice(at) });
  };

  const clearGhost = () => {
    if (ghostTimer.current) clearTimeout(ghostTimer.current);
    ghostTimer.current = null;
    setShownGhost(null);
  };

  useEffect(() => {
    if (!ghost || ghost.req !== ghostReq.current) return;
    ghostSince.current = 0;
    const a = area.current;
    const was = ghostAt.current;
    if (!a || !was) return;
    if (a.value !== was.text || a.selectionStart !== was.at) {
      // Typed on while it was being asked: ask for where the cursor is now.
      if (!ghostTimer.current) ghostTimer.current = setTimeout(askGhost, GHOST_WAIT_MS);
    } else if (ghost.text) {
      setShownGhost({ at: was.at, text: ghost.text });
    }
  }, [ghost]);

  const change = (t) => {
    setText(t);
    if (timer.current) clearTimeout(timer.current);
    timer.current = setTimeout(() => flush(t), 400);
    clearGhost();
    if (canAi && suggest) ghostTimer.current = setTimeout(() => {
      ghostTimer.current = null;
      askGhost();
    }, GHOST_WAIT_MS);
  };

  const html = useMemo(() => {
    const lines = text.split("\n");
    const g = shownGhost;
    const before = g ? text.slice(0, g.at).split("\n") : [];
    const gLine = before.length - 1;
    const gCol = g ? before[gLine].length : 0;
    const grey = (s) => `<span class="vba-ghost" style="color:#9CA3AF">${escapeHtml(s)}</span>`;
    return lines
      .map((l, i) => {
        let body = highlightLine(l) || " ";
        if (g && i === gLine) body = highlightLine(l.slice(0, gCol)) + grey(g.text) + highlightLine(l.slice(gCol));
        if (errorLine && i + 1 === errorLine) return `<span style="background:#FDE2E1;display:inline-block;min-width:100%">${body}</span>`;
        return body;
      })
      .join("\n") + "\n";
  }, [text, errorLine, shownGhost]);

  const numbers = useMemo(() => {
    const out = [];
    const count = text.split("\n").length;
    const g = shownGhost;
    const gLine = g ? text.slice(0, g.at).split("\n").length : 0;
    const extra = g ? g.text.split("\n").length - 1 : 0;
    for (let i = 1; i <= count; i++) {
      out.push(i);
      if (i === gLine) for (let k = 1; k <= extra; k++) out.push(-k); // beside the suggestion's lines
    }
    return out;
  }, [text, shownGhost]);

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

  useEffect(() => {
    if (!gotoLine) return;
    requestAnimationFrame(() => {
      const a = area.current;
      if (!a) return;
      const lines = a.value.split("\n");
      let at = 0;
      for (let i = 0; i < gotoLine - 1 && i < lines.length; i++) at += lines[i].length + 1;
      at += (lines[gotoLine - 1] || "").length;
      a.focus();
      a.setSelectionRange(at, at);
      const lh = (fontSize || 13) * 1.5;
      const top = (gotoLine - 1) * lh;
      if (top < a.scrollTop || top > a.scrollTop + a.clientHeight - lh) a.scrollTop = Math.max(0, top - a.clientHeight / 3);
      syncScroll();
    });
  }, [gotoKey]);

  // Into the textarea as typing is, so that its Ctrl+Z takes it back.
  const typeIn = (a, start, end, insert) => {
    a.focus();
    a.setSelectionRange(start, end);
    const before = a.value;
    const next = before.slice(0, start) + insert + before.slice(end);
    const done = document.execCommand(insert ? "insertText" : "delete", false, insert);
    if (!done || a.value !== next) {
      a.value = next;
      change(next);
      a.setSelectionRange(start + insert.length, start + insert.length);
    }
    return next;
  };

  const takeGhost = (a) => {
    const g = shownGhost;
    clearGhost();
    typeIn(a, g.at, g.at, g.text);
  };

  // Accepts the AI's change: the smallest edit that makes the code what it wrote, applied at once.
  const accept = () => {
    const a = area.current;
    if (!a || !shown || shown.status !== "done") return;
    const old = a.value;
    const next = shown.code;
    let start = 0;
    const limit = Math.min(old.length, next.length);
    while (start < limit && old[start] === next[start]) start++;
    let endOld = old.length;
    let endNew = next.length;
    while (endOld > start && endNew > start && old[endOld - 1] === next[endNew - 1]) {
      endOld--;
      endNew--;
    }
    onAi({ type: "close" });
    if (next !== old) typeIn(a, start, endOld, next.slice(start, endNew));
    clearGhost();
    flush(next);
    if (onSave) onSave(next);
    const line = next.slice(0, start).split("\n").length;
    requestAnimationFrame(() => {
      const lh = (fontSize || 13) * 1.5;
      a.scrollTop = Math.max(0, (line - 1) * lh - a.clientHeight / 3);
      syncScroll();
    });
  };

  // Keys that answer the change on screen, wherever the keyboard is.
  useEffect(() => {
    if (!shown) return undefined;
    const onKey = (e) => {
      const take = (f) => {
        e.preventDefault();
        e.stopPropagation();
        f();
      };
      if ((e.metaKey || e.ctrlKey) && e.key === "Enter" && shown.status === "done") take(accept);
      else if (e.key === "Escape" && !e.isComposing) take(() => onAi({ type: "close" }));
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  });

  // What the bar's request is about: the cursor's line, and the whole lines selected.
  const selection = () => {
    const a = area.current;
    if (!a) return { line: 0, first: 0, last: 0 };
    const lineAt = (i) => a.value.slice(0, i).split("\n").length;
    const s = a.selectionStart;
    const e = a.selectionEnd;
    if (s === e) return { line: lineAt(s), first: 0, last: 0 };
    const last = lineAt(a.value[e - 1] === "\n" ? e - 1 : e);
    return { line: lineAt(s), first: lineAt(s), last };
  };

  // For the AI Hub's assistant opened on the page (sheet_view.jsx): where the person is.
  useEffect(() => {
    const context = () => {
      const a = area.current;
      const line = a ? a.value.slice(0, a.selectionStart).split("\n").length : 0;
      const proc = a ? procedureAt(a.value, line - 1) : "";
      return `The VBA editor is open on module ${module || "(none)"}, the cursor on line ${line}${proc ? `, in ${proc}` : ""}.`;
    };
    window.__vbaContext = context;
    return () => {
      if (window.__vbaContext === context) delete window.__vbaContext;
    };
  }, [module]);

  const replaceRange = (a, start, end, insert, selStart, selEnd) => {
    typeIn(a, start, end, insert);
    a.setSelectionRange(selStart, selEnd);
  };

  const onKeyDown = (e) => {
    const a = area.current;
    if (!a) return;
    const mod = e.ctrlKey || e.metaKey;
    if (shownGhost && e.key === "Tab" && !e.shiftKey) {
      e.preventDefault();
      takeGhost(a);
      return;
    }
    if (shownGhost && e.key === "Escape") {
      e.preventDefault();
      clearGhost();
      return;
    }
    if (canAi && mod && (e.key === "i" || e.key === "I")) {
      e.preventDefault();
      if (bar.current) bar.current.focus();
      return;
    }
    if (canAi && e.altKey && e.code === "KeyV") {
      e.preventDefault();
      if (bar.current) bar.current.toggleMic();
      return;
    }
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

  const code = (
    <div style={{ display: "flex", flex: 1, minHeight: 0, width: "100%", position: "relative" }}>
      <div
        ref={gutter}
        aria-hidden="true"
        style={{ ...common, padding: "6px 6px 6px 4px", width: 46, flexShrink: 0, overflow: "hidden", textAlign: "right", color: "#8A9199", background: "#F3F4F6", borderRight: "1px solid #E1E4E8", userSelect: "none" }}
      >
        {numbers.map((n) => (
          <div key={n} style={{ height: lh, color: n === errorLine ? "#C5221F" : undefined, fontWeight: n === errorLine ? 700 : 400 }}>
            {n > 0 ? n : ""}
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
          onBlur={(e) => {
            flush(e.target.value);
            clearGhost();
          }}
          onSelect={(e) => {
            if (shownGhost && e.target.selectionStart !== shownGhost.at) clearGhost();
          }}
          onScroll={syncScroll}
          onKeyDown={onKeyDown}
          style={{ ...common, position: "absolute", inset: 0, width: "100%", height: "100%", resize: "none", outline: "none", overflow: "auto", background: "transparent", color: "transparent", caretColor: "#111827" }}
        />
      </div>
      {covering && <DraftPane draft={shown} base={text} highlight={highlightLine} common={common} lineHeight={lh} />}
    </div>
  );

  return (
    <div style={{ display: "flex", flexDirection: "column", width: "100%", height: "100%", background: "#FFFFFF", border: "1px solid #C9CED4", borderRadius: 4, overflow: "hidden" }}>
      {canAi && (
        <AiBar
          ref={bar}
          draft={shown}
          hub={hub}
          suggest={suggest}
          fixable={fixable}
          words={words}
          selection={selection}
          code={() => (area.current ? area.current.value : text)}
          apply={accept}
          onAi={onAi}
          onReturn={() => area.current && area.current.focus()}
        />
      )}
      {code}
    </div>
  );
}
