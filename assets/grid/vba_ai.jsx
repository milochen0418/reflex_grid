// The AI in the VBA editor (reflex_grid/vba_ai.py): the bar above the code, where the person types
// or says what they want (Cmd/Ctrl+I, Alt+V for the microphone), and the pane over the code that
// shows the change as it is written and then as a diff to accept (Cmd/Ctrl+Enter) or reject (Esc).
// Speaking: a pause sends what was said; "accept", "reject" and "fix" answer by themselves.
import React, { forwardRef, useEffect, useImperativeHandle, useMemo, useRef, useState } from "react";
import { Ears, commandOf } from "./vba_voice.js";

const VIOLET = "#7C3AED";
const MOD = typeof navigator !== "undefined" && /Mac/.test(navigator.platform) ? "⌘" : "Ctrl+";
const AUTO_SEND_MS = 2200; // after the last words heard
const CONTEXT = 3; // unchanged lines kept around a change
const MAX_LCS = 4000000; // cells of the table below; beyond, the changed part is simply replaced

// Lines of a and b as a list of {t: "=" | "-" | "+", s, o, n}: the common head and tail, and the
// longest common subsequence of what is between.
export function lineDiff(a, b) {
  let head = 0;
  while (head < a.length && head < b.length && a[head] === b[head]) head++;
  let ea = a.length;
  let eb = b.length;
  while (ea > head && eb > head && a[ea - 1] === b[eb - 1]) {
    ea--;
    eb--;
  }
  const out = [];
  let o = 1;
  let n = 1;
  const same = (s) => out.push({ t: "=", s, o: o++, n: n++ });
  const gone = (s) => out.push({ t: "-", s, o: o++ });
  const added = (s) => out.push({ t: "+", s, n: n++ });
  for (let i = 0; i < head; i++) same(a[i]);
  const rows = ea - head;
  const cols = eb - head;
  if (rows * cols > MAX_LCS) {
    for (let i = head; i < ea; i++) gone(a[i]);
    for (let j = head; j < eb; j++) added(b[j]);
  } else {
    const w = cols + 1;
    const table = new Uint32Array((rows + 1) * w);
    for (let i = rows - 1; i >= 0; i--) {
      for (let j = cols - 1; j >= 0; j--) {
        table[i * w + j] = a[head + i] === b[head + j] ? table[(i + 1) * w + j + 1] + 1 : Math.max(table[(i + 1) * w + j], table[i * w + j + 1]);
      }
    }
    let i = 0;
    let j = 0;
    while (i < rows || j < cols) {
      if (i < rows && j < cols && a[head + i] === b[head + j]) {
        same(a[head + i]);
        i++;
        j++;
      } else if (i < rows && (j >= cols || table[(i + 1) * w + j] >= table[i * w + j + 1])) {
        gone(a[head + i]);
        i++;
      } else {
        added(b[head + j]);
        j++;
      }
    }
  }
  for (let i = ea; i < a.length; i++) same(a[i]);
  return out;
}

// Long unchanged runs fold into one row {t: "…", count}.
export function folded(diff) {
  const out = [];
  let i = 0;
  while (i < diff.length) {
    if (diff[i].t !== "=") {
      out.push(diff[i++]);
      continue;
    }
    let end = i;
    while (end < diff.length && diff[end].t === "=") end++;
    const keepHead = i === 0 ? 0 : CONTEXT;
    const keepTail = end === diff.length ? 0 : CONTEXT;
    if (end - i > keepHead + keepTail + 1) {
      out.push(...diff.slice(i, i + keepHead), { t: "…", count: end - i - keepHead - keepTail }, ...diff.slice(end - keepTail, end));
    } else {
      out.push(...diff.slice(i, end));
    }
    i = end;
  }
  return out;
}

const ROW_COLORS = { "+": "#E6F6EA", "-": "#FDEBEA" };
const SIGN_COLORS = { "+": "#1A7F37", "-": "#C5221F" };

// Over the code: what is being written, then the change.
export function DraftPane({ draft, base, highlight, common, lineHeight }) {
  const pane = useRef(null);
  const done = draft.status === "done";
  const rows = useMemo(() => (done ? folded(lineDiff(base.replace(/\n$/, "").split("\n"), String(draft.code || "").replace(/\n$/, "").split("\n"))) : []), [done, base, draft.code]);
  useEffect(() => {
    if (!done && pane.current) pane.current.scrollTop = pane.current.scrollHeight;
  }, [draft.text, done]);
  useEffect(() => {
    if (done && pane.current) pane.current.scrollTop = 0;
  }, [done]);
  const box = { ...common, padding: 0, overflow: "visible" };
  return (
    <div ref={pane} id="vba-ai-pane" style={{ position: "absolute", inset: 0, overflow: "auto", background: "#FFFFFF", zIndex: 2 }}>
      {done ? (
        <div id="vba-ai-diff" style={{ ...box, minWidth: "100%", width: "max-content", padding: "6px 0" }}>
          {rows.map((r, i) =>
            r.t === "…" ? (
              <div key={i} style={{ height: lineHeight, color: "#8A9199", background: "#F3F4F6", padding: "0 10px" }}>
                ⋯ {r.count} unchanged lines
              </div>
            ) : (
              <div key={i} className={"vba-diff-" + (r.t === "+" ? "add" : r.t === "-" ? "del" : "same")} style={{ display: "flex", height: lineHeight, background: ROW_COLORS[r.t] }}>
                <span style={{ width: 44, flexShrink: 0, textAlign: "right", color: "#8A9199", paddingRight: 6, userSelect: "none" }}>{r.t === "+" ? "" : r.o}</span>
                <span style={{ width: 44, flexShrink: 0, textAlign: "right", color: "#8A9199", paddingRight: 6, userSelect: "none" }}>{r.t === "-" ? "" : r.n}</span>
                <span style={{ width: 16, flexShrink: 0, textAlign: "center", fontWeight: 700, color: SIGN_COLORS[r.t], userSelect: "none" }}>{r.t === "=" ? "" : r.t}</span>
                <span style={{ color: "#1F2937", paddingRight: 12 }} dangerouslySetInnerHTML={{ __html: highlight(r.s) || " " }} />
              </div>
            )
          )}
        </div>
      ) : (
        <pre id="vba-ai-writing" style={{ ...box, padding: "6px 10px", color: "#1F2937", minHeight: "100%" }} dangerouslySetInnerHTML={{ __html: String(draft.text || "").split("\n").map((l) => highlight(l) || " ").join("\n") }} />
      )}
    </div>
  );
}

function Sparkles({ size = 14, color = VIOLET }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ flexShrink: 0 }}>
      <path d="M12 3l1.8 5.2L19 10l-5.2 1.8L12 17l-1.8-5.2L5 10l5.2-1.8z" />
      <path d="M19 15l.7 2.3L22 18l-2.3.7L19 21l-.7-2.3L16 18l2.3-.7z" />
    </svg>
  );
}

function Mic({ on }) {
  return (
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke={on ? "#FFFFFF" : "#4B5563"} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <rect x="9" y="2" width="6" height="12" rx="3" />
      <path d="M5 11a7 7 0 0 0 14 0" />
      <path d="M12 18v4" />
    </svg>
  );
}

const button = { border: "1px solid #D5D9DE", background: "#FFFFFF", color: "#1F2937", borderRadius: 4, fontSize: 12, height: 24, padding: "0 8px", cursor: "pointer", whiteSpace: "nowrap", flexShrink: 0 };
const primary = { ...button, background: VIOLET, borderColor: VIOLET, color: "#FFFFFF" };

const join = (a, b) => (!a ? b : !b ? a : /[　-鿿＀-￯]$/.test(a) || /^[　-鿿＀-￯]/.test(b) ? a + b : a + " " + b);

// What the person says or types, and what happens with the change on screen. `selection()` says
// {line, first, last} and `code()` gives the text; `apply()` accepts the change.
export const AiBar = forwardRef(function AiBar({ draft, hub, suggest, fixable, words: known, selection, code, apply, onAi, onReturn }, ref) {
  const [prompt, setPrompt] = useState("");
  const [hearing, setHearing] = useState("");
  const [listening, setListening] = useState(false);
  const [note, setNote] = useState("");
  const [pending, setPending] = useState(0); // the countdown to sending what was said
  const [picked, setPicked] = useState(null);
  const box = useRef(null);
  const ears = useRef(null);
  const timer = useRef(null);
  const status = draft ? draft.status : "";
  const writing = status === "writing" || status === "checking";
  const review = status === "done";
  const latest = useRef({});

  const send = (text) => {
    const words = (text === undefined ? latest.current.prompt : text).trim();
    if (!words) return;
    disarm();
    onAi({ type: "ask", prompt: words, code: code(), refine: latest.current.review, ...selection() });
    setPrompt("");
  };

  const disarm = () => {
    if (timer.current) clearTimeout(timer.current);
    timer.current = null;
    setPending(0);
  };

  const arm = () => {
    disarm();
    setPending(Date.now());
    timer.current = setTimeout(function fire() {
      // While the AI writes, what was said waits for its answer.
      if (latest.current.writing) timer.current = setTimeout(fire, 400);
      else send();
    }, AUTO_SEND_MS);
  };

  // Words from the microphone: an answer to the change on screen, else more of the request.
  const onWords = (words) => {
    const now = latest.current;
    const command = now.prompt.trim() ? "" : commandOf(words);
    if (command === "accept" && now.review) return now.apply();
    if (command === "reject" && (now.review || now.writing)) return onAi({ type: "close" });
    if (command === "fix" && now.fixable && !now.writing) return onAi({ type: "fix" });
    if (command) return undefined; // nothing to answer: a stray word
    setPrompt((p) => {
      const next = join(p, words);
      latest.current.prompt = next;
      return next;
    });
    arm();
    return undefined;
  };

  latest.current = { prompt, writing, review, fixable, apply, known };

  const toggleMic = async () => {
    if (listening) {
      ears.current && ears.current.stop();
      setListening(false);
      setNote("");
      disarm();
      return;
    }
    if (!ears.current) {
      ears.current = new Ears({
        hub,
        hint: () => latest.current.known,
        onHearing: (words) => {
          setHearing(words);
          if (words && timer.current) arm(); // still talking
        },
        onWords: (words) => latest.current.onWords(words),
        onNote: setNote,
      });
    }
    setListening(await ears.current.start());
  };
  latest.current.onWords = onWords;

  useEffect(() => () => {
    if (ears.current) ears.current.stop();
    if (timer.current) clearTimeout(timer.current);
  }, []);

  useImperativeHandle(ref, () => ({
    focus: () => box.current && box.current.focus(),
    toggleMic,
  }));

  // The selection a request will be about, as it was when the box took the keyboard.
  useEffect(() => setPicked(selection()), [status]);

  // The box grows with what is said, to four lines.
  const shown = hearing ? join(prompt, hearing) : prompt;
  useEffect(() => {
    const el = box.current;
    if (!el) return;
    el.style.height = "auto";
    el.style.height = Math.min(el.scrollHeight, 88) + "px";
  }, [shown]);

  const onKeyDown = (e) => {
    if (e.key === "Enter" && !e.shiftKey && !e.nativeEvent.isComposing && e.keyCode !== 229) {
      e.preventDefault();
      send(shown);
    } else if (e.key === "Escape" && pending) {
      e.preventDefault();
      e.stopPropagation();
      disarm();
    } else if (e.key === "Escape" && !review) {
      e.preventDefault();
      onReturn();
    } else if (e.altKey && e.code === "KeyV") {
      e.preventDefault();
      toggleMic();
    }
  };

  const lines = picked && picked.first ? (picked.first === picked.last ? `Line ${picked.first}` : `Lines ${picked.first}–${picked.last}`) : "";
  const fixing = draft && draft.kind === "fix";
  const words = status === "checking" ? "Checking: it didn't compile, asking again…" : fixing ? "Fixing…" : "Writing…";
  return (
    <div id="vba-ai-bar" style={{ flexShrink: 0, background: "#F7F5FF", borderBottom: "1px solid #E4DEFB" }}>
      <style>{`@keyframes vbaAiSend { from { width: 100%; } to { width: 0; } } @keyframes vbaAiPulse { 50% { opacity: 0.55; } }`}</style>
      <div style={{ display: "flex", alignItems: "flex-start", gap: 6, padding: "4px 6px" }}>
        <div style={{ height: 24, display: "flex", alignItems: "center" }}>
          <Sparkles />
        </div>
        <textarea
          ref={box}
          id="vba-ai-prompt"
          rows={1}
          value={shown}
          placeholder={review ? "Ask for more changes, or accept…" : `Ask AI to write or change this code… (${MOD}I)`}
          onChange={(e) => {
            disarm();
            setHearing("");
            setPrompt(e.target.value);
          }}
          onKeyDown={onKeyDown}
          onFocus={() => setPicked(selection())}
          style={{ flex: 1, minWidth: 0, resize: "none", overflow: "auto", border: 0, outline: "none", background: "transparent", color: "#1F2937", fontSize: 13, lineHeight: "22px", padding: "1px 2px", fontFamily: "inherit" }}
        />
        {lines && <span style={{ fontSize: 11, lineHeight: "24px", color: "#6B7280", flexShrink: 0 }} id="vba-ai-lines">{lines}</span>}
        <button
          type="button"
          id="vba-ai-mic"
          data-on={listening ? "1" : "0"}
          title={listening ? "Stop listening (Alt+V)" : "Say what you want (Alt+V). A pause sends it; say accept or reject to answer a change"}
          onClick={toggleMic}
          style={{ ...button, width: 26, padding: 0, display: "flex", alignItems: "center", justifyContent: "center", background: listening ? "#D93025" : "#FFFFFF", borderColor: listening ? "#D93025" : "#D5D9DE", animation: listening ? "vbaAiPulse 1.4s ease-in-out infinite" : "none" }}
        >
          <Mic on={listening} />
        </button>
        {!review && !writing && (
          <button type="button" id="vba-ai-send" style={{ ...primary, opacity: shown.trim() ? 1 : 0.5 }} disabled={!shown.trim()} onClick={() => send(shown)}>
            Generate
          </button>
        )}
        {writing && (
          <button type="button" id="vba-ai-stop" style={button} onClick={() => onAi({ type: "close" })}>
            Stop
          </button>
        )}
        {review && (
          <>
            {shown.trim() && (
              <button type="button" id="vba-ai-send" style={button} onClick={() => send(shown)}>
                Ask
              </button>
            )}
            <button type="button" id="vba-ai-accept" style={primary} title={`${MOD}Enter`} onClick={apply}>
              Accept
            </button>
            <button type="button" id="vba-ai-reject" style={button} title="Esc" onClick={() => onAi({ type: "close" })}>
              Reject
            </button>
          </>
        )}
        {!draft && (
          <button
            type="button"
            id="vba-ai-suggest"
            data-on={suggest ? "1" : "0"}
            title={suggest ? "Grey suggestions are on: Tab takes one. Click to turn them off" : "Grey suggestions are off. Click to turn them on"}
            onClick={() => onAi({ type: "toggle" })}
            style={{ ...button, color: suggest ? VIOLET : "#6B7280", borderColor: suggest ? "#C9BDF5" : "#D5D9DE" }}
          >
            Suggest
          </button>
        )}
      </div>
      {pending ? <div key={pending} id="vba-ai-countdown" style={{ height: 2, background: VIOLET, animation: `vbaAiSend ${AUTO_SEND_MS}ms linear forwards` }} /> : null}
      {note && <div id="vba-ai-note" style={{ fontSize: 11, color: "#6B7280", padding: "0 8px 4px 28px" }}>{note}</div>}
      {writing && (
        <div id="vba-ai-status" style={{ fontSize: 11, color: "#5B21B6", padding: "0 8px 4px 28px" }}>
          {words} {draft.model ? `(${draft.model})` : ""}
        </div>
      )}
      {review && (
        <div id="vba-ai-status" style={{ fontSize: 11, color: draft.problem ? "#B45309" : "#5B21B6", padding: "0 8px 4px 28px" }}>
          {draft.problem ? `It still has a compile error. ${draft.problem}` : "It compiles. Accept puts it in the module; Ctrl+Z in the code takes it back."}
          {draft.model ? ` (${draft.model})` : ""}
        </div>
      )}
      {status === "answer" && (
        <div id="vba-ai-answer" style={{ display: "flex", gap: 8, alignItems: "flex-start", fontSize: 13, color: "#1F2937", padding: "2px 8px 6px 28px", maxHeight: 160, overflow: "auto", whiteSpace: "pre-wrap" }}>
          <div style={{ flex: 1, minWidth: 0 }}>{draft.text}</div>
          <button type="button" id="vba-ai-close" style={button} onClick={() => onAi({ type: "close" })}>
            Close
          </button>
        </div>
      )}
      {status === "error" && (
        <div id="vba-ai-error" style={{ display: "flex", gap: 8, alignItems: "center", fontSize: 12, color: "#C5221F", padding: "0 8px 6px 28px" }}>
          <div style={{ flex: 1, minWidth: 0 }}>{draft.error}</div>
          <button type="button" id="vba-ai-close" style={button} onClick={() => onAi({ type: "close" })}>
            Close
          </button>
        </div>
      )}
    </div>
  );
});
