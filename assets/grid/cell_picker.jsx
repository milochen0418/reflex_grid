// Choosing cells from a sheet: a plain table of what the cells show (nothing here can edit them),
// where pressing and dragging selects, as on the sheet. The server answers with what it made of the
// selection (a click on a cell means the block of data around it), and `selection` follows.
//
// <CellPicker rows={[[text, …], …]} top={r0} left={c0} selection={[r0, c0, r1, c1]} onSelect={(s) => …} />
import React, { useEffect, useRef, useState } from "react";

function letters(c) {
  let s = "";
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

const HEAD = { position: "sticky", background: "#F3F4F6", color: "#4B5563", fontWeight: 500, border: "1px solid #D1D5DB", padding: "2px 6px", userSelect: "none" };

export function CellPicker({ rows, top, left, selection, onSelect }) {
  const [sel, setSel] = useState(null);
  const anchor = useRef(null);
  const latest = useRef(null);
  const report = useRef(onSelect);
  report.current = onSelect;
  useEffect(() => {
    setSel(Array.isArray(selection) && selection.length === 4 ? selection : null);
  }, [selection]);
  useEffect(() => {
    const up = () => {
      if (anchor.current && latest.current) report.current && report.current(latest.current);
      anchor.current = null;
    };
    window.addEventListener("mouseup", up);
    return () => window.removeEventListener("mouseup", up);
  }, []);
  const to = (r, c) => {
    const a = anchor.current;
    const s = [Math.min(a[0], r), Math.min(a[1], c), Math.max(a[0], r), Math.max(a[1], c)];
    latest.current = s;
    setSel(s);
  };
  const list = Array.isArray(rows) ? rows : [];
  const width = list.length ? list[0].length : 0;
  const inside = (r, c) => sel && r >= sel[0] && r <= sel[2] && c >= sel[1] && c <= sel[3];
  return (
    <table style={{ borderCollapse: "separate", borderSpacing: 0, fontSize: 12, color: "#111827", fontFamily: "Calibri, Carlito, 'Segoe UI', Arial, sans-serif" }}>
      <thead>
        <tr>
          <th style={{ ...HEAD, top: 0, left: 0, zIndex: 3 }} />
          {Array.from({ length: width }, (_, i) => (
            <th key={i} style={{ ...HEAD, top: 0, zIndex: 2 }}>{letters(left + i)}</th>
          ))}
        </tr>
      </thead>
      <tbody>
        {list.map((row, i) => (
          <tr key={i}>
            <th style={{ ...HEAD, left: 0, zIndex: 1 }}>{top + i + 1}</th>
            {row.map((text, j) => {
              const r = top + i;
              const c = left + j;
              return (
                <td
                  key={j}
                  data-cell={`${r},${c}`}
                  onMouseDown={(e) => {
                    e.preventDefault();
                    anchor.current = [r, c];
                    to(r, c);
                  }}
                  onMouseEnter={() => anchor.current && to(r, c)}
                  style={{ border: "1px solid #E5E7EB", padding: "1px 6px", minWidth: 56, maxWidth: 220, height: 20, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", cursor: "cell", userSelect: "none", background: inside(r, c) ? "#CFE8F7" : "#FFFFFF" }}
                >
                  {text}
                </td>
              );
            })}
          </tr>
        ))}
      </tbody>
    </table>
  );
}
