import React, { useState, useEffect, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  ChevronLeft,
  ChevronRight,
  Plus,
  X,
  Trash2,
  Stethoscope,
  Truck,
  Building2,
  Clock,
  User,
  MapPin,
  StickyNote,
  Loader2,
  CheckCircle2,
  Circle,
  ArrowRight,
  AlertCircle,
  Printer,
  CalendarRange,
  Search,
  FileSpreadsheet,
  Presentation,
  Copy,
} from "lucide-react";
import { supabase, supabaseConfigured } from "./supabaseClient";

const FONT_IMPORT_ID = "plantoes-fonts";
const TABLE = "entries";

// Converte uma linha da tabela `entries` (Supabase) no formato de registro
// que o resto do componente já usa (camelCase, sem colunas nulas do outro tipo).
function rowToEntry(row) {
  const base = {
    id: row.id,
    type: row.type,
    value: Number(row.value) || 0,
    color: row.color,
    pago: !!row.pago,
    obs: row.obs || "",
  };
  if (row.type === "plantao") {
    return {
      ...base,
      local: row.local || "",
      iH: row.i_h,
      iM: row.i_m,
      fH: row.f_h,
      fM: row.f_m,
      turno: row.turno,
    };
  }
  if (row.type === "evento") {
    return {
      ...base,
      local: row.local || "",
      empresa: row.empresa || "",
      iH: row.i_h,
      iM: row.i_m,
      fH: row.f_h,
      fM: row.f_m,
      turno: row.turno,
    };
  }
  return {
    ...base,
    empresa: row.empresa || "",
    paciente: row.paciente || "",
    origem: row.origem || "",
    destino: row.destino || "",
  };
}

// Converte um registro local + dia em uma linha pronta para gravar no Supabase.
function entryToRow(dayKey, entry) {
  return {
    id: entry.id,
    day_key: dayKey,
    type: entry.type,
    value: entry.value,
    color: entry.color,
    pago: !!entry.pago,
    local: entry.local ?? null,
    i_h: entry.iH ?? null,
    i_m: entry.iM ?? null,
    f_h: entry.fH ?? null,
    f_m: entry.fM ?? null,
    turno: entry.turno ?? null,
    empresa: entry.empresa ?? null,
    paciente: entry.paciente ?? null,
    origem: entry.origem ?? null,
    destino: entry.destino ?? null,
    obs: entry.obs ?? null,
  };
}

function rowsToEntries(rows) {
  const grouped = {};
  for (const row of rows) {
    const entry = rowToEntry(row);
    (grouped[row.day_key] ||= []).push(entry);
  }
  return grouped;
}

const MESES = [
  "Janeiro","Fevereiro","Março","Abril","Maio","Junho",
  "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro",
];
const DIAS_SEMANA = ["dom", "seg", "ter", "qua", "qui", "sex", "sáb"];

const PALETTE = [
  { id: "teal", label: "Teal", base: "#2D6E6E", bg: "#E4F0EF", text: "#215454" },
  { id: "terracota", label: "Terracota", base: "#B5541F", bg: "#F5E6DC", text: "#8C4118" },
  { id: "azul", label: "Azul", base: "#2A5DA8", bg: "#E3EAF6", text: "#1F4278" },
  { id: "verde", label: "Verde", base: "#2F8F52", bg: "#E2F2E7", text: "#206B3C" },
  { id: "roxo", label: "Roxo", base: "#7A4FB0", bg: "#EEE4F6", text: "#5C3A88" },
  { id: "rosa", label: "Rosa", base: "#C24B7C", bg: "#F7E4EE", text: "#96355D" },
  { id: "mostarda", label: "Mostarda", base: "#B8912B", bg: "#F6EFDD", text: "#8C6D1B" },
  { id: "vermelho", label: "Vermelho", base: "#C0392B", bg: "#F8E2DE", text: "#942E22" },
  { id: "cinza", label: "Cinza-azulado", base: "#556575", bg: "#E7EAEC", text: "#3E4B57" },
  { id: "marrom", label: "Marrom", base: "#7A5230", bg: "#EFE3D6", text: "#5C3D22" },
];
const paletteFor = (id) => PALETTE.find((p) => p.id === id) || PALETTE[0];
const defaultColorFor = (type) =>
  type === "plantao" ? "teal" : type === "evento" ? "azul" : "terracota";

function ColorScroll({ options, value, onChange }) {
  const ref = React.useRef(null);
  const timeoutRef = React.useRef(null);
  const itemW = 44;

  React.useEffect(() => {
    const idx = Math.max(0, options.findIndex((o) => o.id === value));
    if (ref.current) ref.current.scrollLeft = idx * itemW;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const settle = () => {
    if (!ref.current) return;
    const idx = Math.max(
      0,
      Math.min(options.length - 1, Math.round(ref.current.scrollLeft / itemW))
    );
    ref.current.scrollTo({ left: idx * itemW, behavior: "smooth" });
    onChange(options[idx].id);
  };

  const handleScroll = () => {
    if (timeoutRef.current) clearTimeout(timeoutRef.current);
    timeoutRef.current = setTimeout(settle, 110);
  };

  const pick = (idx) => {
    if (ref.current) ref.current.scrollTo({ left: idx * itemW, behavior: "smooth" });
    onChange(options[idx].id);
  };

  return (
    <div style={styles.colorScrollWrap}>
      <div style={styles.colorScrollBand} />
      <div
        ref={ref}
        className="plantoes-scroll-col"
        onScroll={handleScroll}
        style={styles.colorScrollRow}
      >
        <div style={{ width: itemW, flexShrink: 0 }} />
        {options.map((c, idx) => (
          <div key={c.id} onClick={() => pick(idx)} style={styles.colorScrollItemWrap}>
            <div
              style={{
                ...styles.colorScrollDot,
                background: c.base,
                ...(c.id === value ? styles.colorScrollDotActive : {}),
              }}
            />
          </div>
        ))}
        <div style={{ width: itemW, flexShrink: 0 }} />
      </div>
    </div>
  );
}

function ScrollColumn({ options, value, onChange }) {
  const ref = React.useRef(null);
  const timeoutRef = React.useRef(null);
  const itemH = 32;

  React.useEffect(() => {
    const idx = Math.max(0, options.indexOf(value));
    if (ref.current) ref.current.scrollTop = idx * itemH;
    // only position on mount
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const settle = () => {
    if (!ref.current) return;
    const idx = Math.max(
      0,
      Math.min(options.length - 1, Math.round(ref.current.scrollTop / itemH))
    );
    ref.current.scrollTo({ top: idx * itemH, behavior: "smooth" });
    onChange(options[idx]);
  };

  const handleScroll = () => {
    if (timeoutRef.current) clearTimeout(timeoutRef.current);
    timeoutRef.current = setTimeout(settle, 110);
  };

  const pick = (opt, idx) => {
    if (ref.current) ref.current.scrollTo({ top: idx * itemH, behavior: "smooth" });
    onChange(opt);
  };

  return (
    <div style={styles.scrollColWrap}>
      <div style={styles.scrollColBand} />
      <div
        ref={ref}
        className="plantoes-scroll-col"
        onScroll={handleScroll}
        style={styles.scrollCol}
      >
        <div style={{ height: itemH }} />
        {options.map((opt, idx) => (
          <div
            key={opt}
            onClick={() => pick(opt, idx)}
            style={{
              ...styles.scrollItem,
              ...(opt === value ? styles.scrollItemActive : {}),
            }}
          >
            {opt}
          </div>
        ))}
        <div style={{ height: itemH }} />
      </div>
    </div>
  );
}

const currency = (n) =>
  new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" }).format(
    Number(n) || 0
  );

const pad = (n) => String(n).padStart(2, "0");
const keyFor = (y, m, d) => `${y}-${pad(m + 1)}-${pad(d)}`;
const HOURS = Array.from({ length: 24 }, (_, i) => pad(i));
const MINUTES = Array.from({ length: 60 }, (_, i) => pad(i));
const parseBRL = (str) => {
  if (str == null) return NaN;
  let s = String(str).trim();
  if (!s) return NaN;
  // if there's a comma, treat it as the decimal separator and strip "." thousand marks
  if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
  return parseFloat(s);
};
const todayKey = () => {
  const t = new Date();
  return keyFor(t.getFullYear(), t.getMonth(), t.getDate());
};
const firstDayKey = (y, m) => keyFor(y, m, 1);
const lastDayKey = (y, m) => keyFor(y, m, new Date(y, m + 1, 0).getDate());
const formatShort = (dayKey) => {
  const [y, m, d] = dayKey.split("-").map(Number);
  return `${pad(d)}/${pad(m)}/${y}`;
};

const normText = (s) =>
  (s || "")
    .toString()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "");

const escapeHtml = (str) =>
  String(str == null ? "" : str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");

function buildReportHTML(data) {
  const rowsHtml = data.rows.length
    ? data.rows
        .map((e) => {
          const desc = e.type === "plantao" ? e.local : e.empresa;
          const horario =
            e.type === "plantao" && e.iH ? `${e.iH}:${e.iM} – ${e.fH}:${e.fM}` : "—";
          return `<tr>
            <td>${escapeHtml(formatShort(e.dayKey))}</td>
            <td>${e.type === "plantao" ? "Plantão" : "Remoção"}</td>
            <td>${escapeHtml(desc)}</td>
            <td>${escapeHtml(horario)}</td>
            <td>${e.pago ? "Pago" : "A receber"}</td>
            <td class="num">${escapeHtml(currency(e.value))}</td>
          </tr>`;
        })
        .join("")
    : `<tr><td colspan="6" class="empty">Nenhum registro neste período.</td></tr>`;

  return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8" />
<title>Resumo de plantões e remoções</title>
<style>
  * { box-sizing: border-box; }
  body {
    font-family: -apple-system, "Segoe UI", Inter, sans-serif;
    color: #1C2B39;
    padding: 32px;
    max-width: 820px;
    margin: 0 auto;
  }
  h1 { font-size: 20px; margin: 0 0 4px; }
  .subtitle { font-size: 13px; color: #5B6B75; margin: 0 0 18px; }
  header { border-bottom: 2px solid #1C2B39; padding-bottom: 10px; margin-bottom: 18px; }
  table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
  th, td { text-align: left; padding: 7px 8px; }
  th {
    border-bottom: 1px solid #1C2B39;
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.03em;
  }
  td { border-bottom: 1px solid #E0DDD3; font-family: "Courier New", monospace; }
  td.num { text-align: right; }
  td.empty { text-align: center; color: #5B6B75; font-family: inherit; padding: 24px 0; }
  .summary { margin-top: 20px; max-width: 320px; margin-left: auto; }
  .summary-row {
    display: flex;
    justify-content: space-between;
    padding: 5px 0;
    font-size: 13px;
    font-family: "Courier New", monospace;
  }
  .summary-total {
    border-top: 2px solid #1C2B39;
    margin-top: 4px;
    padding-top: 8px;
    font-weight: 700;
    font-size: 15px;
  }
  @media print {
    body { padding: 12px; }
  }
</style>
</head>
<body>
  <header>
    <h1>Resumo de plantões e remoções</h1>
    <p class="subtitle">Período: ${escapeHtml(formatShort(data.start))} a ${escapeHtml(
    formatShort(data.end)
  )}</p>
  </header>
  <table>
    <thead>
      <tr>
        <th>Data</th><th>Tipo</th><th>Descrição</th><th>Horário</th><th>Status</th>
        <th class="num">Valor</th>
      </tr>
    </thead>
    <tbody>${rowsHtml}</tbody>
  </table>
  <div class="summary">
    <div class="summary-row"><span>Plantões (${data.plantaoCount})</span><span>${escapeHtml(
    currency(data.plantaoSum)
  )}</span></div>
    <div class="summary-row"><span>Remoções (${data.remocaoCount})</span><span>${escapeHtml(
    currency(data.remocaoSum)
  )}</span></div>
    <div class="summary-row"><span>Recebido</span><span>${escapeHtml(
      currency(data.paidSum)
    )}</span></div>
    <div class="summary-row"><span>A receber</span><span>${escapeHtml(
      currency(data.pendingSum)
    )}</span></div>
    <div class="summary-row summary-total"><span>Total do período</span><span>${escapeHtml(
      currency(data.total)
    )}</span></div>
  </div>
</body>
</html>`;
}

const emptyForm = {
  type: "plantao",
  value: "",
  local: "",
  iH: "07",
  iM: "00",
  fH: "19",
  fM: "00",
  empresa: "",
  paciente: "",
  origem: "",
  destino: "",
  obs: "",
  color: null,
  pago: false,
};

const PDF_PAGE_W = 595;
const PDF_PAGE_H = 842;
const PDF_MARGIN = 40;
const PDF_ROW_H = 15;
const PDF_COL = { data: 40, nome: 110, horario: 300, status: 395, valor: 470 };
const PDF_DARK = [0.11, 0.169, 0.224];
const PDF_GRAY = [0.357, 0.42, 0.459];

const escPdfText = (s) =>
  String(s == null ? "" : s)
    .replace(/\\/g, "\\\\")
    .replace(/\(/g, "\\(")
    .replace(/\)/g, "\\)")
    .replace(/[\r\n\t]/g, " ");

const truncatePdf = (s, n) => {
  const str = s == null ? "" : String(s);
  return str.length > n ? str.slice(0, n - 3) + "..." : str;
};

function pdfTextOp(x, y, size, fontKey, text, color) {
  return `BT /${fontKey} ${size} Tf ${color[0]} ${color[1]} ${color[2]} rg 1 0 0 1 ${x} ${y} Tm (${escPdfText(
    text
  )}) Tj ET\n`;
}

function pdfLineOp(x1, y1, x2, y2) {
  return `0.11 0.169 0.224 RG 0.75 w ${x1} ${y1} m ${x2} ${y2} l S\n`;
}

function buildPdfPages(data) {
  const pages = [];
  let page = [];
  let y = PDF_PAGE_H - 50;

  const addTableHeader = () => {
    page.push(pdfTextOp(PDF_COL.data, y, 9, "F2", "Data", PDF_DARK));
    page.push(pdfTextOp(PDF_COL.nome, y, 9, "F2", "Nome", PDF_DARK));
    page.push(pdfTextOp(PDF_COL.horario, y, 9, "F2", "Horário", PDF_DARK));
    page.push(pdfTextOp(PDF_COL.status, y, 9, "F2", "Status", PDF_DARK));
    page.push(pdfTextOp(PDF_COL.valor, y, 9, "F2", "Valor", PDF_DARK));
    y -= 6;
    page.push(pdfLineOp(PDF_MARGIN, y, PDF_PAGE_W - PDF_MARGIN, y));
    y -= 14;
  };

  const startNewPage = () => {
    pages.push(page);
    page = [];
    y = PDF_PAGE_H - 50;
    addTableHeader();
  };

  page.push(
    pdfTextOp(PDF_MARGIN, y, 16, "F2", "Resumo de plantões e remoções", PDF_DARK)
  );
  y -= 18;
  page.push(
    pdfTextOp(
      PDF_MARGIN,
      y,
      10,
      "F1",
      `Período: ${formatShort(data.start)} a ${formatShort(data.end)}`,
      PDF_GRAY
    )
  );
  y -= 10;
  page.push(pdfLineOp(PDF_MARGIN, y, PDF_PAGE_W - PDF_MARGIN, y));
  y -= 20;
  addTableHeader();

  const rows = data.rows.length ? data.rows : [{ empty: true }];
  rows.forEach((e) => {
    if (y < 80) startNewPage();
    if (e.empty) {
      page.push(
        pdfTextOp(PDF_MARGIN, y, 10, "F1", "Nenhum registro neste período.", PDF_GRAY)
      );
    } else {
      const desc = e.type === "remocao" ? e.empresa || "" : e.local || "";
      const horario =
        e.type !== "remocao" && e.iH ? `${e.iH}:${e.iM}-${e.fH}:${e.fM}` : "-";
      page.push(pdfTextOp(PDF_COL.data, y, 9, "F1", formatShort(e.dayKey), PDF_DARK));
      page.push(
        pdfTextOp(PDF_COL.nome, y, 9, "F1", truncatePdf(desc, 28), PDF_DARK)
      );
      page.push(pdfTextOp(PDF_COL.horario, y, 9, "F1", horario, PDF_DARK));
      page.push(
        pdfTextOp(PDF_COL.status, y, 9, "F1", e.pago ? "Pago" : "A receber", PDF_DARK)
      );
      page.push(pdfTextOp(PDF_COL.valor, y, 9, "F1", currency(e.value), PDF_DARK));
    }
    y -= PDF_ROW_H;
  });

  if (y < 160) startNewPage();
  y -= 10;
  page.push(pdfLineOp(360, y, PDF_PAGE_W - PDF_MARGIN, y));
  y -= 20;
  const summaryLines = [
    [`Plantões (${data.plantaoCount})`, currency(data.plantaoSum)],
    [`Remoções (${data.remocaoCount})`, currency(data.remocaoSum)],
    [`Eventos (${data.eventoCount})`, currency(data.eventoSum)],
    ["Recebido", currency(data.paidSum)],
    ["A receber", currency(data.pendingSum)],
  ];
  summaryLines.forEach(([label, val]) => {
    page.push(pdfTextOp(360, y, 10, "F1", label, PDF_DARK));
    page.push(pdfTextOp(470, y, 10, "F1", val, PDF_DARK));
    y -= 16;
  });
  y -= 4;
  page.push(pdfLineOp(360, y, PDF_PAGE_W - PDF_MARGIN, y));
  y -= 16;
  page.push(pdfTextOp(360, y, 12, "F2", "Total do período", PDF_DARK));
  page.push(pdfTextOp(470, y, 12, "F2", currency(data.total), PDF_DARK));

  pages.push(page);
  return pages;
}

function assemblePdfFromPages(pagesOps) {
  const numPages = pagesOps.length;

  const catalogNum = 1;
  const pagesNum = 2;
  const pageNums = Array.from({ length: numPages }, (_, i) => 3 + i);
  const fontF1Num = 3 + numPages;
  const fontF2Num = fontF1Num + 1;
  const contentNums = Array.from({ length: numPages }, (_, i) => fontF2Num + 1 + i);
  const maxObjNum = fontF2Num + numPages;

  const offsets = [];
  let body = "%PDF-1.4\n" + String.fromCharCode(37, 226, 227, 207, 211) + "\n";

  const addObj = (num, content) => {
    offsets[num] = body.length;
    body += `${num} 0 obj\n${content}\nendobj\n`;
  };

  addObj(catalogNum, `<< /Type /Catalog /Pages ${pagesNum} 0 R >>`);

  const kids = pageNums.map((n) => `${n} 0 R`).join(" ");
  addObj(pagesNum, `<< /Type /Pages /Kids [${kids}] /Count ${numPages} >>`);

  pageNums.forEach((pnum, idx) => {
    const cnum = contentNums[idx];
    addObj(
      pnum,
      `<< /Type /Page /Parent ${pagesNum} 0 R /MediaBox [0 0 ${PDF_PAGE_W} ${PDF_PAGE_H}] /Resources << /Font << /F1 ${fontF1Num} 0 R /F2 ${fontF2Num} 0 R >> >> /Contents ${cnum} 0 R >>`
    );
  });

  addObj(
    fontF1Num,
    `<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>`
  );
  addObj(
    fontF2Num,
    `<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>`
  );

  pagesOps.forEach((ops, idx) => {
    const cnum = contentNums[idx];
    const stream = ops.join("");
    addObj(cnum, `<< /Length ${stream.length} >>\nstream\n${stream}endstream`);
  });

  const xrefStart = body.length;
  let xref = `xref\n0 ${maxObjNum + 1}\n0000000000 65535 f \n`;
  for (let i = 1; i <= maxObjNum; i++) {
    const off = offsets[i] || 0;
    xref += String(off).padStart(10, "0") + " 00000 n \n";
  }
  body += xref;
  body += `trailer\n<< /Size ${maxObjNum + 1} /Root ${catalogNum} 0 R >>\nstartxref\n${xrefStart}\n%%EOF`;

  const bytes = new Uint8Array(body.length);
  for (let i = 0; i < body.length; i++) bytes[i] = body.charCodeAt(i) & 0xff;
  return new Blob([bytes], { type: "application/pdf" });
}

function buildPdfBlob(data) {
  return assemblePdfFromPages(buildPdfPages(data));
}

const EXPORT_COL_LABEL = { data: "Data", nome: "Nome", pago: "Pago", valor: "Valor" };
const EXPORT_COL_ORDER = ["data", "nome", "pago", "valor"];

function getSearchColText(e, key) {
  if (key === "data") return formatShort(e.dayKey);
  if (key === "nome") {
    if (e.type === "remocao") return e.empresa || "remoção";
    if (e.type === "evento") return e.local || "evento";
    return e.local || "plantão";
  }
  if (key === "pago") return e.pago ? "Pago" : "A receber";
  if (key === "valor") return currency(e.value);
  return "";
}

function buildSearchPdfPages(results, colsSelected) {
  const activeCols = EXPORT_COL_ORDER.filter((k) => colsSelected[k]);
  const cols = activeCols.length ? activeCols : ["nome"];
  const usableWidth = PDF_PAGE_W - 2 * PDF_MARGIN;
  const colWidth = usableWidth / cols.length;
  const colX = cols.map((_, i) => PDF_MARGIN + i * colWidth);

  const pages = [];
  let page = [];
  let y = PDF_PAGE_H - 50;

  const addHeader = () => {
    cols.forEach((k, i) => {
      page.push(pdfTextOp(colX[i], y, 9, "F2", EXPORT_COL_LABEL[k], PDF_DARK));
    });
    y -= 6;
    page.push(pdfLineOp(PDF_MARGIN, y, PDF_PAGE_W - PDF_MARGIN, y));
    y -= 14;
  };

  const startNewPage = () => {
    pages.push(page);
    page = [];
    y = PDF_PAGE_H - 50;
    addHeader();
  };

  page.push(pdfTextOp(PDF_MARGIN, y, 16, "F2", "Busca de plantões e remoções", PDF_DARK));
  y -= 20;
  page.push(pdfLineOp(PDF_MARGIN, y, PDF_PAGE_W - PDF_MARGIN, y));
  y -= 20;
  addHeader();

  const rows = results.length ? results : [{ empty: true }];
  rows.forEach((e) => {
    if (y < 80) startNewPage();
    if (e.empty) {
      page.push(
        pdfTextOp(PDF_MARGIN, y, 10, "F1", "Nenhum registro encontrado.", PDF_GRAY)
      );
    } else {
      cols.forEach((k, i) => {
        const text = k === "nome" ? truncatePdf(getSearchColText(e, k), 30) : getSearchColText(e, k);
        page.push(pdfTextOp(colX[i], y, 9, "F1", text, PDF_DARK));
      });
    }
    y -= PDF_ROW_H;
  });

  if (colsSelected.valor && results.length) {
    if (y < 100) startNewPage();
    y -= 10;
    page.push(pdfLineOp(PDF_MARGIN, y, PDF_PAGE_W - PDF_MARGIN, y));
    y -= 18;
    const total = results.reduce((s, e) => s + (Number(e.value) || 0), 0);
    page.push(pdfTextOp(PDF_MARGIN, y, 11, "F2", `Total (${results.length})`, PDF_DARK));
    page.push(
      pdfTextOp(colX[colX.length - 1], y, 11, "F2", currency(total), PDF_DARK)
    );
  }

  pages.push(page);
  return pages;
}

function buildSearchPdfBlob(results, colsSelected) {
  return assemblePdfFromPages(buildSearchPdfPages(results, colsSelected));
}

export default function PlantoesApp() {
  const [entries, setEntries] = useState({});
  const [loading, setLoading] = useState(true);
  const [saveError, setSaveError] = useState(false);
  const [toast, setToast] = useState(null);
  const [cursor, setCursor] = useState(() => {
    const t = new Date();
    return { year: t.getFullYear(), month: t.getMonth() };
  });
  const [selectedDay, setSelectedDay] = useState(null); // "YYYY-MM-DD"
  const [draggedEntry, setDraggedEntry] = useState(null); // { dayKey, id }
  const [dragOverDay, setDragOverDay] = useState(null);
  const [modalOpen, setModalOpen] = useState(false);
  const [printModalOpen, setPrintModalOpen] = useState(false);
  const [printRange, setPrintRange] = useState(() => {
    const t = new Date();
    return {
      start: firstDayKey(t.getFullYear(), t.getMonth()),
      end: lastDayKey(t.getFullYear(), t.getMonth()),
    };
  });
  const [printDraft, setPrintDraft] = useState(null);
  const [pdfLoading, setPdfLoading] = useState(false);
  const [excelLoading, setExcelLoading] = useState(false);
  const [searchStart, setSearchStart] = useState("");
  const [searchEnd, setSearchEnd] = useState("");
  const [searchType, setSearchType] = useState("todos");
  const [searchPaid, setSearchPaid] = useState("todos");
  const [exportCols, setExportCols] = useState({
    data: true,
    nome: true,
    pago: true,
    valor: true,
  });
  const [searchPdfLoading, setSearchPdfLoading] = useState(false);
  const [searchExcelLoading, setSearchExcelLoading] = useState(false);
  const [searchName, setSearchName] = useState("");
  const [activeTab, setActiveTab] = useState("calendario");
  const [editingId, setEditingId] = useState(null);
  const [form, setForm] = useState(emptyForm);
  const [duplicating, setDuplicating] = useState(false);
  const [duplicateTargetDay, setDuplicateTargetDay] = useState("");

  // Inject fonts once
  useEffect(() => {
    if (document.getElementById(FONT_IMPORT_ID)) return;
    const link = document.createElement("link");
    link.id = FONT_IMPORT_ID;
    link.rel = "stylesheet";
    link.href =
      "https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=IBM+Plex+Mono:wght@500;600&family=Inter:wght@400;500;600&display=swap";
    document.head.appendChild(link);
  }, []);

  const showToast = useCallback((msg, type) => {
    setToast({ msg, type, key: Date.now() });
  }, []);

  useEffect(() => {
    if (!toast) return;
    const t = setTimeout(() => setToast(null), 2200);
    return () => clearTimeout(t);
  }, [toast]);

  // Carrega os registros do Supabase e assina atualizações em tempo real,
  // para que mudanças feitas em outro dispositivo apareçam sem recarregar a página.
  useEffect(() => {
    if (!supabaseConfigured) {
      setSaveError(true);
      setLoading(false);
      showToast("Supabase não configurado (veja .env.example)", "error");
      return;
    }

    let cancelled = false;

    (async () => {
      const { data, error } = await supabase.from(TABLE).select("*");
      if (cancelled) return;
      if (error) {
        setSaveError(true);
        showToast("Não foi possível carregar os dados", "error");
      } else {
        setEntries(rowsToEntries(data));
      }
      setLoading(false);
    })();

    const removeFromAllDays = (obj, id) => {
      for (const day of Object.keys(obj)) {
        const filtered = obj[day].filter((e) => e.id !== id);
        if (filtered.length) obj[day] = filtered;
        else delete obj[day];
      }
    };

    const channel = supabase
      .channel("entries-changes")
      .on(
        "postgres_changes",
        { event: "*", schema: "public", table: TABLE },
        (payload) => {
          setEntries((prev) => {
            const next = { ...prev };
            if (payload.eventType === "DELETE") {
              removeFromAllDays(next, payload.old.id);
            } else {
              const row = payload.new;
              removeFromAllDays(next, row.id);
              const entry = rowToEntry(row);
              next[row.day_key] = [...(next[row.day_key] || []), entry];
            }
            return next;
          });
        }
      )
      .subscribe();

    return () => {
      cancelled = true;
      supabase.removeChannel(channel);
    };
  }, [showToast]);

  const daysGrid = useMemo(() => {
    const { year, month } = cursor;
    const first = new Date(year, month, 1);
    const startOffset = first.getDay();
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const cells = [];
    for (let i = 0; i < startOffset; i++) cells.push(null);
    for (let d = 1; d <= daysInMonth; d++) cells.push(d);
    while (cells.length % 7 !== 0) cells.push(null);
    return cells;
  }, [cursor]);

  const monthTotals = useMemo(() => {
    const { year, month } = cursor;
    const prefix = `${year}-${pad(month + 1)}-`;
    let plantaoSum = 0,
      plantaoCount = 0,
      remocaoSum = 0,
      remocaoCount = 0,
      eventoSum = 0,
      eventoCount = 0,
      paidSum = 0,
      pendingSum = 0;
    Object.entries(entries).forEach(([key, list]) => {
      if (!key.startsWith(prefix)) return;
      list.forEach((e) => {
        const v = Number(e.value) || 0;
        if (e.type === "plantao") {
          plantaoSum += v;
          plantaoCount += 1;
        } else if (e.type === "evento") {
          eventoSum += v;
          eventoCount += 1;
        } else {
          remocaoSum += v;
          remocaoCount += 1;
        }
        if (e.pago) paidSum += v;
        else pendingSum += v;
      });
    });
    return {
      plantaoSum,
      plantaoCount,
      remocaoSum,
      remocaoCount,
      eventoSum,
      eventoCount,
      paidSum,
      pendingSum,
      total: plantaoSum + remocaoSum + eventoSum,
    };
  }, [entries, cursor]);

  const computePrintData = useCallback(
    (range) => {
      const { start, end } = range;
      const days = Object.keys(entries)
        .filter((k) => k >= start && k <= end)
        .sort();
      let plantaoSum = 0,
        plantaoCount = 0,
        remocaoSum = 0,
        remocaoCount = 0,
        eventoSum = 0,
        eventoCount = 0,
        paidSum = 0,
        pendingSum = 0;
      const rows = [];
      days.forEach((dayKey) => {
        entries[dayKey].forEach((e) => {
          const v = Number(e.value) || 0;
          if (e.type === "plantao") {
            plantaoSum += v;
            plantaoCount += 1;
          } else if (e.type === "evento") {
            eventoSum += v;
            eventoCount += 1;
          } else {
            remocaoSum += v;
            remocaoCount += 1;
          }
          if (e.pago) paidSum += v;
          else pendingSum += v;
          rows.push({ dayKey, ...e });
        });
      });
      return {
        start,
        end,
        rows,
        plantaoSum,
        plantaoCount,
        remocaoSum,
        remocaoCount,
        eventoSum,
        eventoCount,
        paidSum,
        pendingSum,
        total: plantaoSum + remocaoSum + eventoSum,
      };
    },
    [entries]
  );

  const searchResults = useMemo(() => {
    const nameQ = normText(searchName.trim());
    const results = [];
    Object.keys(entries)
      .filter((dayKey) => !searchStart || dayKey >= searchStart)
      .filter((dayKey) => !searchEnd || dayKey <= searchEnd)
      .sort()
      .forEach((dayKey) => {
        entries[dayKey].forEach((e) => {
          if (searchType !== "todos" && e.type !== searchType) return;
          if (searchPaid === "pago" && !e.pago) return;
          if (searchPaid === "pendente" && e.pago) return;
          if (nameQ) {
            const haystack = normText(
              [
                e.type === "remocao" ? e.empresa : e.local,
                e.empresa,
                e.paciente,
                e.origem,
                e.destino,
                e.obs,
              ]
                .filter(Boolean)
                .join(" ")
            );
            if (!haystack.includes(nameQ)) return;
          }
          results.push({ dayKey, ...e });
        });
      });
    return results.slice(0, 300);
  }, [searchStart, searchEnd, searchType, searchPaid, searchName, entries]);

  const goToSearchResult = (result) => {
    const [y, m] = result.dayKey.split("-").map(Number);
    setCursor({ year: y, month: m - 1 });
    openEditModal(result.dayKey, result);
    setActiveTab("calendario");
  };

  const toggleEntryPago = async (result) => {
    if (!supabaseConfigured) {
      showToast("Supabase não configurado", "error");
      return;
    }
    const novoPago = !result.pago;
    setEntries((prev) => {
      const list = prev[result.dayKey] || [];
      const idx = list.findIndex((e) => e.id === result.id);
      if (idx === -1) return prev;
      const updated = [...list];
      updated[idx] = { ...updated[idx], pago: novoPago };
      return { ...prev, [result.dayKey]: updated };
    });
    const { error } = await supabase.from(TABLE).update({ pago: novoPago }).eq("id", result.id);
    if (error) {
      setSaveError(true);
      showToast("Não foi possível salvar", "error");
    } else {
      showToast(novoPago ? "Marcado como pago" : "Marcado como a receber", "success");
    }
  };

  const toggleExportCol = (key) => {
    setExportCols((prev) => ({ ...prev, [key]: !prev[key] }));
  };

  const generateSearchPDF = () => {
    if (!Object.values(exportCols).some(Boolean)) {
      showToast("Escolha ao menos uma coluna", "error");
      return;
    }
    setSearchPdfLoading(true);
    try {
      const blob = buildSearchPdfBlob(searchResults, exportCols);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `busca-plantoes_${todayKey()}.pdf`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 30000);
      showToast("PDF gerado e baixado", "success");
    } catch (err) {
      showToast("Não foi possível gerar o PDF", "error");
    } finally {
      setSearchPdfLoading(false);
    }
  };

  const generateSearchExcel = () => {
    if (!Object.values(exportCols).some(Boolean)) {
      showToast("Escolha ao menos uma coluna", "error");
      return;
    }
    setSearchExcelLoading(true);
    try {
      const activeCols = EXPORT_COL_ORDER.filter((k) => exportCols[k]);
      const header = activeCols.map((k) => EXPORT_COL_LABEL[k]);
      const rows = searchResults.map((e) =>
        activeCols.map((k) =>
          k === "valor" ? Number(e.value) || 0 : getSearchColText(e, k)
        )
      );
      const sheetData = [header, ...rows];
      if (exportCols.valor && searchResults.length) {
        const total = searchResults.reduce((s, e) => s + (Number(e.value) || 0), 0);
        const totalRow = activeCols.map((k, i) =>
          i === 0 ? `Total (${searchResults.length})` : k === "valor" ? total : ""
        );
        sheetData.push([]);
        sheetData.push(totalRow);
      }
      const ws = XLSX.utils.aoa_to_sheet(sheetData);
      ws["!cols"] = activeCols.map(() => ({ wch: 20 }));
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Busca");
      const wbArray = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([wbArray], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `busca-plantoes_${todayKey()}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 30000);
      showToast("Excel gerado", "success");
    } catch (err) {
      showToast("Não foi possível gerar o Excel", "error");
    } finally {
      setSearchExcelLoading(false);
    }
  };

  const printData = useMemo(
    () => computePrintData(printRange),
    [computePrintData, printRange]
  );

  const openPrintModal = () => {
    setPrintDraft({ start: printRange.start, end: printRange.end });
    setPrintModalOpen(true);
  };

  const applyPreset = (preset) => {
    const t = new Date();
    if (preset === "esteMes") {
      setPrintDraft({
        start: firstDayKey(t.getFullYear(), t.getMonth()),
        end: lastDayKey(t.getFullYear(), t.getMonth()),
      });
    } else if (preset === "mesPassado") {
      let y = t.getFullYear(),
        m = t.getMonth() - 1;
      if (m < 0) {
        m = 11;
        y -= 1;
      }
      setPrintDraft({ start: firstDayKey(y, m), end: lastDayKey(y, m) });
    } else if (preset === "esteAno") {
      setPrintDraft({
        start: keyFor(t.getFullYear(), 0, 1),
        end: keyFor(t.getFullYear(), 11, 31),
      });
    }
  };

  const generatePDF = (range) => {
    setPdfLoading(true);
    try {
      const data = computePrintData(range);
      const blob = buildPdfBlob(data);
      const url = URL.createObjectURL(blob);
      const filename = `resumo-plantoes_${data.start}_a_${data.end}.pdf`;
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 30000);
      showToast("PDF gerado e baixado", "success");
    } catch (err) {
      showToast("Não foi possível gerar o PDF", "error");
    } finally {
      setPdfLoading(false);
    }
  };

  const generateExcel = async (range) => {
    setExcelLoading(true);
    try {
      const data = computePrintData(range);
      const header = ["Data", "Tipo", "Descrição", "Horário", "Status", "Valor (R$)"];
      const tipoLabel = { plantao: "Plantão", remocao: "Remoção", evento: "Evento" };
      const rows = data.rows.map((e) => [
        formatShort(e.dayKey),
        tipoLabel[e.type] || e.type,
        e.type === "remocao" ? e.empresa || "" : e.local || "",
        e.type !== "remocao" && e.iH ? `${e.iH}:${e.iM} – ${e.fH}:${e.fM}` : "",
        e.pago ? "Pago" : "A receber",
        Number(e.value) || 0,
      ]);
      const sheetData = [
        [`Resumo de plantões, remoções e eventos — ${formatShort(data.start)} a ${formatShort(data.end)}`],
        [],
        header,
        ...rows,
        [],
        [`Plantões (${data.plantaoCount})`, data.plantaoSum],
        [`Remoções (${data.remocaoCount})`, data.remocaoSum],
        [`Eventos (${data.eventoCount})`, data.eventoSum],
        ["Recebido", data.paidSum],
        ["A receber", data.pendingSum],
        ["Total do período", data.total],
      ];

      const ws = XLSX.utils.aoa_to_sheet(sheetData);
      ws["!cols"] = [
        { wch: 14 },
        { wch: 10 },
        { wch: 30 },
        { wch: 16 },
        { wch: 12 },
        { wch: 14 },
      ];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Resumo");
      const wbArray = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([wbArray], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const filename = `resumo-plantoes_${data.start}_a_${data.end}.xlsx`;

      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 30000);

      showToast("Excel gerado", "success");
    } catch (err) {
      showToast("Não foi possível gerar o Excel", "error");
    } finally {
      setExcelLoading(false);
    }
  };

  const confirmPrint = (format) => {
    if (!printDraft || !printDraft.start || !printDraft.end) return;
    if (printDraft.start > printDraft.end) {
      showToast("A data inicial deve ser antes da final", "error");
      return;
    }
    setPrintRange(printDraft);
    setPrintModalOpen(false);
    if (format === "excel") generateExcel(printDraft);
    else generatePDF(printDraft);
  };

  const goMonth = (delta) => {
    setCursor((c) => {
      let month = c.month + delta;
      let year = c.year;
      if (month < 0) {
        month = 11;
        year -= 1;
      } else if (month > 11) {
        month = 0;
        year += 1;
      }
      return { year, month };
    });
  };

  const goToday = () => {
    const t = new Date();
    setCursor({ year: t.getFullYear(), month: t.getMonth() });
  };

  const openAddModal = (dayKey) => {
    setSelectedDay(dayKey);
    setEditingId(null);
    setForm(emptyForm);
    setDuplicating(false);
    setModalOpen(true);
  };

  const openEditModal = (dayKey, entry) => {
    setSelectedDay(dayKey);
    setEditingId(entry.id);
    setForm({ ...emptyForm, ...entry, value: String(entry.value) });
    setDuplicating(false);
    setModalOpen(true);
  };

  const closeModal = () => {
    setModalOpen(false);
    setEditingId(null);
    setForm(emptyForm);
    setDuplicating(false);
  };

  const openDuplicate = () => {
    setDuplicateTargetDay(selectedDay);
    setDuplicating(true);
  };

  const confirmDuplicate = async () => {
    if (!supabaseConfigured) {
      showToast("Supabase não configurado", "error");
      return;
    }
    if (!duplicateTargetDay) {
      showToast("Escolha uma data de destino", "error");
      return;
    }
    const original = (entries[selectedDay] || []).find((e) => e.id === editingId);
    if (!original) return;
    const copy = { ...original, id: `${Date.now()}-${Math.random().toString(36).slice(2, 7)}` };

    setEntries((prev) => {
      const list = prev[duplicateTargetDay] ? [...prev[duplicateTargetDay]] : [];
      return { ...prev, [duplicateTargetDay]: [...list, copy] };
    });
    const targetDay = duplicateTargetDay;
    closeModal();

    const { error } = await supabase.from(TABLE).insert(entryToRow(targetDay, copy));
    if (error) {
      setSaveError(true);
      showToast("Não foi possível duplicar", "error");
    } else {
      showToast(`Duplicado para ${formatShort(targetDay)}`, "success");
    }
  };

  const handleSave = async () => {
    if (!selectedDay) return;
    if (!supabaseConfigured) {
      showToast("Supabase não configurado", "error");
      return;
    }
    const value = parseBRL(form.value);
    if (!value || value <= 0) {
      showToast("Informe um valor válido em R$", "error");
      return;
    }
    if (form.type === "plantao" && !form.local.trim()) {
      showToast("Informe o local / hospital", "error");
      return;
    }
    if (form.type === "remocao" && !form.empresa.trim()) {
      showToast("Informe a empresa da remoção", "error");
      return;
    }
    if (form.type === "evento" && !form.local.trim()) {
      showToast("Informe o nome do evento", "error");
      return;
    }
    if (form.type === "evento" && !form.empresa.trim()) {
      showToast("Informe a empresa do evento", "error");
      return;
    }

    const record = {
      id: editingId || `${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
      type: form.type,
      value,
      color: form.color || defaultColorFor(form.type),
      pago: !!form.pago,
      ...(form.type === "plantao"
        ? {
            local: form.local,
            iH: form.iH,
            iM: form.iM,
            fH: form.fH,
            fM: form.fM,
            turno: `${form.iH}:${form.iM} – ${form.fH}:${form.fM}`,
            obs: form.obs,
          }
        : form.type === "evento"
        ? {
            local: form.local,
            empresa: form.empresa,
            iH: form.iH,
            iM: form.iM,
            fH: form.fH,
            fM: form.fM,
            turno: `${form.iH}:${form.iM} – ${form.fH}:${form.fM}`,
            obs: form.obs,
          }
        : {
            empresa: form.empresa,
            paciente: form.paciente,
            origem: form.origem,
            destino: form.destino,
            obs: form.obs,
          }),
    };

    setEntries((prev) => {
      const list = prev[selectedDay] ? [...prev[selectedDay]] : [];
      if (editingId) {
        const idx = list.findIndex((e) => e.id === editingId);
        if (idx >= 0) list[idx] = record;
        else list.push(record);
      } else {
        list.push(record);
      }
      return { ...prev, [selectedDay]: list };
    });
    closeModal();

    const { error } = await supabase.from(TABLE).upsert(entryToRow(selectedDay, record));
    if (error) {
      setSaveError(true);
      showToast("Não foi possível salvar", "error");
    } else {
      showToast(editingId ? "Alterações salvas" : "Registro salvo", "success");
    }
  };

  const handleDelete = async () => {
    if (!selectedDay || !editingId) return;
    if (!supabaseConfigured) {
      showToast("Supabase não configurado", "error");
      return;
    }
    const idToDelete = editingId;
    const day = selectedDay;
    setEntries((prev) => {
      const list = (prev[day] || []).filter((e) => e.id !== idToDelete);
      const next = { ...prev };
      if (list.length) next[day] = list;
      else delete next[day];
      return next;
    });
    closeModal();

    const { error } = await supabase.from(TABLE).delete().eq("id", idToDelete);
    if (error) {
      setSaveError(true);
      showToast("Não foi possível excluir", "error");
    } else {
      showToast("Registro excluído", "success");
    }
  };

  const moveEntry = async (fromDay, id, toDay) => {
    if (!fromDay || !toDay || fromDay === toDay) return;
    if (!supabaseConfigured) {
      showToast("Supabase não configurado", "error");
      return;
    }
    setEntries((prev) => {
      const fromList = prev[fromDay] || [];
      const entry = fromList.find((e) => e.id === id);
      if (!entry) return prev;
      const newFromList = fromList.filter((e) => e.id !== id);
      const toList = prev[toDay] ? [...prev[toDay]] : [];
      toList.push(entry);
      const next = { ...prev };
      if (newFromList.length) next[fromDay] = newFromList;
      else delete next[fromDay];
      next[toDay] = toList;
      return next;
    });

    const { error } = await supabase.from(TABLE).update({ day_key: toDay }).eq("id", id);
    if (error) {
      setSaveError(true);
      showToast("Não foi possível salvar", "error");
    } else {
      showToast("Movido para outra data", "success");
    }
  };

  const isValid =
    form.value &&
    parseBRL(form.value) > 0 &&
    (form.type === "plantao"
      ? form.local.trim()
      : form.type === "evento"
      ? form.local.trim() && form.empresa.trim()
      : form.empresa.trim());

  return (
    <>
    <div style={styles.app} className="app-main">
      <style>{globalCss}</style>

      <header style={styles.header}>
        <div style={styles.headerTop}>
          <div style={styles.brand}>
            <div style={styles.brandMark}>
              <Stethoscope size={18} color="#F7F5F0" strokeWidth={2.2} />
            </div>
            <div>
              <div style={styles.brandTitle}>Plantões</div>
              <div style={styles.brandSub}>controle de escala &amp; financeiro</div>
            </div>
          </div>
          {saveError && (
            <div style={styles.saveWarning}>não foi possível salvar</div>
          )}
        </div>

        <div style={styles.tabRow}>
          <button
            onClick={() => setActiveTab("calendario")}
            className="pill-btn"
            style={{
              ...styles.tabBtn,
              ...(activeTab === "calendario" ? styles.tabBtnActive : {}),
            }}
          >
            <CalendarRange size={14} />
            calendário
          </button>
          <button
            onClick={() => setActiveTab("buscar")}
            className="pill-btn"
            style={{
              ...styles.tabBtn,
              ...(activeTab === "buscar" ? styles.tabBtnActive : {}),
            }}
          >
            <Search size={14} />
            buscar
          </button>
        </div>

        {activeTab === "buscar" && (
        <div style={styles.searchWrap}>
          <div style={styles.searchFiltersRow}>
            <div style={styles.searchInputWrap}>
              <CalendarRange size={15} color="#5B6B75" style={{ flexShrink: 0 }} />
              <input
                type="date"
                style={styles.searchInput}
                value={searchStart}
                onChange={(e) => setSearchStart(e.target.value)}
              />
            </div>
            <span style={styles.searchDateSep}>até</span>
            <div style={styles.searchInputWrap}>
              <input
                type="date"
                style={styles.searchInput}
                value={searchEnd}
                onChange={(e) => setSearchEnd(e.target.value)}
              />
              {(searchStart || searchEnd) && (
                <button
                  className="btn-icon"
                  style={styles.searchClearBtn}
                  onClick={() => {
                    setSearchStart("");
                    setSearchEnd("");
                  }}
                  aria-label="Limpar período"
                >
                  <X size={14} />
                </button>
              )}
            </div>
            <div style={{ ...styles.searchInputWrap, flex: 1 }}>
              <Search size={15} color="#5B6B75" style={{ flexShrink: 0 }} />
              <input
                style={styles.searchInput}
                value={searchName}
                onChange={(e) => setSearchName(e.target.value)}
                placeholder="Nome do plantão ou remoção…"
              />
              {searchName && (
                <button
                  className="btn-icon"
                  style={styles.searchClearBtn}
                  onClick={() => setSearchName("")}
                  aria-label="Limpar nome"
                >
                  <X size={14} />
                </button>
              )}
            </div>
          </div>

          <div style={styles.searchTypeRow}>
            {[
              { id: "todos", label: "todos" },
              { id: "plantao", label: "plantões" },
              { id: "remocao", label: "remoções" },
              { id: "evento", label: "eventos" },
            ].map((t) => (
              <button
                key={t.id}
                onClick={() => setSearchType(t.id)}
                className="pill-btn"
                style={{
                  ...styles.searchTypeBtn,
                  ...(searchType === t.id ? styles.searchTypeBtnActive : {}),
                }}
              >
                {t.label}
              </button>
            ))}
            <span style={styles.searchTypeDivider} />
            {[
              { id: "todos", label: "todos" },
              { id: "pago", label: "pago" },
              { id: "pendente", label: "a receber" },
            ].map((t) => (
              <button
                key={t.id}
                onClick={() => setSearchPaid(t.id)}
                className="pill-btn"
                style={{
                  ...styles.searchTypeBtn,
                  ...(searchPaid === t.id ? styles.searchTypeBtnActive : {}),
                }}
              >
                {t.label}
              </button>
            ))}
          </div>

          <div style={styles.exportPanel}>
            <div style={styles.exportColsRow}>
              {EXPORT_COL_ORDER.map((k) => (
                <label key={k} style={styles.exportColCheckbox}>
                  <input
                    type="checkbox"
                    checked={!!exportCols[k]}
                    onChange={() => toggleExportCol(k)}
                    style={styles.checkboxInput}
                  />
                  {EXPORT_COL_LABEL[k]}
                </label>
              ))}
            </div>
            <div style={styles.exportButtonsRow}>
              <button
                className="btn-lift"
                style={{ ...styles.excelBtn, opacity: searchExcelLoading ? 0.7 : 1 }}
                onClick={generateSearchExcel}
                disabled={searchExcelLoading || searchPdfLoading}
              >
                {searchExcelLoading ? (
                  <Loader2 size={14} className="spin" style={{ marginRight: 6 }} />
                ) : (
                  <FileSpreadsheet size={14} style={{ marginRight: 6 }} />
                )}
                {searchExcelLoading ? "gerando…" : "gerar Excel"}
              </button>
              <button
                className="btn-lift"
                style={{ ...styles.saveBtn, opacity: searchPdfLoading ? 0.7 : 1 }}
                onClick={generateSearchPDF}
                disabled={searchPdfLoading || searchExcelLoading}
              >
                {searchPdfLoading ? (
                  <Loader2 size={14} className="spin" style={{ marginRight: 6 }} />
                ) : (
                  <Printer size={14} style={{ marginRight: 6 }} />
                )}
                {searchPdfLoading ? "gerando…" : "gerar PDF"}
              </button>
            </div>
          </div>

          <div style={styles.searchDropdown}>
              {searchResults.length === 0 ? (
                <div style={styles.searchEmpty}>Nada encontrado.</div>
              ) : (
                <>
                  <div style={styles.searchScrollArea}>
                    <table style={styles.searchTable}>
                      <thead>
                        <tr>
                          <th style={styles.searchTh}>Data</th>
                          <th style={styles.searchTh}>Nome</th>
                          <th style={{ ...styles.searchTh, textAlign: "center" }}>
                            Pago
                          </th>
                          <th style={{ ...styles.searchTh, textAlign: "right" }}>
                            Valor
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {searchResults.map((r) => (
                          <tr
                            key={r.id}
                            style={styles.searchRow}
                            onClick={() => goToSearchResult(r)}
                          >
                            <td style={styles.searchTd}>{formatShort(r.dayKey)}</td>
                            <td style={styles.searchTdName}>
                              {r.type === "remocao"
                                ? r.empresa || "remoção"
                                : r.local || (r.type === "evento" ? "evento" : "plantão")}
                            </td>
                            <td style={{ ...styles.searchTd, textAlign: "center" }}>
                              <button
                                className="btn-icon"
                                style={styles.searchPagoBtn}
                                onClick={(evt) => {
                                  evt.stopPropagation();
                                  toggleEntryPago(r);
                                }}
                                title={r.pago ? "Marcar como a receber" : "Marcar como pago"}
                              >
                                {r.pago ? (
                                  <CheckCircle2 size={15} color="#206B3C" />
                                ) : (
                                  <Circle size={15} color="#8C6D1B" />
                                )}
                              </button>
                            </td>
                            <td
                              style={{
                                ...styles.searchTd,
                                textAlign: "right",
                                ...(!r.pago ? styles.searchValuePending : {}),
                              }}
                            >
                              {currency(r.value)}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <div style={styles.searchTotalsWrap}>
                    {(() => {
                      const totalGroups = [
                        { id: "plantao", label: "Total plantões", style: styles.searchTotalPlantao },
                        { id: "remocao", label: "Total remoções", style: null },
                        { id: "evento", label: "Total eventos", style: styles.searchTotalEvento },
                      ].filter((g) => searchType === "todos" || searchType === g.id);
                      return totalGroups.map((g, idx) => {
                        const items = searchResults.filter((r) => r.type === g.id);
                        const isLast = idx === totalGroups.length - 1;
                        return (
                          <div
                            key={g.id}
                            style={{
                              ...styles.searchTotalRow,
                              ...(g.style || {}),
                              ...(isLast ? { borderRadius: "0 0 10px 10px" } : {}),
                            }}
                          >
                            <span>
                              {g.label} ({items.length})
                            </span>
                            <span>
                              {currency(items.reduce((s, r) => s + (Number(r.value) || 0), 0))}
                            </span>
                          </div>
                        );
                      });
                    })()}
                  </div>
                </>
              )}
            </div>
        </div>
        )}

        {activeTab === "calendario" && (
        <>
        <div style={styles.monthNav}>
          <button className="btn-icon" style={styles.navBtn} onClick={() => goMonth(-1)} aria-label="Mês anterior">
            <ChevronLeft size={18} />
          </button>
          <div style={styles.monthLabel}>
            {MESES[cursor.month]} <span style={styles.year}>{cursor.year}</span>
          </div>
          <button className="btn-icon" style={styles.navBtn} onClick={() => goMonth(1)} aria-label="Próximo mês">
            <ChevronRight size={18} />
          </button>
          <button className="btn-lift" style={styles.todayBtn} onClick={goToday}>
            hoje
          </button>
          <button className="btn-lift" style={styles.printBtn} onClick={openPrintModal}>
            <Printer size={13} />
            gerar PDF
          </button>
        </div>

        <div style={styles.summaryBar}>
          <SummaryChip
            icon={<Stethoscope size={14} />}
            label="plantões"
            count={monthTotals.plantaoCount}
            value={monthTotals.plantaoSum}
            color="#2D6E6E"
            bg="#E4F0EF"
          />
          <SummaryChip
            icon={<Truck size={14} />}
            label="remoções"
            count={monthTotals.remocaoCount}
            value={monthTotals.remocaoSum}
            color="#B5541F"
            bg="#F5E6DC"
          />
          <SummaryChip
            icon={<Presentation size={14} />}
            label="eventos"
            count={monthTotals.eventoCount}
            value={monthTotals.eventoSum}
            color="#1F4278"
            bg="#E3EAF6"
          />
          <div style={styles.totalChip}>
            <div style={styles.totalLabel}>total do mês</div>
            <div style={styles.totalValue}>{currency(monthTotals.total)}</div>
          </div>
        </div>

        <div style={styles.financeRow}>
          <SummaryChip
            icon={<CheckCircle2 size={14} />}
            label="recebido"
            value={monthTotals.paidSum}
            color="#206B3C"
            bg="#E2F2E7"
          />
          <SummaryChip
            icon={<Circle size={14} />}
            label="a receber"
            value={monthTotals.pendingSum}
            color="#8C6D1B"
            bg="#F6EFDD"
          />
        </div>
        </>
        )}
      </header>

      {activeTab === "calendario" && (loading ? (
        <div style={styles.loadingBox}>
          <Loader2 size={22} className="spin" />
          <span>carregando agenda…</span>
        </div>
      ) : (
        <>
          <div style={styles.dragHint}>
            Arraste um plantão ou remoção para outro dia para reagendar.
          </div>
          <div style={styles.weekHeader}>
            {DIAS_SEMANA.map((d) => (
              <div key={d} style={styles.weekHeaderCell}>
                {d}
              </div>
            ))}
          </div>

          <div style={styles.grid} className="cal-grid">
            {daysGrid.map((d, i) => {
              if (d === null) return <div key={i} style={styles.emptyCell} />;
              const dayKey = keyFor(cursor.year, cursor.month, d);
              const list = entries[dayKey] || [];
              const isToday = dayKey === todayKey();
              const isDragOver = dragOverDay === dayKey;
              return (
                <div
                  key={i}
                  className="cal-cell"
                  style={{
                    ...styles.cell,
                    ...(isToday ? styles.cellToday : {}),
                    ...(isDragOver ? styles.cellDragOver : {}),
                  }}
                  onDragOver={(evt) => {
                    evt.preventDefault();
                    if (dragOverDay !== dayKey) setDragOverDay(dayKey);
                  }}
                  onDragLeave={() => {
                    setDragOverDay((prev) => (prev === dayKey ? null : prev));
                  }}
                  onDrop={(evt) => {
                    evt.preventDefault();
                    if (draggedEntry) {
                      moveEntry(draggedEntry.dayKey, draggedEntry.id, dayKey);
                    }
                    setDraggedEntry(null);
                    setDragOverDay(null);
                  }}
                >
                  <div style={styles.cellHeader}>
                    <span style={{ ...styles.cellNum, ...(isToday ? styles.cellNumToday : {}) }}>
                      {d}
                    </span>
                    <button
                      className="btn-icon"
                      style={styles.addBtn}
                      onClick={() => openAddModal(dayKey)}
                      aria-label={`Adicionar em ${d}`}
                    >
                      <Plus size={12} strokeWidth={2.5} />
                    </button>
                  </div>
                  <div style={styles.chipStack}>
                    {list.map((e) => {
                      const c = paletteFor(e.color || defaultColorFor(e.type));
                      const isDragging =
                        draggedEntry && draggedEntry.id === e.id && draggedEntry.dayKey === dayKey;
                      return (
                      <button
                        key={e.id}
                        draggable
                        className="chip-lift"
                        onDragStart={(evt) => {
                          setDraggedEntry({ dayKey, id: e.id });
                          evt.dataTransfer.effectAllowed = "move";
                          evt.dataTransfer.setData("text/plain", e.id);
                        }}
                        onDragEnd={() => {
                          setDraggedEntry(null);
                          setDragOverDay(null);
                        }}
                        style={{
                          ...styles.entryChip,
                          background: c.bg,
                          color: c.text,
                          ...(isDragging ? styles.entryChipDragging : {}),
                        }}
                        onClick={() => openEditModal(dayKey, e)}
                        title={e.type === "remocao" ? e.empresa : e.local}
                      >
                        {e.type === "plantao" ? (
                          <Stethoscope size={10} />
                        ) : e.type === "evento" ? (
                          <Presentation size={10} />
                        ) : (
                          <Truck size={10} />
                        )}
                        <span style={styles.chipText}>
                          {e.type === "remocao"
                            ? e.empresa || "remoção"
                            : `${e.iH ? `${e.iH}:${e.iM} · ` : ""}${e.local || (e.type === "evento" ? "evento" : "plantão")}`}
                        </span>
                        {e.pago ? (
                          <CheckCircle2 size={10} style={{ flexShrink: 0 }} />
                        ) : (
                          <Circle size={10} style={{ flexShrink: 0, opacity: 0.5 }} />
                        )}
                        <span style={styles.chipValue}>{currency(e.value)}</span>
                      </button>
                      );
                    })}
                  </div>
                </div>
              );
            })}
          </div>
        </>
      ))}

      {modalOpen && (
        <div style={styles.overlay} onClick={closeModal}>
          <div style={styles.modal} onClick={(e) => e.stopPropagation()}>
            <div style={styles.modalHeader}>
              <div style={styles.modalTitle}>
                {editingId ? "editar registro" : "novo registro"}
                <div style={styles.modalDate}>{formatFullDate(selectedDay)}</div>
              </div>
              <button className="btn-icon" style={styles.closeBtn} onClick={closeModal}>
                <X size={18} />
              </button>
            </div>

            <div style={styles.typeToggle}>
              <button
                className="btn-lift"
                style={{
                  ...styles.typeBtn,
                  ...(form.type === "plantao" ? styles.typeBtnActivePlantao : {}),
                }}
                onClick={() => setForm((f) => ({ ...f, type: "plantao" }))}
              >
                <Stethoscope size={15} />
                Plantão
              </button>
              <button
                className="btn-lift"
                style={{
                  ...styles.typeBtn,
                  ...(form.type === "remocao" ? styles.typeBtnActiveRemocao : {}),
                }}
                onClick={() => setForm((f) => ({ ...f, type: "remocao" }))}
              >
                <Truck size={15} />
                Remoção
              </button>
              <button
                className="btn-lift"
                style={{
                  ...styles.typeBtn,
                  ...(form.type === "evento" ? styles.typeBtnActiveEvento : {}),
                }}
                onClick={() => setForm((f) => ({ ...f, type: "evento" }))}
              >
                <Presentation size={15} />
                Evento
              </button>
            </div>

            <div style={styles.field}>
              <div style={styles.fieldLabel}>
                <span>Cor da etiqueta</span>
              </div>
              <ColorScroll
                options={PALETTE}
                value={form.color || defaultColorFor(form.type)}
                onChange={(id) => setForm((f) => ({ ...f, color: id }))}
              />
            </div>

            <div style={styles.formBody}>
              {form.type === "plantao" ? (
                <>
                  <div style={styles.field}>
                    <div style={styles.fieldLabel}>
                      <Clock size={14} />
                      <span>Horário do plantão</span>
                    </div>
                    <div style={styles.timeRangeRow}>
                      <div style={styles.timeGroup}>
                        <div style={styles.timeGroupLabel}>início</div>
                        <div style={styles.timeColsRow}>
                          <ScrollColumn
                            options={HOURS}
                            value={form.iH}
                            onChange={(v) => setForm((f) => ({ ...f, iH: v }))}
                          />
                          <span style={styles.timeColon}>:</span>
                          <ScrollColumn
                            options={MINUTES}
                            value={form.iM}
                            onChange={(v) => setForm((f) => ({ ...f, iM: v }))}
                          />
                        </div>
                      </div>
                      <ArrowRight size={16} style={styles.timeArrow} />
                      <div style={styles.timeGroup}>
                        <div style={styles.timeGroupLabel}>fim</div>
                        <div style={styles.timeColsRow}>
                          <ScrollColumn
                            options={HOURS}
                            value={form.fH}
                            onChange={(v) => setForm((f) => ({ ...f, fH: v }))}
                          />
                          <span style={styles.timeColon}>:</span>
                          <ScrollColumn
                            options={MINUTES}
                            value={form.fM}
                            onChange={(v) => setForm((f) => ({ ...f, fM: v }))}
                          />
                        </div>
                      </div>
                    </div>
                  </div>

                  <Field icon={<Building2 size={14} />} label="Local / hospital">
                    <input
                      style={styles.input}
                      value={form.local}
                      onChange={(e) => setForm((f) => ({ ...f, local: e.target.value }))}
                      placeholder="ex: Hospital São Lucas"
                    />
                  </Field>
                </>
              ) : form.type === "evento" ? (
                <>
                  <div style={styles.field}>
                    <div style={styles.fieldLabel}>
                      <Clock size={14} />
                      <span>Horário do evento</span>
                    </div>
                    <div style={styles.timeRangeRow}>
                      <div style={styles.timeGroup}>
                        <div style={styles.timeGroupLabel}>início</div>
                        <div style={styles.timeColsRow}>
                          <ScrollColumn
                            options={HOURS}
                            value={form.iH}
                            onChange={(v) => setForm((f) => ({ ...f, iH: v }))}
                          />
                          <span style={styles.timeColon}>:</span>
                          <ScrollColumn
                            options={MINUTES}
                            value={form.iM}
                            onChange={(v) => setForm((f) => ({ ...f, iM: v }))}
                          />
                        </div>
                      </div>
                      <ArrowRight size={16} style={styles.timeArrow} />
                      <div style={styles.timeGroup}>
                        <div style={styles.timeGroupLabel}>fim</div>
                        <div style={styles.timeColsRow}>
                          <ScrollColumn
                            options={HOURS}
                            value={form.fH}
                            onChange={(v) => setForm((f) => ({ ...f, fH: v }))}
                          />
                          <span style={styles.timeColon}>:</span>
                          <ScrollColumn
                            options={MINUTES}
                            value={form.fM}
                            onChange={(v) => setForm((f) => ({ ...f, fM: v }))}
                          />
                        </div>
                      </div>
                    </div>
                  </div>

                  <Field icon={<Presentation size={14} />} label="Nome do evento">
                    <input
                      style={styles.input}
                      value={form.local}
                      onChange={(e) => setForm((f) => ({ ...f, local: e.target.value }))}
                      placeholder="ex: Congresso Brasileiro de Cardiologia"
                    />
                  </Field>
                  <Field icon={<Building2 size={14} />} label="Nome da empresa">
                    <input
                      style={styles.input}
                      value={form.empresa}
                      onChange={(e) => setForm((f) => ({ ...f, empresa: e.target.value }))}
                      placeholder="ex: Sociedade Brasileira de Cardiologia"
                    />
                  </Field>
                </>
              ) : (
                <>
                  <Field icon={<Building2 size={14} />} label="Empresa da remoção">
                    <input
                      style={styles.input}
                      value={form.empresa}
                      onChange={(e) => setForm((f) => ({ ...f, empresa: e.target.value }))}
                      placeholder="ex: Vida Ambulâncias"
                    />
                  </Field>
                  <Field icon={<User size={14} />} label="Paciente (opcional)">
                    <input
                      style={styles.input}
                      value={form.paciente}
                      onChange={(e) => setForm((f) => ({ ...f, paciente: e.target.value }))}
                      placeholder="identificação opcional"
                    />
                  </Field>
                  <Field icon={<MapPin size={14} />} label="Origem → destino (opcional)">
                    <div style={styles.rowFields}>
                      <input
                        style={styles.input}
                        value={form.origem}
                        onChange={(e) => setForm((f) => ({ ...f, origem: e.target.value }))}
                        placeholder="origem"
                      />
                      <input
                        style={styles.input}
                        value={form.destino}
                        onChange={(e) => setForm((f) => ({ ...f, destino: e.target.value }))}
                        placeholder="destino"
                      />
                    </div>
                  </Field>
                </>
              )}

              <Field icon={<span style={styles.moneySign}>R$</span>} label="Valor recebido">
                <input
                  style={{ ...styles.input, fontFamily: "'IBM Plex Mono', monospace" }}
                  value={form.value}
                  onChange={(e) => setForm((f) => ({ ...f, value: e.target.value }))}
                  placeholder="0,00"
                  inputMode="decimal"
                />
              </Field>

              <label style={styles.checkboxRow}>
                <input
                  type="checkbox"
                  checked={!!form.pago}
                  onChange={(e) => setForm((f) => ({ ...f, pago: e.target.checked }))}
                  style={styles.checkboxInput}
                />
                <span style={styles.checkboxLabel}>
                  {form.pago ? (
                    <CheckCircle2 size={15} color="#206B3C" />
                  ) : (
                    <Circle size={15} color="#8C6D1B" />
                  )}
                  Já foi pago
                </span>
              </label>

              <Field icon={<StickyNote size={14} />} label="Observações (opcional)">
                <textarea
                  style={{ ...styles.input, minHeight: 60, resize: "vertical" }}
                  value={form.obs}
                  onChange={(e) => setForm((f) => ({ ...f, obs: e.target.value }))}
                  placeholder="notas adicionais"
                />
              </Field>
            </div>

            {duplicating ? (
              <div style={styles.duplicateRow}>
                <Field icon={<Copy size={14} />} label="Duplicar para a data">
                  <input
                    type="date"
                    style={styles.input}
                    value={duplicateTargetDay}
                    onChange={(e) => setDuplicateTargetDay(e.target.value)}
                  />
                </Field>
                <div style={styles.modalFooter}>
                  <div style={{ flex: 1 }} />
                  <button className="btn-lift" style={styles.cancelBtn} onClick={() => setDuplicating(false)}>
                    cancelar
                  </button>
                  <button className="btn-lift" style={styles.saveBtn} onClick={confirmDuplicate}>
                    confirmar
                  </button>
                </div>
              </div>
            ) : (
              <div style={styles.modalFooter}>
                {editingId && (
                  <>
                    <button className="btn-lift" style={styles.deleteBtn} onClick={handleDelete}>
                      <Trash2 size={15} />
                      excluir
                    </button>
                    <button className="btn-lift" style={styles.duplicateBtn} onClick={openDuplicate}>
                      <Copy size={15} />
                      duplicar
                    </button>
                  </>
                )}
                <div style={{ flex: 1 }} />
                <button className="btn-lift" style={styles.cancelBtn} onClick={closeModal}>
                  cancelar
                </button>
                <button
                  className="btn-lift"
                  style={{ ...styles.saveBtn, opacity: isValid ? 1 : 0.7 }}
                  onClick={handleSave}
                >
                  salvar
                </button>
              </div>
            )}
          </div>
        </div>
      )}

      {toast && (
        <div
          key={toast.key}
          style={{
            ...styles.toast,
            ...(toast.type === "error" ? styles.toastError : styles.toastSuccess),
          }}
        >
          {toast.type === "error" ? (
            <AlertCircle size={15} />
          ) : (
            <CheckCircle2 size={15} />
          )}
          <span>{toast.msg}</span>
        </div>
      )}
    </div>

    {printModalOpen && printDraft && (
      <div style={styles.overlay} onClick={() => setPrintModalOpen(false)}>
        <div style={styles.modal} onClick={(e) => e.stopPropagation()}>
          <div style={styles.modalHeader}>
            <div style={styles.modalTitle}>
              resumo para impressão
              <div style={styles.modalDate}>escolha o período que quer resumir</div>
            </div>
            <button className="btn-icon" style={styles.closeBtn} onClick={() => setPrintModalOpen(false)}>
              <X size={18} />
            </button>
          </div>

          <div style={styles.presetRow}>
            <button className="pill-btn" style={styles.presetBtn} onClick={() => applyPreset("esteMes")}>
              este mês
            </button>
            <button className="pill-btn" style={styles.presetBtn} onClick={() => applyPreset("mesPassado")}>
              mês passado
            </button>
            <button className="pill-btn" style={styles.presetBtn} onClick={() => applyPreset("esteAno")}>
              este ano
            </button>
          </div>

          <div style={styles.formBody}>
            <Field icon={<CalendarRange size={14} />} label="De">
              <input
                type="date"
                style={styles.input}
                value={printDraft.start}
                onChange={(e) =>
                  setPrintDraft((d) => ({ ...d, start: e.target.value }))
                }
              />
            </Field>
            <Field icon={<CalendarRange size={14} />} label="Até">
              <input
                type="date"
                style={styles.input}
                value={printDraft.end}
                onChange={(e) =>
                  setPrintDraft((d) => ({ ...d, end: e.target.value }))
                }
              />
            </Field>
          </div>

          <div style={styles.popupHint}>
            <Loader2 size={13} style={{ flexShrink: 0, marginTop: 1 }} />
            <span>
              O arquivo é baixado diretamente pelo navegador — na primeira
              vez pode levar alguns segundos para preparar.
            </span>
          </div>

          <div style={styles.modalFooter}>
            <div style={{ flex: 1 }} />
            <button className="btn-lift" style={styles.cancelBtn} onClick={() => setPrintModalOpen(false)}>
              cancelar
            </button>
            <button
              className="btn-lift"
              style={{ ...styles.excelBtn, opacity: excelLoading ? 0.7 : 1 }}
              onClick={() => confirmPrint("excel")}
              disabled={excelLoading || pdfLoading}
            >
              {excelLoading ? (
                <Loader2 size={14} className="spin" style={{ marginRight: 6 }} />
              ) : (
                <FileSpreadsheet size={14} style={{ marginRight: 6 }} />
              )}
              {excelLoading ? "gerando…" : "gerar Excel"}
            </button>
            <button
              className="btn-lift"
              style={{ ...styles.saveBtn, opacity: pdfLoading ? 0.7 : 1 }}
              onClick={() => confirmPrint("pdf")}
              disabled={pdfLoading || excelLoading}
            >
              {pdfLoading ? (
                <Loader2 size={14} className="spin" style={{ marginRight: 6 }} />
              ) : (
                <Printer size={14} style={{ marginRight: 6 }} />
              )}
              {pdfLoading ? "gerando…" : "gerar PDF"}
            </button>
          </div>
        </div>
      </div>
    )}

    <div className="print-report" style={styles.printReport}>
      <div style={styles.printHeader}>
        <div style={styles.printTitle}>Resumo de plantões e remoções</div>
        <div style={styles.printSubtitle}>
          Período: {formatShort(printData.start)} a {formatShort(printData.end)}
        </div>
      </div>

      <table style={styles.printTable}>
        <thead>
          <tr>
            <th style={styles.printTh}>Data</th>
            <th style={styles.printTh}>Tipo</th>
            <th style={styles.printTh}>Descrição</th>
            <th style={styles.printTh}>Horário</th>
            <th style={styles.printTh}>Status</th>
            <th style={{ ...styles.printTh, textAlign: "right" }}>Valor</th>
          </tr>
        </thead>
        <tbody>
          {printData.rows.length === 0 && (
            <tr>
              <td style={styles.printTd} colSpan={6}>
                Nenhum registro neste período.
              </td>
            </tr>
          )}
          {printData.rows.map((e) => (
            <tr key={e.id}>
              <td style={styles.printTd}>{formatShort(e.dayKey)}</td>
              <td style={styles.printTd}>
                {e.type === "plantao" ? "Plantão" : e.type === "evento" ? "Evento" : "Remoção"}
              </td>
              <td style={styles.printTd}>
                {e.type === "remocao" ? e.empresa : e.local}
              </td>
              <td style={styles.printTd}>
                {e.type !== "remocao" && e.iH ? `${e.iH}:${e.iM} – ${e.fH}:${e.fM}` : "—"}
              </td>
              <td style={styles.printTd}>{e.pago ? "Pago" : "A receber"}</td>
              <td style={{ ...styles.printTd, textAlign: "right" }}>
                {currency(e.value)}
              </td>
            </tr>
          ))}
        </tbody>
      </table>

      <div style={styles.printSummary}>
        <div style={styles.printSummaryRow}>
          <span>Plantões ({printData.plantaoCount})</span>
          <span>{currency(printData.plantaoSum)}</span>
        </div>
        <div style={styles.printSummaryRow}>
          <span>Remoções ({printData.remocaoCount})</span>
          <span>{currency(printData.remocaoSum)}</span>
        </div>
        <div style={styles.printSummaryRow}>
          <span>Eventos ({printData.eventoCount})</span>
          <span>{currency(printData.eventoSum)}</span>
        </div>
        <div style={styles.printSummaryRow}>
          <span>Recebido</span>
          <span>{currency(printData.paidSum)}</span>
        </div>
        <div style={styles.printSummaryRow}>
          <span>A receber</span>
          <span>{currency(printData.pendingSum)}</span>
        </div>
        <div style={{ ...styles.printSummaryRow, ...styles.printSummaryTotal }}>
          <span>Total do período</span>
          <span>{currency(printData.total)}</span>
        </div>
      </div>
    </div>
    </>
  );
}

function SummaryChip({ icon, label, count, value, color, bg }) {
  return (
    <div style={{ ...styles.summaryChip, background: bg, color }}>
      <div style={styles.summaryChipIconRow}>
        {icon}
        <span style={styles.summaryChipLabel}>{label}</span>
        {count !== undefined && (
          <span style={styles.summaryChipCount}>{count}</span>
        )}
      </div>
      <div style={styles.summaryChipValue}>{currency(value)}</div>
    </div>
  );
}

function Field({ icon, label, children }) {
  return (
    <div style={styles.field}>
      <div style={styles.fieldLabel}>
        {icon}
        <span>{label}</span>
      </div>
      {children}
    </div>
  );
}

function formatFullDate(dayKey) {
  if (!dayKey) return "";
  const [y, m, d] = dayKey.split("-").map(Number);
  const date = new Date(y, m - 1, d);
  const dias = ["domingo","segunda-feira","terça-feira","quarta-feira","quinta-feira","sexta-feira","sábado"];
  return `${dias[date.getDay()]}, ${d} de ${MESES[m - 1].toLowerCase()}`;
}

const globalCss = `
  * { box-sizing: border-box; }
  .spin { animation: plantoes-spin 1s linear infinite; }
  @keyframes plantoes-spin { to { transform: rotate(360deg); } }
  @media (prefers-reduced-motion: reduce) {
    .spin { animation: none; }
  }
  .plantoes-scroll-col::-webkit-scrollbar { display: none; }
  .plantoes-scroll-col { scrollbar-width: none; }
  @keyframes plantoes-toast-in {
    from { opacity: 0; transform: translate(-50%, 8px); }
    to { opacity: 1; transform: translate(-50%, 0); }
  }
  @media print {
    .app-main { display: none !important; }
    .print-report { display: block !important; }
  }
  @media screen {
    .print-report { display: none; }
  }
  @media (max-width: 480px) {
    .cal-grid { gap: 3px !important; }
    .cal-cell { min-height: 74px !important; padding: 4px !important; }
  }
  input:focus, textarea:focus, button:focus-visible {
    outline: 2px solid #2D6E6E;
    outline-offset: 1px;
  }

  /* --- sophisticated button interactions --- */
  .btn-lift {
    transition: transform 0.16s cubic-bezier(.4,0,.2,1), box-shadow 0.16s cubic-bezier(.4,0,.2,1), filter 0.16s ease;
  }
  .btn-lift:hover:not(:disabled) {
    transform: translateY(-1.5px);
    filter: brightness(1.05);
    box-shadow: 0 6px 14px rgba(28,43,57,0.16);
  }
  .btn-lift:active:not(:disabled) {
    transform: translateY(0);
    filter: brightness(0.96);
    box-shadow: 0 1px 3px rgba(28,43,57,0.14);
    transition-duration: 0.06s;
  }
  .btn-lift:disabled {
    cursor: not-allowed;
  }

  .btn-icon {
    transition: transform 0.15s cubic-bezier(.4,0,.2,1), box-shadow 0.15s ease, background-color 0.15s ease;
  }
  .btn-icon:hover:not(:disabled) {
    transform: scale(1.1);
    box-shadow: 0 3px 9px rgba(28,43,57,0.2);
    background-color: #F1EFE9;
  }
  .btn-icon:active:not(:disabled) {
    transform: scale(0.9);
    transition-duration: 0.06s;
  }

  .pill-btn {
    transition: transform 0.14s ease, box-shadow 0.14s ease, filter 0.14s ease;
  }
  .pill-btn:hover:not(:disabled) {
    box-shadow: inset 0 0 0 1px rgba(28,43,57,0.2);
    filter: brightness(0.98);
  }
  .pill-btn:active:not(:disabled) {
    transform: scale(0.95);
    transition-duration: 0.06s;
  }

  .chip-lift {
    transition: transform 0.14s ease, box-shadow 0.14s ease;
  }
  .chip-lift:hover {
    transform: translateY(-1.5px);
    box-shadow: 0 4px 10px rgba(28,43,57,0.18);
  }
  .chip-lift:active {
    transform: translateY(0) scale(0.98);
    transition-duration: 0.06s;
  }
`;

const styles = {
  presetRow: { display: "flex", gap: 6, marginBottom: 14, flexWrap: "wrap" },
  searchWrap: { position: "relative", marginBottom: 12 },
  searchFiltersRow: { display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" },
  searchDateSep: { fontSize: 11.5, color: "#5B6B75", flexShrink: 0 },
  searchTypeRow: { display: "flex", gap: 6, marginTop: 8, flexWrap: "wrap", alignItems: "center" },
  searchTypeDivider: {
    width: 1,
    height: 18,
    background: "#E0DDD3",
    margin: "0 2px",
  },
  exportPanel: {
    marginTop: 12,
    padding: "10px 12px",
    background: "#F1EFE9",
    borderRadius: 10,
    display: "flex",
    flexWrap: "wrap",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 10,
  },
  exportColsRow: { display: "flex", gap: 12, flexWrap: "wrap" },
  exportColCheckbox: {
    display: "flex",
    alignItems: "center",
    gap: 5,
    fontSize: 12,
    color: "#1C2B39",
    cursor: "pointer",
  },
  exportButtonsRow: { display: "flex", gap: 8, flexWrap: "wrap" },
  searchTypeBtn: {
    border: "1px solid #E0DDD3",
    background: "#fff",
    borderRadius: 999,
    padding: "5px 12px",
    fontSize: 11.5,
    cursor: "pointer",
    color: "#5B6B75",
  },
  searchTypeBtnActive: {
    background: "#1C2B39",
    borderColor: "#1C2B39",
    color: "#F7F5F0",
    fontWeight: 600,
  },
  searchInputWrap: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    border: "1px solid #E0DDD3",
    background: "#fff",
    borderRadius: 10,
    padding: "8px 12px",
  },
  searchInput: {
    flex: 1,
    border: "none",
    outline: "none",
    fontSize: 13.5,
    fontFamily: "'Inter', sans-serif",
    color: "#1C2B39",
    background: "transparent",
  },
  searchClearBtn: {
    border: "none",
    background: "#F1EFE9",
    color: "#5B6B75",
    borderRadius: 6,
    width: 20,
    height: 20,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    cursor: "pointer",
    flexShrink: 0,
  },
  searchDropdown: {
    marginTop: 12,
    background: "#fff",
    border: "1px solid #E0DDD3",
    borderRadius: 10,
    boxShadow: "0 4px 14px rgba(28,43,57,0.06)",
    display: "flex",
    flexDirection: "column",
    overflow: "hidden",
  },
  searchScrollArea: { maxHeight: 460, overflowY: "auto", overflowX: "auto" },
  tabRow: { display: "flex", gap: 6, marginBottom: 14, flexWrap: "wrap" },
  tabBtn: {
    display: "flex",
    alignItems: "center",
    gap: 6,
    border: "1px solid #E0DDD3",
    background: "#fff",
    borderRadius: 8,
    padding: "7px 14px",
    fontSize: 12.5,
    cursor: "pointer",
    color: "#5B6B75",
  },
  tabBtnActive: {
    background: "#1C2B39",
    borderColor: "#1C2B39",
    color: "#F7F5F0",
    fontWeight: 600,
  },
  searchEmpty: { padding: "14px 12px", fontSize: 12.5, color: "#5B6B75" },
  searchTable: { width: "100%", minWidth: 460, borderCollapse: "collapse", fontSize: 12.5 },
  searchTh: {
    textAlign: "left",
    padding: "8px 12px",
    fontSize: 10.5,
    color: "#5B6B75",
    textTransform: "uppercase",
    letterSpacing: "0.03em",
    borderBottom: "1px solid #E0DDD3",
    position: "sticky",
    top: 0,
    background: "#fff",
  },
  searchRow: { cursor: "pointer" },
  searchPagoBtn: {
    border: "none",
    background: "transparent",
    cursor: "pointer",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    padding: 2,
  },
  searchTd: {
    padding: "8px 12px",
    borderBottom: "1px solid #F1EFE9",
    fontFamily: "'IBM Plex Mono', monospace",
    color: "#1C2B39",
  },
  searchValuePending: {
    color: "#E0554A",
    fontWeight: 700,
  },
  searchTdName: {
    padding: "8px 12px",
    borderBottom: "1px solid #F1EFE9",
    fontFamily: "'Inter', sans-serif",
    color: "#1C2B39",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
    maxWidth: 180,
  },
  searchTotalsWrap: { display: "flex", flexDirection: "column" },
  searchTotalPlantao: {
    background: "#E4F0EF",
    color: "#215454",
    borderRadius: 0,
  },
  searchTotalEvento: {
    background: "#E3EAF6",
    color: "#1F4278",
    borderRadius: 0,
  },
  searchTotalRow: {
    display: "flex",
    justifyContent: "space-between",
    padding: "9px 12px",
    fontSize: 12.5,
    fontWeight: 700,
    background: "#F5E6DC",
    color: "#8C4118",
    borderTop: "1px solid #E0DDD3",
    borderRadius: "0 0 10px 10px",
    fontFamily: "'IBM Plex Mono', monospace",
  },
  searchItem: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    width: "100%",
    border: "none",
    background: "transparent",
    padding: "9px 12px",
    borderBottom: "1px solid #F1EFE9",
    cursor: "pointer",
    textAlign: "left",
  },
  searchItemDot: { width: 8, height: 8, borderRadius: "50%", flexShrink: 0 },
  searchItemDate: {
    fontFamily: "'IBM Plex Mono', monospace",
    fontSize: 11.5,
    color: "#5B6B75",
    flexShrink: 0,
  },
  searchItemDesc: {
    flex: 1,
    fontSize: 12.5,
    color: "#1C2B39",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  searchItemValue: {
    fontFamily: "'IBM Plex Mono', monospace",
    fontSize: 12,
    fontWeight: 600,
    color: "#1C2B39",
    flexShrink: 0,
  },
  popupHint: {
    display: "flex",
    alignItems: "flex-start",
    gap: 6,
    fontSize: 11.5,
    color: "#8C6D1B",
    background: "#F6EFDD",
    borderRadius: 8,
    padding: "8px 10px",
    marginTop: 12,
  },
  presetBtn: {
    border: "1px solid #E0DDD3",
    background: "#F1EFE9",
    borderRadius: 8,
    padding: "6px 10px",
    fontSize: 12,
    cursor: "pointer",
    color: "#1C2B39",
  },
  printReport: {
    display: "none",
    padding: 24,
    fontFamily: "'Inter', sans-serif",
    color: "#1C2B39",
  },
  printHeader: { marginBottom: 16, borderBottom: "2px solid #1C2B39", paddingBottom: 10 },
  printTitle: { fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 20 },
  printSubtitle: { fontSize: 12.5, color: "#5B6B75", marginTop: 3 },
  printTable: { width: "100%", borderCollapse: "collapse", fontSize: 12 },
  printTh: {
    textAlign: "left",
    borderBottom: "1px solid #1C2B39",
    padding: "6px 8px",
    fontSize: 11,
    textTransform: "uppercase",
    letterSpacing: "0.03em",
  },
  printTd: {
    borderBottom: "1px solid #E0DDD3",
    padding: "6px 8px",
    fontFamily: "'IBM Plex Mono', monospace",
    fontSize: 12,
  },
  printSummary: { marginTop: 18, maxWidth: 320, marginLeft: "auto" },
  printSummaryRow: {
    display: "flex",
    justifyContent: "space-between",
    padding: "5px 0",
    fontSize: 13,
    fontFamily: "'IBM Plex Mono', monospace",
  },
  printSummaryTotal: {
    borderTop: "2px solid #1C2B39",
    marginTop: 4,
    paddingTop: 8,
    fontWeight: 700,
    fontSize: 15,
  },
  app: {
    fontFamily: "'Inter', -apple-system, sans-serif",
    background: "#F7F5F0",
    color: "#1C2B39",
    minHeight: "100vh",
    padding: "20px 16px 60px",
    maxWidth: 1000,
    margin: "0 auto",
  },
  header: { marginBottom: 18 },
  headerTop: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    marginBottom: 14,
    flexWrap: "wrap",
    gap: 8,
  },
  brand: { display: "flex", alignItems: "center", gap: 10 },
  brandMark: {
    width: 34,
    height: 34,
    borderRadius: 9,
    background: "#1C2B39",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    flexShrink: 0,
  },
  brandTitle: {
    fontFamily: "'Space Grotesk', sans-serif",
    fontWeight: 700,
    fontSize: 19,
    lineHeight: 1.1,
    letterSpacing: "-0.01em",
  },
  brandSub: { fontSize: 11, color: "#5B6B75", marginTop: 1 },
  saveWarning: {
    fontSize: 11,
    color: "#B5541F",
    background: "#F5E6DC",
    padding: "4px 8px",
    borderRadius: 6,
  },
  monthNav: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    marginBottom: 14,
    flexWrap: "wrap",
  },
  navBtn: {
    width: 30,
    height: 30,
    borderRadius: 8,
    border: "1px solid #E0DDD3",
    background: "#fff",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    cursor: "pointer",
    color: "#1C2B39",
  },
  monthLabel: {
    fontFamily: "'Space Grotesk', sans-serif",
    fontWeight: 600,
    fontSize: 17,
    minWidth: 150,
  },
  year: { color: "#5B6B75", fontWeight: 500 },
  todayBtn: {
    marginLeft: "auto",
    border: "1px solid #E0DDD3",
    background: "#fff",
    borderRadius: 8,
    padding: "6px 12px",
    fontSize: 12.5,
    cursor: "pointer",
    color: "#1C2B39",
    boxShadow: "0 1px 2px rgba(28,43,57,0.06)",
  },
  printBtn: {
    display: "flex",
    alignItems: "center",
    gap: 5,
    border: "1px solid #1C2B39",
    background: "#1C2B39",
    color: "#F7F5F0",
    borderRadius: 8,
    padding: "6px 12px",
    fontSize: 12.5,
    cursor: "pointer",
    boxShadow: "0 2px 6px rgba(28,43,57,0.22)",
  },
  summaryBar: {
    display: "flex",
    gap: 8,
    flexWrap: "wrap",
  },
  financeRow: {
    display: "flex",
    gap: 8,
    flexWrap: "wrap",
    marginTop: 8,
  },
  summaryChip: {
    borderRadius: 10,
    padding: "8px 12px",
    minWidth: 130,
  },
  summaryChipIconRow: {
    display: "flex",
    alignItems: "center",
    gap: 6,
    fontSize: 11.5,
    fontWeight: 600,
    textTransform: "uppercase",
    letterSpacing: "0.02em",
  },
  summaryChipLabel: {},
  summaryChipCount: {
    marginLeft: "auto",
    fontFamily: "'IBM Plex Mono', monospace",
    fontSize: 11,
    opacity: 0.75,
  },
  summaryChipValue: {
    fontFamily: "'IBM Plex Mono', monospace",
    fontSize: 15,
    fontWeight: 600,
    marginTop: 3,
  },
  totalChip: {
    borderRadius: 10,
    padding: "8px 14px",
    background: "#1C2B39",
    color: "#F7F5F0",
    marginLeft: "auto",
  },
  totalLabel: {
    fontSize: 11,
    textTransform: "uppercase",
    letterSpacing: "0.03em",
    opacity: 0.75,
  },
  totalValue: {
    fontFamily: "'IBM Plex Mono', monospace",
    fontSize: 17,
    fontWeight: 600,
    marginTop: 2,
  },
  loadingBox: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    gap: 8,
    padding: 60,
    color: "#5B6B75",
    fontSize: 13,
  },
  weekHeader: {
    display: "grid",
    gridTemplateColumns: "repeat(7, 1fr)",
    marginBottom: 4,
  },
  weekHeaderCell: {
    textAlign: "center",
    fontSize: 11,
    color: "#5B6B75",
    textTransform: "uppercase",
    letterSpacing: "0.04em",
    padding: "4px 0",
  },
  grid: {
    display: "grid",
    gridTemplateColumns: "repeat(7, 1fr)",
    gap: 6,
  },
  emptyCell: { minHeight: 90 },
  cell: {
    background: "#fff",
    border: "1px solid #E0DDD3",
    borderRadius: 10,
    padding: 6,
    minHeight: 90,
    display: "flex",
    flexDirection: "column",
    gap: 4,
  },
  cellToday: {
    borderColor: "#2D6E6E",
    boxShadow: "0 0 0 1px #2D6E6E inset",
  },
  cellDragOver: {
    borderColor: "#2D6E6E",
    background: "#E4F0EF",
    boxShadow: "0 0 0 2px #2D6E6E inset",
  },
  dragHint: {
    fontSize: 11,
    color: "#5B6B75",
    marginBottom: 8,
    textAlign: "center",
  },
  cellHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
  },
  cellNum: {
    fontFamily: "'IBM Plex Mono', monospace",
    fontSize: 12,
    color: "#5B6B75",
  },
  cellNumToday: {
    color: "#2D6E6E",
    fontWeight: 700,
  },
  addBtn: {
    width: 18,
    height: 18,
    borderRadius: 5,
    border: "none",
    background: "#F1EFE9",
    color: "#5B6B75",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    cursor: "pointer",
  },
  chipStack: { display: "flex", flexDirection: "column", gap: 3, overflow: "hidden" },
  entryChip: {
    display: "flex",
    alignItems: "center",
    gap: 4,
    border: "none",
    borderRadius: 5,
    padding: "3px 5px",
    fontSize: 10,
    cursor: "grab",
    textAlign: "left",
    width: "100%",
  },
  entryChipDragging: {
    opacity: 0.4,
  },
  chipPlantao: { background: "#E4F0EF", color: "#215454" },
  chipRemocao: { background: "#F5E6DC", color: "#8C4118" },
  chipText: {
    flex: 1,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  chipValue: {
    fontFamily: "'IBM Plex Mono', monospace",
    fontSize: 9.5,
    fontWeight: 600,
    flexShrink: 0,
  },
  overlay: {
    position: "fixed",
    inset: 0,
    background: "rgba(28,43,57,0.45)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    padding: 16,
    zIndex: 50,
  },
  modal: {
    background: "#FFFDF9",
    borderRadius: 14,
    width: "100%",
    maxWidth: 420,
    maxHeight: "90vh",
    overflowY: "auto",
    padding: 18,
  },
  modalHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "flex-start",
    marginBottom: 14,
  },
  modalTitle: {
    fontFamily: "'Space Grotesk', sans-serif",
    fontWeight: 600,
    fontSize: 16,
    textTransform: "lowercase",
  },
  modalDate: {
    fontSize: 11.5,
    color: "#5B6B75",
    marginTop: 2,
    textTransform: "capitalize",
  },
  closeBtn: {
    border: "none",
    background: "transparent",
    cursor: "pointer",
    color: "#5B6B75",
  },
  typeToggle: {
    display: "flex",
    gap: 6,
    marginBottom: 16,
    background: "#F1EFE9",
    borderRadius: 10,
    padding: 4,
  },
  typeBtn: {
    flex: 1,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    gap: 6,
    padding: "8px 0",
    borderRadius: 8,
    border: "none",
    background: "transparent",
    color: "#5B6B75",
    fontSize: 13,
    fontWeight: 600,
    cursor: "pointer",
  },
  typeBtnActivePlantao: { background: "#2D6E6E", color: "#fff", boxShadow: "0 2px 6px rgba(45,110,110,0.35)" },
  typeBtnActiveRemocao: { background: "#B5541F", color: "#fff", boxShadow: "0 2px 6px rgba(181,84,31,0.35)" },
  typeBtnActiveEvento: { background: "#2A5DA8", color: "#fff", boxShadow: "0 2px 6px rgba(42,93,168,0.35)" },
  formBody: { display: "flex", flexDirection: "column", gap: 12 },
  field: { display: "flex", flexDirection: "column", gap: 5 },
  fieldLabel: {
    display: "flex",
    alignItems: "center",
    gap: 5,
    fontSize: 11.5,
    color: "#5B6B75",
    fontWeight: 500,
  },
  moneySign: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, fontWeight: 700 },
  colorScrollWrap: {
    position: "relative",
    width: 132,
    height: 56,
    borderRadius: 10,
    border: "1px solid #E0DDD3",
    background: "#fff",
    overflow: "hidden",
    margin: "2px 0",
  },
  colorScrollBand: {
    position: "absolute",
    top: 0,
    bottom: 0,
    left: 44,
    width: 44,
    borderLeft: "1px solid #2D6E6E",
    borderRight: "1px solid #2D6E6E",
    background: "rgba(45,110,110,0.06)",
    pointerEvents: "none",
    zIndex: 1,
  },
  colorScrollRow: {
    display: "flex",
    height: 56,
    alignItems: "center",
    overflowX: "auto",
    scrollSnapType: "x mandatory",
  },
  colorScrollItemWrap: {
    width: 44,
    flexShrink: 0,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    scrollSnapAlign: "center",
    cursor: "pointer",
  },
  colorScrollDot: {
    width: 26,
    height: 26,
    borderRadius: "50%",
    border: "2px solid transparent",
  },
  colorScrollDotActive: {
    border: "2px solid #1C2B39",
    boxShadow: "0 0 0 2px #FFFDF9",
    transform: "scale(1.1)",
  },
  toast: {
    position: "fixed",
    bottom: 22,
    left: "50%",
    transform: "translateX(-50%)",
    display: "flex",
    alignItems: "center",
    gap: 7,
    padding: "10px 16px",
    borderRadius: 10,
    fontSize: 13,
    fontWeight: 600,
    boxShadow: "0 6px 18px rgba(28,43,57,0.18)",
    zIndex: 200,
    animation: "plantoes-toast-in 0.2s ease-out",
  },
  toastSuccess: { background: "#1C2B39", color: "#E4F0EF" },
  toastError: { background: "#8C2C1F", color: "#F8E2DE" },
  swatchRow: { display: "flex", flexWrap: "wrap", gap: 8, padding: "2px 0 10px" },
  swatch: {
    width: 24,
    height: 24,
    borderRadius: "50%",
    border: "2px solid transparent",
    cursor: "pointer",
    padding: 0,
  },
  swatchActive: {
    border: "2px solid #1C2B39",
    boxShadow: "0 0 0 2px #FFFDF9",
    transform: "scale(1.08)",
  },
  checkboxRow: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    cursor: "pointer",
    padding: "8px 10px",
    background: "#F1EFE9",
    borderRadius: 8,
    width: "fit-content",
  },
  checkboxInput: {
    width: 15,
    height: 15,
    accentColor: "#1C2B39",
    cursor: "pointer",
  },
  checkboxLabel: {
    display: "flex",
    alignItems: "center",
    gap: 6,
    fontSize: 13,
    fontWeight: 500,
    color: "#1C2B39",
  },
  timeRangeRow: {
    display: "flex",
    alignItems: "flex-start",
    gap: 10,
    flexWrap: "wrap",
  },
  timeGroup: { display: "flex", flexDirection: "column", gap: 4, alignItems: "center" },
  timeGroupLabel: {
    fontSize: 10.5,
    color: "#5B6B75",
    textTransform: "uppercase",
    letterSpacing: "0.04em",
  },
  timeColsRow: { display: "flex", alignItems: "center", gap: 3 },
  timeColon: { fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, color: "#5B6B75" },
  timeArrow: { marginTop: 26, color: "#5B6B75", flexShrink: 0 },
  scrollColWrap: {
    position: "relative",
    width: 50,
    height: 96,
    borderRadius: 8,
    border: "1px solid #E0DDD3",
    background: "#fff",
    overflow: "hidden",
  },
  scrollColBand: {
    position: "absolute",
    top: 32,
    left: 0,
    right: 0,
    height: 32,
    borderTop: "1px solid #2D6E6E",
    borderBottom: "1px solid #2D6E6E",
    background: "rgba(45,110,110,0.06)",
    pointerEvents: "none",
    zIndex: 1,
  },
  scrollCol: {
    height: 96,
    overflowY: "auto",
    scrollSnapType: "y mandatory",
    textAlign: "center",
    position: "relative",
  },
  scrollItem: {
    height: 32,
    lineHeight: "32px",
    fontFamily: "'IBM Plex Mono', monospace",
    fontSize: 14,
    color: "#B7BEC2",
    scrollSnapAlign: "center",
    cursor: "pointer",
    userSelect: "none",
  },
  scrollItemActive: {
    color: "#1C2B39",
    fontWeight: 700,
  },
  rowFields: { display: "flex", gap: 6 },
  input: {
    border: "1px solid #E0DDD3",
    borderRadius: 8,
    padding: "9px 10px",
    fontSize: 13.5,
    fontFamily: "'Inter', sans-serif",
    color: "#1C2B39",
    background: "#fff",
    width: "100%",
  },
  modalFooter: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    marginTop: 18,
    flexWrap: "wrap",
  },
  deleteBtn: {
    display: "flex",
    alignItems: "center",
    gap: 5,
    border: "1px solid #E9C9B4",
    color: "#B5541F",
    background: "transparent",
    borderRadius: 8,
    padding: "8px 12px",
    fontSize: 12.5,
    cursor: "pointer",
  },
  duplicateBtn: {
    display: "flex",
    alignItems: "center",
    gap: 5,
    border: "1px solid #C2D6EA",
    color: "#1F4278",
    background: "transparent",
    borderRadius: 8,
    padding: "8px 12px",
    fontSize: 12.5,
    cursor: "pointer",
  },
  duplicateRow: {
    display: "flex",
    flexDirection: "column",
  },
  cancelBtn: {
    border: "1px solid #E0DDD3",
    background: "transparent",
    borderRadius: 8,
    padding: "8px 14px",
    fontSize: 12.5,
    cursor: "pointer",
    color: "#5B6B75",
  },
  saveBtn: {
    border: "none",
    background: "#1C2B39",
    color: "#fff",
    borderRadius: 8,
    padding: "8px 16px",
    fontSize: 12.5,
    fontWeight: 600,
    cursor: "pointer",
    boxShadow: "0 2px 6px rgba(28,43,57,0.25)",
  },
  excelBtn: {
    display: "flex",
    alignItems: "center",
    border: "1px solid #206B3C",
    background: "#fff",
    color: "#206B3C",
    borderRadius: 8,
    padding: "8px 16px",
    fontSize: 12.5,
    fontWeight: 600,
    cursor: "pointer",
    boxShadow: "0 1px 3px rgba(32,107,60,0.12)",
  },
};
