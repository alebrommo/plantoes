import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import {
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
  Tooltip,
  Legend,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
} from "recharts";
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
  BarChart3,
  Mail,
  Lock,
  LogOut,
  Bold,
  Italic,
  List,
} from "lucide-react";
import { supabase, supabaseConfigured } from "./supabaseClient";

const FONT_IMPORT_ID = "plantoes-fonts";
const TABLE = "entries";
const CACHE_KEY = "plantoes-cache";
const QUEUE_KEY = "plantoes-offline-queue";

// Cópia local dos registros, usada para abrir o app sem internet.
// Cache e fila são separados por usuário (uid), para um computador
// compartilhado não misturar dados de contas diferentes.
function loadCache(uid) {
  if (!uid) return {};
  try {
    const raw = localStorage.getItem(`${CACHE_KEY}-${uid}`);
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}

function saveCache(uid, entries) {
  if (!uid) return;
  try {
    localStorage.setItem(`${CACHE_KEY}-${uid}`, JSON.stringify(entries));
  } catch {
    // localStorage indisponível/cheio — cache é best-effort, ignora
  }
}

// Fila de alterações feitas offline, sincronizadas quando a conexão volta.
function loadQueue(uid) {
  if (!uid) return [];
  try {
    const raw = localStorage.getItem(`${QUEUE_KEY}-${uid}`);
    return raw ? JSON.parse(raw) : [];
  } catch {
    return [];
  }
}

function saveQueue(uid, queue) {
  if (!uid) return;
  try {
    localStorage.setItem(`${QUEUE_KEY}-${uid}`, JSON.stringify(queue));
  } catch {
    // ignore
  }
}

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
function entryToRow(dayKey, entry, userId) {
  return {
    id: entry.id,
    day_key: dayKey,
    user_id: userId,
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

const TYPE_CHART_COLORS = { plantao: "#2D6E6E", remocao: "#B5541F", evento: "#2A5DA8" };

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

function ColorGrid({ options, value, onChange }) {
  return (
    <div style={styles.colorGrid}>
      {options.map((c) => (
        <button
          key={c.id}
          type="button"
          className="btn-lift"
          onClick={() => onChange(c.id)}
          title={c.label}
          aria-label={c.label}
          style={{
            ...styles.colorGridDot,
            background: c.base,
            ...(c.id === value ? styles.colorGridDotActive : {}),
          }}
        />
      ))}
    </div>
  );
}

function stripHtml(html) {
  if (!html) return "";
  const div = document.createElement("div");
  div.innerHTML = html;
  return div.textContent || "";
}

function RichTextEditor({ value, onChange, placeholder }) {
  const ref = useRef(null);

  const wrapSelection = (before, after = before) => {
    const el = ref.current;
    if (!el) return;
    const start = el.selectionStart;
    const end = el.selectionEnd;
    const text = value || "";
    const selected = text.slice(start, end);
    const next = text.slice(0, start) + before + selected + after + text.slice(end);
    onChange(next);
    requestAnimationFrame(() => {
      el.focus();
      el.setSelectionRange(start + before.length, start + before.length + selected.length);
    });
  };

  const prefixLines = (prefix) => {
    const el = ref.current;
    if (!el) return;
    const start = el.selectionStart;
    const end = el.selectionEnd;
    const text = value || "";
    const lineStart = text.lastIndexOf("\n", start - 1) + 1;
    let lineEnd = text.indexOf("\n", end);
    if (lineEnd === -1) lineEnd = text.length;
    const block = text.slice(lineStart, lineEnd);
    const prefixed = block
      .split("\n")
      .map((line) => (line.startsWith(prefix) ? line : prefix + line))
      .join("\n");
    const next = text.slice(0, lineStart) + prefixed + text.slice(lineEnd);
    onChange(next);
    requestAnimationFrame(() => {
      el.focus();
      el.setSelectionRange(lineStart, lineStart + prefixed.length);
    });
  };

  return (
    <div style={styles.richWrap}>
      <div style={styles.richToolbar}>
        <button type="button" className="btn-icon" style={styles.richToolbarBtn} onMouseDown={(e) => { e.preventDefault(); wrapSelection("**"); }} title="Negrito">
          <Bold size={13} />
        </button>
        <button type="button" className="btn-icon" style={styles.richToolbarBtn} onMouseDown={(e) => { e.preventDefault(); wrapSelection("_"); }} title="Itálico">
          <Italic size={13} />
        </button>
        <button type="button" className="btn-icon" style={styles.richToolbarBtn} onMouseDown={(e) => { e.preventDefault(); prefixLines("- "); }} title="Tópicos">
          <List size={13} />
        </button>
      </div>
      <textarea
        ref={ref}
        style={styles.richEditable}
        value={value || ""}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
      />
    </div>
  );
}

const currency = (n) =>
  new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" }).format(
    Number(n) || 0
  );

const pad = (n) => String(n).padStart(2, "0");
const keyFor = (y, m, d) => `${y}-${pad(m + 1)}-${pad(d)}`;
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

const formatShortWithWeekday = (dayKey) => {
  const [y, m, d] = dayKey.split("-").map(Number);
  const weekday = DIAS_SEMANA[new Date(y, m - 1, d).getDay()];
  return `${weekday}, ${pad(d)}/${pad(m)}/${y}`;
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

const EXPORT_COL_LABEL = { data: "Data", nome: "Nome", empresa: "Empresa", pago: "Pago", valor: "Valor" };
const EXPORT_COL_ORDER = ["data", "nome", "empresa", "pago", "valor"];

function getSearchColText(e, key) {
  if (key === "data") return formatShort(e.dayKey);
  if (key === "nome") {
    if (e.type === "remocao") return e.empresa || "remoção";
    if (e.type === "evento") return e.local || "evento";
    return e.local || "plantão";
  }
  if (key === "empresa") return e.type === "plantao" ? "—" : e.empresa || "—";
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
  // undefined = ainda verificando; null = sem sessão; objeto = logado
  const [session, setSession] = useState(undefined);
  const [passwordRecovery, setPasswordRecovery] = useState(false);
  const userId = session?.user?.id ?? null;

  useEffect(() => {
    if (!supabaseConfigured) {
      setSession(null);
      return;
    }
    let mounted = true;
    supabase.auth.getSession().then(({ data }) => {
      if (mounted) setSession(data.session ?? null);
    });
    const { data: listener } = supabase.auth.onAuthStateChange((event, newSession) => {
      setSession(newSession);
      if (event === "PASSWORD_RECOVERY") {
        setPasswordRecovery(true);
      }
    });
    return () => {
      mounted = false;
      listener.subscription.unsubscribe();
    };
  }, []);

  const [entries, setEntries] = useState({});
  const [loading, setLoading] = useState(true);
  const [saveError, setSaveError] = useState(false);
  const [toast, setToast] = useState(null);
  const [cursor, setCursor] = useState(() => {
    const t = new Date();
    return { year: t.getFullYear(), month: t.getMonth() };
  });
  const [selectedDay, setSelectedDay] = useState(null); // "YYYY-MM-DD"
  const [viewDay, setViewDay] = useState(() => todayKey()); // dia exibido no painel abaixo do calendário
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
    empresa: true,
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

  // Histórico de desfazer (Ctrl+Z): cada ação sabe reverter a si mesma,
  // tanto no estado local quanto no Supabase.
  const undoStackRef = useRef([]);

  const pushUndo = useCallback((action) => {
    undoStackRef.current = [...undoStackRef.current, action].slice(-20);
  }, []);

  const undo = useCallback(async () => {
    const stack = undoStackRef.current;
    if (stack.length === 0) {
      showToast("Nada para desfazer", "error");
      return;
    }
    const action = stack[stack.length - 1];
    undoStackRef.current = stack.slice(0, -1);
    try {
      await action.run();
      showToast(action.label || "Ação desfeita", "success");
    } catch (err) {
      showToast("Não foi possível desfazer", "error");
    }
  }, [showToast]);

  useEffect(() => {
    function handleKeyDown(evt) {
      const isUndo = (evt.ctrlKey || evt.metaKey) && !evt.shiftKey && evt.key.toLowerCase() === "z";
      if (!isUndo) return;
      const tag = document.activeElement && document.activeElement.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA") return;
      evt.preventDefault();
      undo();
    }
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [undo]);

  // Modo offline: mantém uma cópia local dos registros e uma fila de
  // alterações que não conseguiram ser enviadas ao Supabase, para sincronizar
  // sozinho quando a conexão voltar.
  const [isOnline, setIsOnline] = useState(navigator.onLine);
  const [pendingCount, setPendingCount] = useState(0);

  useEffect(() => {
    setPendingCount(loadQueue(userId).length);
  }, [userId]);

  const enqueueOffline = useCallback(
    (op) => {
      const queue = [...loadQueue(userId), op];
      saveQueue(userId, queue);
      setPendingCount(queue.length);
    },
    [userId]
  );

  // Tenta a operação no Supabase; se não houver internet ou a rede falhar,
  // guarda na fila offline em vez de mostrar erro.
  const runOrQueue = useCallback(
    async (onlineOp, queueOp) => {
      if (!navigator.onLine) {
        enqueueOffline(queueOp);
        return { status: "queued" };
      }
      try {
        const { error } = await onlineOp();
        if (error) return { status: "error", error };
        return { status: "ok" };
      } catch {
        // fetch lança exceção em falha real de rede (offline, DNS etc.)
        enqueueOffline(queueOp);
        return { status: "queued" };
      }
    },
    [enqueueOffline]
  );

  const flushOfflineQueue = useCallback(async () => {
    if (!userId) return;
    let queue = loadQueue(userId);
    if (queue.length === 0) return;
    while (queue.length > 0) {
      const op = queue[0];
      try {
        const { error } =
          op.kind === "delete"
            ? await supabase.from(TABLE).delete().eq("id", op.row.id)
            : await supabase.from(TABLE).upsert(op.row);
        if (error) throw error;
        queue = queue.slice(1);
        saveQueue(userId, queue);
        setPendingCount(queue.length);
      } catch {
        break;
      }
    }
    if (queue.length === 0) {
      const { data, error } = await supabase.from(TABLE).select("*");
      if (!error && data) {
        const next = rowsToEntries(data);
        setEntries(next);
        saveCache(userId, next);
      }
      showToast("Sincronizado", "success");
    }
  }, [showToast, userId]);

  useEffect(() => {
    function handleOnline() {
      setIsOnline(true);
      if (loadQueue(userId).length > 0) {
        showToast("Conectado — sincronizando alterações…", "success");
        flushOfflineQueue();
      }
    }
    function handleOffline() {
      setIsOnline(false);
      showToast("Sem conexão — as alterações serão salvas localmente", "error");
    }
    window.addEventListener("online", handleOnline);
    window.addEventListener("offline", handleOffline);
    return () => {
      window.removeEventListener("online", handleOnline);
      window.removeEventListener("offline", handleOffline);
    };
  }, [flushOfflineQueue, showToast, userId]);

  // Salva uma cópia local sempre que os registros mudam (depois do carregamento
  // inicial), para o app conseguir abrir com os últimos dados mesmo sem internet.
  useEffect(() => {
    if (loading || !userId) return;
    saveCache(userId, entries);
  }, [entries, loading, userId]);

  // Carrega os registros do Supabase e assina atualizações em tempo real,
  // para que mudanças feitas em outro dispositivo apareçam sem recarregar a página.
  useEffect(() => {
    if (!supabaseConfigured) {
      setSaveError(true);
      setLoading(false);
      showToast("Supabase não configurado (veja .env.example)", "error");
      return;
    }
    if (!userId) {
      // Sem sessão: nada pra carregar (tela de login é exibida em vez do app).
      setEntries({});
      setLoading(false);
      return;
    }

    let cancelled = false;
    setLoading(true);

    (async () => {
      try {
        if (!navigator.onLine) throw new Error("offline");
        const { data, error } = await supabase.from(TABLE).select("*");
        if (cancelled) return;
        if (error) throw error;
        const next = rowsToEntries(data);
        setEntries(next);
        saveCache(userId, next);
      } catch {
        if (cancelled) return;
        const cached = loadCache(userId);
        setEntries(cached);
        showToast(
          Object.keys(cached).length > 0
            ? "Sem conexão — mostrando dados salvos localmente"
            : "Sem conexão e sem dados salvos localmente ainda",
          "error"
        );
      }
      if (!cancelled) setLoading(false);
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
  }, [showToast, userId]);

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

  // Todos os registros já lançados, de qualquer mês — base das estatísticas gerais.
  const allEntries = useMemo(() => {
    const list = [];
    for (const [dayKey, dayList] of Object.entries(entries)) {
      for (const e of dayList) list.push({ dayKey, ...e });
    }
    return list;
  }, [entries]);

  const statsGeral = useMemo(() => {
    const total = allEntries.reduce((s, e) => s + (Number(e.value) || 0), 0);
    const recebido = allEntries
      .filter((e) => e.pago)
      .reduce((s, e) => s + (Number(e.value) || 0), 0);
    return { count: allEntries.length, total, recebido, aReceber: total - recebido };
  }, [allEntries]);

  const statsPorTipo = useMemo(() => {
    const tipos = [
      { id: "plantao", label: "Plantão" },
      { id: "remocao", label: "Remoção" },
      { id: "evento", label: "Evento" },
    ];
    return tipos.map((t) => {
      const items = allEntries.filter((e) => e.type === t.id);
      const total = items.reduce((s, e) => s + (Number(e.value) || 0), 0);
      const recebido = items
        .filter((e) => e.pago)
        .reduce((s, e) => s + (Number(e.value) || 0), 0);
      return {
        ...t,
        count: items.length,
        total,
        recebido,
        aReceber: total - recebido,
        media: items.length ? total / items.length : 0,
      };
    });
  }, [allEntries]);

  const statsPorEmpresa = useMemo(() => {
    const groups = new Map();
    allEntries.forEach((e) => {
      if (e.type === "plantao") return;
      const key = e.empresa || "Sem empresa";
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(e);
    });
    return Array.from(groups.entries())
      .map(([empresa, items]) => {
        const total = items.reduce((s, e) => s + (Number(e.value) || 0), 0);
        const recebido = items
          .filter((e) => e.pago)
          .reduce((s, e) => s + (Number(e.value) || 0), 0);
        return { empresa, count: items.length, total, recebido, aReceber: total - recebido };
      })
      .sort((a, b) => b.total - a.total);
  }, [allEntries]);

  const statsPorMes = useMemo(() => {
    const groups = new Map();
    allEntries.forEach((e) => {
      const key = e.dayKey.slice(0, 7);
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(e);
    });
    return Array.from(groups.entries())
      .map(([mesKey, items]) => {
        const total = items.reduce((s, e) => s + (Number(e.value) || 0), 0);
        const recebido = items
          .filter((e) => e.pago)
          .reduce((s, e) => s + (Number(e.value) || 0), 0);
        const [y, m] = mesKey.split("-").map(Number);
        return {
          mesKey,
          label: `${MESES[m - 1]} ${y}`,
          plantoes: items.filter((e) => e.type === "plantao").length,
          remocoes: items.filter((e) => e.type === "remocao").length,
          eventos: items.filter((e) => e.type === "evento").length,
          total,
          recebido,
          aReceber: total - recebido,
        };
      })
      .sort((a, b) => b.mesKey.localeCompare(a.mesKey));
  }, [allEntries]);

  const statsPorMesChart = useMemo(
    () => [...statsPorMes].slice(0, 12).reverse(),
    [statsPorMes]
  );

  const statsPorEmpresaChart = useMemo(
    () => statsPorEmpresa.slice(0, 8),
    [statsPorEmpresa]
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
                stripHtml(e.obs),
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
    const outcome = await runOrQueue(
      () => supabase.from(TABLE).update({ pago: novoPago }).eq("id", result.id),
      { kind: "upsert", row: entryToRow(result.dayKey, { ...result, pago: novoPago }, userId) }
    );
    if (outcome.status === "error") {
      setSaveError(true);
      showToast("Não foi possível salvar", "error");
      return;
    }
    showToast(
      outcome.status === "queued"
        ? "Salvo offline — sincroniza quando reconectar"
        : novoPago
        ? "Marcado como pago"
        : "Marcado como a receber",
      "success"
    );
    pushUndo({
      label: "Status de pagamento restaurado",
      run: async () => {
        setEntries((prev) => {
          const list = prev[result.dayKey] || [];
          const idx = list.findIndex((e) => e.id === result.id);
          if (idx === -1) return prev;
          const updated = [...list];
          updated[idx] = { ...updated[idx], pago: result.pago };
          return { ...prev, [result.dayKey]: updated };
        });
        await runOrQueue(
          () => supabase.from(TABLE).update({ pago: result.pago }).eq("id", result.id),
          { kind: "upsert", row: entryToRow(result.dayKey, { ...result, pago: result.pago }, userId) }
        );
      },
    });
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
    setViewDay(todayKey());
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

    const outcome = await runOrQueue(
      () => supabase.from(TABLE).insert(entryToRow(targetDay, copy, userId)),
      { kind: "upsert", row: entryToRow(targetDay, copy, userId) }
    );
    if (outcome.status === "error") {
      setSaveError(true);
      showToast("Não foi possível duplicar", "error");
      return;
    }
    showToast(
      outcome.status === "queued"
        ? `Duplicado offline para ${formatShort(targetDay)} — sincroniza quando reconectar`
        : `Duplicado para ${formatShort(targetDay)}`,
      "success"
    );
    pushUndo({
      label: "Duplicação desfeita",
      run: async () => {
        setEntries((prev) => {
          const list = (prev[targetDay] || []).filter((e) => e.id !== copy.id);
          const next = { ...prev };
          if (list.length) next[targetDay] = list;
          else delete next[targetDay];
          return next;
        });
        await runOrQueue(
          () => supabase.from(TABLE).delete().eq("id", copy.id),
          { kind: "delete", row: { id: copy.id } }
        );
      },
    });
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

    const day = selectedDay;
    const wasEditing = !!editingId;
    const original = wasEditing ? (entries[day] || []).find((e) => e.id === editingId) : null;

    setEntries((prev) => {
      const list = prev[day] ? [...prev[day]] : [];
      if (editingId) {
        const idx = list.findIndex((e) => e.id === editingId);
        if (idx >= 0) list[idx] = record;
        else list.push(record);
      } else {
        list.push(record);
      }
      return { ...prev, [day]: list };
    });
    closeModal();

    const outcome = await runOrQueue(
      () => supabase.from(TABLE).upsert(entryToRow(day, record, userId)),
      { kind: "upsert", row: entryToRow(day, record, userId) }
    );
    if (outcome.status === "error") {
      setSaveError(true);
      showToast("Não foi possível salvar", "error");
      return;
    }
    showToast(
      outcome.status === "queued"
        ? "Salvo offline — sincroniza quando reconectar"
        : wasEditing
        ? "Alterações salvas"
        : "Registro salvo",
      "success"
    );
    if (wasEditing && original) {
      pushUndo({
        label: "Edição desfeita",
        run: async () => {
          setEntries((prev) => {
            const list = (prev[day] || []).map((e) => (e.id === original.id ? original : e));
            return { ...prev, [day]: list };
          });
          await runOrQueue(
            () => supabase.from(TABLE).upsert(entryToRow(day, original, userId)),
            { kind: "upsert", row: entryToRow(day, original, userId) }
          );
        },
      });
    } else {
      pushUndo({
        label: "Criação desfeita",
        run: async () => {
          setEntries((prev) => {
            const list = (prev[day] || []).filter((e) => e.id !== record.id);
            const next = { ...prev };
            if (list.length) next[day] = list;
            else delete next[day];
            return next;
          });
          await runOrQueue(
            () => supabase.from(TABLE).delete().eq("id", record.id),
            { kind: "delete", row: { id: record.id } }
          );
        },
      });
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
    const deletedEntry = (entries[day] || []).find((e) => e.id === idToDelete);
    setEntries((prev) => {
      const list = (prev[day] || []).filter((e) => e.id !== idToDelete);
      const next = { ...prev };
      if (list.length) next[day] = list;
      else delete next[day];
      return next;
    });
    closeModal();

    const outcome = await runOrQueue(
      () => supabase.from(TABLE).delete().eq("id", idToDelete),
      { kind: "delete", row: { id: idToDelete } }
    );
    if (outcome.status === "error") {
      setSaveError(true);
      showToast("Não foi possível excluir", "error");
      return;
    }
    showToast(
      outcome.status === "queued" ? "Excluído offline — sincroniza quando reconectar" : "Registro excluído",
      "success"
    );
    if (deletedEntry) {
      pushUndo({
        label: "Exclusão desfeita",
        run: async () => {
          setEntries((prev) => {
            const list = prev[day] ? [...prev[day]] : [];
            return { ...prev, [day]: [...list, deletedEntry] };
          });
          await runOrQueue(
            () => supabase.from(TABLE).insert(entryToRow(day, deletedEntry, userId)),
            { kind: "upsert", row: entryToRow(day, deletedEntry, userId) }
          );
        },
      });
    }
  };

  const moveEntry = async (fromDay, id, toDay) => {
    if (!fromDay || !toDay || fromDay === toDay) return;
    if (!supabaseConfigured) {
      showToast("Supabase não configurado", "error");
      return;
    }
    const movedEntry = (entries[fromDay] || []).find((e) => e.id === id);
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

    const outcome = await runOrQueue(
      () => supabase.from(TABLE).update({ day_key: toDay }).eq("id", id),
      { kind: "upsert", row: entryToRow(toDay, movedEntry || { id }, userId) }
    );
    if (outcome.status === "error") {
      setSaveError(true);
      showToast("Não foi possível salvar", "error");
      return;
    }
    showToast(
      outcome.status === "queued" ? "Movido offline — sincroniza quando reconectar" : "Movido para outra data",
      "success"
    );
    pushUndo({
      label: "Movimentação desfeita",
      run: async () => {
        setEntries((prev) => {
          const fromList = prev[toDay] || [];
          const entry = fromList.find((e) => e.id === id);
          if (!entry) return prev;
          const newFromList = fromList.filter((e) => e.id !== id);
          const toList = prev[fromDay] ? [...prev[fromDay]] : [];
          toList.push(entry);
          const next = { ...prev };
          if (newFromList.length) next[toDay] = newFromList;
          else delete next[toDay];
          next[fromDay] = toList;
          return next;
        });
        await runOrQueue(
          () => supabase.from(TABLE).update({ day_key: fromDay }).eq("id", id),
          { kind: "upsert", row: entryToRow(fromDay, movedEntry || { id }, userId) }
        );
      },
    });
  };

  const isValid =
    form.value &&
    parseBRL(form.value) > 0 &&
    (form.type === "plantao"
      ? form.local.trim()
      : form.type === "evento"
      ? form.local.trim() && form.empresa.trim()
      : form.empresa.trim());

  if (session === undefined) {
    return (
      <div style={styles.authScreen}>
        <style>{globalCss}</style>
        <div style={styles.loadingBox}>
          <Loader2 size={22} className="spin" />
          <span>verificando sessão…</span>
        </div>
      </div>
    );
  }

  if (passwordRecovery) {
    return (
      <>
        <style>{globalCss}</style>
        <ResetPasswordScreen showToast={showToast} onDone={() => setPasswordRecovery(false)} />
      </>
    );
  }

  if (!session) {
    return (
      <>
        <style>{globalCss}</style>
        <AuthScreen showToast={showToast} />
      </>
    );
  }

  return (
    <>
    <div style={styles.app} className="app-main">
      <style>{globalCss}</style>

      <header style={styles.header}>
        <div style={styles.headerTop}>
          <div style={styles.brand}>
            <Logo size={34} />
            <div>
              <div style={styles.brandTitle}>Plantões</div>
              <div style={styles.brandSub}>controle de escala &amp; financeiro</div>
            </div>
          </div>
          {!isOnline && (
            <div style={styles.offlineWarning}>
              <Circle size={8} fill="#8C6D1B" />
              offline{pendingCount > 0 ? ` — ${pendingCount} pendente${pendingCount > 1 ? "s" : ""}` : ""}
            </div>
          )}
          {saveError && (
            <div style={styles.saveWarning}>não foi possível salvar</div>
          )}
          <button
            className="btn-lift"
            style={styles.logoutBtn}
            onClick={() => supabase.auth.signOut()}
            title={session?.user?.email}
          >
            <LogOut size={13} />
            sair
          </button>
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
          <button
            onClick={() => setActiveTab("estatisticas")}
            className="pill-btn"
            style={{
              ...styles.tabBtn,
              ...(activeTab === "estatisticas" ? styles.tabBtnActive : {}),
            }}
          >
            <BarChart3 size={14} />
            estatísticas
          </button>
        </div>

        {activeTab === "estatisticas" && (
          <div style={styles.searchWrap}>
            <div style={styles.summaryBar}>
              <StatCard label="registros no total" value={String(statsGeral.count)} />
              <StatCard label="valor total" value={currency(statsGeral.total)} />
              <StatCard
                label="recebido"
                value={currency(statsGeral.recebido)}
                color="#206B3C"
                bg="#E2F2E7"
              />
              <StatCard
                label="a receber"
                value={currency(statsGeral.aReceber)}
                color="#8C6D1B"
                bg="#F6EFDD"
              />
            </div>

            <p style={styles.statsSectionTitle}>Por tipo</p>
            <div style={styles.searchDropdown}>
              {statsPorTipo.some((t) => t.total > 0) && (
                <div style={styles.chartWrap}>
                  <ResponsiveContainer width="100%" height={130}>
                    <PieChart>
                      <Pie
                        data={statsPorTipo.filter((t) => t.total > 0)}
                        dataKey="total"
                        nameKey="label"
                        innerRadius={34}
                        outerRadius={58}
                        paddingAngle={2}
                      >
                        {statsPorTipo
                          .filter((t) => t.total > 0)
                          .map((t) => (
                            <Cell key={t.id} fill={TYPE_CHART_COLORS[t.id]} />
                          ))}
                      </Pie>
                      <Tooltip formatter={(v) => currency(v)} />
                      <Legend iconSize={8} wrapperStyle={{ fontSize: 11 }} />
                    </PieChart>
                  </ResponsiveContainer>
                </div>
              )}
              <div style={styles.searchScrollArea}>
                <table style={styles.searchTable}>
                  <thead>
                    <tr>
                      <th style={styles.searchTh}>Tipo</th>
                      <th style={{ ...styles.searchTh, textAlign: "right" }}>Qtd.</th>
                      <th style={{ ...styles.searchTh, textAlign: "right" }}>Total</th>
                      <th style={{ ...styles.searchTh, textAlign: "right" }}>Recebido</th>
                      <th style={{ ...styles.searchTh, textAlign: "right" }}>A receber</th>
                      <th style={{ ...styles.searchTh, textAlign: "right" }}>Média</th>
                    </tr>
                  </thead>
                  <tbody>
                    {statsPorTipo.map((t) => (
                      <tr key={t.id}>
                        <td style={styles.searchTdName}>{t.label}</td>
                        <td style={{ ...styles.searchTd, textAlign: "right" }}>{t.count}</td>
                        <td style={{ ...styles.searchTd, textAlign: "right" }}>{currency(t.total)}</td>
                        <td style={{ ...styles.searchTd, textAlign: "right" }}>{currency(t.recebido)}</td>
                        <td style={{ ...styles.searchTd, textAlign: "right" }}>{currency(t.aReceber)}</td>
                        <td style={{ ...styles.searchTd, textAlign: "right" }}>{currency(t.media)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            <p style={styles.statsSectionTitle}>Por empresa</p>
            <div style={styles.searchDropdown}>
              {statsPorEmpresaChart.length > 0 && (
                <div style={styles.chartWrap}>
                  <ResponsiveContainer width="100%" height={Math.min(180, Math.max(90, statsPorEmpresaChart.length * 26))}>
                    <BarChart data={statsPorEmpresaChart} layout="vertical" margin={{ top: 2, left: 8, right: 12, bottom: 2 }}>
                      <CartesianGrid strokeDasharray="3 3" horizontal={false} />
                      <XAxis type="number" tickFormatter={(v) => currency(v)} fontSize={10} height={20} />
                      <YAxis type="category" dataKey="empresa" width={90} fontSize={10} />
                      <Tooltip formatter={(v) => currency(v)} />
                      <Bar dataKey="total" name="Total" fill="#1C2B39" radius={[0, 3, 3, 0]} barSize={14} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              )}
              {statsPorEmpresa.length === 0 ? (
                <div style={styles.searchEmpty}>Nenhum registro com empresa ainda.</div>
              ) : (
                <div style={styles.searchScrollArea}>
                  <table style={styles.searchTable}>
                    <thead>
                      <tr>
                        <th style={styles.searchTh}>Empresa</th>
                        <th style={{ ...styles.searchTh, textAlign: "right" }}>Qtd.</th>
                        <th style={{ ...styles.searchTh, textAlign: "right" }}>Total</th>
                        <th style={{ ...styles.searchTh, textAlign: "right" }}>Recebido</th>
                        <th style={{ ...styles.searchTh, textAlign: "right" }}>A receber</th>
                      </tr>
                    </thead>
                    <tbody>
                      {statsPorEmpresa.map((e) => (
                        <tr key={e.empresa}>
                          <td style={styles.searchTdName}>{e.empresa}</td>
                          <td style={{ ...styles.searchTd, textAlign: "right" }}>{e.count}</td>
                          <td style={{ ...styles.searchTd, textAlign: "right" }}>{currency(e.total)}</td>
                          <td style={{ ...styles.searchTd, textAlign: "right" }}>{currency(e.recebido)}</td>
                          <td style={{ ...styles.searchTd, textAlign: "right" }}>{currency(e.aReceber)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>

            <p style={styles.statsSectionTitle}>Por mês</p>
            <div style={styles.searchDropdown}>
              {statsPorMesChart.length > 0 && (
                <div style={styles.chartWrap}>
                  <ResponsiveContainer width="100%" height={150}>
                    <BarChart data={statsPorMesChart} margin={{ top: 2, left: 4, right: 8, bottom: 2 }}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} />
                      <XAxis dataKey="label" fontSize={10} />
                      <YAxis tickFormatter={(v) => currency(v)} fontSize={10} width={60} />
                      <Tooltip formatter={(v) => currency(v)} />
                      <Legend iconSize={8} wrapperStyle={{ fontSize: 11 }} />
                      <Bar dataKey="recebido" name="Recebido" stackId="v" fill="#2F8F52" barSize={16} />
                      <Bar dataKey="aReceber" name="A receber" stackId="v" fill="#B8912B" radius={[3, 3, 0, 0]} barSize={16} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              )}
              {statsPorMes.length === 0 ? (
                <div style={styles.searchEmpty}>Nenhum registro ainda.</div>
              ) : (
                <div style={styles.searchScrollArea}>
                  <table style={styles.searchTable}>
                    <thead>
                      <tr>
                        <th style={styles.searchTh}>Mês</th>
                        <th style={{ ...styles.searchTh, textAlign: "right" }}>Plantões</th>
                        <th style={{ ...styles.searchTh, textAlign: "right" }}>Remoções</th>
                        <th style={{ ...styles.searchTh, textAlign: "right" }}>Eventos</th>
                        <th style={{ ...styles.searchTh, textAlign: "right" }}>Total</th>
                        <th style={{ ...styles.searchTh, textAlign: "right" }}>Recebido</th>
                        <th style={{ ...styles.searchTh, textAlign: "right" }}>A receber</th>
                      </tr>
                    </thead>
                    <tbody>
                      {statsPorMes.map((m) => (
                        <tr key={m.mesKey}>
                          <td style={styles.searchTdName}>{m.label}</td>
                          <td style={{ ...styles.searchTd, textAlign: "right" }}>{m.plantoes}</td>
                          <td style={{ ...styles.searchTd, textAlign: "right" }}>{m.remocoes}</td>
                          <td style={{ ...styles.searchTd, textAlign: "right" }}>{m.eventos}</td>
                          <td style={{ ...styles.searchTd, textAlign: "right" }}>{currency(m.total)}</td>
                          <td style={{ ...styles.searchTd, textAlign: "right" }}>{currency(m.recebido)}</td>
                          <td style={{ ...styles.searchTd, textAlign: "right" }}>{currency(m.aReceber)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </div>
        )}

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
                placeholder="Nome do plantão, remoção ou empresa…"
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
                          <th style={{ ...styles.searchTh, ...styles.searchThBold }}>Data</th>
                          <th style={{ ...styles.searchTh, ...styles.searchThBold }}>Nome</th>
                          <th style={{ ...styles.searchTh, ...styles.searchThBold }}>Empresa</th>
                          <th style={{ ...styles.searchTh, ...styles.searchThBold, textAlign: "center" }}>
                            Pago
                          </th>
                          <th style={{ ...styles.searchTh, ...styles.searchThBold, textAlign: "right" }}>
                            Valor
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {searchResults.map((r, idx) => (
                          <tr
                            key={r.id}
                            style={{
                              ...styles.searchRow,
                              background: idx % 2 === 1 ? "#F4F2ED" : "#fff",
                            }}
                            onClick={() => goToSearchResult(r)}
                          >
                            <td style={styles.searchTd}>{formatShortWithWeekday(r.dayKey)}</td>
                            <td style={styles.searchTdName}>
                              {r.type === "remocao"
                                ? r.empresa || "remoção"
                                : r.local || (r.type === "evento" ? "evento" : "plantão")}
                            </td>
                            <td style={styles.searchTdName}>
                              {r.type === "plantao" ? "—" : r.empresa || "—"}
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

            {searchResults.length > 0 && (
              <div style={{ ...styles.summaryBar, marginTop: 12 }}>
                <SummaryChip
                  icon={<Circle size={14} />}
                  label="a receber"
                  value={searchResults
                    .filter((r) => !r.pago)
                    .reduce((s, r) => s + (Number(r.value) || 0), 0)}
                  color="#8C6D1B"
                  bg="#F6EFDD"
                />
                <SummaryChip
                  icon={<CheckCircle2 size={14} />}
                  label="pago"
                  value={searchResults
                    .filter((r) => r.pago)
                    .reduce((s, r) => s + (Number(r.value) || 0), 0)}
                  color="#206B3C"
                  bg="#E2F2E7"
                />
                <SummaryChip
                  icon={<Search size={14} />}
                  label="total"
                  value={searchResults.reduce((s, r) => s + (Number(r.value) || 0), 0)}
                  color="#F7F5F0"
                  bg="#1C2B39"
                />
              </div>
            )}
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
              const isSelected = dayKey === viewDay;
              return (
                <div
                  key={`${i}-${isSelected}`}
                  className="cal-cell"
                  style={{
                    ...styles.cell,
                    ...(isToday ? styles.cellToday : {}),
                    ...(isSelected ? styles.cellSelected : {}),
                    ...(isDragOver ? styles.cellDragOver : {}),
                  }}
                  onClick={() => setViewDay((prev) => (prev === dayKey ? null : dayKey))}
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

          <div style={styles.dayPanel}>
            <div style={styles.dayPanelHeader}>
              <span style={styles.dayPanelTitle}>
                {viewDay ? formatFullDate(viewDay) : "nenhum dia selecionado"}
              </span>
              {viewDay && (
                <button
                  type="button"
                  className="btn-lift"
                  style={styles.dayPanelAddBtn}
                  onClick={() => openAddModal(viewDay)}
                >
                  <Plus size={13} /> adicionar
                </button>
              )}
            </div>
            {!viewDay ? (
              <div style={styles.dayPanelEmpty}>Clique em um dia do calendário para ver os detalhes.</div>
            ) : (entries[viewDay] || []).length === 0 ? (
              <div style={styles.dayPanelEmpty}>Nenhum registro neste dia.</div>
            ) : (
              <div style={styles.dayPanelList}>
                {(entries[viewDay] || []).map((e) => (
                  <button
                    key={e.id}
                    type="button"
                    className="btn-lift"
                    style={styles.dayPanelItem}
                    onClick={() => openEditModal(viewDay, e)}
                  >
                    <span style={styles.dayPanelItemIcon}>
                      {e.type === "plantao" ? (
                        <Stethoscope size={15} />
                      ) : e.type === "evento" ? (
                        <Presentation size={15} />
                      ) : (
                        <Truck size={15} />
                      )}
                    </span>
                    <span style={styles.dayPanelItemBody}>
                      <span style={styles.dayPanelItemTitle}>
                        {e.type === "remocao"
                          ? e.empresa || "remoção"
                          : e.local || (e.type === "evento" ? "evento" : "plantão")}
                      </span>
                      <span style={styles.dayPanelItemSub}>
                        {e.iH ? `${e.iH}:${e.iM} – ${e.fH}:${e.fM}` : "sem horário"}
                        {e.empresa && e.type !== "remocao" ? ` · ${e.empresa}` : ""}
                      </span>
                    </span>
                    <span style={styles.dayPanelItemRight}>
                      <span style={styles.dayPanelItemValue}>{currency(e.value)}</span>
                      <span
                        style={{
                          ...styles.dayPanelBadge,
                          ...(e.pago ? styles.dayPanelBadgePago : styles.dayPanelBadgePendente),
                        }}
                      >
                        {e.pago ? "pago" : "a receber"}
                      </span>
                    </span>
                  </button>
                ))}
              </div>
            )}
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
              <ColorGrid
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
                        <input
                          type="time"
                          style={styles.timeInput}
                          value={form.iH && form.iM ? `${form.iH}:${form.iM}` : ""}
                          onChange={(e) => {
                            const [h, m] = e.target.value.split(":");
                            setForm((f) => ({ ...f, iH: h || "", iM: m || "" }));
                          }}
                        />
                      </div>
                      <ArrowRight size={16} style={styles.timeArrow} />
                      <div style={styles.timeGroup}>
                        <div style={styles.timeGroupLabel}>fim</div>
                        <input
                          type="time"
                          style={styles.timeInput}
                          value={form.fH && form.fM ? `${form.fH}:${form.fM}` : ""}
                          onChange={(e) => {
                            const [h, m] = e.target.value.split(":");
                            setForm((f) => ({ ...f, fH: h || "", fM: m || "" }));
                          }}
                        />
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
                        <input
                          type="time"
                          style={styles.timeInput}
                          value={form.iH && form.iM ? `${form.iH}:${form.iM}` : ""}
                          onChange={(e) => {
                            const [h, m] = e.target.value.split(":");
                            setForm((f) => ({ ...f, iH: h || "", iM: m || "" }));
                          }}
                        />
                      </div>
                      <ArrowRight size={16} style={styles.timeArrow} />
                      <div style={styles.timeGroup}>
                        <div style={styles.timeGroupLabel}>fim</div>
                        <input
                          type="time"
                          style={styles.timeInput}
                          value={form.fH && form.fM ? `${form.fH}:${form.fM}` : ""}
                          onChange={(e) => {
                            const [h, m] = e.target.value.split(":");
                            setForm((f) => ({ ...f, fH: h || "", fM: m || "" }));
                          }}
                        />
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
                <RichTextEditor
                  value={form.obs}
                  onChange={(html) => setForm((f) => ({ ...f, obs: html }))}
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

function StatCard({ label, value, color = "#1C2B39", bg = "#F1EFE9" }) {
  return (
    <div style={{ ...styles.summaryChip, background: bg, color }}>
      <div style={styles.summaryChipIconRow}>
        <span style={styles.summaryChipLabel}>{label}</span>
      </div>
      <div style={styles.summaryChipValue}>{value}</div>
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

function Logo({ size = 40 }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 120 120"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
      role="img"
      aria-label="Logotipo Plantões"
    >
      <rect x="4" y="4" width="112" height="112" rx="28" fill="#1C2B39" />
      <polyline
        points="18,60 34,60 40,50 46,70 52,26 58,88 64,60 102,60"
        fill="none"
        stroke="#F7F5F0"
        strokeWidth="6"
        strokeLinejoin="round"
        strokeLinecap="round"
      />
      <circle cx="102" cy="60" r="6" fill="#2D6E6E" />
    </svg>
  );
}

function translateAuthError(err) {
  const msg = (err?.message || "").toLowerCase();
  const code = err?.code || err?.error_code || "";
  if (code === "invalid_credentials" || msg.includes("invalid login credentials")) {
    return "E-mail ou senha incorretos. Confira os dados e tente novamente.";
  }
  if (code === "email_not_confirmed" || msg.includes("email not confirmed")) {
    return "Este e-mail ainda não foi confirmado. Verifique sua caixa de entrada (e o spam) e clique no link de confirmação antes de entrar.";
  }
  if (code === "user_already_exists" || msg.includes("already registered") || msg.includes("already been registered")) {
    return "Este e-mail já está cadastrado. Tente entrar ou use \"Esqueci minha senha\".";
  }
  if (msg.includes("user not found")) {
    return "Não existe conta cadastrada com este e-mail.";
  }
  if (code === "over_email_send_rate_limit" || msg.includes("rate limit")) {
    return "Muitas tentativas em pouco tempo. Aguarde alguns minutos e tente de novo.";
  }
  if (msg.includes("password should be at least")) {
    return "A senha precisa ter pelo menos 6 caracteres.";
  }
  if (msg.includes("unable to validate email") || msg.includes("invalid email")) {
    return "E-mail inválido. Confira se foi digitado corretamente.";
  }
  if (msg.includes("failed to fetch") || msg.includes("network")) {
    return "Não foi possível conectar. Verifique sua internet e tente novamente.";
  }
  return err?.message || "Não foi possível continuar. Tente novamente.";
}

function AuthScreen({ showToast }) {
  const [mode, setMode] = useState("login"); // "login" | "signup" | "forgot"
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [info, setInfo] = useState("");
  const [error, setError] = useState("");

  const switchMode = (next) => {
    setMode(next);
    setInfo("");
    setError("");
  };

  const handleSubmit = async (evt) => {
    evt.preventDefault();
    setError("");
    if (!supabaseConfigured) {
      setError("Supabase não configurado");
      return;
    }
    if (!email.trim()) {
      setError("Preencha o e-mail");
      return;
    }
    if (mode !== "forgot" && !password) {
      setError("Preencha a senha");
      return;
    }
    if (mode === "signup" && password.length < 6) {
      setError("A senha precisa ter pelo menos 6 caracteres");
      return;
    }
    setSubmitting(true);
    setInfo("");
    try {
      if (mode === "login") {
        const { error: err } = await supabase.auth.signInWithPassword({
          email: email.trim(),
          password,
        });
        if (err) throw err;
      } else if (mode === "signup") {
        const { data, error: err } = await supabase.auth.signUp({
          email: email.trim(),
          password,
        });
        if (err) throw err;
        if (data.session) {
          showToast("Conta criada!", "success");
        } else {
          setInfo("Conta criada! Verifique seu e-mail para confirmar antes de entrar.");
        }
      } else {
        const { error: err } = await supabase.auth.resetPasswordForEmail(email.trim(), {
          redirectTo: window.location.origin,
        });
        if (err) throw err;
        setInfo("Link de recuperação enviado! Confira seu e-mail (e a caixa de spam).");
      }
    } catch (err) {
      setError(translateAuthError(err));
      showToast(translateAuthError(err), "error");
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <div style={styles.authScreen}>
      <div style={styles.authCard}>
        <div style={styles.authLogoWrap}>
          <Logo size={64} />
        </div>
        <h1 style={styles.authTitle}>Plantões</h1>
        <p style={styles.authSubtitle}>controle de escala &amp; financeiro</p>

        {mode !== "forgot" && (
          <div style={styles.typeToggle}>
            <button
              type="button"
              className="btn-lift"
              style={{
                ...styles.typeBtn,
                ...(mode === "login" ? styles.typeBtnActivePlantao : {}),
              }}
              onClick={() => switchMode("login")}
            >
              Entrar
            </button>
            <button
              type="button"
              className="btn-lift"
              style={{
                ...styles.typeBtn,
                ...(mode === "signup" ? styles.typeBtnActiveEvento : {}),
              }}
              onClick={() => switchMode("signup")}
            >
              Criar conta
            </button>
          </div>
        )}

        <form onSubmit={handleSubmit} style={styles.authForm}>
          <Field icon={<Mail size={14} />} label="E-mail">
            <input
              type="email"
              autoComplete="email"
              style={styles.input}
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              placeholder="seuemail@exemplo.com"
            />
          </Field>

          {mode !== "forgot" && (
            <Field icon={<Lock size={14} />} label="Senha">
              <input
                type="password"
                autoComplete={mode === "login" ? "current-password" : "new-password"}
                style={styles.input}
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                placeholder="mínimo 6 caracteres"
              />
            </Field>
          )}

          {mode === "login" && (
            <button type="button" style={styles.authForgotLink} onClick={() => switchMode("forgot")}>
              Esqueci minha senha
            </button>
          )}

          {error && <p style={styles.authError}>{error}</p>}
          {info && <p style={styles.authInfo}>{info}</p>}

          <button
            type="submit"
            className="btn-lift"
            style={{ ...styles.saveBtn, ...styles.authSubmitBtn, opacity: submitting ? 0.7 : 1 }}
            disabled={submitting}
          >
            {submitting ? (
              <Loader2 size={15} className="spin" />
            ) : mode === "login" ? (
              "Entrar"
            ) : mode === "signup" ? (
              "Criar conta"
            ) : (
              "Enviar link de recuperação"
            )}
          </button>

          {mode === "forgot" && (
            <button type="button" style={styles.authForgotLink} onClick={() => switchMode("login")}>
              Voltar para o login
            </button>
          )}
        </form>
      </div>
    </div>
  );
}

function ResetPasswordScreen({ showToast, onDone }) {
  const [password, setPassword] = useState("");
  const [confirm, setConfirm] = useState("");
  const [submitting, setSubmitting] = useState(false);

  const handleSubmit = async (evt) => {
    evt.preventDefault();
    if (password.length < 6) {
      showToast("A senha precisa ter pelo menos 6 caracteres", "error");
      return;
    }
    if (password !== confirm) {
      showToast("As senhas não coincidem", "error");
      return;
    }
    setSubmitting(true);
    try {
      const { error } = await supabase.auth.updateUser({ password });
      if (error) throw error;
      showToast("Senha atualizada!", "success");
      onDone();
    } catch (err) {
      showToast(err.message || "Não foi possível atualizar a senha", "error");
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <div style={styles.authScreen}>
      <div style={styles.authCard}>
        <div style={styles.authLogoWrap}>
          <Logo size={64} />
        </div>
        <h1 style={styles.authTitle}>Nova senha</h1>
        <p style={styles.authSubtitle}>Escolha uma nova senha para sua conta</p>
        <form onSubmit={handleSubmit} style={styles.authForm}>
          <Field icon={<Lock size={14} />} label="Nova senha">
            <input
              type="password"
              autoComplete="new-password"
              style={styles.input}
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              placeholder="mínimo 6 caracteres"
            />
          </Field>
          <Field icon={<Lock size={14} />} label="Confirmar nova senha">
            <input
              type="password"
              autoComplete="new-password"
              style={styles.input}
              value={confirm}
              onChange={(e) => setConfirm(e.target.value)}
              placeholder="repita a senha"
            />
          </Field>
          <button
            type="submit"
            className="btn-lift"
            style={{ ...styles.saveBtn, ...styles.authSubmitBtn, opacity: submitting ? 0.7 : 1 }}
            disabled={submitting}
          >
            {submitting ? <Loader2 size={15} className="spin" /> : "Salvar nova senha"}
          </button>
        </form>
      </div>
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
    .cal-cell { height: 74px !important; padding: 4px !important; }
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
  chartWrap: {
    padding: "10px 6px 2px",
    borderBottom: "1px solid #EFEBE1",
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
  statsSectionTitle: {
    fontSize: 12.5,
    fontWeight: 700,
    color: "#1C2B39",
    margin: "16px 0 6px",
    textTransform: "uppercase",
    letterSpacing: "0.03em",
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
  searchThBold: {
    fontSize: 12.5,
    fontWeight: 700,
    color: "#1C2B39",
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
  offlineWarning: {
    display: "flex",
    alignItems: "center",
    gap: 5,
    fontSize: 11,
    fontWeight: 600,
    color: "#8C6D1B",
    background: "#F6EFDD",
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
  authScreen: {
    minHeight: "100vh",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    background: "#F7F5F0",
    padding: 20,
    fontFamily: "'Inter', -apple-system, sans-serif",
  },
  authCard: {
    background: "#FFFDF9",
    borderRadius: 16,
    padding: "32px 28px",
    width: "100%",
    maxWidth: 380,
    boxShadow: "0 10px 30px rgba(28,43,57,0.12)",
  },
  authLogoWrap: {
    display: "flex",
    justifyContent: "center",
    marginBottom: 16,
  },
  authTitle: {
    fontFamily: "'Space Grotesk', sans-serif",
    fontWeight: 700,
    fontSize: 24,
    color: "#1C2B39",
    margin: "0 0 4px",
    textAlign: "center",
  },
  authSubtitle: {
    fontSize: 13,
    color: "#5B6B75",
    margin: "0 0 24px",
    textAlign: "center",
  },
  authForm: {
    display: "flex",
    flexDirection: "column",
    gap: 4,
    marginTop: 20,
  },
  authInfo: {
    fontSize: 12.5,
    color: "#206B3C",
    background: "#E2F2E7",
    borderRadius: 8,
    padding: "8px 10px",
    marginBottom: 4,
  },
  authError: {
    fontSize: 12.5,
    color: "#A02B2B",
    background: "#FBE4E4",
    borderRadius: 8,
    padding: "8px 10px",
    marginBottom: 4,
  },
  authSubmitBtn: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    width: "100%",
    padding: "10px 16px",
    marginTop: 8,
  },
  authForgotLink: {
    border: "none",
    background: "transparent",
    color: "#5B6B75",
    fontSize: 12,
    textDecoration: "underline",
    cursor: "pointer",
    padding: "4px 0",
    textAlign: "center",
  },
  logoutBtn: {
    display: "flex",
    alignItems: "center",
    gap: 5,
    border: "1px solid #E0DDD3",
    background: "#fff",
    borderRadius: 8,
    padding: "6px 10px",
    fontSize: 12,
    cursor: "pointer",
    color: "#5B6B75",
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
  emptyCell: { height: 92 },
  cell: {
    background: "#fff",
    border: "1px solid #E0DDD3",
    borderRadius: 10,
    padding: 6,
    height: 92,
    display: "flex",
    flexDirection: "column",
    gap: 4,
    overflow: "hidden",
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
  cellSelected: {
    background: "#E7F0FB",
    borderColor: "#A9C7ED",
    boxShadow: "0 0 0 2px #A9C7ED inset",
  },
  dayPanel: {
    marginTop: 14,
    background: "#fff",
    border: "1px solid #E0DDD3",
    borderRadius: 10,
    boxShadow: "0 4px 14px rgba(28,43,57,0.06)",
    padding: 12,
  },
  dayPanelHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    marginBottom: 10,
  },
  dayPanelTitle: {
    fontSize: 13.5,
    fontWeight: 700,
    color: "#1C2B39",
    textTransform: "capitalize",
  },
  dayPanelAddBtn: {
    display: "flex",
    alignItems: "center",
    gap: 4,
    fontSize: 12,
    fontWeight: 600,
    color: "#fff",
    background: "#1C2B39",
    border: "none",
    borderRadius: 8,
    padding: "6px 10px",
    cursor: "pointer",
  },
  dayPanelEmpty: {
    fontSize: 12.5,
    color: "#8A8578",
    padding: "10px 2px",
  },
  dayPanelList: {
    display: "flex",
    flexDirection: "column",
    gap: 6,
  },
  dayPanelItem: {
    display: "flex",
    alignItems: "center",
    gap: 10,
    width: "100%",
    textAlign: "left",
    border: "1px solid #E7E3D8",
    background: "#FCFBF8",
    borderRadius: 8,
    padding: "8px 10px",
    cursor: "pointer",
  },
  dayPanelItemIcon: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    width: 28,
    height: 28,
    borderRadius: 7,
    background: "#F1EFE9",
    color: "#1C2B39",
    flexShrink: 0,
  },
  dayPanelItemBody: {
    display: "flex",
    flexDirection: "column",
    gap: 2,
    flex: 1,
    minWidth: 0,
  },
  dayPanelItemTitle: {
    fontSize: 13,
    fontWeight: 600,
    color: "#1C2B39",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  dayPanelItemSub: {
    fontSize: 11.5,
    color: "#8A8578",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  dayPanelItemRight: {
    display: "flex",
    flexDirection: "column",
    alignItems: "flex-end",
    gap: 3,
    flexShrink: 0,
  },
  dayPanelItemValue: {
    fontSize: 13,
    fontWeight: 700,
    color: "#1C2B39",
  },
  dayPanelBadge: {
    fontSize: 10.5,
    fontWeight: 600,
    borderRadius: 6,
    padding: "2px 6px",
  },
  dayPanelBadgePago: {
    color: "#206B3C",
    background: "#E2F2E7",
  },
  dayPanelBadgePendente: {
    color: "#8C6D1B",
    background: "#F6EFDD",
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
  chipStack: {
    display: "flex",
    flexDirection: "column",
    gap: 3,
    overflowY: "auto",
    flex: 1,
    minHeight: 0,
  },
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
    flexShrink: 0,
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
  colorGrid: {
    display: "flex",
    flexWrap: "wrap",
    gap: 10,
    padding: "4px 0",
  },
  colorGridDot: {
    width: 30,
    height: 30,
    borderRadius: "50%",
    border: "2px solid transparent",
    cursor: "pointer",
    padding: 0,
  },
  colorGridDotActive: {
    border: "2px solid #1C2B39",
    boxShadow: "0 0 0 2px #FFFDF9, 0 0 0 4px #1C2B39",
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
  timeArrow: { marginTop: 26, color: "#5B6B75", flexShrink: 0 },
  timeInput: {
    border: "1px solid #E0DDD3",
    borderRadius: 8,
    padding: "9px 10px",
    fontSize: 13.5,
    fontFamily: "'Inter', sans-serif",
    color: "#1C2B39",
    background: "#fff",
    width: 130,
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
  richWrap: {
    border: "1px solid #E0DDD3",
    borderRadius: 8,
    background: "#fff",
    overflow: "hidden",
  },
  richToolbar: {
    display: "flex",
    alignItems: "center",
    gap: 2,
    padding: "5px 6px",
    borderBottom: "1px solid #E0DDD3",
    background: "#FBFAF7",
  },
  richToolbarBtn: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    width: 26,
    height: 26,
    border: "none",
    background: "transparent",
    borderRadius: 6,
    color: "#5B6B75",
    cursor: "pointer",
  },
  richEditable: {
    display: "block",
    width: "100%",
    minHeight: 60,
    padding: "9px 10px",
    fontSize: 13.5,
    fontFamily: "'Inter', sans-serif",
    color: "#1C2B39",
    border: "none",
    outline: "none",
    resize: "vertical",
    boxSizing: "border-box",
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
