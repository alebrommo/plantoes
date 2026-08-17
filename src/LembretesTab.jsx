import { useState } from "react";
import { Bell, Plus, Trash2, CheckCircle2, Circle } from "lucide-react";

function pad(n) {
  return String(n).padStart(2, "0");
}

function todayKey() {
  const t = new Date();
  return `${t.getFullYear()}-${pad(t.getMonth() + 1)}-${pad(t.getDate())}`;
}

function formatDayShort(dayKey) {
  const [y, m, d] = dayKey.split("-").map(Number);
  const date = new Date(y, m - 1, d);
  const weekday = date.toLocaleDateString("pt-BR", { weekday: "short" }).replace(".", "");
  return { weekday, dayMonth: `${pad(d)}/${pad(m)}` };
}

// Aba "lembretes": lista simples de avisos com data, cada um marcável como feito.
export default function LembretesTab({ lembretes, onAdd, onToggle, onDelete, styles }) {
  const [newDate, setNewDate] = useState(todayKey());
  const [newTexto, setNewTexto] = useState("");

  const items = Object.entries(lembretes)
    .flatMap(([dayKey, list]) => list.map((l) => ({ ...l, dayKey })))
    .sort((a, b) => (a.dayKey < b.dayKey ? -1 : a.dayKey > b.dayKey ? 1 : 0));

  const hoje = todayKey();

  function handleAdd() {
    if (!newTexto.trim() || !newDate) return;
    onAdd(newDate, newTexto);
    setNewTexto("");
  }

  return (
    <div style={styles.searchWrap}>
      <div style={styles.lembretesAddRow}>
        <input
          type="date"
          style={styles.lembretesDateInput}
          value={newDate}
          onChange={(e) => setNewDate(e.target.value)}
        />
        <input
          type="text"
          style={styles.lembretesTextInput}
          placeholder="Novo lembrete…"
          value={newTexto}
          onChange={(e) => setNewTexto(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") handleAdd();
          }}
        />
        <button type="button" className="btn-lift" style={styles.lembretesAddBtn} onClick={handleAdd}>
          <Plus size={14} /> adicionar
        </button>
      </div>

      {items.length === 0 ? (
        <div style={styles.searchEmpty}>
          <Bell size={18} style={{ opacity: 0.5, marginBottom: 4 }} />
          <div>Nenhum lembrete cadastrado ainda.</div>
        </div>
      ) : (
        <div style={styles.lembretesListWrap}>
          {items.map((l) => {
            const { weekday, dayMonth } = formatDayShort(l.dayKey);
            const overdue = !l.feito && l.dayKey < hoje;
            return (
              <div key={l.id} style={styles.lembretesRow}>
                <span
                  style={{
                    ...styles.lembretesRowDate,
                    ...(overdue ? styles.lembretesRowDateOverdue : {}),
                  }}
                >
                  <span>{weekday}</span>
                  <span>{dayMonth}</span>
                </span>
                <button
                  type="button"
                  className="btn-icon"
                  style={{ display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}
                  onClick={() => onToggle(l.dayKey, l.id)}
                  aria-label={l.feito ? "Marcar como pendente" : "Marcar como feito"}
                >
                  {l.feito ? <CheckCircle2 size={18} color="#5C3A88" /> : <Circle size={18} color="#8A8578" />}
                </button>
                <span
                  style={{
                    ...styles.lembretesRowText,
                    ...(l.feito ? styles.lembretesRowTextDone : {}),
                  }}
                >
                  {l.texto}
                </span>
                <button
                  type="button"
                  className="btn-icon"
                  style={{ display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, color: "#B5541F" }}
                  onClick={() => onDelete(l.dayKey, l.id)}
                  aria-label="Excluir lembrete"
                >
                  <Trash2 size={14} />
                </button>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
