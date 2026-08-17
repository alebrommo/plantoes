import { useState } from "react";
import { Bell, Plus, Trash2, CheckCircle2, Circle, ChevronDown, ChevronRight } from "lucide-react";

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

// Aba "lembretes": lista simples de avisos com data, cada um marcável como feito
// e com uma checklist própria de subtópicos.
export default function LembretesTab({
  lembretes,
  onAdd,
  onToggle,
  onDelete,
  onChangeData,
  onAddSubitem,
  onToggleSubitem,
  onDeleteSubitem,
  styles,
}) {
  const [newDate, setNewDate] = useState(todayKey());
  const [newTexto, setNewTexto] = useState("");
  const [expandedId, setExpandedId] = useState(null);
  const [subitemDraft, setSubitemDraft] = useState("");
  const [editingDateId, setEditingDateId] = useState(null);

  const items = Object.entries(lembretes)
    .flatMap(([dayKey, list]) => list.map((l) => ({ ...l, dayKey })))
    .sort((a, b) => (a.dayKey < b.dayKey ? -1 : a.dayKey > b.dayKey ? 1 : 0));

  const hoje = todayKey();

  function handleAdd() {
    if (!newTexto.trim() || !newDate) return;
    onAdd(newDate, newTexto);
    setNewTexto("");
  }

  function handleAddSubitem(l) {
    if (!subitemDraft.trim()) return;
    onAddSubitem(l.dayKey, l.id, subitemDraft);
    setSubitemDraft("");
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
            const isExpanded = expandedId === l.id;
            const subitens = l.subitens || [];
            const feitosCount = subitens.filter((s) => s.feito).length;
            return (
              <div key={l.id}>
                <div style={styles.lembretesRow}>
                  {editingDateId === l.id ? (
                    <input
                      type="date"
                      autoFocus
                      style={styles.lembretesRowDateInput}
                      defaultValue={l.dayKey}
                      onChange={(e) => {
                        onChangeData(l.dayKey, l.id, e.target.value);
                        setEditingDateId(null);
                      }}
                      onBlur={() => setEditingDateId(null)}
                    />
                  ) : (
                    <button
                      type="button"
                      className="btn-icon"
                      style={{
                        ...styles.lembretesRowDate,
                        ...(overdue ? styles.lembretesRowDateOverdue : {}),
                      }}
                      onClick={() => setEditingDateId(l.id)}
                      title="Mudar a data"
                    >
                      <span>{weekday}</span>
                      <span>{dayMonth}</span>
                    </button>
                  )}
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
                  {subitens.length > 0 && (
                    <span style={styles.lembretesProgressBadge}>
                      {feitosCount}/{subitens.length}
                    </span>
                  )}
                  <button
                    type="button"
                    className="btn-icon"
                    style={styles.lembretesExpandBtn}
                    onClick={() => setExpandedId((prev) => (prev === l.id ? null : l.id))}
                    aria-label={isExpanded ? "Ocultar subtópicos" : "Ver subtópicos"}
                  >
                    {isExpanded ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
                  </button>
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

                {isExpanded && (
                  <div style={styles.lembretesSubitensWrap}>
                    {subitens.map((s) => (
                      <div key={s.id} style={styles.lembretesSubitemRow}>
                        <button
                          type="button"
                          className="btn-icon"
                          style={{ display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}
                          onClick={() => onToggleSubitem(l.dayKey, l.id, s.id)}
                          aria-label={s.feito ? "Marcar subtópico como pendente" : "Marcar subtópico como feito"}
                        >
                          {s.feito ? <CheckCircle2 size={15} color="#5C3A88" /> : <Circle size={15} color="#8A8578" />}
                        </button>
                        <span
                          style={{
                            ...styles.lembretesSubitemText,
                            ...(s.feito ? styles.lembretesSubitemTextDone : {}),
                          }}
                        >
                          {s.texto}
                        </span>
                        <button
                          type="button"
                          className="btn-icon"
                          style={styles.lembretesSubitemDelBtn}
                          onClick={() => onDeleteSubitem(l.dayKey, l.id, s.id)}
                          aria-label="Excluir subtópico"
                        >
                          <Trash2 size={12} />
                        </button>
                      </div>
                    ))}
                    <div style={styles.lembretesSubitemAddRow}>
                      <input
                        type="text"
                        style={styles.lembretesSubitemInput}
                        placeholder="Novo subtópico…"
                        value={subitemDraft}
                        onChange={(e) => setSubitemDraft(e.target.value)}
                        onKeyDown={(e) => {
                          if (e.key === "Enter") handleAddSubitem(l);
                        }}
                      />
                      <button
                        type="button"
                        className="btn-icon"
                        style={styles.lembretesSubitemAddBtn}
                        onClick={() => handleAddSubitem(l)}
                        aria-label="Adicionar subtópico"
                      >
                        <Plus size={15} />
                      </button>
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
