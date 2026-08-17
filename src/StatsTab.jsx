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
import { styles, StatCard, TYPE_CHART_COLORS } from "./PlantoesApp";

// Aba "estatísticas" isolada num arquivo próprio, carregado sob demanda,
// para não incluir a biblioteca de gráficos (recharts) no pacote principal do app.
export default function StatsTab({
  statsGeral,
  statsPorTipo,
  statsPorEmpresaChart,
  statsPorEmpresa,
  statsPorMesChart,
  statsPorMes,
  fmtValue,
}) {
  return (
    <div style={styles.searchWrap}>
      <div style={styles.summaryBar}>
        <StatCard label="registros no total" value={String(statsGeral.count)} />
        <StatCard label="valor total" value={fmtValue(statsGeral.total)} />
        <StatCard
          label="recebido"
          value={fmtValue(statsGeral.recebido)}
          color="#206B3C"
          bg="#E2F2E7"
        />
        <StatCard
          label="a receber"
          value={fmtValue(statsGeral.aReceber)}
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
                <Tooltip formatter={(v) => fmtValue(v)} />
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
                  <td style={{ ...styles.searchTd, textAlign: "right" }}>{fmtValue(t.total)}</td>
                  <td style={{ ...styles.searchTd, textAlign: "right" }}>{fmtValue(t.recebido)}</td>
                  <td style={{ ...styles.searchTd, textAlign: "right" }}>{fmtValue(t.aReceber)}</td>
                  <td style={{ ...styles.searchTd, textAlign: "right" }}>{fmtValue(t.media)}</td>
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
                <XAxis type="number" tickFormatter={(v) => fmtValue(v)} fontSize={10} height={20} />
                <YAxis type="category" dataKey="empresa" width={90} fontSize={10} />
                <Tooltip formatter={(v) => fmtValue(v)} />
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
                    <td style={{ ...styles.searchTd, textAlign: "right" }}>{fmtValue(e.total)}</td>
                    <td style={{ ...styles.searchTd, textAlign: "right" }}>{fmtValue(e.recebido)}</td>
                    <td style={{ ...styles.searchTd, textAlign: "right" }}>{fmtValue(e.aReceber)}</td>
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
                <YAxis tickFormatter={(v) => fmtValue(v)} fontSize={10} width={60} />
                <Tooltip formatter={(v) => fmtValue(v)} />
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
                    <td style={{ ...styles.searchTd, textAlign: "right" }}>{fmtValue(m.total)}</td>
                    <td style={{ ...styles.searchTd, textAlign: "right" }}>{fmtValue(m.recebido)}</td>
                    <td style={{ ...styles.searchTd, textAlign: "right" }}>{fmtValue(m.aReceber)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}
