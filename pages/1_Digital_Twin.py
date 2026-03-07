"""
Solar Flex — Digital Twin
Gemelo digital en tiempo real de cada instalación BIPV
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from data.solar_flex_simulator import INSTALLATIONS, generate_timeseries, compute_kpis

st.set_page_config(page_title="Digital Twin — Solar Flex", page_icon="🔬", layout="wide")

st.markdown("""
<style>
[data-testid="stMetricValue"] { font-size: 1.5rem; font-weight: 700; }
.kpi-card { background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:12px; text-align:center; }
.level-badge { display:inline-block; padding:3px 10px; border-radius:12px; font-size:0.8rem; font-weight:600; }
.level-1 { background:#dbeafe; color:#1d4ed8; }
.level-2 { background:#dcfce7; color:#15803d; }
.level-3 { background:#fef3c7; color:#92400e; }
</style>
""", unsafe_allow_html=True)

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🔬 Digital Twin")
    inst_names = {i["id"]: i["name"] for i in INSTALLATIONS}
    selected_id = st.selectbox("Seleccionar instalación", list(inst_names.keys()),
                               format_func=lambda x: inst_names[x])
    days = st.selectbox("Periodo de análisis", [7, 14, 30, 60, 90], index=2)
    freq = st.selectbox("Resolución temporal", {"15 min": 15, "1 hora": 60, "1 día": 1440}.keys(), index=1)
    freq_map = {"15 min": 15, "1 hora": 60, "1 día": 1440}
    freq_min = freq_map[freq]

    st.markdown("---")
    st.markdown("**Nivel Digital Twin**")
    st.markdown('<span class="level-badge level-1">Nivel 1 — Shadow</span> Real-time', unsafe_allow_html=True)
    st.markdown('<span class="level-badge level-2">Nivel 2 — Predictivo</span> 2027', unsafe_allow_html=True)
    st.markdown('<span class="level-badge level-3">Nivel 3 — Prescriptivo</span> 2028', unsafe_allow_html=True)

inst = next(i for i in INSTALLATIONS if i["id"] == selected_id)

st.markdown(f"# 🔬 Digital Twin — {inst['name']}")
st.markdown(f"**{inst['id']}** · {inst['location']} · {inst['kwp']} kWp · Instalado: {inst['install_date']} · Panel: {inst['panel_model']}")

status_colors = {"operativo": "🟢", "alerta": "🟡", "mantenimiento": "🔵"}
st.markdown(f"Estado: {status_colors.get(inst['status'], '⚪')} **{inst['status'].upper()}**")
st.markdown("---")

@st.cache_data(ttl=180)
def load_data(inst_id, d, f):
    return generate_timeseries(inst_id, days=d, freq_minutes=f)

with st.spinner("Sincronizando gemelo digital..."):
    df = load_data(selected_id, days, freq_min)
    kpis = compute_kpis(df, inst["kwp"])

# ── KPIs en tiempo real ────────────────────────────────────────────────────────
st.subheader("📊 KPIs en tiempo real")
c1, c2, c3, c4, c5, c6 = st.columns(6)
pr_delta = kpis["pr"] - inst["pr_target"]
c1.metric("Performance Ratio", f"{kpis['pr']:.1%}", delta=f"{pr_delta*100:+.1f}pp vs objetivo", delta_color="normal")
c2.metric("Energía producida", f"{kpis['energia_kwh']:,.0f} kWh", delta=f"{days}d acumulado")
c3.metric("Específico energético", f"{kpis['especifico_kwh_kwp']:.0f} kWh/kWp")
c4.metric("T° célula máx.", f"{kpis['t_celula_max']:.1f} °C", delta_color="inverse",
          delta="⚠ alto" if kpis["t_celula_max"] > 75 else "✓ normal")
c5.metric("Strain adhesivo máx.", f"{kpis['strain_max_ue']:.0f} µε",
          delta="⚠ revisar" if kpis["strain_max_ue"] > 350 else "✓ OK", delta_color="inverse" if kpis["strain_max_ue"] > 350 else "off")
c6.metric("CO₂ evitado", f"{kpis['co2_evitado_t']:.3f} t", delta=f"Factor REE: 0.181 kg/kWh")

st.markdown("---")

# ── Gráfico principal: Potencia + Irradiancia ──────────────────────────────────
st.subheader("⚡ Producción energética y recurso solar")
fig_energy = make_subplots(
    rows=2, cols=1, shared_xaxes=True,
    row_heights=[0.65, 0.35],
    subplot_titles=["Potencia AC (kW) vs Irradiancia POA (W/m²)", "Temperatura: Célula vs Ambiente (°C)"],
    vertical_spacing=0.08,
)
fig_energy.add_trace(go.Scatter(
    x=df["timestamp"], y=df["potencia_ac_kw"],
    name="Potencia AC (kW)", fill="tozeroy",
    fillcolor="rgba(255,165,0,0.15)", line=dict(color="#f97316", width=1.5),
), row=1, col=1)
fig_energy.add_trace(go.Scatter(
    x=df["timestamp"], y=df["irradiancia_poa"],
    name="Irradiancia POA (W/m²)", yaxis="y2",
    line=dict(color="#fbbf24", width=1, dash="dot"),
), row=1, col=1)
fig_energy.add_trace(go.Scatter(
    x=df["timestamp"], y=df["temperatura_celula"],
    name="T° célula (°C)", line=dict(color="#dc2626", width=1.5),
), row=2, col=1)
fig_energy.add_trace(go.Scatter(
    x=df["timestamp"], y=df["temperatura_ambiente"],
    name="T° ambiente (°C)", line=dict(color="#3b82f6", width=1),
), row=2, col=1)

fig_energy.update_layout(height=450, plot_bgcolor="white", hovermode="x unified",
                          legend=dict(orientation="h", y=1.06))
fig_energy.update_yaxes(gridcolor="#f1f5f9")
st.plotly_chart(fig_energy, use_container_width=True)

# ── Dominio estructural-adhesivo ───────────────────────────────────────────────
st.subheader("🏗️ Monitorización estructural — Sistema adhesivo")
st.info("Este bloque es diferenciador clave Solar Flex: monitorización del sistema adhesivo en tiempo real para garantía activa de instalación.")

fig_struct = make_subplots(
    rows=1, cols=2,
    subplot_titles=["Strain Adhesivo (µε) — Zonas 1 y 2", "Humedad Interfase (%) y Viento (m/s)"],
)
fig_struct.add_trace(go.Scatter(
    x=df["timestamp"], y=df["strain_adhesivo_1_ue"],
    name="Strain zona 1", line=dict(color="#7c3aed", width=1.5),
), row=1, col=1)
fig_struct.add_trace(go.Scatter(
    x=df["timestamp"], y=df["strain_adhesivo_2_ue"],
    name="Strain zona 2", line=dict(color="#a855f7", width=1, dash="dash"),
), row=1, col=1)
# Umbral crítico
fig_struct.add_hline(y=350, line_dash="dash", line_color="#dc2626",
                     annotation_text="Umbral crítico 350 µε", row=1, col=1)

fig_struct.add_trace(go.Scatter(
    x=df["timestamp"], y=df["humedad_interfase_pct"],
    name="HR interfase (%)", fill="tozeroy",
    fillcolor="rgba(6,182,212,0.15)", line=dict(color="#06b6d4", width=1.5),
), row=1, col=2)
fig_struct.add_trace(go.Scatter(
    x=df["timestamp"], y=df["velocidad_viento_ms"] * 10,
    name="Viento ×10 (m/s)", line=dict(color="#64748b", width=1, dash="dot"),
), row=1, col=2)

fig_struct.update_layout(height=320, plot_bgcolor="white", hovermode="x unified",
                          legend=dict(orientation="h", y=1.12))
fig_struct.update_yaxes(gridcolor="#f1f5f9")
st.plotly_chart(fig_struct, use_container_width=True)

# ── Performance Ratio histórico ────────────────────────────────────────────────
st.subheader("📈 Performance Ratio diario")
df_daily = df.set_index("timestamp").resample("D").apply({
    "potencia_ac_kw": "sum",
    "irradiancia_poa": "sum",
}).reset_index()
df_daily["energia_kwh"] = df_daily["potencia_ac_kw"] * (15 / 60) if freq_min <= 60 else df_daily["potencia_ac_kw"]
df_daily["irr_kwh_m2"] = df_daily["irradiancia_poa"] * (15 / 60 if freq_min <= 60 else 1) / 1000
df_daily["pr_diario"] = (df_daily["energia_kwh"] / (df_daily["irr_kwh_m2"] * inst["kwp"])).clip(0, 1)

fig_pr = go.Figure()
fig_pr.add_trace(go.Scatter(
    x=df_daily["timestamp"], y=df_daily["pr_diario"] * 100,
    fill="tozeroy", fillcolor="rgba(34,197,94,0.15)",
    line=dict(color="#22c55e", width=2), name="PR diario (%)",
))
fig_pr.add_hline(y=inst["pr_target"] * 100, line_dash="dash",
                 line_color="#1e3a8a", annotation_text=f"Objetivo {inst['pr_target']:.0%}")
fig_pr.add_hline(y=inst["pr_target"] * 93, line_dash="dot",
                 line_color="#dc2626", annotation_text="Umbral alerta (−7%)")
fig_pr.update_layout(height=250, plot_bgcolor="white", yaxis_title="PR (%)", yaxis_range=[60, 95])
fig_pr.update_yaxes(gridcolor="#f1f5f9")
st.plotly_chart(fig_pr, use_container_width=True)

# ── Consumo vs Producción (autoconsumo) ───────────────────────────────────────
st.subheader("🏭 Balance energético edificio — Autoconsumo")
fig_balance = go.Figure()
fig_balance.add_trace(go.Scatter(
    x=df["timestamp"], y=df["consumo_edificio_kw"],
    name="Consumo edificio (kW)", fill="tozeroy",
    fillcolor="rgba(239,68,68,0.1)", line=dict(color="#ef4444", width=1.5),
))
fig_balance.add_trace(go.Scatter(
    x=df["timestamp"], y=df["potencia_ac_kw"],
    name="Producción FV (kW)", fill="tozeroy",
    fillcolor="rgba(34,197,94,0.15)", line=dict(color="#22c55e", width=1.5),
))
fig_balance.add_trace(go.Scatter(
    x=df["timestamp"], y=df["soc_bateria_pct"],
    name="SoC Batería (%)", yaxis="y2", line=dict(color="#8b5cf6", width=1, dash="dot"),
))
fig_balance.update_layout(
    height=300, plot_bgcolor="white", hovermode="x unified",
    legend=dict(orientation="h", y=1.1),
    yaxis2=dict(overlaying="y", side="right", title="SoC Batería (%)", range=[0, 110]),
)
fig_balance.update_yaxes(gridcolor="#f1f5f9")
st.plotly_chart(fig_balance, use_container_width=True)

# ── Info técnica instalación ───────────────────────────────────────────────────
with st.expander("ℹ️ Información técnica de la instalación"):
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown(f"""
**ID único:** `{inst['id']}`
**Modelo panel:** {inst['panel_model']}
**Potencia pico:** {inst['kwp']} kWp
**Área cubierta:** {inst['area_m2']} m²
**Densidad:** {inst['kwp']*1000/inst['area_m2']:.0f} Wp/m²
**Sistema adhesivo:** {inst['adhesive']}
        """)
    with col_b:
        st.markdown(f"""
**Fecha instalación:** {inst['install_date']}
**Localización:** {inst['location']}
**Tipo de edificio:** {inst['type']}
**PR objetivo:** {inst['pr_target']:.0%}
**Protocolo conectividad:** LoRaWAN + 4G Cat-M1
**Gateway Edge:** Advantech UNO-2484G
        """)

st.markdown("---")
st.caption("Digital Twin Nivel 1 (Shadow) — Solar Flex IoT Platform · Datos actualizados cada 15 minutos · TLS 1.3 cifrado extremo a extremo")
