"""
Solar Flex IoT Platform — Fleet Overview Dashboard
Arquitectura IoT, Sensórica & Digital Twin para Economía Circular
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from data.solar_flex_simulator import get_fleet_summary, get_maintenance_alerts, INSTALLATIONS

st.set_page_config(
    page_title="Solar Flex IoT Platform",
    page_icon="☀️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Estilos ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stMetricValue"] { font-size: 1.6rem; font-weight: 700; }
.alert-alta  { background:#fee2e2; border-left:4px solid #dc2626; padding:10px 14px; border-radius:6px; margin:6px 0; }
.alert-media { background:#fef3c7; border-left:4px solid #d97706; padding:10px 14px; border-radius:6px; margin:6px 0; }
.alert-baja  { background:#dcfce7; border-left:4px solid #16a34a; padding:10px 14px; border-radius:6px; margin:6px 0; }
.status-operativo    { color:#16a34a; font-weight:700; }
.status-alerta       { color:#d97706; font-weight:700; }
.status-mantenimiento{ color:#2563eb; font-weight:700; }
</style>
""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.image("https://via.placeholder.com/200x60/1e3a5f/ffffff?text=SOLAR+FLEX", use_container_width=True)
    st.markdown("### 🛰️ IoT Platform v1.0")
    st.markdown("---")
    st.markdown("**Navegación**")
    st.page_link("streamlit_app.py", label="🏠 Vista General Flota", icon="🏠")
    st.page_link("pages/1_Digital_Twin.py", label="🔬 Digital Twin", icon="🔬")
    st.page_link("pages/2_Pasaporte_Digital.py", label="📋 Pasaporte Digital", icon="📋")
    st.page_link("pages/3_Mantenimiento_Predictivo.py", label="🔧 Mantenimiento Predictivo", icon="🔧")
    st.page_link("pages/4_Informe_ESG.py", label="🌱 Informe ESG", icon="🌱")
    st.page_link("pages/5_Arquitectura_Sensorica.py", label="📡 Arquitectura Sensórica", icon="📡")
    st.markdown("---")
    st.caption(f"Última actualización: {pd.Timestamp.now().strftime('%d/%m/%Y %H:%M')}")

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("# ☀️ Solar Flex — Vista General de Flota")
st.markdown("**Plataforma IoT · Digital Twin · Pasaporte Digital ESPR** | Roof-as-a-Service")
st.markdown("---")

# ── Carga de datos ────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_fleet():
    return get_fleet_summary()

@st.cache_data(ttl=300)
def load_alerts():
    return get_maintenance_alerts()

with st.spinner("Cargando datos de flota..."):
    fleet_df = load_fleet()
    alerts = load_alerts()

# ── KPIs globales ─────────────────────────────────────────────────────────────
total_kwp = fleet_df["kWp"].sum()
total_energia = fleet_df["Energía_kWh_30d"].sum()
total_co2 = fleet_df["CO2_evitado_t_30d"].sum()
pr_medio = fleet_df["PR_real"].mean()
n_alertas = len([a for a in alerts if a["severidad"] == "ALTA"])
n_instalaciones = len(fleet_df)

col1, col2, col3, col4, col5, col6 = st.columns(6)
col1.metric("Instalaciones activas", n_instalaciones, delta="5 total")
col2.metric("Potencia total", f"{total_kwp:.0f} kWp", delta=f"{total_kwp/1000:.2f} MWp")
col3.metric("Energía últimos 30d", f"{total_energia/1000:.1f} MWh", delta=f"+{total_energia/1000*0.03:.1f} MWh vs mes ant.")
col4.metric("PR medio flota", f"{pr_medio:.1%}", delta=f"{(pr_medio-0.81)*100:.1f}pp vs objetivo")
col5.metric("CO₂ evitado 30d", f"{total_co2:.1f} t", delta="↓ carbono")
col6.metric("Alertas activas", n_alertas, delta="ALTA severidad", delta_color="inverse")

st.markdown("---")

# ── Mapa de instalaciones ─────────────────────────────────────────────────────
col_map, col_table = st.columns([1.2, 1])

with col_map:
    st.subheader("📍 Mapa de instalaciones")
    color_map_status = {"operativo": "#16a34a", "alerta": "#d97706", "mantenimiento": "#2563eb"}
    fleet_df["color"] = fleet_df["Estado"].map(color_map_status)
    fleet_df["pr_pct"] = (fleet_df["PR_real"] * 100).round(1)

    fig_map = px.scatter_mapbox(
        fleet_df, lat="lat", lon="lon",
        size="kWp", color="Estado",
        color_discrete_map={"operativo": "#16a34a", "alerta": "#d97706", "mantenimiento": "#2563eb"},
        hover_name="Instalación",
        hover_data={"lat": False, "lon": False, "kWp": True, "pr_pct": True, "Energía_kWh_30d": True},
        size_max=30, zoom=5,
        mapbox_style="carto-positron",
        labels={"pr_pct": "PR (%)", "Energía_kWh_30d": "Energía 30d (kWh)"},
    )
    fig_map.update_layout(margin={"r": 0, "t": 0, "l": 0, "b": 0}, height=380)
    st.plotly_chart(fig_map, use_container_width=True)

with col_table:
    st.subheader("📊 Estado de instalaciones")
    display_cols = ["Instalación", "kWp", "Estado", "PR_real", "Energía_kWh_30d", "CO2_evitado_t_30d"]
    display_df = fleet_df[display_cols].copy()
    display_df["PR_real"] = display_df["PR_real"].apply(lambda x: f"{x:.1%}")
    display_df["Energía_kWh_30d"] = display_df["Energía_kWh_30d"].apply(lambda x: f"{x:,.0f}")
    display_df["CO2_evitado_t_30d"] = display_df["CO2_evitado_t_30d"].apply(lambda x: f"{x:.2f} t")
    display_df.columns = ["Instalación", "kWp", "Estado", "PR", "Energía 30d (kWh)", "CO₂ evitado"]

    def style_estado(val):
        colors = {"operativo": "color: #16a34a; font-weight:bold",
                  "alerta": "color: #d97706; font-weight:bold",
                  "mantenimiento": "color: #2563eb; font-weight:bold"}
        return colors.get(val, "")

    st.dataframe(
        display_df.style.applymap(style_estado, subset=["Estado"]),
        use_container_width=True, height=360
    )

st.markdown("---")

# ── Gráficos de flota ─────────────────────────────────────────────────────────
col_pr, col_energia = st.columns(2)

with col_pr:
    st.subheader("Performance Ratio por instalación")
    fig_pr = go.Figure()
    for _, row in fleet_df.iterrows():
        color = color_map_status.get(row["Estado"], "#64748b")
        fig_pr.add_trace(go.Bar(
            x=[row["Instalación"].split(" ")[0] + "..."],
            y=[row["PR_real"] * 100],
            name=row["Instalación"],
            marker_color=color,
            showlegend=False,
            hovertemplate=f"<b>{row['Instalación']}</b><br>PR: {row['PR_real']:.1%}<br>Objetivo: {row['PR_objetivo']:.1%}<extra></extra>",
        ))
    # Línea objetivo
    fig_pr.add_hline(y=81, line_dash="dash", line_color="#94a3b8", annotation_text="Objetivo 81%")
    fig_pr.update_layout(
        yaxis_title="Performance Ratio (%)", yaxis_range=[70, 90],
        height=300, margin=dict(t=20, b=40), plot_bgcolor="white",
    )
    fig_pr.update_yaxes(gridcolor="#f1f5f9")
    st.plotly_chart(fig_pr, use_container_width=True)

with col_energia:
    st.subheader("Energía producida (últimos 30 días)")
    fig_en = px.bar(
        fleet_df, x="Instalación", y="Energía_kWh_30d",
        color="Tipo", color_discrete_map={"Industrial": "#1e3a8a", "Comercial": "#0891b2", "Terciario": "#0d9488"},
        labels={"Energía_kWh_30d": "kWh", "Instalación": ""},
        height=300,
    )
    fig_en.update_xaxes(tickangle=-30)
    fig_en.update_layout(margin=dict(t=20, b=80), plot_bgcolor="white", legend_title="Tipo")
    fig_en.update_yaxes(gridcolor="#f1f5f9")
    st.plotly_chart(fig_en, use_container_width=True)

st.markdown("---")

# ── Alertas activas ───────────────────────────────────────────────────────────
st.subheader("🚨 Alertas de mantenimiento predictivo")

if not alerts:
    st.success("✅ No hay alertas activas en la flota.")
else:
    for alert in sorted(alerts, key=lambda x: {"ALTA": 0, "MEDIA": 1, "BAJA": 2}.get(x["severidad"], 3)):
        sev_class = {"ALTA": "alert-alta", "MEDIA": "alert-media", "BAJA": "alert-baja"}.get(alert["severidad"], "alert-baja")
        icon = {"ALTA": "🔴", "MEDIA": "🟡", "BAJA": "🟢"}.get(alert["severidad"], "⚪")
        st.markdown(
            f'<div class="{sev_class}">'
            f'<strong>{icon} [{alert["severidad"]}] {alert["tipo"]} — {alert["instalacion"]}</strong><br>'
            f'{alert["mensaje"]}'
            f'</div>',
            unsafe_allow_html=True,
        )

st.markdown("---")

# ── Distribución por tipo y degradación ──────────────────────────────────────
col_pie, col_deg = st.columns(2)

with col_pie:
    st.subheader("Distribución de potencia por tipo")
    pie_data = fleet_df.groupby("Tipo")["kWp"].sum().reset_index()
    fig_pie = px.pie(pie_data, values="kWp", names="Tipo",
                     color_discrete_map={"Industrial": "#1e3a8a", "Comercial": "#0891b2", "Terciario": "#0d9488"},
                     height=280)
    fig_pie.update_layout(margin=dict(t=20))
    st.plotly_chart(fig_pie, use_container_width=True)

with col_deg:
    st.subheader("Degradación PR vs objetivo (%)")
    fig_deg = px.bar(
        fleet_df, x="Instalación", y="degradacion_pct",
        color="degradacion_pct",
        color_continuous_scale=["#16a34a", "#d97706", "#dc2626"],
        range_color=[0, 15],
        labels={"degradacion_pct": "Degradación (%)", "Instalación": ""},
        height=280,
    )
    fig_deg.add_hline(y=7, line_dash="dash", line_color="#dc2626", annotation_text="Umbral alerta 7%")
    fig_deg.update_xaxes(tickangle=-30)
    fig_deg.update_layout(margin=dict(t=20, b=80), plot_bgcolor="white", coloraxis_showscale=False)
    fig_deg.update_yaxes(gridcolor="#f1f5f9")
    st.plotly_chart(fig_deg, use_container_width=True)

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption(
    "**Solar Flex IoT Platform** · Arquitectura: LoRaWAN + Edge Gateway + Data Lakehouse · "
    "Cumplimiento: ESPR/DPP · CSRD · EPBD · RENOCICLA | "
    "© 2025 Solar Flex Technologies S.L."
)
