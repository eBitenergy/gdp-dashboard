"""
Solar Flex — Mantenimiento Predictivo
Detección automática de anomalías, alertas y tickets O&M
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from data.solar_flex_simulator import INSTALLATIONS, generate_timeseries, compute_kpis, get_maintenance_alerts

st.set_page_config(page_title="Mantenimiento Predictivo — Solar Flex", page_icon="🔧", layout="wide")

st.markdown("""
<style>
.alert-card-alta  { background:#fee2e2; border-left:5px solid #dc2626; padding:14px 16px;
                    border-radius:8px; margin:8px 0; }
.alert-card-media { background:#fef3c7; border-left:5px solid #d97706; padding:14px 16px;
                    border-radius:8px; margin:8px 0; }
.alert-card-baja  { background:#dcfce7; border-left:5px solid #16a34a; padding:14px 16px;
                    border-radius:8px; margin:8px 0; }
.ticket-box { background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px;
              padding:16px; margin:8px 0; }
.anomaly-tag { display:inline-block; padding:2px 8px; border-radius:10px; font-size:0.75rem;
               font-weight:600; margin:2px; }
.tag-struct { background:#ede9fe; color:#7c3aed; }
.tag-perf   { background:#dbeafe; color:#1d4ed8; }
.tag-thermal{ background:#fee2e2; color:#dc2626; }
.tag-soiling{ background:#fef3c7; color:#92400e; }
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 🔧 Mantenimiento Predictivo")
    st.markdown("---")
    st.markdown("**Tipos de anomalías detectadas:**")
    st.markdown('<span class="anomaly-tag tag-struct">🏗️ Estructural</span>', unsafe_allow_html=True)
    st.markdown('<span class="anomaly-tag tag-perf">⚡ Rendimiento</span>', unsafe_allow_html=True)
    st.markdown('<span class="anomaly-tag tag-thermal">🌡️ Térmica</span>', unsafe_allow_html=True)
    st.markdown('<span class="anomaly-tag tag-soiling">🌫️ Soiling</span>', unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("**Ahorro O&M estimado:**")
    st.metric("vs. mantenimiento correctivo", "30-40%", delta="coste reducido")
    st.markdown("---")
    st.markdown("**SLA respuesta:**")
    st.markdown("🔴 ALTA: 24h")
    st.markdown("🟡 MEDIA: 72h")
    st.markdown("🟢 BAJA: 7 días")

st.markdown("# 🔧 Mantenimiento Predictivo")
st.markdown("Detección automática de anomalías · Alertas en tiempo real · Tickets O&M")
st.markdown("---")

@st.cache_data(ttl=180)
def load_all_data():
    all_data = {}
    for inst in INSTALLATIONS:
        df = generate_timeseries(inst["id"], days=30, freq_minutes=60)
        kpis = compute_kpis(df, inst["kwp"])
        all_data[inst["id"]] = {"df": df, "kpis": kpis, "inst": inst}
    return all_data

@st.cache_data(ttl=180)
def load_alerts():
    return get_maintenance_alerts()

with st.spinner("Ejecutando modelos de detección de anomalías..."):
    all_data = load_all_data()
    alerts = load_alerts()

# ── Resumen de estado de flota ────────────────────────────────────────────────
st.subheader("📊 Estado de salud de la flota")

health_data = []
for inst in INSTALLATIONS:
    d = all_data[inst["id"]]
    kpis = d["kpis"]
    df = d["df"]
    strain_max = max(df["strain_adhesivo_1_ue"].max(), df["strain_adhesivo_2_ue"].max())
    t_max = df["temperatura_celula"].max()
    pr_ratio = kpis["pr"] / inst["pr_target"]
    soiling_events = (df["pm25_ug_m3"] > 50).sum()

    # Score de salud 0-100
    score_pr = min(100, pr_ratio * 100)
    score_struct = max(0, 100 - strain_max / 5)
    score_thermal = max(0, 100 - max(0, t_max - 65) * 4)
    score_overall = (score_pr * 0.5 + score_struct * 0.3 + score_thermal * 0.2)

    health_data.append({
        "Instalación": inst["name"],
        "PR_score": round(score_pr, 1),
        "Struct_score": round(score_struct, 1),
        "Thermal_score": round(score_thermal, 1),
        "Health_score": round(score_overall, 1),
        "Estado": inst["status"],
        "Strain_max": round(strain_max, 0),
        "T_max": round(t_max, 1),
        "Alertas": len([a for a in alerts if a["id"] == inst["id"]]),
    })

health_df = pd.DataFrame(health_data)

col1, col2, col3 = st.columns(3)

with col1:
    # Gauge chart de salud media
    health_mean = health_df["Health_score"].mean()
    fig_gauge = go.Figure(go.Indicator(
        mode="gauge+number",
        value=health_mean,
        gauge={
            "axis": {"range": [0, 100]},
            "bar": {"color": "#22c55e" if health_mean > 75 else "#d97706" if health_mean > 60 else "#dc2626"},
            "steps": [
                {"range": [0, 60], "color": "#fee2e2"},
                {"range": [60, 80], "color": "#fef3c7"},
                {"range": [80, 100], "color": "#dcfce7"},
            ],
        },
        title={"text": "Salud media flota (%)"},
        number={"suffix": "%"},
    ))
    fig_gauge.update_layout(height=250)
    st.plotly_chart(fig_gauge, use_container_width=True)

with col2:
    st.markdown("**Scores por instalación**")
    fig_health = px.bar(
        health_df, x="Health_score", y="Instalación", orientation="h",
        color="Health_score",
        color_continuous_scale=["#dc2626", "#d97706", "#22c55e"],
        range_color=[50, 100],
        labels={"Health_score": "Score salud (%)", "Instalación": ""},
        height=240,
    )
    fig_health.add_vline(x=75, line_dash="dash", line_color="#64748b", annotation_text="Umbral OK")
    fig_health.update_layout(plot_bgcolor="white", coloraxis_showscale=False, margin=dict(t=10))
    fig_health.update_xaxes(gridcolor="#f1f5f9", range=[0, 110])
    st.plotly_chart(fig_health, use_container_width=True)

with col3:
    st.markdown("**Breakdown por dominio**")
    radar_df = health_df[["Instalación", "PR_score", "Struct_score", "Thermal_score"]].copy()
    radar_df["Instalación"] = radar_df["Instalación"].apply(lambda x: x.split(" ")[0])

    fig_radar = go.Figure()
    categories = ["Rendimiento PV", "Estructura/Adhesivo", "Temperatura"]
    for _, row in radar_df.iterrows():
        fig_radar.add_trace(go.Scatterpolar(
            r=[row["PR_score"], row["Struct_score"], row["Thermal_score"], row["PR_score"]],
            theta=categories + [categories[0]],
            name=row["Instalación"], fill="toself", opacity=0.5,
        ))
    fig_radar.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 110])),
        height=250, margin=dict(t=20),
        legend=dict(font=dict(size=9)),
    )
    st.plotly_chart(fig_radar, use_container_width=True)

st.markdown("---")

# ── Alertas activas con detalle ───────────────────────────────────────────────
st.subheader(f"🚨 Alertas activas ({len(alerts)} total)")

if not alerts:
    st.success("✅ Sin alertas activas. Todos los sistemas operando dentro de parámetros normales.")
else:
    for alert in sorted(alerts, key=lambda x: {"ALTA": 0, "MEDIA": 1, "BAJA": 2}.get(x["severidad"], 3)):
        css = {"ALTA": "alert-card-alta", "MEDIA": "alert-card-media", "BAJA": "alert-card-baja"}.get(alert["severidad"], "alert-card-baja")
        icon = {"ALTA": "🔴", "MEDIA": "🟡", "BAJA": "🟢"}.get(alert["severidad"])
        sla = {"ALTA": "24h", "MEDIA": "72h", "BAJA": "7 días"}.get(alert["severidad"])

        st.markdown(f"""
<div class="{css}">
    <div style="display:flex; justify-content:space-between; align-items:flex-start;">
        <div>
            <strong>{icon} [{alert['severidad']}] {alert['tipo']} — {alert['instalacion']}</strong><br>
            <span style="font-size:0.9rem;">{alert['mensaje']}</span><br>
            <span style="font-size:0.8rem; color:#64748b;">
                Sensor: <code>{alert['sensor']}</code> |
                Valor: <strong>{alert['valor']:.1f}</strong> |
                Umbral: {alert['umbral']:.1f} |
                ID: <code>{alert['id']}</code>
            </span>
        </div>
        <div style="text-align:right; font-size:0.8rem; color:#64748b; min-width:80px;">
            SLA: <strong>{sla}</strong><br>
            <span style="background:#e2e8f0; padding:2px 6px; border-radius:4px;">ABIERTO</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown("---")

# ── Análisis anomalía: Strain adhesivo ────────────────────────────────────────
st.subheader("🏗️ Análisis detallado: Integridad adhesiva")
st.markdown("Modelo de detección de despegue incipiente basado en strain gauge galgas extensométricas (HBK).")

alert_inst = next((i for i in INSTALLATIONS if i["status"] == "alerta"), INSTALLATIONS[0])
df_alert = all_data[alert_inst["id"]]["df"]

fig_strain_detail = make_subplots(
    rows=2, cols=1, shared_xaxes=True,
    subplot_titles=[
        f"Strain Adhesivo µε — {alert_inst['name']}",
        "Correlación Strain vs Temperatura célula (factor de correlación esperado)"
    ],
    row_heights=[0.6, 0.4], vertical_spacing=0.1,
)
fig_strain_detail.add_trace(go.Scatter(
    x=df_alert["timestamp"], y=df_alert["strain_adhesivo_1_ue"],
    name="Strain zona 1 (µε)", line=dict(color="#7c3aed", width=1.5),
), row=1, col=1)
fig_strain_detail.add_trace(go.Scatter(
    x=df_alert["timestamp"], y=df_alert["strain_adhesivo_2_ue"],
    name="Strain zona 2 (µε)", line=dict(color="#a855f7", width=1, dash="dash"),
), row=1, col=1)
fig_strain_detail.add_hline(y=350, line_dash="dash", line_color="#dc2626",
                             annotation_text="Umbral alerta 350 µε", row=1, col=1)
fig_strain_detail.add_hline(y=200, line_dash="dot", line_color="#d97706",
                             annotation_text="Umbral vigilancia 200 µε", row=1, col=1)

# Diferencia relativa entre zonas (anomalía si divergen >30%)
strain_ratio = (df_alert["strain_adhesivo_1_ue"] / (df_alert["strain_adhesivo_2_ue"] + 0.1) - 1) * 100
fig_strain_detail.add_trace(go.Scatter(
    x=df_alert["timestamp"], y=strain_ratio,
    name="Asimetría zona1/zona2 (%)", fill="tozeroy",
    fillcolor="rgba(239,68,68,0.1)", line=dict(color="#ef4444", width=1),
), row=2, col=1)
fig_strain_detail.add_hline(y=30, line_dash="dash", line_color="#dc2626",
                             annotation_text="Umbral asimetría 30%", row=2, col=1)

fig_strain_detail.update_layout(height=380, plot_bgcolor="white", hovermode="x unified")
fig_strain_detail.update_yaxes(gridcolor="#f1f5f9")
st.plotly_chart(fig_strain_detail, use_container_width=True)

# ── Análisis rendimiento: IV curve / PR ───────────────────────────────────────
st.subheader("⚡ Análisis de rendimiento — Detección de módulos degradados")
tab_pr, tab_soiling = st.tabs(["Performance Ratio por instalación", "Análisis Soiling (PM2.5)"])

with tab_pr:
    # Comparativa PR últimas 4 semanas vs semana anterior
    st.markdown("Comparativa PR semanal para detectar degradación acelerada:")
    pr_comp = []
    for inst in INSTALLATIONS:
        df_i = all_data[inst["id"]]["df"]
        w_recent = df_i[df_i["timestamp"] >= df_i["timestamp"].max() - pd.Timedelta(days=7)]
        w_prev = df_i[
            (df_i["timestamp"] >= df_i["timestamp"].max() - pd.Timedelta(days=14)) &
            (df_i["timestamp"] < df_i["timestamp"].max() - pd.Timedelta(days=7))
        ]
        kpi_r = compute_kpis(w_recent, inst["kwp"])
        kpi_p = compute_kpis(w_prev, inst["kwp"])
        pr_comp.append({
            "Instalación": inst["name"].split(" ")[0],
            "PR semana actual": kpi_r["pr"] * 100,
            "PR semana anterior": kpi_p["pr"] * 100,
            "Variación (pp)": (kpi_r["pr"] - kpi_p["pr"]) * 100,
        })

    pr_comp_df = pd.DataFrame(pr_comp)
    fig_pr_comp = go.Figure()
    fig_pr_comp.add_trace(go.Bar(
        x=pr_comp_df["Instalación"], y=pr_comp_df["PR semana anterior"],
        name="Semana anterior", marker_color="#94a3b8",
    ))
    fig_pr_comp.add_trace(go.Bar(
        x=pr_comp_df["Instalación"], y=pr_comp_df["PR semana actual"],
        name="Semana actual",
        marker_color=["#22c55e" if v >= 0 else "#ef4444" for v in pr_comp_df["Variación (pp)"]],
    ))
    fig_pr_comp.update_layout(
        barmode="group", height=280, plot_bgcolor="white",
        yaxis_title="PR (%)", yaxis_range=[70, 90],
    )
    fig_pr_comp.update_yaxes(gridcolor="#f1f5f9")
    st.plotly_chart(fig_pr_comp, use_container_width=True)

with tab_soiling:
    st.markdown("Correlación entre PM2.5 (suciedad) y pérdida de rendimiento:")
    inst_sel = INSTALLATIONS[0]
    df_soil = all_data[inst_sel["id"]]["df"].copy()
    df_soil_d = df_soil.set_index("timestamp").resample("D").agg({
        "pm25_ug_m3": "mean", "potencia_ac_kw": "sum", "irradiancia_poa": "sum"
    }).reset_index()
    df_soil_d["pr_d"] = (df_soil_d["potencia_ac_kw"] / (df_soil_d["irradiancia_poa"] / 1000 * inst_sel["kwp"] + 0.01)).clip(0, 1)

    fig_soil = make_subplots(specs=[[{"secondary_y": True}]])
    fig_soil.add_trace(go.Bar(
        x=df_soil_d["timestamp"], y=df_soil_d["pm25_ug_m3"],
        name="PM2.5 (µg/m³)", marker_color="rgba(148,163,184,0.6)",
    ), secondary_y=False)
    fig_soil.add_trace(go.Scatter(
        x=df_soil_d["timestamp"], y=df_soil_d["pr_d"] * 100,
        name="PR diario (%)", line=dict(color="#f97316", width=2),
    ), secondary_y=True)
    fig_soil.update_layout(height=260, plot_bgcolor="white", hovermode="x unified")
    fig_soil.update_yaxes(title_text="PM2.5 (µg/m³)", secondary_y=False, gridcolor="#f1f5f9")
    fig_soil.update_yaxes(title_text="PR (%)", secondary_y=True, range=[60, 95])
    st.plotly_chart(fig_soil, use_container_width=True)
    st.caption("Cuando PM2.5 > 50 µg/m³ durante >3 días consecutivos → recomendación automática de limpieza.")

st.markdown("---")

# ── Tickets O&M ────────────────────────────────────────────────────────────────
st.subheader("🎫 Tickets de mantenimiento O&M")
tickets = pd.DataFrame([
    {"#": "TKT-2026-0089", "Instalación": "Plataforma Logística Zaragoza", "Tipo": "ESTRUCTURAL",
     "Descripción": "Strain zona 1 >350 µε. Inspección adhesiva zonas NE recomendada.",
     "Prioridad": "🔴 ALTA", "Estado": "Abierto", "Creado": "2026-03-06", "SLA": "2026-03-07"},
    {"#": "TKT-2026-0087", "Instalación": "Plataforma Logística Zaragoza", "Tipo": "RENDIMIENTO",
     "Descripción": "PR 73.2% (-8.5% vs objetivo). Revisar string 2 y conexiones DC.",
     "Prioridad": "🟡 MEDIA", "Estado": "En curso", "Creado": "2026-03-05", "SLA": "2026-03-08"},
    {"#": "TKT-2026-0081", "Instalación": "Fábrica Sector Automoción Valencia", "Tipo": "PREVENTIVO",
     "Descripción": "Mantenimiento preventivo anual programado. Limpieza + termografía.",
     "Prioridad": "🟢 BAJA", "Estado": "Programado", "Creado": "2026-02-28", "SLA": "2026-03-14"},
    {"#": "TKT-2026-0076", "Instalación": "Nave Logística Guadalajara", "Tipo": "SOILING",
     "Descripción": "PM2.5 elevado 5 días consecutivos. Limpieza preventiva recomendada.",
     "Prioridad": "🟢 BAJA", "Estado": "Cerrado ✅", "Creado": "2026-02-20", "SLA": "2026-02-27"},
])
st.dataframe(tickets, use_container_width=True, hide_index=True)

col_t1, col_t2, col_t3 = st.columns(3)
col_t1.metric("Tickets abiertos", 2)
col_t2.metric("En curso", 1)
col_t3.metric("Cerrados este mes", 1)

st.markdown("---")
st.caption("Motor de detección: reglas basadas en umbrales + Z-score anomalías | Reducción O&M correctivo estimada: 30-40% | Solar Flex IoT Platform")
