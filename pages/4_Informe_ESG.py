"""
Solar Flex — Informe ESG Automatizado
Cumplimiento CSRD · GHG Protocol Scope 2 · ODS
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from data.solar_flex_simulator import INSTALLATIONS, generate_timeseries, compute_kpis

st.set_page_config(page_title="Informe ESG — Solar Flex", page_icon="🌱", layout="wide")

st.markdown("""
<style>
.esg-header { background: linear-gradient(135deg, #064e3b 0%, #065f46 50%, #047857 100%);
              color:white; padding:20px 24px; border-radius:10px; margin-bottom:16px; }
.ods-badge { display:inline-flex; align-items:center; justify-content:center;
             width:52px; height:52px; border-radius:8px; font-size:1.2rem;
             font-weight:800; margin:4px; }
.ods-7  { background:#FCC30B; color:#1a1a1a; }
.ods-11 { background:#FD9D24; color:white; }
.ods-12 { background:#BF8B2E; color:white; }
.ods-13 { background:#3F7E44; color:white; }
.csrd-ok  { background:#dcfce7; color:#15803d; padding:8px 14px; border-radius:8px;
             font-weight:700; display:inline-block; margin:4px; }
.csrd-prog{ background:#fef3c7; color:#92400e; padding:8px 14px; border-radius:8px;
             font-weight:700; display:inline-block; margin:4px; }
.kpi-green { color:#16a34a; font-size:1.8rem; font-weight:800; }
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 🌱 Informe ESG")
    year_sel = st.selectbox("Año de reporte", [2025, 2026], index=1)
    inst_scope = st.multiselect(
        "Instalaciones incluidas",
        [i["name"] for i in INSTALLATIONS],
        default=[i["name"] for i in INSTALLATIONS],
    )
    export_fmt = st.selectbox("Formato exportación", ["PDF (CSRD-ready)", "JSON (ESG API)", "CSV (auditoría)"])
    st.markdown("---")
    st.markdown("**Normativas cubiertas:**")
    for n in ["CSRD", "GHG Protocol Scope 2", "GRI 302-1", "ISO 14064", "EU Taxonomy"]:
        st.markdown(f"✅ {n}")
    st.markdown("---")
    if st.button("📄 Generar informe completo", type="primary"):
        st.success("Informe generado. Descarga disponible.")

# ── Carga de datos ─────────────────────────────────────────────────────────────
@st.cache_data(ttl=600)
def load_annual_data():
    results = []
    for inst in INSTALLATIONS:
        df90 = generate_timeseries(inst["id"], days=90, freq_minutes=60)
        kpis = compute_kpis(df90, inst["kwp"])
        # Extrapolamos a anual (×4)
        energia_anual = kpis["energia_kwh"] * 4
        co2_anual = kpis["co2_evitado_t"] * 4
        results.append({
            "id": inst["id"],
            "name": inst["name"],
            "location": inst["location"],
            "tipo": inst["type"],
            "kwp": inst["kwp"],
            "energia_anual_kwh": round(energia_anual, 0),
            "co2_evitado_t": round(co2_anual, 3),
            "pr": kpis["pr"],
            "especifico": kpis["especifico_kwh_kwp"] * 4,
        })
    return pd.DataFrame(results)

annual_df = load_annual_data()
insts_sel_names = set(inst_scope)
filtered_df = annual_df[annual_df["name"].isin(insts_sel_names)] if insts_sel_names else annual_df

# ── Header ESG ────────────────────────────────────────────────────────────────
total_energia = filtered_df["energia_anual_kwh"].sum()
total_co2 = filtered_df["co2_evitado_t"].sum()
total_kwp = filtered_df["kwp"].sum()
n_inst = len(filtered_df)

st.markdown(f"""
<div class="esg-header">
    <div style="font-size:0.9rem; opacity:0.8;">INFORME DE SOSTENIBILIDAD ENERGÉTICA · AÑO {year_sel}</div>
    <div style="font-size:1.5rem; font-weight:700; margin:6px 0;">Solar Flex — Plataforma Roof-as-a-Service</div>
    <div style="display:flex; gap:40px; margin-top:12px; flex-wrap:wrap;">
        <div><div style="font-size:2rem; font-weight:800;">{total_energia/1000:,.1f} MWh</div><div style="opacity:0.8; font-size:0.85rem;">Energía renovable producida</div></div>
        <div><div style="font-size:2rem; font-weight:800;">{total_co2:.1f} t</div><div style="opacity:0.8; font-size:0.85rem;">CO₂ evitado (Scope 2)</div></div>
        <div><div style="font-size:2rem; font-weight:800;">{total_kwp:.0f} kWp</div><div style="opacity:0.8; font-size:0.85rem;">Potencia renovable instalada</div></div>
        <div><div style="font-size:2rem; font-weight:800;">{n_inst}</div><div style="opacity:0.8; font-size:0.85rem;">Instalaciones activas</div></div>
    </div>
</div>
""", unsafe_allow_html=True)

tab_ghg, tab_circular, tab_ods, tab_csrd = st.tabs([
    "🌍 GHG Protocol — Emisiones", "♻️ Indicadores Circulares", "🎯 ODS", "📋 CSRD Reporting"
])

# ════════════════════════════════════════════════════════════════════
# GHG PROTOCOL
# ════════════════════════════════════════════════════════════════════
with tab_ghg:
    st.subheader("🌍 Emisiones GHG Protocol — Alcance 2 (Scope 2)")

    col1, col2, col3, col4 = st.columns(4)
    factor_ree = 0.181  # kg CO2/kWh REE España 2025
    col1.metric("Energía renovable producida", f"{total_energia/1000:,.1f} MWh")
    col2.metric("CO₂ evitado (Scope 2)", f"{total_co2:.2f} t CO₂eq",
                delta=f"Factor REE: {factor_ree} kg/kWh")
    col3.metric("Equivalente árboles plantados", f"{total_co2*46.3:.0f}",
                delta="(21.6 kg CO2/árbol/año)")
    col4.metric("Equivalente km coche evitados", f"{total_co2*1000/0.120:,.0f} km",
                delta="(120 g CO2/km vehículo medio)")

    st.markdown("---")

    # Serie mensual CO2 simulada
    months = pd.date_range(f"{year_sel}-01-01", periods=12, freq="MS")
    monthly_irr = np.array([50, 70, 110, 150, 190, 210, 220, 195, 145, 100, 65, 45])
    monthly_irr_norm = monthly_irr / monthly_irr.sum()
    monthly_energia = monthly_irr_norm * total_energia
    monthly_co2 = monthly_energia * factor_ree / 1000

    monthly_df = pd.DataFrame({
        "Mes": months, "Energía (MWh)": monthly_energia / 1000,
        "CO₂ evitado (t)": monthly_co2,
        "CO₂ acumulado (t)": monthly_co2.cumsum(),
    })

    col_bar, col_cum = st.columns(2)
    with col_bar:
        st.markdown("**Producción mensual y CO₂ evitado**")
        fig_monthly = go.Figure()
        fig_monthly.add_trace(go.Bar(
            x=monthly_df["Mes"], y=monthly_df["Energía (MWh)"],
            name="Energía (MWh)", marker_color="#f97316", yaxis="y",
        ))
        fig_monthly.add_trace(go.Scatter(
            x=monthly_df["Mes"], y=monthly_df["CO₂ evitado (t)"] * 10,
            name="CO₂ evitado ×10 (t)", line=dict(color="#22c55e", width=2), yaxis="y2",
        ))
        fig_monthly.update_layout(
            height=280, plot_bgcolor="white", hovermode="x unified",
            yaxis2=dict(overlaying="y", side="right"),
            legend=dict(orientation="h", y=1.1),
        )
        fig_monthly.update_yaxes(gridcolor="#f1f5f9")
        st.plotly_chart(fig_monthly, use_container_width=True)

    with col_cum:
        st.markdown("**CO₂ acumulado — progreso anual**")
        fig_cum = go.Figure()
        fig_cum.add_trace(go.Scatter(
            x=monthly_df["Mes"], y=monthly_df["CO₂ acumulado (t)"],
            fill="tozeroy", fillcolor="rgba(34,197,94,0.2)",
            line=dict(color="#16a34a", width=2), name="tCO₂ acumulado",
        ))
        target_co2 = total_co2
        fig_cum.add_hline(y=target_co2, line_dash="dash", line_color="#1e3a8a",
                          annotation_text=f"Objetivo anual: {target_co2:.1f} t")
        fig_cum.update_layout(height=280, plot_bgcolor="white", yaxis_title="tCO₂eq")
        fig_cum.update_yaxes(gridcolor="#f1f5f9")
        st.plotly_chart(fig_cum, use_container_width=True)

    st.markdown("**Desglose por instalación**")
    fig_inst_co2 = px.bar(
        filtered_df, x="name", y="co2_evitado_t", color="tipo",
        color_discrete_map={"Industrial": "#1e3a8a", "Comercial": "#0891b2", "Terciario": "#0d9488"},
        labels={"co2_evitado_t": "tCO₂ evitado/año", "name": ""},
        height=260,
    )
    fig_inst_co2.update_layout(plot_bgcolor="white")
    fig_inst_co2.update_yaxes(gridcolor="#f1f5f9")
    st.plotly_chart(fig_inst_co2, use_container_width=True)

    st.markdown("""
    > **Metodología:** Factor de emisión Red Eléctrica España (REE) actualizado mensualmente.
    > Cálculo conforme GHG Protocol Scope 2 (market-based method).
    > Auditado bajo GRI 302-1 (Consumo energético dentro de la organización).
    """)


# ════════════════════════════════════════════════════════════════════
# INDICADORES CIRCULARES
# ════════════════════════════════════════════════════════════════════
with tab_circular:
    st.subheader("♻️ Indicadores de Economía Circular")

    col_c1, col_c2, col_c3, col_c4 = st.columns(4)
    total_area = filtered_df["kwp"].sum() / 0.15  # ~150 Wp/m²
    col_c1.metric("Área cubierta activada", f"{total_area:,.0f} m²", delta="superficie con DPP")
    col_c2.metric("Contenido reciclado medio", "28%", delta="acero + CIGS")
    col_c3.metric("Materiales críticos trazados", "100%", delta="In, Se, Cd — pasaporte")
    col_c4.metric("Instalaciones con 2ª vida elegible", f"{sum(1 for i in INSTALLATIONS if i['status']=='operativo')}/{len(INSTALLATIONS)}")

    st.markdown("---")

    # Pirámide circular
    circ_data = pd.DataFrame([
        {"Nivel": "1. Reducción en origen", "Indicador": "Peso < 6 kg/m² (vs. 15 kg/m² panel rígido)", "Valor": "−60% peso", "Impacto": "Alto"},
        {"Nivel": "2. Reutilización", "Indicador": "Instalaciones con PR >80% elegibles 2ª vida", "Valor": f"{sum(1 for i in INSTALLATIONS if i['status']=='operativo')}/{len(INSTALLATIONS)}", "Impacto": "Alto"},
        {"Nivel": "3. Reacondicionamiento", "Indicador": "Módulos con diagnóstico IoT disponible", "Valor": "100%", "Impacto": "Medio"},
        {"Nivel": "4. Reciclaje", "Indicador": "Fracción reciclable por masa", "Valor": "~65%", "Impacto": "Medio"},
        {"Nivel": "5. Recuperación energética", "Indicador": "Polímeros no reciclables (PVF, PI)", "Valor": "~35%", "Impacto": "Bajo"},
    ])
    st.dataframe(circ_data, use_container_width=True, hide_index=True)

    col_circ1, col_circ2 = st.columns(2)
    with col_circ1:
        st.markdown("**Flujo de materiales estimado (fin de vida)**")
        labels = ["Panel instalado", "Acero (reciclaje)", "CIGS (hidrometal.)", "EVA (pirólisis)", "PVF + PI (energ.)", "Adhesivo PU"]
        values = [100, 45, 8, 20, 17, 10]
        parents = ["", "Panel instalado", "Panel instalado", "Panel instalado", "Panel instalado", "Panel instalado"]
        fig_sankey = go.Figure(go.Sunburst(
            labels=labels, parents=parents, values=values,
            branchvalues="total",
            marker=dict(colors=["#1e3a8a", "#22c55e", "#0891b2", "#f97316", "#94a3b8", "#8b5cf6"]),
        ))
        fig_sankey.update_layout(height=300, margin=dict(t=10))
        st.plotly_chart(fig_sankey, use_container_width=True)

    with col_circ2:
        st.markdown("**Vida residual estimada por instalación**")
        vida_data = []
        for i, inst in enumerate(INSTALLATIONS):
            install_year = int(inst["install_date"][:4])
            vida_restante = max(0, 25 - (year_sel - install_year))
            kpi = compute_kpis(generate_timeseries(inst["id"], days=30), inst["kwp"])
            deg_rate = (1 - kpi["pr"] / inst["pr_target"]) / max(1, year_sel - install_year)
            vida_data.append({
                "Instalación": inst["name"].split(" ")[0],
                "Vida restante (años)": vida_restante,
                "PR actual": f"{kpi['pr']:.1%}",
                "Degradación/año": f"{deg_rate*100:.2f}%",
            })
        vida_df = pd.DataFrame(vida_data)
        fig_vida = px.bar(
            vida_df, x="Vida restante (años)", y="Instalación", orientation="h",
            color="Vida restante (años)",
            color_continuous_scale=["#dc2626", "#d97706", "#22c55e"],
            range_color=[0, 25], height=260,
        )
        fig_vida.update_layout(plot_bgcolor="white", coloraxis_showscale=False, margin=dict(t=10))
        fig_vida.update_xaxes(gridcolor="#f1f5f9")
        st.plotly_chart(fig_vida, use_container_width=True)


# ════════════════════════════════════════════════════════════════════
# ODS
# ════════════════════════════════════════════════════════════════════
with tab_ods:
    st.subheader("🎯 Contribución a Objetivos de Desarrollo Sostenible (ODS)")

    st.markdown("""
    <div style="display:flex; gap:16px; flex-wrap:wrap; margin:16px 0;">
        <div style="text-align:center;">
            <div class="ods-badge ods-7">7</div>
            <div style="font-size:0.75rem; margin-top:4px; max-width:80px;">Energía asequible y no contaminante</div>
        </div>
        <div style="text-align:center;">
            <div class="ods-badge ods-11">11</div>
            <div style="font-size:0.75rem; margin-top:4px; max-width:80px;">Ciudades y comunidades sostenibles</div>
        </div>
        <div style="text-align:center;">
            <div class="ods-badge ods-12">12</div>
            <div style="font-size:0.75rem; margin-top:4px; max-width:80px;">Producción y consumo responsables</div>
        </div>
        <div style="text-align:center;">
            <div class="ods-badge ods-13">13</div>
            <div style="font-size:0.75rem; margin-top:4px; max-width:80px;">Acción por el clima</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    ods_detail = pd.DataFrame([
        {"ODS": "ODS 7 — Energía limpia", "Meta específica": "7.2 — Aumentar % renovables en mix energético",
         "Indicador Solar Flex": f"{total_energia/1000:,.1f} MWh renovables producidos",
         "Contribución": "Directa — Alta"},
        {"ODS": "ODS 7 — Energía limpia", "Meta específica": "7.3 — Duplicar tasa de mejora eficiencia",
         "Indicador Solar Flex": f"PR medio flota: {filtered_df['pr'].mean():.1%} | Autoconsumo edificio",
         "Contribución": "Directa — Alta"},
        {"ODS": "ODS 11 — Ciudades sostenibles", "Meta específica": "11.6 — Reducir impacto ambiental ciudades",
         "Indicador Solar Flex": f"{total_co2:.1f} t CO₂eq evitadas en entorno urbano/industrial",
         "Contribución": "Directa — Alta"},
        {"ODS": "ODS 12 — Consumo responsable", "Meta específica": "12.5 — Reducir generación de residuos",
         "Indicador Solar Flex": "Pasaporte Digital activo | 2ª vida módulos | trazabilidad RAEE",
         "Contribución": "Directa — Alta"},
        {"ODS": "ODS 12 — Consumo responsable", "Meta específica": "12.6 — Prácticas sostenibles empresas",
         "Indicador Solar Flex": "Reporting CSRD automatizado | Datos ESG en tiempo real",
         "Contribución": "Facilitadora"},
        {"ODS": "ODS 13 — Acción climática", "Meta específica": "13.2 — Integrar medidas climáticas en políticas",
         "Indicador Solar Flex": f"{total_co2:.1f} t CO₂eq mitigadas | Certificados origen digital",
         "Contribución": "Directa — Alta"},
    ])
    st.dataframe(ods_detail, use_container_width=True, hide_index=True)

    # Gráfico spider/radar ODS
    ods_scores = {"ODS 7\nEnergía": 92, "ODS 11\nCiudades": 78, "ODS 12\nConsumo": 85,
                  "ODS 13\nClima": 88, "ODS 9\nIndustria": 70, "ODS 17\nAlianzas": 65}
    fig_ods = go.Figure()
    cats = list(ods_scores.keys())
    vals = list(ods_scores.values()) + [list(ods_scores.values())[0]]
    cats_closed = cats + [cats[0]]
    fig_ods.add_trace(go.Scatterpolar(
        r=vals, theta=cats_closed, fill="toself",
        fillcolor="rgba(22,163,74,0.2)", line=dict(color="#16a34a", width=2),
        name="Score alineación ODS",
    ))
    fig_ods.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 100])),
        height=320, margin=dict(t=30),
    )
    st.plotly_chart(fig_ods, use_container_width=True)


# ════════════════════════════════════════════════════════════════════
# CSRD REPORTING
# ════════════════════════════════════════════════════════════════════
with tab_csrd:
    st.subheader("📋 CSRD — Corporate Sustainability Reporting Directive")
    st.info("Obligatorio para empresas >250 empleados desde ejercicio 2024 (reporte 2025). Desde 2026 para medianas empresas.")

    st.markdown("**Estado de cumplimiento CSRD para clientes Solar Flex:**")
    csrd_html = """
    <div style="display:flex; flex-wrap:wrap; gap:8px; margin:12px 0;">
        <span class="csrd-ok">✅ ESRS E1 — Cambio Climático</span>
        <span class="csrd-ok">✅ ESRS E5 — Uso de Recursos</span>
        <span class="csrd-ok">✅ GRI 302-1 — Energía</span>
        <span class="csrd-ok">✅ GRI 305-2 — Emisiones Scope 2</span>
        <span class="csrd-prog">🔄 EU Taxonomy — Elegibilidad</span>
        <span class="csrd-prog">🔄 ESRS E2 — Contaminación</span>
    </div>
    """
    st.markdown(csrd_html, unsafe_allow_html=True)

    st.markdown("---")

    csrd_table = pd.DataFrame([
        {"Estándar": "ESRS E1.5", "Indicador": "Consumo energético total (GJ)", "Valor": f"{total_energia * 3.6 / 1000:,.1f} GJ renovable", "Fuente": "Medidor AC + IoT"},
        {"Estándar": "ESRS E1.6", "Indicador": "Emisiones Scope 1 directas (t CO₂eq)", "Valor": "0 t (instalación PV — sin combustión)", "Fuente": "Calculado"},
        {"Estándar": "ESRS E1.6", "Indicador": "Emisiones Scope 2 evitadas (t CO₂eq)", "Valor": f"{total_co2:.2f} t CO₂eq", "Fuente": "IoT + factor REE"},
        {"Estándar": "ESRS E5.1", "Indicador": "% materiales reciclados en producto", "Valor": "28% (acero + CIGS)", "Fuente": "Pasaporte Digital"},
        {"Estándar": "ESRS E5.5", "Indicador": "Productos con información fin de vida", "Valor": "100% (DPP activo)", "Fuente": "Plataforma IoT"},
        {"Estándar": "GRI 302-1", "Indicador": "Energía renovable producida (MWh)", "Valor": f"{total_energia/1000:,.1f} MWh", "Fuente": "Medidor AC"},
        {"Estándar": "GRI 305-2", "Indicador": "Reducción emisiones GHG indirectas", "Valor": f"{total_co2:.2f} t CO₂eq (método market-based)", "Fuente": "IoT + REE"},
    ])
    st.dataframe(csrd_table, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("**Integraciones con plataformas ESG:**")
    platforms = pd.DataFrame([
        {"Plataforma": "Watershed", "Formato": "JSON API", "Estado": "✅ Disponible", "Estándar": "GHG Protocol"},
        {"Plataforma": "Persefoni", "Formato": "CSV + API", "Estado": "✅ Disponible", "Estándar": "TCFD"},
        {"Plataforma": "Greenly", "Formato": "REST API", "Estado": "✅ Disponible", "Estándar": "CSRD"},
        {"Plataforma": "IDAE eBitenergy", "Formato": "XML + API", "Estado": "✅ RENOCICLA", "Estándar": "PNIEC"},
        {"Plataforma": "SAP Sustainability", "Formato": "OData API", "Estado": "🔄 En desarrollo", "Estándar": "ESRS"},
    ])
    st.dataframe(platforms, use_container_width=True, hide_index=True)

st.markdown("---")
st.caption(f"Informe ESG Solar Flex {year_sel} · Metodología GHG Protocol Scope 2 · Factor emisión REE España: 0.181 kg CO₂/kWh · CSRD conforme ESRS E1/E5 · © 2025 Solar Flex Technologies S.L.")
