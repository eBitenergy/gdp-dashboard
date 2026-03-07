"""
Solar Flex — Pasaporte Digital de Producto (DPP)
Cumplimiento ESPR Reglamento (UE) 2024/1781
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from data.solar_flex_simulator import INSTALLATIONS, generate_timeseries, compute_kpis, DPP_MATERIALS

st.set_page_config(page_title="Pasaporte Digital — Solar Flex", page_icon="📋", layout="wide")

st.markdown("""
<style>
.dpp-header { background: linear-gradient(135deg, #1e3a5f 0%, #0891b2 100%);
              color:white; padding:20px; border-radius:10px; margin-bottom:16px; }
.dpp-id { font-family: monospace; font-size:1.3rem; font-weight:700; letter-spacing:2px; }
.cert-badge { display:inline-block; background:#dbeafe; color:#1d4ed8; padding:4px 10px;
              border-radius:12px; font-size:0.8rem; font-weight:600; margin:2px; }
.material-row { padding:6px 0; border-bottom:1px solid #f1f5f9; }
.qr-placeholder { background:#f8fafc; border:2px dashed #cbd5e1; border-radius:8px;
                  padding:40px; text-align:center; font-size:2rem; }
.espr-badge { background:#dcfce7; color:#15803d; padding:8px 14px; border-radius:8px;
              font-weight:700; font-size:0.9rem; }
</style>
""", unsafe_allow_html=True)

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📋 Pasaporte Digital (DPP)")
    inst_names = {i["id"]: i["name"] for i in INSTALLATIONS}
    selected_id = st.selectbox("Seleccionar instalación", list(inst_names.keys()),
                               format_func=lambda x: inst_names[x])
    st.markdown("---")
    st.markdown("**Marco normativo**")
    st.markdown("📜 ESPR Reg. (UE) 2024/1781")
    st.markdown("📜 CSRD (Sostenibilidad)")
    st.markdown("📜 EPBD (Building Log Book)")
    st.markdown("📜 Baterías EU 2023/1542")
    st.markdown("---")
    st.markdown("""
    <div class="espr-badge">✅ Conforme ESPR 2026</div>
    """, unsafe_allow_html=True)

inst = next(i for i in INSTALLATIONS if i["id"] == selected_id)

# ── Header DPP ────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="dpp-header">
    <div style="display:flex; justify-content:space-between; align-items:center;">
        <div>
            <div style="font-size:0.85rem; opacity:0.8; margin-bottom:4px;">PASAPORTE DIGITAL DE PRODUCTO — ESPR (UE) 2024/1781</div>
            <div class="dpp-id">🪪 {inst['id']}</div>
            <div style="margin-top:8px; font-size:1.1rem; font-weight:600;">{inst['name']}</div>
            <div style="opacity:0.8; font-size:0.9rem;">{inst['location']} · {inst['kwp']} kWp · {inst['panel_model']}</div>
        </div>
        <div style="text-align:center; opacity:0.9;">
            <div style="font-size:3rem;">📱</div>
            <div style="font-size:0.75rem;">QR + NFC disponible</div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

tab1, tab2, tab3, tab4 = st.tabs(["📦 Bloque A — Identidad", "📈 Bloque B — Historial de Vida", "♻️ Fin de Vida", "🔗 Verificación Blockchain"])

# ════════════════════════════════════════════════════════════════════
# BLOQUE A — Identidad y Composición (estático)
# ════════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("Bloque A — Identidad y Composición (datos estáticos de fabricación)")
    st.info("Creado en fabricación. Inmutable. Hash certificado en blockchain Polygon.")

    col_id, col_qr = st.columns([2, 1])

    with col_id:
        st.markdown("#### 🏷️ Identificación única")
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown(f"""
| Campo | Valor |
|---|---|
| **ID único** | `{inst['id']}` |
| **Modelo panel** | {inst['panel_model']} |
| **Fecha fabricación** | {inst['install_date']} |
| **Planta fabricación** | Solar Flex Prod. — Guadalajara |
| **Lote** | LOTE-2025-{inst['id'][-5:]} |
| **Ref. Modelo de Utilidad** | OEPM alfaDOC U202400{inst['id'][-3:]} |
            """)
        with col_b:
            st.markdown(f"""
| Campo | Valor |
|---|---|
| **Adhesivo estructural** | {inst['adhesive']} |
| **Potencia pico** | {inst['kwp']} kWp |
| **Área instalada** | {inst['area_m2']} m² |
| **Peso específico** | {5.8} kg/m² |
| **Tipo edificio** | {inst['type']} |
| **Coordenadas** | {inst['lat']:.4f}°N, {inst['lon']:.4f}°E |
            """)

        st.markdown("#### 🏅 Certificaciones")
        certs = ["IEC 61215", "IEC 61730", "BROOF(t1)", "TÜV Rheinland", "CE", "RoHS", "RENOCICLA Apto"]
        cert_html = " ".join([f'<span class="cert-badge">{c}</span>' for c in certs])
        st.markdown(cert_html, unsafe_allow_html=True)

    with col_qr:
        st.markdown("#### 📱 Acceso digital")
        st.markdown(f"""
<div class="qr-placeholder">
    ▓▓▓▓▓▓▓<br>
    ▓ ░░░ ▓<br>
    ▓ ░▓░ ▓<br>
    ▓ ░░░ ▓<br>
    ▓▓▓▓▓▓▓
</div>
        """, unsafe_allow_html=True)
        st.markdown(f"🔗 `dpp.solarflex.es/{inst['id']}`")
        st.markdown("📡 NFC: **ST25DV-I2C** (IP67)")
        st.markdown("⛓️ Hash Polygon: `0x7f3a...b94e`")

    st.markdown("---")
    st.markdown("#### 🧱 Composición de materiales por capa")

    mat_data = DPP_MATERIALS.get(inst["panel_model"], DPP_MATERIALS["FlexCIGS-200W"])
    mat_df = pd.DataFrame(mat_data["capas"])
    mat_df["reciclable"] = mat_df["reciclable"].apply(lambda x: "♻️ Sí" if x else "❌ No")
    mat_df["pct_reciclado_contenido"] = mat_df["pct_reciclado_contenido"].apply(lambda x: f"{x}%" if x > 0 else "—")
    mat_df.columns = ["Capa", "Material", "Espesor (mm)", "Reciclable", "% Contenido reciclado"]
    st.dataframe(mat_df, use_container_width=True, hide_index=True)

    # Gráfico de composición
    st.markdown("#### 📊 Perfil de materiales")
    mat_raw = pd.DataFrame(mat_data["capas"])
    fig_mat = px.bar(
        mat_raw, x="espesor_mm", y="capa", orientation="h",
        color="reciclable", color_discrete_map={True: "#22c55e", False: "#f97316"},
        labels={"espesor_mm": "Espesor (mm)", "capa": "", "reciclable": "¿Reciclable?"},
        height=300,
    )
    fig_mat.update_layout(plot_bgcolor="white", margin=dict(t=10))
    fig_mat.update_xaxes(gridcolor="#f1f5f9")
    st.plotly_chart(fig_mat, use_container_width=True)


# ════════════════════════════════════════════════════════════════════
# BLOQUE B — Historial de Vida (dinámico, actualizado por IoT)
# ════════════════════════════════════════════════════════════════════
with tab2:
    st.subheader("Bloque B — Historial de Vida (datos dinámicos IoT)")
    st.info("Actualizado en tiempo real desde sensores embebidos. Cada evento crítico genera hash en blockchain.")

    @st.cache_data(ttl=300)
    def load_history(iid):
        return generate_timeseries(iid, days=90, freq_minutes=60)

    df90 = load_history(selected_id)
    kpis90 = compute_kpis(df90, inst["kwp"])

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Energía acumulada (90d)", f"{kpis90['energia_kwh']:,.0f} kWh")
    col2.metric("PR medio histórico", f"{kpis90['pr']:.1%}")
    deg_real = (1 - kpis90["pr"] / inst["pr_target"]) * 100
    col3.metric("Degradación real", f"{deg_real:.1f}%", delta="vs. curva garantía fabricante",
                delta_color="inverse" if deg_real > 5 else "off")
    co2_total = kpis90["co2_evitado_t"]
    col4.metric("CO₂ evitado acumulado", f"{co2_total:.2f} t CO₂", delta="tCO₂eq")

    st.markdown("---")

    # Performance histórico
    df_daily = df90.set_index("timestamp").resample("D").agg({
        "potencia_ac_kw": "sum", "irradiancia_poa": "sum",
        "temperatura_celula": "max", "strain_adhesivo_1_ue": "max",
    }).reset_index()
    df_daily["energia_kwh_d"] = df_daily["potencia_ac_kw"] * 1.0
    df_daily["irr_d"] = df_daily["irradiancia_poa"] / 1000
    df_daily["pr_d"] = (df_daily["energia_kwh_d"] / (df_daily["irr_d"] * inst["kwp"])).clip(0, 1)
    df_daily["co2_d"] = df_daily["energia_kwh_d"] * 0.181 / 1000
    df_daily["co2_acum"] = df_daily["co2_d"].cumsum()

    col_pr, col_co2 = st.columns(2)

    with col_pr:
        st.markdown("**PR diario histórico vs curva garantía**")
        fig_pr_hist = go.Figure()
        fig_pr_hist.add_trace(go.Scatter(
            x=df_daily["timestamp"], y=df_daily["pr_d"] * 100,
            fill="tozeroy", fillcolor="rgba(34,197,94,0.15)",
            line=dict(color="#22c55e", width=1.5), name="PR real (%)",
        ))
        fig_pr_hist.add_hline(y=inst["pr_target"] * 100, line_dash="dash",
                               line_color="#1e3a8a", annotation_text="Garantía fabricante")
        fig_pr_hist.update_layout(height=250, plot_bgcolor="white",
                                   yaxis_title="PR (%)", yaxis_range=[60, 95])
        fig_pr_hist.update_yaxes(gridcolor="#f1f5f9")
        st.plotly_chart(fig_pr_hist, use_container_width=True)

    with col_co2:
        st.markdown("**CO₂ evitado acumulado (tCO₂)**")
        fig_co2 = go.Figure()
        fig_co2.add_trace(go.Scatter(
            x=df_daily["timestamp"], y=df_daily["co2_acum"],
            fill="tozeroy", fillcolor="rgba(34,197,94,0.2)",
            line=dict(color="#16a34a", width=2), name="tCO₂ acumulado",
        ))
        fig_co2.update_layout(height=250, plot_bgcolor="white",
                               yaxis_title="tCO₂ evitado")
        fig_co2.update_yaxes(gridcolor="#f1f5f9")
        st.plotly_chart(fig_co2, use_container_width=True)

    st.markdown("**Evolución strain adhesivo — diagnóstico integridad**")
    fig_strain = go.Figure()
    fig_strain.add_trace(go.Scatter(
        x=df_daily["timestamp"], y=df_daily["strain_adhesivo_1_ue"],
        line=dict(color="#7c3aed", width=1.5), name="Strain máx. diario (µε)",
    ))
    fig_strain.add_hline(y=350, line_dash="dash", line_color="#dc2626",
                          annotation_text="Umbral alerta despegue incipiente")
    fig_strain.add_hrect(y0=0, y1=200, fillcolor="rgba(34,197,94,0.05)", line_width=0, annotation_text="Zona OK")
    fig_strain.add_hrect(y0=200, y1=350, fillcolor="rgba(251,191,36,0.07)", line_width=0, annotation_text="Zona vigilancia")
    fig_strain.update_layout(height=220, plot_bgcolor="white", yaxis_title="µε")
    fig_strain.update_yaxes(gridcolor="#f1f5f9")
    st.plotly_chart(fig_strain, use_container_width=True)

    st.markdown("**Registro de eventos críticos**")
    events = pd.DataFrame([
        {"Fecha": "2025-04-12", "Tipo": "CLIMÁTICO", "Descripción": "Viento máximo 23.4 m/s — carga estructural registrada", "Severidad": "Media"},
        {"Fecha": "2025-06-28", "Tipo": "RENDIMIENTO", "Descripción": "PR diario 71.2% — limpieza realizada al día siguiente (+8% recuperado)", "Severidad": "Baja"},
        {"Fecha": inst["install_date"], "Tipo": "INSTALACIÓN", "Descripción": "Instalación completada. Ensayo adhesivo: OK. PR inicial: 82.3%", "Severidad": "Info"},
    ])
    st.dataframe(events, use_container_width=True, hide_index=True)


# ════════════════════════════════════════════════════════════════════
# FIN DE VIDA
# ════════════════════════════════════════════════════════════════════
with tab3:
    st.subheader("♻️ Información de Fin de Vida — Economía Circular")

    mat_data = DPP_MATERIALS.get(inst["panel_model"], DPP_MATERIALS["FlexCIGS-200W"])
    fov = mat_data["fin_de_vida"]

    col_inst, col_recycl = st.columns(2)
    with col_inst:
        st.markdown("#### 🔧 Instrucciones de desinstalación")
        st.markdown(f"""
**Procedimiento recomendado:**
{fov['instrucciones']}

**Herramientas necesarias:**
- Espátula calefactable (>60°C)
- EPIs: guantes resistentes a solventes
- Recipientes RAEE homologados

**Tiempo estimado:** 2-3 h por cada 100 m²
        """)

    with col_recycl:
        st.markdown("#### ♻️ Reciclaje y recuperación")
        st.markdown("**Gestores RAEE autorizados:**")
        for g in fov["gestores_autorizados"]:
            st.markdown(f"- ✅ {g}")

        st.markdown("**Materiales críticos a recuperar:**")
        for m in fov["contenido_critico"]:
            st.markdown(f"- ⚠️ {m}")

        st.markdown("**Recuperabilidad estimada:**")
        rec_data = pd.DataFrame([
            {"Fracción": "Acero galvanizado", "Recuperable (%)": 95, "Método": "Fundición"},
            {"Fracción": "CIGS (metales críticos)", "Recuperable (%)": 75, "Método": "Hidrometalurgia"},
            {"Fracción": "EVA encapsulante", "Recuperable (%)": 30, "Método": "Pirólisis"},
            {"Fracción": "Poliimida sustrato", "Recuperable (%)": 0, "Método": "Valorización energética"},
            {"Fracción": "PVF tedlar", "Recuperable (%)": 0, "Método": "Valorización energética"},
            {"Fracción": "Adhesivo PU", "Recuperable (%)": 20, "Método": "Pirólisis/GQ"},
        ])
        st.dataframe(rec_data, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("#### 🔄 Evaluación automática segunda vida")
    kpis_cur = compute_kpis(generate_timeseries(selected_id, days=30), inst["kwp"])
    pr_actual = kpis_cur["pr"]
    pr_segunda_vida_umbral = 0.80

    if pr_actual >= pr_segunda_vida_umbral:
        st.success(f"✅ **Apto para segunda vida** — PR actual {pr_actual:.1%} ≥ umbral {pr_segunda_vida_umbral:.0%}")
        st.markdown("Instalación elegible para: reutilización en edificio de menor exigencia, comunidad energética rural, o reacondicionamiento para exportación.")
    else:
        st.warning(f"♻️ **Recomendado para reciclaje** — PR {pr_actual:.1%} < {pr_segunda_vida_umbral:.0%}. Contactar con gestor RAEE.")

    fig_vida = go.Figure(go.Indicator(
        mode="gauge+number+delta",
        value=pr_actual * 100,
        delta={"reference": pr_segunda_vida_umbral * 100, "valueformat": ".1f"},
        gauge={
            "axis": {"range": [50, 100]},
            "bar": {"color": "#22c55e" if pr_actual >= pr_segunda_vida_umbral else "#f97316"},
            "steps": [
                {"range": [50, 70], "color": "#fee2e2"},
                {"range": [70, 80], "color": "#fef3c7"},
                {"range": [80, 100], "color": "#dcfce7"},
            ],
            "threshold": {"line": {"color": "#dc2626", "width": 3}, "thickness": 0.8, "value": pr_segunda_vida_umbral * 100},
        },
        title={"text": "PR actual (%) — Umbral segunda vida: 80%"},
        number={"suffix": "%"},
    ))
    fig_vida.update_layout(height=280)
    st.plotly_chart(fig_vida, use_container_width=True)


# ════════════════════════════════════════════════════════════════════
# BLOCKCHAIN
# ════════════════════════════════════════════════════════════════════
with tab4:
    st.subheader("⛓️ Verificación blockchain — Trazabilidad inmutable")
    st.info("Cada evento crítico del pasaporte genera un hash en Polygon (PoS). Coste por transacción: <0.01€.")

    blockchain_records = pd.DataFrame([
        {"#": 1, "Fecha": inst["install_date"], "Evento": "CREACIÓN PASAPORTE",
         "Hash Tx": "0x7f3a8c2d1e...b94e", "Bloque": "48291043", "Verificado": "✅"},
        {"#": 2, "Fecha": "2025-04-12", "Evento": "EVENTO CLIMÁTICO — Viento 23.4 m/s",
         "Hash Tx": "0x3b2f1a9c4d...7e21", "Bloque": "51038294", "Verificado": "✅"},
        {"#": 3, "Fecha": "2025-06-28", "Evento": "MANTENIMIENTO — Limpieza módulos",
         "Hash Tx": "0x9d4c2b8f3a...1c57", "Bloque": "55192847", "Verificado": "✅"},
        {"#": 4, "Fecha": "2025-09-15", "Evento": "ACTUALIZACIÓN PR — Degradación anual: 0.48%/año",
         "Hash Tx": "0x1e7b3d5f2a...8d34", "Bloque": "59847203", "Verificado": "✅"},
        {"#": 5, "Fecha": "2026-01-01", "Evento": "CERTIFICADO ORIGEN DIGITAL — 18,432 kWh renovables",
         "Hash Tx": "0x4c8a2f9e1b...3f76", "Bloque": "64203847", "Verificado": "✅"},
    ])
    st.dataframe(blockchain_records, use_container_width=True, hide_index=True)

    col_bc1, col_bc2, col_bc3 = st.columns(3)
    col_bc1.metric("Red blockchain", "Polygon PoS")
    col_bc2.metric("Transacciones registradas", len(blockchain_records))
    col_bc3.metric("Coste total blockchain", f"€{len(blockchain_records)*0.008:.3f}")

    st.markdown("---")
    st.markdown("""
    **Arquitectura de trazabilidad:**
    - Identificador único físico: QR dinámico + NFC ST25DV-I2C (IP67)
    - Hash de composición de materiales: SHA-256 → Polygon
    - Eventos de vida: firmados con clave privada del operador → Polygon
    - Acceso público: `dpp.solarflex.es/<ID>` (sin registro)
    - API privada: integraciones ERP/ESG/RAEE via REST + JWT
    """)

st.markdown("---")
st.caption("Pasaporte Digital Solar Flex · Conforme ESPR Reg. (UE) 2024/1781 · CSRD · Modelo de utilidad OEPM alfaDOC · © 2025 Solar Flex Technologies S.L.")
