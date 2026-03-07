"""
Solar Flex — Arquitectura Sensórica e IoT
Documentación técnica del stack tecnológico completo
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from data.solar_flex_simulator import SENSOR_SPECS

st.set_page_config(page_title="Arquitectura Sensórica — Solar Flex", page_icon="📡", layout="wide")

st.markdown("""
<style>
.layer-card { border-radius:10px; padding:16px 20px; margin:8px 0; }
.layer-4 { background: linear-gradient(90deg, #1e3a5f, #1e40af); color:white; }
.layer-3 { background: linear-gradient(90deg, #0c4a6e, #0369a1); color:white; }
.layer-2 { background: linear-gradient(90deg, #164e63, #0891b2); color:white; }
.layer-1 { background: linear-gradient(90deg, #134e4a, #0d9488); color:white; }
.sensor-tag { display:inline-block; padding:3px 9px; border-radius:12px; font-size:0.78rem;
              font-weight:600; margin:2px; background:#e2e8f0; color:#374151; }
.domain-pv     { background:#fef3c7; color:#92400e; }
.domain-struct { background:#ede9fe; color:#5b21b6; }
.domain-amb    { background:#dcfce7; color:#065f46; }
.domain-build  { background:#dbeafe; color:#1e40af; }
.tech-card { background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px;
             padding:14px; margin:6px 0; }
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 📡 Arquitectura Sensórica")
    view = st.radio("Vista", ["Capas del sistema", "Catálogo sensores", "Stack tecnológico", "Roadmap"])
    st.markdown("---")
    st.markdown("**Instalación referencia:**")
    st.markdown("Nave Logística Guadalajara")
    st.markdown("120 kWp · 800 m²")
    st.markdown("---")
    st.markdown("**Protocolos activos:**")
    for p in ["LoRaWAN", "Modbus RS485", "MQTT", "4G Cat-M1"]:
        st.markdown(f"✅ {p}")

st.markdown("# 📡 Arquitectura Sensórica e IoT")
st.markdown("Stack tecnológico completo: Sensórica → Conectividad → Cloud → Inteligencia")
st.markdown("---")

if view == "Capas del sistema":
    st.subheader("🏗️ Arquitectura en 4 capas")

    layers = [
        ("CAPA 4 — INTELIGENCIA & NEGOCIO", "layer-4",
         "Digital Twin · Pasaporte Digital (ESPR) · Reporting ESG · Marketplace O&M · Comunidades Energéticas",
         "☁️ SaaS Platform", ["Thingsboard", "Eclipse Ditto", "Grafana", "MLflow", "Polygon blockchain"]),
        ("CAPA 3 — PLATAFORMA CLOUD", "layer-3",
         "Data Lakehouse · Stream Processing · AI/ML · REST APIs · Time-Series DB",
         "☁️ OVHcloud / Azure Spain Central", ["InfluxDB/TimescaleDB", "Apache Kafka", "FastAPI", "MinIO/S3", "Neo4j"]),
        ("CAPA 2 — CONECTIVIDAD", "layer-2",
         "Edge Gateway local · LoRaWAN backhaul · 4G/NB-IoT · TLS 1.3 · Certificados x.509",
         "📡 Edge + WAN", ["LoRaWAN Chirpstack", "4G Cat-M1 SIM", "MQTT Broker (EMQX)", "Advantech UNO / Siemens IOT2050"]),
        ("CAPA 1 — SENSÓRICA EMBEBIDA", "layer-1",
         "Sensores físicos en cubierta y panel: PV, Estructural-Adhesivo, Ambiental, Building",
         "🔌 On-site Hardware", ["Piranómetro ISO 9060", "Strain gauge HBK", "MEMS acelerómetro", "Modbus RS485", "I²C / NB-IoT"]),
    ]

    for title, css, desc, infra, tech in layers:
        st.markdown(f"""
<div class="layer-card {css}">
    <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:10px;">
        <div style="flex:1;">
            <div style="font-size:1rem; font-weight:700; margin-bottom:4px;">{title}</div>
            <div style="font-size:0.88rem; opacity:0.9;">{desc}</div>
            <div style="margin-top:8px; font-size:0.78rem; opacity:0.8;">{infra}</div>
        </div>
        <div style="font-size:0.78rem; opacity:0.85; min-width:180px;">
            <strong>Tech stack:</strong><br>{'<br>'.join([f'• {t}' for t in tech])}
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

    st.markdown("---")
    st.subheader("🔄 Flujo de datos")

    # Diagrama de flujo simplificado con plotly
    fig_flow = go.Figure()
    nodes = {
        "Sensores": (1, 1), "Edge Gateway": (3, 1), "MQTT Broker": (5, 2),
        "InfluxDB": (7, 2), "Stream\nProcessing": (7, 1), "ML / AI": (9, 2),
        "Digital Twin": (9, 1), "Dashboard": (11, 1.5), "APIs": (11, 0.5),
    }
    colors = {
        "Sensores": "#0d9488", "Edge Gateway": "#0891b2", "MQTT Broker": "#0369a1",
        "InfluxDB": "#1d4ed8", "Stream\nProcessing": "#1d4ed8", "ML / AI": "#7c3aed",
        "Digital Twin": "#7c3aed", "Dashboard": "#1e3a8a", "APIs": "#1e3a8a",
    }
    edges = [
        ("Sensores", "Edge Gateway"), ("Edge Gateway", "MQTT Broker"),
        ("MQTT Broker", "InfluxDB"), ("MQTT Broker", "Stream\nProcessing"),
        ("InfluxDB", "ML / AI"), ("Stream\nProcessing", "Digital Twin"),
        ("ML / AI", "Dashboard"), ("Digital Twin", "Dashboard"),
        ("InfluxDB", "APIs"),
    ]

    for n, (x, y) in nodes.items():
        fig_flow.add_trace(go.Scatter(
            x=[x], y=[y], mode="markers+text",
            marker=dict(size=40, color=colors[n], symbol="square"),
            text=[n], textposition="middle center",
            textfont=dict(color="white", size=9, family="monospace"),
            hoverinfo="skip", showlegend=False,
        ))
    for src, dst in edges:
        sx, sy = nodes[src]
        dx, dy = nodes[dst]
        fig_flow.add_annotation(
            x=dx - 0.4, y=dy, ax=sx + 0.4, ay=sy,
            xref="x", yref="y", axref="x", ayref="y",
            arrowhead=2, arrowsize=1, arrowwidth=2, arrowcolor="#94a3b8",
        )
    fig_flow.update_layout(
        height=240, plot_bgcolor="white",
        xaxis=dict(showgrid=False, zeroline=False, showticklabels=False, range=[0, 12]),
        yaxis=dict(showgrid=False, zeroline=False, showticklabels=False, range=[0, 3]),
        margin=dict(t=10, b=10),
    )
    st.plotly_chart(fig_flow, use_container_width=True)

elif view == "Catálogo sensores":
    st.subheader("🔌 Catálogo de sensores por dominio")

    domain_colors = {
        "Energético-PV": "domain-pv",
        "Estructural": "domain-struct",
        "Ambiental": "domain-amb",
        "Building": "domain-build",
    }

    sensors_catalog = pd.DataFrame([
        # PV
        {"Dominio": "Energético-PV", "Sensor": "Piranómetro/Irradiámetro", "Variable": "Irradiancia GHI/POA (W/m²)",
         "Especificación": "ISO 9060 Clase A, 0–2000 W/m²", "Protocolo": "Analógico 0-10V / RS485",
         "Ubicación": "Plano del panel", "Proveedor": "HT Instruments / Kipp&Zonen"},
        {"Dominio": "Energético-PV", "Sensor": "Sensor IV (I-V curve tracer)", "Variable": "Curva I-V por string",
         "Especificación": "Resolución 0.1% Isc", "Protocolo": "Modbus RS485",
         "Ubicación": "Por string (8-12 módulos)", "Proveedor": "Chauvin Arnoux / Ametek"},
        {"Dominio": "Energético-PV", "Sensor": "Sensor corriente DC", "Variable": "Isc, Imp string (A)",
         "Especificación": "Hall effect, 0-30A ±0.5%", "Protocolo": "I²C / RS485",
         "Ubicación": "Caja de string", "Proveedor": "LEM / Honeywell"},
        {"Dominio": "Energético-PV", "Sensor": "Sensor temperatura célula", "Variable": "Tc (°C)",
         "Especificación": "PT100/NTC, -40 a +90°C", "Protocolo": "I²C / 1-Wire",
         "Ubicación": "Reverso módulo flexible", "Proveedor": "Heraeus / TE Connectivity"},
        {"Dominio": "Energético-PV", "Sensor": "Sensor potencia AC", "Variable": "kWh producidos (neto)",
         "Especificación": "Clase 0.5, bidireccional", "Protocolo": "Modbus TCP / DLMS",
         "Ubicación": "Punto conexión red", "Proveedor": "Eastron / Carlo Gavazzi"},
        # Estructural
        {"Dominio": "Estructural", "Sensor": "Strain gauge (galga extensométrica)", "Variable": "Tensión mecánica adhesivo (µε)",
         "Especificación": "Piezorresistiva, 0-1000 µε ±0.5%", "Protocolo": "Analógico diferencial / Wheatstone",
         "Ubicación": "Unión adhesiva panel-cubierta", "Proveedor": "HBK (Hottinger) / Vishay"},
        {"Dominio": "Estructural", "Sensor": "Acelerómetro MEMS", "Variable": "Vibración, carga viento (mm/s)",
         "Especificación": "3 ejes, 0.01–50 Hz, ±2g", "Protocolo": "I²C (ADXL345) / SPI",
         "Ubicación": "Estructura cubierta", "Proveedor": "Analog Devices / STMicroelectronics"},
        {"Dominio": "Estructural", "Sensor": "Sensor humedad interfase", "Variable": "HR en capa adhesiva (%)",
         "Especificación": "Capacitivo, 0–100% HR ±2%", "Protocolo": "I²C / SHT40",
         "Ubicación": "Borde panel (zona sellado)", "Proveedor": "Sensirion / Honeywell"},
        {"Dominio": "Estructural", "Sensor": "Celda de carga thin-film", "Variable": "Presión sobre superficie (kg/m²)",
         "Especificación": "Thin-film, 0-100 kg/m²", "Protocolo": "Analógico",
         "Ubicación": "Puntos de carga críticos", "Proveedor": "Tekscan / Futek"},
        # Ambiental
        {"Dominio": "Ambiental", "Sensor": "Anemómetro ultrasónico", "Variable": "Velocidad y dirección viento",
         "Especificación": "0-60 m/s ±0.1 m/s, sin partes móviles", "Protocolo": "Modbus RS485",
         "Ubicación": "Mástil cubierta", "Proveedor": "Gill Instruments / Vaisala"},
        {"Dominio": "Ambiental", "Sensor": "Pluviómetro de báscula", "Variable": "Precipitación (mm/h)",
         "Especificación": "Resolución 0.2 mm/pulso", "Protocolo": "Pulso digital",
         "Ubicación": "Cubierta horizontal", "Proveedor": "Davis Instruments / Onset"},
        {"Dominio": "Ambiental", "Sensor": "Sensor T/HR ambiental", "Variable": "T°C exterior, HR%",
         "Especificación": "±0.2°C, ±1.8% HR", "Protocolo": "I²C / SHT40",
         "Ubicación": "Caseta ventilada cubierta", "Proveedor": "Sensirion / Bosch BME280"},
        {"Dominio": "Ambiental", "Sensor": "Sensor partículas PM2.5/PM10", "Variable": "Calidad aire / soiling",
         "Especificación": "Láser, 0-999 µg/m³", "Protocolo": "UART / SDS011",
         "Ubicación": "Zona cubierta", "Proveedor": "Nova Fitness / Plantower"},
        # Building
        {"Dominio": "Building", "Sensor": "Analizador de red AC", "Variable": "Consumo edificio, FP, armónicos",
         "Especificación": "Clase 0.5S, TRMS, 3-fases", "Protocolo": "Modbus TCP / DLMS",
         "Ubicación": "Cuadro general BT", "Proveedor": "Carlo Gavazzi / Janitza"},
        {"Dominio": "Building", "Sensor": "BMS batería (sensor)", "Variable": "SoC, SoH, T°, ciclos",
         "Especificación": "CAN Bus / UART, integrado BMS", "Protocolo": "CAN Bus / MQTT",
         "Ubicación": "Sistema BESS", "Proveedor": "BYD / CATL / Pylontech"},
    ])

    domains = sensors_catalog["Dominio"].unique()
    domain_tabs = st.tabs([f"{'🌞' if d=='Energético-PV' else '🏗️' if d=='Estructural' else '🌤️' if d=='Ambiental' else '🏭'} {d}" for d in domains])

    for tab, domain in zip(domain_tabs, domains):
        with tab:
            domain_df = sensors_catalog[sensors_catalog["Dominio"] == domain].drop(columns=["Dominio"])
            st.dataframe(domain_df, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("**Resumen de cobertura sensórica**")
    counts = sensors_catalog.groupby("Dominio").size().reset_index(name="Nº sensores")
    fig_counts = px.bar(counts, x="Dominio", y="Nº sensores", color="Dominio",
                        color_discrete_map={"Energético-PV": "#f97316", "Estructural": "#7c3aed",
                                            "Ambiental": "#22c55e", "Building": "#3b82f6"},
                        height=200)
    fig_counts.update_layout(plot_bgcolor="white", showlegend=False, margin=dict(t=10))
    fig_counts.update_yaxes(gridcolor="#f1f5f9")
    st.plotly_chart(fig_counts, use_container_width=True)

elif view == "Stack tecnológico":
    st.subheader("🛠️ Stack tecnológico recomendado")

    tech_table = pd.DataFrame([
        {"Dominio": "Sensores PV", "Partner/Tecnología": "HT Instruments / Chauvin Arnoux", "Justificación": "IV tracers homologados, integrables vía Modbus", "Coste estimado": "€200-800/sensor", "Open-source": "No"},
        {"Dominio": "Sensores estructurales", "Partner/Tecnología": "HBK (Hottinger Baldwin Messtechnik)", "Justificación": "Referencia mundial en strain gauges industriales", "Coste estimado": "€50-300/sensor", "Open-source": "No"},
        {"Dominio": "Gateway Edge", "Partner/Tecnología": "Advantech UNO-2484G / Siemens SIMATIC IOT2050", "Justificación": "Robustez industrial, -40/+70°C, certificación CE/ATEX", "Coste estimado": "€200-500/unidad", "Open-source": "No"},
        {"Dominio": "Conectividad LoRa", "Partner/Tecnología": "RAK Wireless / Kerlink iStation", "Justificación": "Módulos LoRaWAN industriales, económicos, IP67", "Coste estimado": "€30-150/nodo", "Open-source": "No"},
        {"Dominio": "LNS LoRaWAN", "Partner/Tecnología": "Chirpstack (open-source)", "Justificación": "Soberanía de datos, sin dependencia TTN, on-premise", "Coste estimado": "€0 (infra propia)", "Open-source": "Sí"},
        {"Dominio": "MQTT Broker", "Partner/Tecnología": "EMQX / Eclipse Mosquitto", "Justificación": "Alto rendimiento, clustering, QoS 0/1/2", "Coste estimado": "€0 community", "Open-source": "Sí"},
        {"Dominio": "Time-Series DB", "Partner/Tecnología": "InfluxDB / TimescaleDB", "Justificación": "Optimizados para series temporales IoT, compresión nativa", "Coste estimado": "€0 OSS / €500+/mes cloud", "Open-source": "Sí"},
        {"Dominio": "Object Storage", "Partner/Tecnología": "MinIO (S3-compatible)", "Justificación": "Soberanía europea, compatible AWS S3, on-premise", "Coste estimado": "€0 OSS", "Open-source": "Sí"},
        {"Dominio": "Plataforma IoT", "Partner/Tecnología": "Thingsboard Community", "Justificación": "IoT + Dashboard + Rules Engine integrado, muy económico", "Coste estimado": "€0 community", "Open-source": "Sí"},
        {"Dominio": "Digital Twin", "Partner/Tecnología": "Eclipse Ditto (Fase 2)", "Justificación": "Soberanía europea, estándar W3C WoT, escalable", "Coste estimado": "€0 OSS + infra", "Open-source": "Sí"},
        {"Dominio": "Visualización", "Partner/Tecnología": "Grafana", "Justificación": "Open-source, integración nativa InfluxDB, alertas", "Coste estimado": "€0 OSS", "Open-source": "Sí"},
        {"Dominio": "ML Predictivo", "Partner/Tecnología": "Python + MLflow + scikit-learn", "Justificación": "Stack estándar, auditables para convocatorias I+D", "Coste estimado": "€0 OSS", "Open-source": "Sí"},
        {"Dominio": "Blockchain", "Partner/Tecnología": "Polygon (PoS)", "Justificación": "<0.01€/tx, EVM compatible, bajo consumo energético", "Coste estimado": "<€0.01/tx", "Open-source": "Sí"},
        {"Dominio": "Pasaporte Digital", "Partner/Tecnología": "Circularise / desarrollo propio", "Justificación": "Plataformas europeas DPP especializadas ESPR", "Coste estimado": "€5.000-20.000 dev", "Open-source": "No"},
    ])

    col_full, col_oss = st.columns([3, 1])
    with col_full:
        st.dataframe(tech_table, use_container_width=True, hide_index=True)
    with col_oss:
        oss_count = tech_table["Open-source"].value_counts()
        fig_oss = px.pie(
            values=oss_count.values, names=oss_count.index,
            color_discrete_map={"Sí": "#22c55e", "No": "#94a3b8"},
            title="Open-source vs Propietario",
            height=220,
        )
        fig_oss.update_layout(margin=dict(t=40, b=0))
        st.plotly_chart(fig_oss, use_container_width=True)

    st.markdown("---")
    st.markdown("**Costes estimados por instalación (CAPEX hardware IoT)**")
    cost_data = pd.DataFrame([
        {"Componente": "Sensores PV (irradiancia + T° célula + corriente)", "Pequeña (<100m²)": 200, "Mediana (100-500m²)": 450, "Grande (>500m²)": 900},
        {"Componente": "Sensores estructurales (2x strain + humedad)", "Pequeña (<100m²)": 180, "Mediana (100-500m²)": 350, "Grande (>500m²)": 700},
        {"Componente": "Sensores ambientales (anemóm. + T/HR + PM)", "Pequeña (<100m²)": 120, "Mediana (100-500m²)": 250, "Grande (>500m²)": 400},
        {"Componente": "Edge Gateway + conectividad", "Pequeña (<100m²)": 150, "Mediana (100-500m²)": 300, "Grande (>500m²)": 500},
        {"Componente": "Instalación y puesta en marcha", "Pequeña (<100m²)": 200, "Mediana (100-500m²)": 400, "Grande (>500m²)": 800},
        {"Componente": "NFC/QR Pasaporte Digital", "Pequeña (<100m²)": 30, "Mediana (100-500m²)": 50, "Grande (>500m²)": 100},
    ])
    cost_data["TOTAL"] = cost_data[["Pequeña (<100m²)", "Mediana (100-500m²)", "Grande (>500m²)"]].sum()
    st.dataframe(cost_data.style.format({"Pequeña (<100m²)": "€{:.0f}", "Mediana (100-500m²)": "€{:.0f}",
                                          "Grande (>500m²)": "€{:.0f}", "TOTAL": "€{:.0f}"}),
                 use_container_width=True, hide_index=True)

else:  # Roadmap
    st.subheader("🗺️ Roadmap de implementación 2025–2028")

    phases = [
        {
            "fase": "FASE 0 — Diseño e integración mínima", "periodo": "Q4 2025",
            "color": "#0d9488",
            "items": [
                "Definir arquitectura sensórica para piloto Guadalajara",
                "Seleccionar 5-8 sensores core (irradiancia, T° célula, corriente/tensión, 2× strain gauge)",
                "Instalar Edge Gateway básico (Advantech UNO)",
                "Desplegar Thingsboard en OVHcloud",
                "Dashboard básico operativo con KPIs en tiempo real",
            ],
            "coste": "€3.000–5.000 hardware + €2.000 desarrollo",
            "digital_twin": "—",
        },
        {
            "fase": "FASE 1 — MVP Digital Twin + Pasaporte Digital", "periodo": "Q1–Q2 2026",
            "color": "#0891b2",
            "items": [
                "Digital Twin Nivel 1 (Shadow) operativo",
                "Pasaporte Digital QR/NFC en todos los módulos instalados",
                "Reporting ESG automatizado v1.0 (CO₂, energía, PR)",
                "Integración con IDAE para reporting RENOCICLA",
                "API REST pública para integraciones ERP/ESG",
            ],
            "coste": "€25.000–40.000 (subvencionable RENOCICLA/CDTI)",
            "digital_twin": "Nivel 1 — Shadow",
        },
        {
            "fase": "FASE 2 — Predictivo y escalado comercial", "periodo": "Q3–Q4 2026",
            "color": "#1d4ed8",
            "items": [
                "Modelos ML para degradación y mantenimiento predictivo",
                "Benchmarking de flota (primeras 20+ instalaciones)",
                "Pipeline comercial diferenciado: 'Solar Flex + plataforma digital'",
                "Gestión básica de Comunidades Energéticas (CAE)",
                "Migración a Eclipse Ditto + InfluxDB (soberanía)",
            ],
            "coste": "€40.000–70.000 (I+D + desarrollo)",
            "digital_twin": "Nivel 2 — Predictivo",
        },
        {
            "fase": "FASE 3 — Ecosistema completo", "periodo": "2027",
            "color": "#7c3aed",
            "items": [
                "Pasaporte Digital conforme ESPR (cuando sea obligatorio)",
                "Blockchain certificados de origen digital",
                "Gestión CAE automatizada (coeficientes, facturación)",
                "Marketplace O&M terceros",
                "Modelos físicos (physics-based) de degradación adhesiva",
            ],
            "coste": "€60.000–100.000",
            "digital_twin": "Nivel 2–3",
        },
        {
            "fase": "FASE 4 — Roof-as-a-Service Digital", "periodo": "2028+",
            "color": "#1e3a8a",
            "items": [
                "Digital Twin Nivel 3 prescriptivo (optimización automática)",
                "VPP / participación mercados flexibilidad REE",
                "Monetización datos B2B (utilities, aseguradoras, ESG)",
                "Tokenización energética (MiCA + regulación EU)",
                "Expansión europeo: FR, DE, IT, PT",
            ],
            "coste": "€100.000+ (financiación EIC / inversión Serie A)",
            "digital_twin": "Nivel 3 — Prescriptivo",
        },
    ]

    for phase in phases:
        with st.expander(f"📅 {phase['fase']} — {phase['periodo']}", expanded=(phase['periodo'] in ['Q4 2025', 'Q1–Q2 2026'])):
            col_items, col_meta = st.columns([2, 1])
            with col_items:
                st.markdown("**Entregables clave:**")
                for item in phase["items"]:
                    st.markdown(f"- {item}")
            with col_meta:
                st.markdown(f"**Periodo:** {phase['periodo']}")
                st.markdown(f"**Inversión est.:** {phase['coste']}")
                st.markdown(f"**Digital Twin:** {phase['digital_twin']}")

    st.markdown("---")
    st.markdown("**Modelo de negocio digital (ingresos recurrentes)**")
    revenue_data = pd.DataFrame([
        {"Línea": "SaaS Monitorización", "€/mes/instalación": "50-150", "Para 200 inst. (2027)": "€120k-360k/año", "Margen bruto": "75%"},
        {"Línea": "O&M Predictivo", "€/mes/instalación": "40-165", "Para 200 inst. (2027)": "€100k-400k/año", "Margen bruto": "60%"},
        {"Línea": "Reporting ESG Premium", "€/mes/instalación": "17-40", "Para 200 inst. (2027)": "€40k-100k/año", "Margen bruto": "85%"},
        {"Línea": "Licencia API datos", "€/mes/cliente B2B": "830-4.200", "Para 5 licencias": "€50k-250k/año", "Margen bruto": "90%"},
        {"Línea": "Certificados Origen Digital", "% valor certificado": "5-10%", "Estimado 2028": "€30k-100k/año", "Margen bruto": "70%"},
    ])
    st.dataframe(revenue_data, use_container_width=True, hide_index=True)

    # Proyección ingresos
    years = [2025, 2026, 2027, 2028, 2029]
    revenue_proj = [5, 80, 350, 900, 2000]
    fig_rev = go.Figure()
    fig_rev.add_trace(go.Bar(
        x=years, y=revenue_proj, marker_color=["#94a3b8", "#0891b2", "#1d4ed8", "#7c3aed", "#1e3a8a"],
        text=[f"€{v}k" for v in revenue_proj], textposition="outside",
        name="Ingresos recurrentes (€k/año)",
    ))
    fig_rev.update_layout(
        height=260, plot_bgcolor="white", yaxis_title="€k/año",
        title="Proyección ingresos recurrentes SaaS+O&M+ESG",
    )
    fig_rev.update_yaxes(gridcolor="#f1f5f9")
    st.plotly_chart(fig_rev, use_container_width=True)

st.markdown("---")
st.caption("Solar Flex IoT Platform · Arquitectura de referencia v1.0 · Marzo 2026 · Confidencial — Uso interno y presentaciones inversores")
