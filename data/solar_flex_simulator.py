"""
Solar Flex IoT Data Simulator
Genera datos realistas de instalaciones BIPV para el dashboard
"""
import numpy as np
import pandas as pd
from datetime import datetime, timedelta


INSTALLATIONS = [
    {"id": "SF-2025-IND-00247", "name": "Nave Logística Guadalajara", "location": "Guadalajara",
     "lat": 40.633, "lon": -3.167, "kwp": 120.0, "area_m2": 800, "type": "Industrial",
     "install_date": "2025-03-15", "adhesive": "AlfaDOC-A3", "panel_model": "FlexCIGS-200W",
     "status": "operativo", "pr_target": 0.82},
    {"id": "SF-2025-COM-00312", "name": "Centro Comercial Alcalá", "location": "Alcalá de Henares",
     "lat": 40.482, "lon": -3.363, "kwp": 85.0, "area_m2": 560, "type": "Comercial",
     "install_date": "2025-05-20", "adhesive": "AlfaDOC-A3", "panel_model": "FlexCIGS-200W",
     "status": "operativo", "pr_target": 0.80},
    {"id": "SF-2025-IND-00389", "name": "Plataforma Logística Zaragoza", "location": "Zaragoza",
     "lat": 41.656, "lon": -0.876, "kwp": 200.0, "area_m2": 1350, "type": "Industrial",
     "install_date": "2025-07-10", "adhesive": "AlfaDOC-A3", "panel_model": "FlexOPV-150W",
     "status": "alerta", "pr_target": 0.81},
    {"id": "SF-2025-TER-00421", "name": "Hotel Costa Brava", "location": "Girona",
     "lat": 41.983, "lon": 2.825, "kwp": 45.0, "area_m2": 300, "type": "Terciario",
     "install_date": "2025-09-01", "adhesive": "AlfaDOC-B1", "panel_model": "FlexASi-100W",
     "status": "operativo", "pr_target": 0.79},
    {"id": "SF-2024-IND-00156", "name": "Fábrica Sector Automoción Valencia", "location": "Valencia",
     "lat": 39.470, "lon": -0.376, "kwp": 350.0, "area_m2": 2300, "type": "Industrial",
     "install_date": "2024-11-08", "adhesive": "AlfaDOC-A3", "panel_model": "FlexCIGS-200W",
     "status": "mantenimiento", "pr_target": 0.83},
]

SENSOR_SPECS = {
    "irradiancia_ghi": {"unit": "W/m²", "min": 0, "max": 1200, "domain": "Energético-PV"},
    "irradiancia_poa": {"unit": "W/m²", "min": 0, "max": 1350, "domain": "Energético-PV"},
    "temperatura_celula": {"unit": "°C", "min": -5, "max": 85, "domain": "Energético-PV"},
    "temperatura_ambiente": {"unit": "°C", "min": -5, "max": 42, "domain": "Ambiental"},
    "corriente_string_1": {"unit": "A", "min": 0, "max": 12, "domain": "Energético-PV"},
    "corriente_string_2": {"unit": "A", "min": 0, "max": 12, "domain": "Energético-PV"},
    "tension_dc": {"unit": "V", "min": 0, "max": 800, "domain": "Energético-PV"},
    "potencia_ac": {"unit": "kW", "min": 0, "max": 500, "domain": "Energético-PV"},
    "strain_adhesivo_1": {"unit": "με", "min": 0, "max": 500, "domain": "Estructural"},
    "strain_adhesivo_2": {"unit": "με", "min": 0, "max": 500, "domain": "Estructural"},
    "vibracion_rms": {"unit": "mm/s", "min": 0, "max": 15, "domain": "Estructural"},
    "humedad_interfase": {"unit": "%HR", "min": 20, "max": 95, "domain": "Estructural"},
    "velocidad_viento": {"unit": "m/s", "min": 0, "max": 25, "domain": "Ambiental"},
    "precipitacion": {"unit": "mm/h", "min": 0, "max": 30, "domain": "Ambiental"},
    "pm25": {"unit": "µg/m³", "min": 2, "max": 80, "domain": "Ambiental"},
    "consumo_edificio": {"unit": "kW", "min": 10, "max": 800, "domain": "Building"},
    "soc_bateria": {"unit": "%", "min": 5, "max": 100, "domain": "Building"},
}


def _solar_profile(hour: float) -> float:
    """Curva de irradiancia normalizada para hora del día."""
    if hour < 6 or hour > 20:
        return 0.0
    peak = 13.0
    sigma = 3.5
    return np.exp(-((hour - peak) ** 2) / (2 * sigma ** 2))


def generate_timeseries(installation_id: str, days: int = 30, freq_minutes: int = 15) -> pd.DataFrame:
    """Genera series temporales de sensores para una instalación."""
    np.random.seed(hash(installation_id) % 2**31)
    inst = next((i for i in INSTALLATIONS if i["id"] == installation_id), INSTALLATIONS[0])

    end = datetime.now().replace(minute=0, second=0, microsecond=0)
    start = end - timedelta(days=days)
    idx = pd.date_range(start, end, freq=f"{freq_minutes}min")
    n = len(idx)

    hours = np.array([t.hour + t.minute / 60 for t in idx])
    days_arr = np.array([(t - start).days for t in idx])
    solar = np.array([_solar_profile(h) for h in hours])

    # Degradación acumulada (0.5% año)
    degradation = 1 - 0.005 * (days_arr / 365)

    # Irradiancia con nubes y ruido
    cloud_factor = np.clip(1 - 0.3 * np.abs(np.sin(days_arr * 0.7 + np.random.rand(n) * 0.5)), 0.3, 1.0)
    irr_noise = 1 + np.random.normal(0, 0.05, n)
    irradiancia_ghi = np.clip(solar * 950 * cloud_factor * irr_noise, 0, 1100)
    irradiancia_poa = irradiancia_ghi * 1.07  # tilt correction

    # Temperatura célula: T_amb + k * POA
    t_amb = 18 + 12 * np.sin(2 * np.pi * (days_arr / 365 - 0.3)) + np.random.normal(0, 2, n)
    t_celula = np.clip(t_amb + 0.028 * irradiancia_poa, -5, 85)

    # Potencia AC
    kwp = inst["kwp"]
    pr_base = inst["pr_target"]
    # Degradación leve en instalación "alerta"
    if inst["status"] == "alerta":
        pr_base *= 0.91
    potencia_ac = np.clip(irradiancia_poa / 1000 * kwp * pr_base * degradation + np.random.normal(0, 0.5, n), 0, kwp)

    # Corrientes strings
    i_string_1 = np.clip(potencia_ac * 0.52 / (400 * 1e-3 + 1e-9), 0, 12)
    i_string_2 = np.clip(potencia_ac * 0.48 / (400 * 1e-3 + 1e-9), 0, 12)

    # Strain adhesivo (µε) — aumenta con temperatura y viento
    viento = np.abs(np.random.normal(3, 2, n)) + 2 * np.abs(np.sin(days_arr * 0.3))
    strain_1 = np.clip(80 + 0.8 * t_celula + 5 * viento + np.random.normal(0, 10, n), 0, 500)
    strain_2 = np.clip(75 + 0.75 * t_celula + 4.5 * viento + np.random.normal(0, 10, n), 0, 500)
    if inst["status"] == "alerta":
        # String localizada con strain elevado (posible despegue incipiente)
        anomaly_mask = (days_arr > days - 5)
        strain_1[anomaly_mask] *= 1.6

    # Humedad interfase
    precip = np.clip(np.random.exponential(0.5, n) * (np.random.rand(n) > 0.85), 0, 30)
    humedad = np.clip(40 + 20 * (precip > 0).astype(float) + np.random.normal(0, 5, n), 20, 95)

    # Consumo edificio (perfil laboral)
    consumo_base = kwp * 1.8
    is_working = ((hours >= 7) & (hours <= 19)).astype(float)
    consumo_edificio = np.clip(consumo_base * (0.4 + 0.6 * is_working) + np.random.normal(0, consumo_base * 0.05, n), 10, kwp * 5)

    # SoC batería (ciclo diario)
    soc = 50 + 40 * np.sin(2 * np.pi * (hours / 24) - np.pi / 2) + np.random.normal(0, 3, n)
    soc = np.clip(soc, 5, 100)

    df = pd.DataFrame({
        "timestamp": idx,
        "irradiancia_ghi": np.round(irradiancia_ghi, 1),
        "irradiancia_poa": np.round(irradiancia_poa, 1),
        "temperatura_celula": np.round(t_celula, 1),
        "temperatura_ambiente": np.round(t_amb, 1),
        "potencia_ac_kw": np.round(potencia_ac, 2),
        "corriente_string_1": np.round(i_string_1, 3),
        "corriente_string_2": np.round(i_string_2, 3),
        "tension_dc_v": np.round(np.clip(400 + np.random.normal(0, 5, n), 380, 420), 1),
        "strain_adhesivo_1_ue": np.round(strain_1, 1),
        "strain_adhesivo_2_ue": np.round(strain_2, 1),
        "velocidad_viento_ms": np.round(viento, 2),
        "precipitacion_mm_h": np.round(precip, 2),
        "humedad_interfase_pct": np.round(humedad, 1),
        "pm25_ug_m3": np.round(np.clip(np.random.exponential(15, n), 2, 80), 1),
        "consumo_edificio_kw": np.round(consumo_edificio, 2),
        "soc_bateria_pct": np.round(soc, 1),
    })
    return df


def compute_kpis(df: pd.DataFrame, kwp: float) -> dict:
    """Calcula KPIs energéticos principales."""
    hours_step = 0.25  # 15 min = 0.25 h
    energia_kwh = (df["potencia_ac_kw"] * hours_step).sum()
    irr_kwh_m2 = (df["irradiancia_poa"] * hours_step / 1000).sum()
    pr = energia_kwh / (irr_kwh_m2 * kwp) if irr_kwh_m2 > 0 else 0
    especifico = energia_kwh / kwp if kwp > 0 else 0
    t_cel_max = df["temperatura_celula"].max()
    strain_max = max(df["strain_adhesivo_1_ue"].max(), df["strain_adhesivo_2_ue"].max())
    co2_evitado = energia_kwh * 0.181 / 1000  # tCO2 (factor REE ~0.181 kg/kWh 2025)
    return {
        "energia_kwh": round(energia_kwh, 1),
        "pr": round(pr, 3),
        "especifico_kwh_kwp": round(especifico, 1),
        "t_celula_max": round(t_cel_max, 1),
        "strain_max_ue": round(strain_max, 1),
        "co2_evitado_t": round(co2_evitado, 3),
        "horas_datos": len(df) * 0.25,
    }


def get_fleet_summary() -> pd.DataFrame:
    """Resumen de flota para el dashboard principal."""
    rows = []
    for inst in INSTALLATIONS:
        df = generate_timeseries(inst["id"], days=30)
        kpis = compute_kpis(df, inst["kwp"])
        degradation_pct = round((1 - kpis["pr"] / inst["pr_target"]) * 100, 1)
        rows.append({
            "ID": inst["id"],
            "Instalación": inst["name"],
            "Ubicación": inst["location"],
            "Tipo": inst["type"],
            "kWp": inst["kwp"],
            "Estado": inst["status"],
            "PR_real": kpis["pr"],
            "PR_objetivo": inst["pr_target"],
            "Energía_kWh_30d": kpis["energia_kwh"],
            "Especifico_kWh_kWp": kpis["especifico_kwh_kwp"],
            "Strain_max_ue": kpis["strain_max_ue"],
            "CO2_evitado_t_30d": kpis["co2_evitado_t"],
            "lat": inst["lat"],
            "lon": inst["lon"],
            "degradacion_pct": degradation_pct,
        })
    return pd.DataFrame(rows)


def get_maintenance_alerts() -> list:
    """Genera alertas de mantenimiento predictivo."""
    alerts = []
    for inst in INSTALLATIONS:
        df = generate_timeseries(inst["id"], days=7)
        strain_max = df["strain_adhesivo_1_ue"].max()
        t_max = df["temperatura_celula"].max()
        kpis = compute_kpis(df, inst["kwp"])

        if strain_max > 350:
            alerts.append({
                "id": inst["id"], "instalacion": inst["name"],
                "tipo": "ESTRUCTURAL", "severidad": "ALTA",
                "mensaje": f"Strain adhesivo zona 1 alcanzó {strain_max:.0f} µε (umbral: 350 µε). Inspección inmediata recomendada.",
                "sensor": "strain_adhesivo_1", "valor": strain_max, "umbral": 350,
            })
        if t_max > 75:
            alerts.append({
                "id": inst["id"], "instalacion": inst["name"],
                "tipo": "TÉRMICA", "severidad": "MEDIA",
                "mensaje": f"Temperatura de célula máxima {t_max:.1f}°C. Verificar ventilación y adhesivo.",
                "sensor": "temperatura_celula", "valor": t_max, "umbral": 75,
            })
        if kpis["pr"] < inst["pr_target"] * 0.93:
            alerts.append({
                "id": inst["id"], "instalacion": inst["name"],
                "tipo": "RENDIMIENTO", "severidad": "MEDIA",
                "mensaje": f"PR real {kpis['pr']:.3f} por debajo del 93% del objetivo ({inst['pr_target']:.2f}). Revisar módulos y conexiones.",
                "sensor": "pr", "valor": kpis["pr"], "umbral": inst["pr_target"] * 0.93,
            })

    return alerts


DPP_MATERIALS = {
    "FlexCIGS-200W": {
        "capas": [
            {"capa": "Encapsulante frontal", "material": "EVA (Etileno Vinil Acetato)", "espesor_mm": 0.5, "reciclable": True, "pct_reciclado_contenido": 0},
            {"capa": "Célula PV", "material": "CIGS (Cu-In-Ga-Se)", "espesor_mm": 0.003, "reciclable": True, "pct_reciclado_contenido": 15},
            {"capa": "Sustrato flexible", "material": "Poliimida (Kapton)", "espesor_mm": 0.1, "reciclable": False, "pct_reciclado_contenido": 0},
            {"capa": "Encapsulante posterior", "material": "EVA", "espesor_mm": 0.5, "reciclable": True, "pct_reciclado_contenido": 0},
            {"capa": "Tedlar posterior", "material": "PVF (Fluoruro de Polivinilo)", "espesor_mm": 0.15, "reciclable": False, "pct_reciclado_contenido": 0},
            {"capa": "Adhesivo estructural", "material": "Poliuretano bicomponente AlfaDOC-A3", "espesor_mm": 3.0, "reciclable": False, "pct_reciclado_contenido": 0},
            {"capa": "Panel sándwich soporte", "material": "Acero galvanizado + núcleo PU", "espesor_mm": 60.0, "reciclable": True, "pct_reciclado_contenido": 35},
        ],
        "peso_kg_m2": 5.8,
        "certificaciones": ["IEC 61215", "IEC 61730", "BROOF(t1)", "TÜV Rheinland", "CE"],
        "fin_de_vida": {
            "instrucciones": "Retirar con espátula térmica (T>60°C). Separar módulo PV del soporte. Entregar a gestor RAEE autorizado.",
            "gestores_autorizados": ["Recytel", "Ambilamp", "Recupel"],
            "contenido_critico": ["Indio (In)", "Selenio (Se)", "Cadmio (Cd) <trazas>"],
        }
    }
}
