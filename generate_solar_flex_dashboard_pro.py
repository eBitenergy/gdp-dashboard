"""Generate a Solarflex and Livoltek sales management workbook.

This script builds a multi-sheet Excel file that simulates a sales pipeline
control panel with dashboards, KPIs, and supporting data sheets. The generated
file includes:
- A dashboard sheet with CTA-style buttons and KPI summaries.
- A master pipeline with sample opportunities and calculated financials.
- BESS requirements derived from BESS-related opportunities.
- Solarflex roof (asbestos) projects with basic cost estimates.
- Livoltek equipment pricing with discounts and net pricing.

Run the script directly to produce ``Pipeline_Maestro_Solarflex_Livoltek.xlsx``.
"""

from __future__ import annotations

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
HEADER_FILL = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
SUB_HEADER_FILL = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
CENTER_ALIGN = Alignment(horizontal="center", vertical="center")
LEFT_ALIGN = Alignment(horizontal="left", vertical="center")
CURRENCY_FORMAT = "#,##0.00 €"
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)


def create_workbook(filename: str = "Pipeline_Maestro_Solarflex_Livoltek.xlsx") -> str:
    wb = Workbook()
    _build_dashboard(wb)
    pipeline_data = _build_pipeline_master(wb)
    _build_bess_requirements(wb, pipeline_data)
    _build_solarflex_roof_projects(wb, pipeline_data)
    _build_livoltek_equipment(wb)

    _autosize_columns(wb)
    wb.save(filename)
    return filename


def _build_dashboard(workbook: Workbook) -> None:
    ws_dash = workbook.active
    ws_dash.title = "CONTROL_DASHBOARD"

    ws_dash.merge_cells("B2:H2")
    ws_dash["B2"] = "SISTEMA DE GESTIÓN DE VENTAS - SOLARFLEX & LIVOLTEK"
    ws_dash["B2"].font = Font(size=18, bold=True, color="1F497D")
    ws_dash["B2"].alignment = CENTER_ALIGN

    ws_dash["C5"] = "ACCIONES RÁPIDAS (VBA)"
    ws_dash["C5"].font = Font(bold=True)

    buttons = [
        (6, "ACTUALIZAR PIPELINE", "Recalcula márgenes, semáforos y probabilidades"),
        (8, "GENERAR FORECAST", "Crea reporte trimestral en hoja nueva"),
        (10, "EXPORTAR BESS", "Genera archivo para Ingeniería"),
        (12, "ACTUALIZAR GRÁFICOS", "Refresca datos del Dashboard visual"),
    ]

    for row, text, desc in buttons:
        button_cell = ws_dash.cell(row=row, column=3)
        button_cell.value = text
        button_cell.font = Font(bold=True, color="FFFFFF")
        button_cell.fill = HEADER_FILL
        button_cell.alignment = CENTER_ALIGN
        button_cell.border = THIN_BORDER

        desc_cell = ws_dash.cell(row=row, column=4)
        desc_cell.value = desc
        desc_cell.font = Font(italic=True, color="555555")
        desc_cell.alignment = Alignment(vertical="center")

    ws_dash["G5"] = "RESUMEN EJECUTIVO"
    ws_dash["G5"].font = Font(bold=True)

    kpis = [
        (6, "Total Pipeline (€)", "=SUM(Pipeline_Master!I:I)"),
        (7, "Total Potencia (kWp)", "=SUM(Pipeline_Master!C:C)"),
        (8, "Oportunidades BESS", "=COUNTA(BESS_Requirements!A:A)-1"),
        (9, "Proyectos Amianto", "=COUNTA(Solarflex_Roof_Projects!A:A)-1"),
        (
            10,
            "Ventas Q1 2024 Est.",
            '=SUMIFS(Pipeline_Master!I:I, Pipeline_Master!F:F, ">=2024-01-01", Pipeline_Master!F:F, "<=2024-03-31")',
        ),
    ]

    for row, label, formula in kpis:
        ws_dash.cell(row=row, column=7).value = label
        ws_dash.cell(row=row, column=7).font = Font(bold=True)
        value_cell = ws_dash.cell(row=row, column=8)
        value_cell.value = formula
        value_cell.number_format = CURRENCY_FORMAT if "Total" in label or "Ventas" in label else "#,##0"

    ws_dash.column_dimensions["B"].width = 5
    ws_dash.column_dimensions["C"].width = 25
    ws_dash.column_dimensions["D"].width = 40
    ws_dash.column_dimensions["F"].width = 5
    ws_dash.column_dimensions["G"].width = 25
    ws_dash.column_dimensions["H"].width = 20


def _build_pipeline_master(workbook: Workbook) -> list[tuple]:
    ws_pipeline = workbook.create_sheet("Pipeline_Master")

    headers = [
        "Proyecto",
        "Segmento",
        "kWp",
        "Tecnología",
        "Tipo Proyecto",
        "Fecha Decisión",
        "Estado",
        "Probabilidad (%)",
        "CapEx (€)",
        "GM (%)",
        "Ingresos Esperados (€)",
        "Coste Equipos (€)",
        "Margen (€)",
        "Proveedor",
        "Notas",
    ]
    ws_pipeline.append(headers)

    for col_num, header in enumerate(headers, 1):
        cell = ws_pipeline.cell(row=1, column=col_num)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER_ALIGN

    data = _pipeline_seed_data()

    for row_data in data:
        proj, segment, kwp, tech, project_type, decision_date, status, probability, capex, provider = row_data
        gm_pct = 0.30
        ingresos_esp = capex * (probability / 100)
        coste_eq = capex * 0.50
        margen_eur = capex * gm_pct

        ws_pipeline.append(
            [
                proj,
                segment,
                kwp,
                tech,
                project_type,
                decision_date,
                status,
                probability,
                capex,
                gm_pct,
                ingresos_esp,
                coste_eq,
                margen_eur,
                provider,
                "",
            ]
        )

    for col in ["I", "K", "L", "M"]:
        for cell in ws_pipeline[col]:
            if cell.row > 1:
                cell.number_format = CURRENCY_FORMAT

    return data


def _pipeline_seed_data() -> list[tuple]:
    return [
        ("Injection", "Industrial", 400, "PV+BESS", "PPA", "2024-01-31", "Firmado", 100, 226000, "Livoltek+Solarflex"),
        ("Ilermilk Clota", "Industrial", 1300, "PV", "Venta directa", "2024-02-15", "Firmado", 100, 708500, "Solarflex"),
        ("Ilermilk MT", "Industrial", 125.4, "PV", "Venta directa", "2024-02-20", "Firmado", 100, 125400, "Solarflex"),
        ("Dx Aragón", "Industrial", 446, "PV", "Venta directa", "2024-03-01", "Facturado", 100, 236380, "Solarflex"),
        ("Gravisol Lérida", "Industrial", 999, "PV", "Venta directa", "2024-03-10", "Firmado", 100, 649350, "Solarflex"),
        ("Igeasa Odena", "Industrial", 300, "PV", "Venta directa", "2024-06-15", "En negociación", 70, 195000, "Solarflex"),
        ("Camping Garrofer", "Hostelería", 100, "PV", "Venta directa", "2024-06-30", "En negociación", 60, 65000, "Solarflex"),
        ("Industrias Lorenzo", "Industrial", 90, "PV", "PPA", "2024-04-15", "PPA en negociación", 70, 78300, "Solarflex"),
        ("Abzak 1", "Industrial", 500, "PV", "Oferta", "2024-04-30", "Rechazada", 10, 336000, "Solarflex"),
        ("Abzak 2 - Piera", "Industrial", 250, "PV", "Oferta", "2024-04-30", "Rechazada", 10, 168000, "Solarflex"),
        ("Abzak 3 - Breda", "Industrial", 200, "PV", "Oferta", "2024-04-30", "Rechazada", 10, 134400, "Solarflex"),
        ("Nad Sabadell", "Industrial", 80, "PV", "Oferta", "2024-04-30", "Rechazada", 10, 52000, "Solarflex"),
        ("Metaru", "Industrial", 250, "PV", "Oferta", "2024-05-31", "Seguimiento", 40, 175000, "Solarflex"),
        ("Bau Centre", "Terciario", 99, "PV", "Venta directa", "2024-06-17", "En negociación", 70, 69300, "Solarflex"),
        ("Tunkers", "Industrial", 99, "PV", "Venta directa", "2024-06-18", "En negociación", 70, 69300, "Solarflex"),
        ("Araknow", "Industrial", 50, "PV", "Venta directa", "2024-06-19", "En negociación", 70, 28250, "Solarflex"),
        ("Ftv Sta Margarida", "Industrial", 669, "PV", "Venta directa", "2024-06-20", "En negociación", 70, 434850, "Solarflex"),
        ("Forrajes Porvenir", "Industrial", 649, "PV", "PPA", "2024-07-15", "PPA en negociación", 70, 366685, "Solarflex"),
        ("Zincados Canovelles", "Industrial", 160, "Solarflex Amianto", "Amianto", "2024-07-31", "Estudio técnico", 60, 166400, "Solarflex"),
        ("Ced Menarguens", "Comunidad", 450, "PV", "Comunidad", "2024-08-15", "En negociación", 60, 270000, "Solarflex"),
        ("Ced Tornabous", "Comunidad", 80, "PV", "Comunidad", "2024-08-20", "En negociación", 60, 52000, "Solarflex"),
        ("Mas Alborna", "Industrial", 52, "Solarflex Amianto", "Amianto", "2024-08-31", "En negociación", 60, 81400, "Solarflex"),
        ("Valira", "Industrial", 150, "Solarflex Amianto", "Amianto", "2024-09-05", "En negociación", 60, 167400, "Solarflex"),
        ("La Coma", "Industrial", 109, "Solarflex Amianto", "Amianto", "2024-09-15", "En negociación", 60, 107530, "Solarflex"),
        ("TEGONSA", "Area 8", 600, "PV+BESS", "Venta directa", "2024-09-15", "En negociación", 50, 339000, "Livoltek+Solarflex"),
        ("TOTAL", "Area 8", 125, "PV", "Venta directa", "2024-09-16", "En negociación", 40, 70625, "Solarflex"),
        ("CARMEN GARCIA", "Area 8", 250, "PV", "Venta directa", "2024-09-17", "En negociación", 40, 141250, "Solarflex"),
        ("GEMMA GARCIA", "Area 8", 200, "PV", "Venta directa", "2024-09-18", "En negociación", 40, 113000, "Solarflex"),
        ("DOMUS", "Area 8", 500, "PV", "Venta directa", "2024-09-19", "En negociación", 40, 282500, "Solarflex"),
        ("EARPRO", "Area 8", 200, "PV", "Venta directa", "2024-09-20", "En negociación", 40, 113000, "Solarflex"),
        ("TACSA", "Area 8", 500, "PV", "Venta directa", "2024-09-21", "Descartado", 10, 282500, "Solarflex"),
        ("GENEBRE", "Area 8", 300, "PV", "Venta directa", "2024-09-22", "En negociación", 40, 169500, "Solarflex"),
        ("VIVES CORTADA", "Area 8", 200, "PV", "Venta directa", "2024-09-23", "En negociación", 40, 113000, "Solarflex"),
        ("LLORENS", "Area 8", 225, "PV", "Venta directa", "2024-09-24", "En negociación", 40, 127125, "Solarflex"),
        ("SHINE", "Area 8", 150, "PV", "Venta directa", "2024-09-25", "En negociación", 40, 84750, "Solarflex"),
        ("FAE", "Area 8", 700, "PV+BESS", "Venta directa", "2024-09-26", "En negociación", 50, 395500, "Livoltek"),
        ("ERLISO", "Area 8", 150, "PV", "Venta directa", "2024-09-27", "En negociación", 40, 84750, "Solarflex"),
        ("MULTICINES", "Area 8", 300, "PV", "Venta directa", "2024-09-28", "En negociación", 40, 169500, "Solarflex"),
        ("AMAC", "Area 8", 400, "PV", "Venta directa", "2024-09-29", "En negociación", 40, 226000, "Solarflex"),
        ("PORCHE", "Area 8", 400, "PV", "Venta directa", "2024-09-30", "En negociación", 40, 226000, "Solarflex"),
        ("ALTEX", "Area 8", 200, "PV", "Venta directa", "2024-10-01", "En negociación", 40, 113000, "Solarflex"),
        ("SIVILA", "Area 8", 300, "PV", "Venta directa", "2024-10-02", "En negociación", 40, 169500, "Solarflex"),
        ("RENTAURO", "Area 8", 100, "PV", "Venta directa", "2024-10-03", "En negociación", 40, 56500, "Solarflex"),
        ("EQUIVALENZA", "Area 8", 300, "PV", "Venta directa", "2024-10-04", "En negociación", 40, 169500, "Solarflex"),
        ("PLAT. CONS", "Area 8", 900, "PV", "Venta directa", "2024-10-05", "En negociación", 40, 508500, "Solarflex"),
        ("QUALITY", "Area 8", 1000, "PV", "Venta directa", "2024-10-06", "En negociación", 40, 565000, "Solarflex"),
        ("SALVAT", "Area 8", 1000, "PV", "Venta directa", "2024-10-07", "En negociación", 40, 565000, "Solarflex"),
        ("HIPUR SL", "Agrupación", 140, "PV", "Previsión", "2024-03-31", "Pipeline", 30, 79100, "Solarflex"),
        ("Lubricantes SA", "Agrupación", 150, "PV", "Previsión", "2024-04-15", "Pipeline", 30, 84750, "Solarflex"),
        ("Finish Metal", "Agrupación", 28, "PV", "Previsión", "2024-04-30", "Pipeline", 30, 18200, "Solarflex"),
        ("Thornytex", "Agrupación", 22.4, "PV", "Previsión", "2024-05-15", "Pipeline", 30, 15680, "Solarflex"),
        ("CARBONNEL", "Agrupación", 70, "PV", "Previsión", "2024-05-31", "Pipeline", 30, 30800, "Solarflex"),
        ("Benzinera Castell.", "Agrupación", 32, "PV", "Previsión", "2024-06-15", "Pipeline", 30, 22400, "Solarflex"),
        ("R-CARNES OLESA", "Agrupación", 44, "PV", "Previsión", "2024-06-30", "Pipeline", 30, 30800, "Solarflex"),
        ("Normevi SL", "Agrupación", 90, "PV", "Previsión", "2024-07-15", "Pipeline", 30, 58500, "Solarflex"),
        ("MOTO SAE 2019", "Agrupación", 30, "PV", "Previsión", "2024-07-31", "Pipeline", 30, 21000, "Solarflex"),
        ("GLOBAL SMM", "Agrupación", 227, "PV", "Previsión", "2024-08-31", "Pipeline", 30, 147550, "Solarflex"),
        ("INGALCINC", "Agrupación", 162, "PV", "Previsión", "2024-09-30", "Pipeline", 30, 105300, "Solarflex"),
    ]


def _build_bess_requirements(workbook: Workbook, data: list[tuple]) -> None:
    ws_bess = workbook.create_sheet("BESS_Requirements")
    headers_bess = [
        "Proyecto",
        "Potencia PV (kWp)",
        "BESS kWh Nec. (2x)",
        "Autonomía (h)",
        "Marca",
        "Coste BESS Est. (€)",
        "Notas",
    ]
    ws_bess.append(headers_bess)

    for col_num, header in enumerate(headers_bess, 1):
        cell = ws_bess.cell(row=1, column=col_num)
        cell.font = HEADER_FONT
        cell.fill = SUB_HEADER_FILL
        cell.alignment = CENTER_ALIGN

    for row_data in data:
        proj, _, kwp, tech, *_rest = row_data
        if "BESS" in tech or proj in {"Injection", "TEGONSA", "FAE", "Forrajes Porvenir"}:
            bess_kwh = float(kwp) * 2
            coste_bess = bess_kwh * 400
            ws_bess.append([proj, kwp, bess_kwh, 4, "Livoltek", coste_bess, "Candidato Industrial"])

    for cell in ws_bess["F"]:
        if cell.row > 1:
            cell.number_format = CURRENCY_FORMAT


def _build_solarflex_roof_projects(workbook: Workbook, data: list[tuple]) -> None:
    ws_roof = workbook.create_sheet("Solarflex_Roof_Projects")
    headers_sf = [
        "Proyecto",
        "Amianto",
        "m2 Techo Est.",
        "Coste Techo (€)",
        "Coste Módulos Flex (€)",
        "Total (€)",
        "Payback (Años)",
    ]
    ws_roof.append(headers_sf)

    for col_num, header in enumerate(headers_sf, 1):
        cell = ws_roof.cell(row=1, column=col_num)
        cell.font = HEADER_FONT
        cell.fill = SUB_HEADER_FILL
        cell.alignment = CENTER_ALIGN

    for row_data in data:
        proj, _, kwp, tech, *_rest, capex, provider = row_data
        if "Amianto" in tech or "Solarflex" in provider:
            if "Amianto" in tech:
                m2 = float(kwp) * 5
                coste_techo = m2 * 70
                coste_mod = float(kwp) * 800
                total = coste_techo + coste_mod
                payback = round(total / (capex or 1), 1)
                ws_roof.append([proj, "SÍ", m2, coste_techo, coste_mod, total, payback])

    for col in ["D", "E", "F"]:
        for cell in ws_roof[col]:
            if cell.row > 1:
                cell.number_format = CURRENCY_FORMAT


def _build_livoltek_equipment(workbook: Workbook) -> None:
    ws_equipment = workbook.create_sheet("Livoltek_Equipos")
    headers_lv = ["Equipo", "Modelo", "PVP Tarifa (€)", "Dto (%)", "Neto (€)"]
    ws_equipment.append(headers_lv)

    for col_num, header in enumerate(headers_lv, 1):
        cell = ws_equipment.cell(row=1, column=col_num)
        cell.font = HEADER_FONT
        cell.fill = SUB_HEADER_FILL
        cell.alignment = CENTER_ALIGN

    equipos = [
        ("Inversor Híbrido 10kW", "LVT-HYB-10K", 2800, 0.45, 1540),
        ("Inversor Híbrido 20kW", "LVT-HYB-20K", 4200, 0.45, 2310),
        ("Batería HV 100kWh", "LVT-RACK-100", 38000, 0.40, 22800),
        ("Cargador DC 60kW", "LVT-DC-60", 18500, 0.40, 11100),
    ]

    for eq in equipos:
        ws_equipment.append(eq)

    for row in ws_equipment.iter_rows(min_row=2):
        row[2].number_format = CURRENCY_FORMAT
        row[3].number_format = "0%"
        row[4].number_format = CURRENCY_FORMAT


def _autosize_columns(workbook: Workbook) -> None:
    for sheet in workbook.worksheets:
        for column in sheet.columns:
            col_letter = get_column_letter(column[0].column)
            sheet.column_dimensions[col_letter].width = 18


if __name__ == "__main__":
    output_file = create_workbook()
    print(f"Archivo generado: {output_file}")
