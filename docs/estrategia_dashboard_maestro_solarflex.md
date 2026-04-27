# Estrategia del Dashboard Maestro de Inversión — Solarflexible

## 1) Objetivo de negocio
Diseñar un **dashboard maestro para inversores** que permita:
- Visualizar el estado real del proyecto (avance, riesgos y uso de capital).
- Proyectar inversión, retorno y escenarios.
- Trazar cada dato hasta su fuente documental (contratos, CAPEX, PPA, permisos).
- Estructurar el trabajo en **tickets ejecutables** para el equipo.

---

## 2) Usuarios y decisiones clave
### Perfil principal: Inversores / Comité de inversión
Decisiones que deben poder tomar con el dashboard:
1. ¿Cuánto capital adicional se necesita y en qué fecha?
2. ¿Cuál es el retorno esperado por escenario (base, optimista, estrés)?
3. ¿Qué riesgos impactan más en VAN/TIR/payback?
4. ¿Qué evidencias documentales respaldan cada hipótesis?

### Perfil secundario: PMO / Finanzas / Operaciones
1. Priorizar tickets y bloqueos.
2. Validar calidad de datos.
3. Preparar reporting mensual para due diligence.

---

## 3) Estructura funcional del dashboard maestro
## Módulo A — Resumen ejecutivo (1 pantalla)
- Inversión comprometida vs. ejecutada.
- VAN, TIR, payback, DSCR (si aplica) y sensibilidad principal.
- Semáforo de riesgos (regulatorio, construcción, precio energía, O&M).
- Indicador de confianza del dato (completitud y fecha de actualización).

## Módulo B — Pipeline y cartera de proyectos
- Estado por fase: prefactibilidad, permisos, EPC, operación.
- CAPEX estimado vs. CAPEX validado.
- Priorización por rentabilidad ajustada al riesgo.

## Módulo C — Modelo financiero
- Flujo de caja mensual/anual.
- Supuestos configurables: inflación, degradación, precio energía, curtailment, tipo de cambio, coste de deuda.
- Escenarios y simulaciones (base / optimista / estrés).

## Módulo D — Retorno e indicadores para inversor
- VAN/TIR por proyecto y consolidado.
- Payback simple y descontado.
- Múltiplos de inversión y curva J (si aplica al vehículo).

## Módulo E — Riesgo y cumplimiento
- Registro de riesgos con probabilidad/impacto.
- Matriz de mitigación y responsable.
- Estado de cumplimiento documental (auditoría).

## Módulo F — Data room y anexos
- Repositorio vinculado por proyecto:
  - contratos EPC/O&M,
  - PPA/offtake,
  - permisos,
  - estudios técnicos,
  - pólizas,
  - actas de comité.
- Versionado y trazabilidad (quién subió qué y cuándo).

---

## 4) Modelo de datos mínimo viable
### Entidades principales
- `proyecto`
- `hito`
- `capex`
- `opex`
- `produccion_energia`
- `ingresos`
- `financiacion`
- `riesgo`
- `documento`
- `supuesto_modelo`

### Reglas de calidad de dato
- Cada métrica crítica debe tener: `fecha_corte`, `fuente`, `responsable`.
- No publicar KPI financiero sin validación de Finanzas.
- Control de integridad: proyecto sin CAPEX o sin producción estimada = alerta roja.

---

## 5) Roadmap en 3 fases
## Fase 1 (0–4 semanas): Fundaciones
- Definir KPIs oficiales y diccionario de datos.
- Crear plantilla de ingestión y validación.
- Construir versión inicial del resumen ejecutivo + cartera.

## Fase 2 (5–8 semanas): Modelo inversor
- Integrar flujo de caja y escenarios.
- Añadir módulo de sensibilidad y comparación de escenarios.
- Implementar primeros anexos documentales trazables.

## Fase 3 (9–12 semanas): Escalado y gobernanza
- Automatizar refresco de datos.
- Añadir scoring de riesgo ajustado.
- Establecer comité de revisión mensual con checklist.

---

## 6) Estructura de tickets recomendada (backlog)
Formato sugerido: `AREA-TIPO-###`
- `DATA-ING-001` Ingesta de CAPEX histórico.
- `DATA-QA-002` Reglas de validación de supuestos financieros.
- `FIN-MOD-003` Cálculo VAN/TIR por proyecto.
- `FIN-SCN-004` Escenarios base/optimista/estrés.
- `DASH-UX-005` Vista ejecutiva para inversores.
- `RISK-REG-006` Registro y semáforo de riesgos.
- `DOC-DRM-007` Vinculación de anexos por proyecto.
- `GOV-REP-008` Reporte mensual para comité.

### Plantilla de ticket (estándar)
- **Objetivo**
- **Valor para inversor**
- **Entradas de datos**
- **Reglas de negocio**
- **Métrica de éxito (DoD)**
- **Dependencias**
- **Riesgos**
- **Owner y fecha objetivo**

---

## 7) KPIs obligatorios para comité
1. Capital requerido a 3/6/12 meses.
2. VAN/TIR consolidado y por proyecto.
3. Payback descontado.
4. Desviación CAPEX vs presupuesto.
5. Cumplimiento de hitos críticos.
6. Riesgo agregado y top 5 riesgos.
7. % documentación completa por proyecto.

---

## 8) Gobierno y operación
- Cadencia semanal: actualización operativa.
- Cadencia mensual: cierre financiero y comité de inversión.
- Roles:
  - **Product Owner (Finanzas)**: define KPI y prioridades.
  - **Data Owner**: garantiza calidad de datos.
  - **PMO**: gestiona tickets y bloqueos.
  - **Sponsor ejecutivo**: aprueba cambios de alcance.

---

## 9) Riesgos de implementación y mitigación
- **Datos incompletos** → política de “no KPI sin fuente”.
- **Supuestos inconsistentes** → catálogo único de supuestos versionado.
- **Exceso de complejidad** → iniciar con MVP (20% métricas que explican 80% decisiones).
- **Baja adopción** → diseñar pantallas orientadas a decisiones, no a tablas extensas.

---

## 10) Próximo paso accionable (esta semana)
1. Taller de 90 min con Finanzas + PMO + Operaciones.
2. Congelar diccionario de 25–30 campos críticos.
3. Crear backlog inicial de 12 tickets priorizados.
4. Publicar MVP del resumen ejecutivo con datos reales de 1–2 proyectos piloto.
