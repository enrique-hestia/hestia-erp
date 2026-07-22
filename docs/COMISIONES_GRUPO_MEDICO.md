# Comisiones — Grupo Médico (Médicos Externos + Daniel)

> Spec de la regla de tiers, descuentos, notas de crédito y cargos para el
> "Grupo Médico". Capturado 2026-07-22 de la tabla que envió el usuario.
> **PENDIENTE que el usuario confirme la transcripción de las matrices** (vienen
> de una imagen). Ideal: pegar el spreadsheet original.

## Fórmula maestra (verificada)

    Percepción Hestia(tier) = Precio Médico Externo(tier) − Comisión Daniel(tier)

Tres conceptos por procedimiento:
- **Descuento al médico** = `Base − Precio Médico(tier)`. NO es efectivo, es precio más bajo. Único beneficio del médico.
- **Comisión Daniel** = tabla verde. SÍ es efectivo que Hestia paga (egreso), por tier y tratamiento.
- **Percepción Hestia** = el neto.

## Conteo para tier — SIMPLIFICADO (usuario 2026-07-22)

En vez de pesos por tratamiento: cada producto trae un flag **`cuentaParaTier`** (sí/no).
- Hay productos que **cuentan para descuento pero NO para tier**.
- El tier se calcula contando SOLO los procedimientos que cuentan para tier (los "activos"),
  y se promedia/mide sobre esos. Los que no cuentan igual reciben el precio del tier vigente
  (descuento) pero no empujan la meta.
- => Config: un booleano por producto. Meta 10/15/18/22/25 se mide sobre procedimientos con `cuentaParaTier=true`.

## Metas por tier (conteo de procedimientos que cuentan para tier)

| Tier | Meta (u. ponderadas) |
|---|---|
| 0 | base, sin descuento |
| 1 | 10 |
| 2 | 15 |
| 3 | 18 |
| 4 | 22 |
| 5 | 25 |

## Reglas confirmadas (respuestas del usuario 2026-07-22)

1. **Alcance del tier:** TOTAL de procedimientos del periodo → UN solo tier para todos sus precios, pero **ponderado** (cada tratamiento pesa distinto — PESOS PENDIENTES de que el usuario los dé).
2. **Momento:** al **cierre del periodo (mensual)**. Durante el mes el médico paga a su **tier de entrada** (el que ganó el mes pasado); al cierre se calcula el **tier ganado** y se genera UNA nota de crédito o UN cargo por el delta contra los procedimientos del periodo.
3. **Comisión Daniel:** aplica a **todo el Grupo Médico**, usando el **tier ganado**, pagada como egreso.

## Mecánica true-up (subir/bajar)

- Entra el periodo pagando **tier de entrada** (heredado del periodo anterior).
- Al cierre se mide el **tier ganado** este periodo.
- El tier ganado fija el **precio del siguiente periodo** Y dispara ajuste sobre los procedimientos de **este** periodo:
  - **Subió:** pagó de más → **nota de crédito (abono)** = `Precio(entrada) − Precio(ganado)` × cada procedimiento (por tratamiento). Ej. Captura+Vitr T1→T2 = $637/procedimiento.
  - **Bajó:** pagó de menos → **cargo (cobro)** = mismo delta, sobre los procedimientos del periodo.
- Neto: al final se le cobra el **tier realmente ganado** ese periodo. Con médicos nunca hay efectivo; con Daniel sí.

## Matriz — Precio a Médico Externo (Base | T0 | T1 | T2 | T3 | T4 | T5)

| Tratamiento | Base | T0 | T1 | T2 | T3 | T4 | T5 |
|---|--:|--:|--:|--:|--:|--:|--:|
| Captura + Vitrificación | 23,600 | 23,600 | 22,302 | 21,665 | 21,094 | 20,249 | 20,001 |
| Captura, Fert. y Vitrificación | 40,250 | 40,250 | 38,036 | 36,950 | 35,975 | 34,535 | 34,112 |
| Captura, Fert. y Transf. Fresco | 46,700 | 46,700 | 44,132 | 42,871 | 41,740 | 40,069 | 39,578 |
| Captura, Fert. y Transf. Diferida | 53,000 | 53,000 | 50,085 | 48,654 | 47,371 | 45,474 | 44,918 |
| Desvitrificación + Transferencia | 19,250 | 19,250 | 18,191 | 17,672 | 17,206 | 16,517 | 16,314 |
| Desvit., Fert., Transf. y Vitrif. | 26,700 | 26,700 | 25,232 | 24,511 | 23,864 | 22,909 | 22,628 |
| Histeroscopia Con | 23,000 | 23,000 | 21,850 | 20,700 | 20,010 | 19,090 | 18,400 |
| Histeroscopia Sin | 20,000 | 20,000* | 19,000 | 18,000 | 17,400 | 16,600 | 16,000 |
| Recepción de Células | 5,400 | 5,400 | 5,103 | 4,957 | 4,827 | 4,633 | 4,577 |

\* La imagen muestra Histeroscopia Sin T0 = 20,001 — probable typo de 20,000. CONFIRMAR.

## Matriz — Comisión Daniel (efectivo) (T0..T5; Base y T0 = 0)

| Tratamiento | T1 | T2 | T3 | T4 | T5 |
|---|--:|--:|--:|--:|--:|
| Captura + Vitrificación | 1,062 | 1,605 | 1,742 | 1,841 | 2,301 |
| Captura, Fert. y Vitrificación | 1,811 | 2,737 | 2,970 | 3,140 | 3,924 |
| Captura, Fert. y Transf. Fresco | 2,102 | 3,176 | 3,446 | 3,643 | 4,553 |
| Captura, Fert. y Transf. Diferida | 2,385 | 3,604 | 3,911 | 4,134 | 5,167 |
| Desvitrificación + Transferencia | 866 | 1,309 | 1,421 | 1,502 | 1,877 |
| Desvit., Fert., Transf. y Vitrif. | 1,202 | 1,816 | 1,970 | 2,083 | 2,603 |
| Histeroscopia Con | 1,150 | 1,750 | 2,070 | 2,530 | 2,875 |
| Histeroscopia Sin | 1,000 | 1,500 | 1,800 | 2,200 | 2,500 |
| Recepción de Células | 243 | 367 | 399 | 421 | 527 |

## PENDIENTE para implementar
- **Pesos por tratamiento** hacia la meta (regla "ponderado").
- **Tier de entrada inicial** de cada médico al arranque (periodo 0).
- Confirmar periodo = mes calendario.
- Confirmar transcripción de matrices / pegar el spreadsheet original.
- Histeroscopia y Recepción: ¿cómo se distinguen Con/Sin al capturar (2 productos del catálogo)?
- Ligar la nota de crédito al motor de CxC / próximo procedimiento; Daniel → egreso.
- Relación con `COBRANZA_DESC_CFG` (motor % actual) — Grupo Médico es perfil NUEVO (precio absoluto por tratamiento×tier + true-up + comisión efectivo).
