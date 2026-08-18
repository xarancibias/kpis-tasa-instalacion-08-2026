# Tasa de Instalación BAF — ELECNOR

KPI operativo/comercial que mide el porcentaje de órdenes de venta BAF (banda ancha fija) que terminaron con una instalación técnica completada, por agencia zonal. Se actualiza mes a mes y acumula histórico para comparar períodos.

**Dashboard en vivo:** https://xarancibias.github.io/kpis-tasa-instalacion-08-2026/

## Definición del KPI

**Numerador** — archivo `pXX_TOA_ELECNOR_AAAA-MM.xlsx` (sistema TOA), agrupado por `XA_ORIGINAL_AGENCY_NAME`, contando `APPT_NUMBER_CD`, filtrado por:
- `STATUS_CD = complete`
- `XA_WORK_TYPE_ID` cuyos dos primeros caracteres correspondan a **"AT"** (productos BAF)

**Denominador** — archivo `Maestro_Ventas_Elecnor_AAAA-MM.xlsx`, agrupado por `AGENCIA_ZONAL_INSTALACION`, contando `ORDEN`:

```
Denominador = Q(TIPO = Emisión) − Q(TIPO = Cancelada)
```

**Meta:** 80,7%. Estados: Cumple meta (≥ meta) · Cerca de meta (dentro de −5 pts) · Crítico (bajo −5 pts) · Verificar dato (tasa ≥ 100%, señal de posible desfase entre fuentes).

## Dashboard

[`dashboard.html`](dashboard.html) / [`index.html`](index.html) — vista gerencial con selector de período, semáforo por agencia, gráfico vs. meta, tabla de detalle y **histórico mensual** (evolución de la tasa por agencia a través de los períodos cargados). Es la misma página, publicada automáticamente en GitHub Pages en el link de arriba.

## Cómo actualizar un nuevo período

Todo el proceso está automatizado en [`scripts/Actualizar-KPI.ps1`](scripts/Actualizar-KPI.ps1). Con los dos archivos fuente del mes (TOA y Maestro de Ventas) a mano:

```powershell
cd "C:\analisis\BBDD\p12_Tasa_Instalacion_BAF_2026-08"
.\scripts\Actualizar-KPI.ps1 `
  -ArchivoTOA "C:\Users\ximel\Downloads\p13_TOA_ELECNOR_2026-09.xlsx" `
  -ArchivoMaestroVentas "C:\Users\ximel\Downloads\Maestro_Ventas_Elecnor_2026-09.xlsx" `
  -Periodo "2026-09" `
  -Push
```

El script:
1. Exporta ambos archivos a CSV y calcula numerador/denominador por agencia con la misma metodología de arriba.
2. Agrega (o reemplaza, si vuelves a correr el mismo período) la fila del período en [`historico/kpi_historico.csv`](historico/kpi_historico.csv) — la fuente única de verdad del histórico.
3. Guarda el detalle del período en `periodos/<AAAA-MM>/datos/`.
4. Regenera `reportes/Reporte_Tasa_Instalacion_BAF_<AAAA-MM>.xlsx` (fórmulas vivas + gráfico nativo) dentro de `periodos/<AAAA-MM>/reportes/`.
5. Regenera `dashboard.html` / `index.html` con el histórico completo (todos los períodos quedan disponibles en el selector).
6. Hace `git commit`; con `-Push` además hace `git push` — GitHub Pages se actualiza solo, ~1 minuto después.

Si omites `-Push`, el commit queda listo localmente y puedes revisar el dashboard antes de publicarlo; luego simplemente `git push origin master`.

**Parámetro opcional:** `-Meta 80.7` (por defecto) si la meta cambia de período a período.

## Estructura del repositorio

```
├── README.md
├── dashboard.html / index.html          # dashboard (lee historico/kpi_historico.csv embebido)
├── .nojekyll
├── historico/
│   └── kpi_historico.csv                # una fila por (período, agencia) — fuente única del histórico
├── periodos/
│   └── 2026-08/
│       ├── datos/                       # detalle filtrado del numerador y denominador de ese período
│       ├── reportes/                    # Reporte_Tasa_Instalacion_BAF_2026-08.xlsx
│       └── graficos/                    # gráfico exportado como PNG
└── scripts/
    ├── Actualizar-KPI.ps1               # automatiza todo lo anterior para un período nuevo
    └── dashboard_template.html          # plantilla del dashboard (con marcador __KPI_DATA_JSON__)
```

Cada mes nuevo agrega su propia carpeta bajo `periodos/`, y su fila correspondiente en `historico/kpi_historico.csv`.

## Notas y advertencias (agosto 2026)

- **San Antonio superó el 100%** en agosto 2026, lo cual no es matemáticamente posible en una tasa real. Causa probable: desfase temporal entre fuentes — el numerador (p12) solo cubre el 1-15 de agosto, mientras el denominador (Maestro_Ventas) incluye órdenes cuya emisión original fue en julio y que se instalaron/completaron en agosto. Antes de reportar esta cifra a negocio se recomienda validar cruzando por número de orden individual entre ambos sistemas.
- Los nombres de agencia coinciden textualmente entre ambos archivos (`VALPARAISO`, `VINA DEL MAR`, `SAN ANTONIO`), pero no se validó que identifiquen exactamente la misma unidad operativa en ambos sistemas de origen.

## Requisitos para correr el script

- Windows con Excel instalado (el script usa Excel vía COM para leer los `.xlsx` fuente).
- Git configurado con acceso de push al repositorio (`gh auth status` para verificar).
