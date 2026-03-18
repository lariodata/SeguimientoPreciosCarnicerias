# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**Workflow Seguimiento de Precios** is an internal pricing-tracking system for Rafaela Alimentos S.A. It automates weekly extraction, transformation, and reporting of product prices, costs, and competitor pricing data. The workflow is driven by a VBA-enabled Excel workbook (`WFW_SPC.xlsm`) that invokes Python scripts, which query SQL Server stored procedures, transform data with pandas, and write results back into Excel sheets via xlwings.

## Architecture

### VBA → Python → SQL Server → Excel pipeline

1. **`WFW_SPC.xlsm`** (Excel/VBA frontend): Users trigger operations from Excel buttons. VBA macros call Python scripts via `sys.argv` passing parameters (number of weeks, list IDs, dates, file paths).
2. **Python scripts** (`00 - Workflow/scripts/`): Each script connects to SQL Server (`lariosql70`, database `DW`), executes stored procedures, pivots/transforms data with pandas, and writes formatted results back into specific Excel sheets using xlwings.
3. **SQL Server stored procedures** (`00 - Workflow/sql/`): Server-side logic for aggregating prices and costs. Key SPs: `SP_CONS_ListarPreciosUltimasSemanas`, `SP_CONS_ListarCostosUltimasSemanas`, `SP_CONS_ListarPreciosCompetenciaSemanal`.
4. **Email sending**: Two scripts generate per-cluster Excel attachments and send them via Office 365 SMTP.

### Key Scripts

| Script | Purpose | Excel Sheet(s) |
|---|---|---|
| `cargar_ListasPrecios.py` | Load price lists by `ID_LISTA` | Multiple sheets via `MAPA_HOJAS` dict |
| `cargar_costos.py` | Load unit costs | "Costos Unitarios" |
| `cargar_costos_comerciales.py` | Load commercial costs | "Costos Comerciales" |
| `cargar_precios_competencia.py` | Load competitor prices from SP | "Competencias" |
| `migrar_precios_competencia.py` | Import competitor prices from `Template_Competencias.xlsx` into SQL | Writes to `Lista_Precios_Competencia` table |
| `migrar_ListaPrecios_*.py` | Migrate price data from Excel sheets into SQL | Read from old sheets, insert into `Lista_Precios_Odoo` |
| `enviar_correos_carnicerias.py` | Generate per-cluster Excel files and email to butcher shops | Reads "Clusters" + "Parametros" sheets |
| `enviar_correos_ventas.py` | Generate per-cluster Excel files and email to sales team | Reads "Clusters" + "Parametros" sheets |
| `ETL_SeguimientoPrecios.py` | ETL process (belongs to Procesos project, runs as scheduled task) | Executes GOLD stored procedures |

### Common Patterns Across Scripts

- **Parameters from VBA**: Scripts receive arguments via `sys.argv` (weeks count, list IDs, dates in `YYYY-MM-DD`, file paths).
- **Status files**: Scripts write completion/error status to `logs/estado_*.txt` so VBA can poll results.
- **xlwings COM interaction**: All scripts use `xw.apps.active` to connect to the already-open Excel instance, then `app.books["WFW_SPC.xlsm"]`.
- **Pivot + variation columns**: Price/cost scripts pivot by week, then insert interleaved "%" variation columns between consecutive weeks.
- **Conditional formatting**: Yellow (`#FFFF99`) for positive variations, red (`#FFC7CE`) for negative. Rows with `UNIDAD_ME` starting with "1-" get light blue highlighting.

## Development Setup

```bash
# Virtual environment (Python 3.13)
python -m venv .venv
.venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

### Key Dependencies

- `pandas`, `numpy` — data transformation
- `pyodbc` — SQL Server connectivity (uses `{SQL Server}` driver)
- `xlwings` — COM-based Excel interaction (requires Excel to be running on Windows)
- `openpyxl`, `xlsxwriter` — Excel file read/write
- `python-dotenv` — environment config from `.env`
- `pywin32` — Windows COM support (required by xlwings)

### Running Scripts
#### Comment folder scripts
In the folder scripts we have all files with process scrits that are called from excel file "WFW_SCP". This excel file contain the app with the workflow. For example: @cargar_costos is called from the workflo to get the unit costs.

**Script naming clarification:**
- `cargar_ListasPrecios.py` is referred to as **"listado"** or **"listas de precios"** in conversations
Scripts are normally invoked from VBA but can be run manually:

```bash
# Price lists (args: weeks, list_id)
python "00 - Workflow/scripts/cargar_ListasPrecios.py" 4 2

# Costs (args: weeks)
python "00 - Workflow/scripts/cargar_costos.py" 4

# Competitor prices import (args: excel_file_path)
python "00 - Workflow/scripts/migrar_precios_competencia.py" "path/to/Template_Competencias.xlsx"

# Email sending (args: date YYYY-MM-DD)
python "00 - Workflow/scripts/enviar_correos_carnicerias.py" 2026-03-09
```

**Important**: Most scripts require Excel to be open with `WFW_SPC.xlsm` loaded, as they interact via xlwings COM.

## Project Structure

```
00 - Workflow/
  ├── .env              # SMTP credentials (gitignored)
  ├── WFW_SPC.xlsm      # Main Excel workbook (VBA frontend)
  ├── Template_Competencias.xlsx  # Input template for competitor prices
  ├── scripts/           # Python scripts (all script logic)
  ├── sql/               # SQL Server stored procedure definitions
  ├── logs/              # Status files and logs (gitignored)
  ├── temp/              # Generated Excel attachments for email (gitignored)
  └── bkp/               # Workbook backups (gitignored)
Avances/                 # Weekly progress presentations (.pptx)
```

## Important Notes

- The `ETL_SeguimientoPrecios.py` script belongs to a separate "Procesos" project and imports from a `core` package that lives outside this repository.
- SQL Server connection details are hardcoded in most scripts (server: `lariosql70`, DB: `DW`). Only the ETL script uses `ConnectionFactory` from the external core package.
- The email scripts read recipient addresses and mail parameters from the "Parametros" sheet in the workbook (different cell ranges for each mail type).
- Cluster definitions (Rafaela, Casilda, Maria Luisa, Mayorista, HORECA, Estancia Rafaela) are business-domain groupings used throughout the system.


