## Context

Sura Logistics requires an automated application to analyze Excel spreadsheets containing location inventory records and identify which locations need coffee bag replenishments. The system must accommodate variations in spreadsheet structures, validate inventory numbers, compute replenishment needs with packaging rules, and provide both web/visual and exportable outputs.

## Goals / Non-Goals

**Goals:**
- Provide reliable ingestion and parsing for `.xlsx` and `.xls` workbooks.
- Flexible column-mapping to recognize standard inventory headers (Location ID, Location Name, Current Stock, Minimum Threshold / Daily Demand).
- Fast, pure-function demand calculation engine that identifies shortage locations and calculates required coffee bag quantities.
- Interactive web dashboard (Next.js / TypeScript) with drag-and-drop Excel upload and instant results visualization.
- Export capabilities for dispatch manifests (Excel/CSV).
- Optional CLI command (`npm run analyze <file>`) for quick automated or terminal-based execution.

**Non-Goals:**
- Full ERP integration (SAP/Oracle/NetSuite) in this initial release.
- Multi-warehouse automatic order routing or carrier dispatch integration (saved for future changes).
- Machine learning-based demand forecasting (heuristic thresholds and runout calculation suffice for v1).

## Decisions

### 1. Technology Foundation: TypeScript & Next.js (App Router)
- **Rationale**: Combines a robust typed backend API runtime with a reactive modern UI in a unified TypeScript codebase.
- **Alternatives considered**:
  - Separate Express API + React SPA: Adds unnecessary repo complexity and dual deployment overhead for an internal tool.
  - Python / FastAPI: Good for data processing, but user specifically requested TypeScript/Node.js.

### 2. Spreadsheet Engine: `exceljs`
- **Rationale**: Clean streaming support, modern TypeScript typings, zero CVE history compared to older `xlsx` (SheetJS) community versions, and excellent support for reading styles, headers, and generating formatted exports.
- **Alternatives considered**:
  - `xlsx` (SheetJS): Fast, but licensing and maintenance concerns in npm ecosystem.
  - Python openpyxl / pandas: Excluded per TypeScript requirement.

### 3. Core Engine Architecture: Pure Domain Layer
- **Architecture**: Decoupled domain engine (`src/core/analyzer/`):
  - `Parser`: Converts Excel rows into typed `LocationRecord` objects.
  - `Normalizer`: Standardizes location names, numbers, and validates schemas using Zod.
  - `Calculator`: Pure function `(records: LocationRecord[], options: AnalysisOptions) => AnalysisResult`.
  - `Exporter`: Generates exportable summary workbook.
- **Rationale**: Allows 100% unit testability without mock databases or HTTP servers, and enables seamless reuse between Web UI and CLI scripts.

### 4. Storage and State Management
- **Decision**: Ephemeral processing for upload-and-analyze workflow with optional persistence (Prisma + PostgreSQL / SQLite for historical runs).
- **Rationale**: Users need immediate analysis upon file drop without mandatory account setups, while architecture permits logging runs for audit trails.

## Risks / Trade-offs

- **[Risk] Spreadsheet Column Mismatch**: Excel templates from different stores or operations teams might use different header labels (e.g., "Store #" vs "Location ID", "Bags" vs "Current Stock").
  - *Mitigation*: Implement intelligent header mapping with fuzzy matching and a fallback column-mapping step in the UI.
- **[Risk] Corrupted or Large Excel Files**: Uploading multi-megabyte or malformed workbooks could cause memory spikes or server hangs.
  - *Mitigation*: Enforce file size limits (e.g. 25MB), validate MIME types, and use streamed parsing if row counts exceed 10,000.
- **[Risk] Missing Threshold Data**: Some locations in the sheet might lack a defined minimum threshold.
  - *Mitigation*: Allow a global default minimum threshold configuration (e.g., default 10 bags) with override alerts in the UI.
