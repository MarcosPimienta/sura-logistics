## Why

Sura Logistics requires an automated solution to ingest location inventory data from Excel files and determine which operating locations require coffee bag replenishments. Currently, manual inspection of spreadsheet data is slow, prone to oversight, and delays coffee bag distribution. This initial system architecture establishes the data ingestion pipeline, inventory threshold calculation engine, and replenishment reporting interface.

## What Changes

- Initialize the project foundation using TypeScript / Node.js.
- Introduce Excel file parsing and validation to extract inventory records and consumption data per location.
- Implement a demand calculation engine to evaluate current stock against minimum thresholds and identify locations in deficit.
- Provide clear reporting and dispatch recommendations showing required coffee bag quantities per location.
- Support file upload via API/UI and export of calculated replenishment manifests.

## Capabilities

### New Capabilities
- `excel-intake`: Parsing, validation, and extraction of location inventory spreadsheets (.xlsx, .xls) with error reporting for malformed rows.
- `demand-analysis`: Inventory threshold evaluation, deficit detection, and coffee bag replenishment demand calculation per location.
- `replenishment-reporting`: Generation of location replenishment summaries, visual dashboard/views, and exportable dispatch recommendations (CSV/Excel).

### Modified Capabilities
<!-- None: This is the initial system architecture for a new repository. -->

## Impact

- **Tech Stack**: TypeScript, Node.js (with xlsx/exceljs parser, PostgreSQL/Prisma or embedded storage, validation via Zod).
- **APIs**: File upload endpoints, analysis trigger, and replenishment recommendation query endpoints.
- **Dependencies**: Adds spreadsheet parsing library (e.g., `xlsx` or `exceljs`), schema validation (`zod`), and backend framework.
