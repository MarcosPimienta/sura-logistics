## 1. Project Initialization & Tooling

- [x] 1.1 Initialize TypeScript Node.js project (`package.json`, `tsconfig.json`, build and test scripts)
- [x] 1.2 Install dependencies (`exceljs`, `zod`, `vitest` for automated testing)

## 2. Excel Parsing & Ingestion Layer (excel-intake)

- [x] 2.1 Define typed schemas and models (`LocationRecord`, `InventoryItem`, `ValidationResult`) using Zod
- [x] 2.2 Implement Excel workbook reader using `exceljs` supporting `.xlsx` files
- [x] 2.3 Implement flexible column header normalization (Location ID, Location Name, Current Stock, Threshold)
- [x] 2.4 Add schema validation and error reporting for invalid rows or missing required columns

## 3. Demand Analysis Engine (demand-analysis)

- [x] 3.1 Implement shortage evaluation logic comparing current stock against minimum threshold per location
- [x] 3.2 Implement replenishment calculation with optional packaging unit / pack-size rounding
- [x] 3.3 Add unit tests verifying calculations across edge cases (zero stock, surplus, exact match, negative values)

## 4. Reporting & Export (replenishment-reporting)

- [x] 4.1 Implement replenishment summary generator sorting locations by shortage priority (critical first)
- [x] 4.2 Implement manifest export generating formatted Excel/CSV files for dispatch
- [x] 4.3 Implement CLI entry point (`npm run analyze -- <path-to-excel>`) for command-line usage

## 5. Web Dashboard Interface

- [x] 5.1 Create web application with drag-and-drop Excel file upload interface
- [x] 5.2 Build interactive results table highlighting locations needing coffee bags with status badges
- [x] 5.3 Add export button allowing immediate download of the calculated replenishment manifest

## 6. Verification & Documentation

- [x] 6.1 Generate sample coffee bag inventory spreadsheet (`sample-inventory.xlsx`) for testing
- [x] 6.2 Run end-to-end verification test validating parsed output and demand calculations
- [x] 6.3 Update documentation with spreadsheet format requirements and run instructions
