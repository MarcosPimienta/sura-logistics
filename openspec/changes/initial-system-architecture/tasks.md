## 1. Project Initialization & Tooling

- [ ] 1.1 Initialize TypeScript Node.js project (`package.json`, `tsconfig.json`, build and test scripts)
- [ ] 1.2 Install dependencies (`exceljs`, `zod`, `vitest` for automated testing)

## 2. Excel Parsing & Ingestion Layer (excel-intake)

- [ ] 2.1 Define typed schemas and models (`LocationRecord`, `InventoryItem`, `ValidationResult`) using Zod
- [ ] 2.2 Implement Excel workbook reader using `exceljs` supporting `.xlsx` files
- [ ] 2.3 Implement flexible column header normalization (Location ID, Location Name, Current Stock, Threshold)
- [ ] 2.4 Add schema validation and error reporting for invalid rows or missing required columns

## 3. Demand Analysis Engine (demand-analysis)

- [ ] 3.1 Implement shortage evaluation logic comparing current stock against minimum threshold per location
- [ ] 3.2 Implement replenishment calculation with optional packaging unit / pack-size rounding
- [ ] 3.3 Add unit tests verifying calculations across edge cases (zero stock, surplus, exact match, negative values)

## 4. Reporting & Export (replenishment-reporting)

- [ ] 4.1 Implement replenishment summary generator sorting locations by shortage priority (critical first)
- [ ] 4.2 Implement manifest export generating formatted Excel/CSV files for dispatch
- [ ] 4.3 Implement CLI entry point (`npm run analyze -- <path-to-excel>`) for command-line usage

## 5. Web Dashboard Interface

- [ ] 5.1 Create web application with drag-and-drop Excel file upload interface
- [ ] 5.2 Build interactive results table highlighting locations needing coffee bags with status badges
- [ ] 5.3 Add export button allowing immediate download of the calculated replenishment manifest

## 6. Verification & Documentation

- [ ] 6.1 Generate sample coffee bag inventory spreadsheet (`sample-inventory.xlsx`) for testing
- [ ] 6.2 Run end-to-end verification test validating parsed output and demand calculations
- [ ] 6.3 Update documentation with spreadsheet format requirements and run instructions
