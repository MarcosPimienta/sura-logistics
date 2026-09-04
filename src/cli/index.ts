import fs from 'node:fs';
import path from 'node:path';
import { parseExcelWorkbook } from '../core/parser.js';
import { analyzeInventory } from '../core/analyzer.js';
import { generateExcelReport, generateCsvReport } from '../core/exporter.js';

async function main() {
  const args = process.argv.slice(2);
  const filePath = args[0];

  if (!filePath || filePath === '--help' || filePath === '-h') {
    console.log(`
\x1b[1;36m☕ Sura Logistics — Coffee Bag Demand Analyzer\x1b[0m

\x1b[1mUsage:\x1b[0m
  npm run analyze -- <path-to-excel-file> [options]

\x1b[1mOptions:\x1b[0m
  --export-csv       Save output as CSV alongside input file
  --export-excel     Save formatted output as Excel alongside input file
  --default-pack <N> Specify fallback pack size if not in sheet (default: 1)
  --help, -h         Show this help message

\x1b[1mExample:\x1b[0m
  npm run analyze -- sample-inventory.xlsx --export-excel
`);
    process.exit(filePath ? 0 : 1);
  }

  const resolvedPath = path.resolve(process.cwd(), filePath);
  if (!fs.existsSync(resolvedPath)) {
    console.error(`\x1b[1;31mError:\x1b[0m File not found: ${resolvedPath}`);
    process.exit(1);
  }

  console.log(`\x1b[1;34m» Ingesting workbook:\x1b[0m ${path.basename(resolvedPath)}`);
  const startTime = Date.now();

  try {
    const ingestion = await parseExcelWorkbook(resolvedPath);
    const summary = analyzeInventory(ingestion.records, ingestion.errors);
    const elapsed = Date.now() - startTime;

    console.log(`\n\x1b[1;32m✓ Analysis Complete in ${elapsed}ms\x1b[0m`);
    console.log(`\x1b[90mSheet:\x1b[0m ${ingestion.sheetName} | \x1b[90mRead:\x1b[0m ${ingestion.totalRowsRead} rows | \x1b[90mValid:\x1b[0m ${ingestion.validRowsCount} locations\n`);

    if (ingestion.errors.length > 0) {
      console.log(`\x1b[1;33m⚠ Ingestion Warnings (${ingestion.errors.length}):\x1b[0m`);
      ingestion.errors.slice(0, 5).forEach(err => {
        console.log(`  - Row ${err.rowNumber} (${err.locationName || err.locationId || 'Unknown'}): ${err.message}`);
      });
      if (ingestion.errors.length > 5) {
        console.log(`  ... and ${ingestion.errors.length - 5} more`);
      }
      console.log('');
    }

    console.log(`\x1b[1m━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\x1b[0m`);
    console.log(`\x1b[1m📋 REPLENISHMENT SUMMARY\x1b[0m`);
    console.log(`━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
    console.log(`Total Locations Analyzed:         ${summary.totalLocations}`);
    console.log(`\x1b[1;31mLocations Requiring Coffee Bags:  ${summary.locationsNeedingStock}\x1b[0m (Critical: ${summary.criticalCount}, Warning: ${summary.warningCount})`);
    console.log(`Locations with Adequate Stock:    ${summary.sufficientCount}`);
    console.log(`\x1b[1;32mTotal Coffee Bags to Dispatch:    ${summary.totalBagsNeeded} bags\x1b[0m`);
    console.log(`━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`);

    const deficitLocations = summary.results.filter(r => r.deficit > 0);

    if (deficitLocations.length === 0) {
      console.log(`\x1b[1;32m🎉 All locations currently have sufficient coffee bag stock!\x1b[0m\n`);
    } else {
      console.log(`\x1b[1mLOCATIONS IN NEED OF REPLENISHMENT:\x1b[0m\n`);
      console.log(
        `${'STATUS'.padEnd(11)} ${'LOCATION ID'.padEnd(14)} ${'LOCATION NAME'.padEnd(26)} ${'STOCK'.padEnd(8)} ${'MIN'.padEnd(8)} ${'DEFICIT'.padEnd(10)} ${'DISPATCH'.padEnd(10)}`
      );
      console.log('─'.repeat(90));

      deficitLocations.forEach(loc => {
        const statusColor = loc.urgency === 'CRITICAL' ? '\x1b[1;31m' : '\x1b[1;33m';
        const statusLabel = `${statusColor}${loc.urgency.padEnd(10)}\x1b[0m`;
        const id = loc.locationId.slice(0, 12).padEnd(14);
        const name = loc.locationName.slice(0, 24).padEnd(26);
        const stock = String(loc.currentStock).padEnd(8);
        const min = String(loc.minimumThreshold).padEnd(8);
        const deficit = `${loc.deficit} bags`.padEnd(10);
        const dispatch = `\x1b[1m${loc.recommendedReplenishment} bags\x1b[0m`;

        console.log(`${statusLabel} ${id} ${name} ${stock} ${min} ${deficit} ${dispatch}`);
      });
      console.log('');
    }

    // Export if requested
    const exportCsv = args.includes('--export-csv');
    const exportExcel = args.includes('--export-excel');

    if (exportCsv) {
      const outPath = resolvedPath.replace(/\.[^/.]+$/, '') + '-manifest.csv';
      fs.writeFileSync(outPath, generateCsvReport(summary));
      console.log(`\x1b[1;32m✓ CSV manifest exported to:\x1b[0m ${outPath}`);
    }

    if (exportExcel) {
      const outPath = resolvedPath.replace(/\.[^/.]+$/, '') + '-manifest.xlsx';
      const buffer = await generateExcelReport(summary);
      fs.writeFileSync(outPath, buffer);
      console.log(`\x1b[1;32m✓ Excel manifest exported to:\x1b[0m ${outPath}`);
    }
  } catch (err: unknown) {
    console.error(`\x1b[1;31mError analyzing spreadsheet:\x1b[0m`, (err as Error).message);
    process.exit(1);
  }
}

main();
