import ExcelJS from 'exceljs';
import { AnalysisSummary, LocationAnalysisResult } from './types.js';

export async function generateExcelReport(summary: AnalysisSummary): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Sura Logistics Demand Engine';
  workbook.created = new Date();

  // 1. Replenishment Dispatch Sheet
  const sheet = workbook.addWorksheet('Replenishment Manifest', {
    views: [{ showGridLines: true }],
  });

  // Title Banner
  sheet.mergeCells('A1:I1');
  const titleCell = sheet.getCell('A1');
  titleCell.value = 'SURA LOGISTICS — COFFEE BAG REPLENISHMENT MANIFEST';
  titleCell.font = { name: 'Arial', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  titleCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF1E293B' }, // Slate 800
  };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  sheet.getRow(1).height = 35;

  // Metadata summary
  sheet.getCell('A2').value = `Generated: ${new Date(summary.generatedAt).toLocaleString()}`;
  sheet.getCell('A2').font = { italic: true, size: 9, color: { argb: 'FF64748B' } };

  sheet.getCell('A3').value = `Total Locations: ${summary.totalLocations}`;
  sheet.getCell('C3').value = `Locations in Deficit: ${summary.locationsNeedingStock}`;
  sheet.getCell('F3').value = `Total Coffee Bags Required: ${summary.totalBagsNeeded}`;
  sheet.getRow(3).font = { bold: true, size: 10 };

  sheet.addRow([]); // Blank row at row 4

  // Table Headers
  const headerRow = sheet.getRow(5);
  headerRow.values = [
    'Priority Status',
    'Location ID',
    'Location Name',
    'Region',
    'Current Stock',
    'Min Threshold',
    'Deficit',
    'Pack Size',
    'Recommended Dispatch'
  ];
  headerRow.height = 24;
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
  headerRow.alignment = { horizontal: 'center', vertical: 'middle' };

  for (let c = 1; c <= 9; c++) {
    const cell = headerRow.getCell(c);
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF0F172A' },
    };
  }

  // Populate data rows
  summary.results.forEach((item, index) => {
    const rowNumber = 6 + index;
    const row = sheet.getRow(rowNumber);

    row.values = [
      item.urgency,
      item.locationId,
      item.locationName,
      item.region || '—',
      item.currentStock,
      item.minimumThreshold,
      item.deficit,
      item.packagingUnit,
      item.recommendedReplenishment
    ];

    const statusCell = row.getCell(1);
    if (item.urgency === 'CRITICAL') {
      statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } }; // light red
      statusCell.font = { color: { argb: 'FF991B1B' }, bold: true };
    } else if (item.urgency === 'WARNING') {
      statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } }; // light amber
      statusCell.font = { color: { argb: 'FF92400E' }, bold: true };
    } else if (item.urgency === 'SUFFICIENT') {
      statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } }; // light green
      statusCell.font = { color: { argb: 'FF166534' } };
    } else {
      statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E7FF' } }; // light blue
      statusCell.font = { color: { argb: 'FF3730A3' } };
    }

    // Bold highlight for dispatch quantity
    const dispatchCell = row.getCell(9);
    if (item.recommendedReplenishment > 0) {
      dispatchCell.font = { bold: true, color: { argb: 'FFB91C1C' } };
    }
  });

  // Adjust column widths automatically
  sheet.columns = [
    { width: 16 }, // Priority Status
    { width: 15 }, // Location ID
    { width: 30 }, // Location Name
    { width: 18 }, // Region
    { width: 16 }, // Current Stock
    { width: 16 }, // Min Threshold
    { width: 12 }, // Deficit
    { width: 12 }, // Pack Size
    { width: 24 }, // Recommended Dispatch
  ];

  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

export function generateCsvReport(summary: AnalysisSummary): string {
  const headers = [
    'Priority Status',
    'Location ID',
    'Location Name',
    'Region',
    'Current Stock',
    'Min Threshold',
    'Deficit',
    'Pack Size',
    'Recommended Dispatch'
  ];

  const rows = summary.results.map(r => [
    r.urgency,
    `"${r.locationId.replace(/"/g, '""')}"`,
    `"${r.locationName.replace(/"/g, '""')}"`,
    `"${(r.region || '').replace(/"/g, '""')}"`,
    r.currentStock,
    r.minimumThreshold,
    r.deficit,
    r.packagingUnit,
    r.recommendedReplenishment
  ]);

  return [headers.join(','), ...rows.map(r => r.join(','))].join('\n');
}
