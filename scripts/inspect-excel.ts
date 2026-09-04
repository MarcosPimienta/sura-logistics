import ExcelJS from 'exceljs';
import path from 'node:path';

async function inspectWorkbook() {
  const filePath = path.resolve('/home/fenix3819/sura-logistics/Formulario RFI AYC 2026.xlsx');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  console.log(`=== WORKBOOK INSPECTION ===`);
  console.log(`Total worksheets: ${workbook.worksheets.length}`);

  workbook.worksheets.forEach((ws, idx) => {
    console.log(`\n--- Sheet ${idx + 1}: "${ws.name}" (Rows: ${ws.rowCount}, Columns: ${ws.columnCount}) ---`);
    const previewRows = Math.min(15, ws.rowCount);
    for (let r = 1; r <= previewRows; r++) {
      const row = ws.getRow(r);
      const values: any[] = [];
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        let val = cell.value;
        if (val && typeof val === 'object') {
          if ('result' in val) val = (val as any).result;
          else if ('text' in val) val = (val as any).text;
        }
        values[colNumber] = val;
      });
      // Compact display
      const filtered = values.filter(v => v !== undefined && v !== null && v !== '');
      if (filtered.length > 0) {
        console.log(`Row ${r}:`, JSON.stringify(values.slice(1, 20)));
      }
    }
  });
}

inspectWorkbook().catch(err => {
  console.error('Inspection error:', err);
});
