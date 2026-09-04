import ExcelJS from 'exceljs';
import path from 'node:path';

async function createSampleExcel() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Sura Logistics';
  const sheet = workbook.addWorksheet('Coffee Inventory');

  sheet.columns = [
    { header: 'Location ID', key: 'locationId', width: 16 },
    { header: 'Location Name', key: 'locationName', width: 32 },
    { header: 'Current Stock (Bags)', key: 'currentStock', width: 22 },
    { header: 'Min Threshold', key: 'minimumThreshold', width: 16 },
    { header: 'Pack Size', key: 'packagingUnit', width: 14 },
    { header: 'Daily Consumption', key: 'dailyConsumption', width: 18 },
    { header: 'Region', key: 'region', width: 20 },
  ];

  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF1E293B' },
  };

  const rows = [
    { locationId: 'LOC-101', locationName: 'Sura Headquarters - Torre 1', currentStock: 2, minimumThreshold: 30, packagingUnit: 5, dailyConsumption: 6, region: 'Bogota - Calle 26' },
    { locationId: 'LOC-102', locationName: 'Sura Sede Norte Cafe', currentStock: 0, minimumThreshold: 20, packagingUnit: 5, dailyConsumption: 4, region: 'Bogota - Usaquen' },
    { locationId: 'LOC-103', locationName: 'Centro de Servicios Salitre', currentStock: 8, minimumThreshold: 25, packagingUnit: 5, dailyConsumption: 5, region: 'Bogota - Salitre' },
    { locationId: 'LOC-104', locationName: 'Sucursal Plaza Claro', currentStock: 14, minimumThreshold: 15, packagingUnit: 5, dailyConsumption: 3, region: 'Bogota - Salitre' },
    { locationId: 'LOC-105', locationName: 'Sede Medellin Poblado', currentStock: 3, minimumThreshold: 25, packagingUnit: 5, dailyConsumption: 5, region: 'Medellin' },
    { locationId: 'LOC-106', locationName: 'Sede Medellin Centro', currentStock: 22, minimumThreshold: 20, packagingUnit: 5, dailyConsumption: 4, region: 'Medellin' },
    { locationId: 'LOC-107', locationName: 'Sucursal Cali Granada', currentStock: 5, minimumThreshold: 20, packagingUnit: 5, dailyConsumption: 3, region: 'Cali' },
    { locationId: 'LOC-108', locationName: 'Sucursal Barranquilla Prado', currentStock: 18, minimumThreshold: 15, packagingUnit: 5, dailyConsumption: 2, region: 'Barranquilla' },
    { locationId: 'LOC-109', locationName: 'Sura Bucaramanga Cabecera', currentStock: 1, minimumThreshold: 15, packagingUnit: 5, dailyConsumption: 3, region: 'Bucaramanga' },
    { locationId: 'LOC-110', locationName: 'Hub Logistico Principal', currentStock: 120, minimumThreshold: 50, packagingUnit: 10, dailyConsumption: 10, region: 'Bogota - Funza' },
  ];

  rows.forEach(r => sheet.addRow(r));

  const outPath = path.resolve(process.cwd(), 'sample-inventory.xlsx');
  await workbook.xlsx.writeFile(outPath);
  console.log(`✓ Sample spreadsheet created at: ${outPath}`);
}

createSampleExcel().catch(console.error);
