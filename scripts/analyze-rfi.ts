import ExcelJS from 'exceljs';
import path from 'node:path';

async function analyzeRFI() {
  const filePath = path.resolve('/home/fenix3819/sura-logistics/Formulario RFI AYC 2026.xlsx');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const sheet = workbook.getWorksheet('Datos logísticos');
  if (!sheet) {
    throw new Error('Sheet "Datos logísticos" not found');
  }

  console.log(`Analyzing sheet: "${sheet.name}" with ${sheet.rowCount} rows`);

  // Detect header row
  const headerRow = sheet.getRow(1);
  const colIndex: Record<string, number> = {};
  headerRow.eachCell((cell, colNum) => {
    const val = String(cell.value || '').trim().toLowerCase();
    colIndex[val] = colNum;
  });

  console.log('Headers found:', colIndex);

  const ciaCol = colIndex['compañía'] || 1;
  const cecoCol = colIndex['centro de costos2'] || 2;
  const sedeCol = colIndex['sede'] || 3;
  const matCol = colIndex['material'] || 4;
  const descCol = colIndex['descripción'] || 5;
  const cantCol = colIndex['cantidad'] || 7;
  const dirCol = colIndex['direccion'] || 8;
  const respCol = colIndex['persona a cargo'] || 9;
  const telCol = colIndex['telefono'] || 10;
  const mesCol = colIndex['mes'] || 11;

  interface SedeDemand {
    sede: string;
    ceco: string;
    compania: string;
    direccion: string;
    personaACargo: string;
    telefono: string;
    months: Record<string, number>;
    materials: Record<string, { desc: string; totalQty: number }>;
    totalCoffeeBags: number;
  }

  const sedesMap = new Map<string, SedeDemand>();
  const coffeeMaterials = new Map<string, { desc: string; count: number; totalQty: number }>();
  let totalCoffeeRows = 0;
  let totalCoffeeQty = 0;

  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const desc = String(row.getCell(descCol).value || '').trim();
    const mat = String(row.getCell(matCol).value || '').trim();
    const qty = Number(row.getCell(cantCol).value) || 0;
    const mes = String(row.getCell(mesCol).value || '').trim();
    const sede = String(row.getCell(sedeCol).value || '').trim();
    const ceco = String(row.getCell(cecoCol).value || '').trim();
    const cia = String(row.getCell(ciaCol).value || '').trim();
    const dir = String(row.getCell(dirCol).value || '').trim();
    const resp = String(row.getCell(respCol).value || '').trim();
    const tel = String(row.getCell(telCol).value || '').trim();

    // Check if item is coffee
    const isCoffee = desc.toLowerCase().includes('café') || desc.toLowerCase().includes('cafe') || mat === '308521' || mat === '308610';

    if (isCoffee) {
      totalCoffeeRows++;
      totalCoffeeQty += qty;

      if (!coffeeMaterials.has(mat)) {
        coffeeMaterials.set(mat, { desc, count: 0, totalQty: 0 });
      }
      const mInfo = coffeeMaterials.get(mat)!;
      mInfo.count++;
      mInfo.totalQty += qty;

      if (!sedesMap.has(sede)) {
        sedesMap.set(sede, {
          sede,
          ceco,
          compania: cia,
          direccion: dir,
          personaACargo: resp,
          telefono: tel,
          months: {},
          materials: {},
          totalCoffeeBags: 0,
        });
      }

      const sData = sedesMap.get(sede)!;
      sData.totalCoffeeBags += qty;
      sData.months[mes] = (sData.months[mes] || 0) + qty;
      
      if (!sData.materials[mat]) {
        sData.materials[mat] = { desc, totalQty: 0 };
      }
      sData.materials[mat].totalQty += qty;
    }
  }

  console.log(`\n======================================================`);
  console.log(`COFFEE MATERIALS FOUND:`);
  console.log(`======================================================`);
  coffeeMaterials.forEach((info, mat) => {
    console.log(`Material ${mat}: "${info.desc}" -> Rows: ${info.count}, Total Units: ${info.totalQty}`);
  });

  console.log(`\n======================================================`);
  console.log(`SUMMARY:`);
  console.log(`Total Coffee Records (Deliveries/Orders): ${totalCoffeeRows}`);
  console.log(`Total Unique Sedes (Locations) Requiring Coffee: ${sedesMap.size}`);
  console.log(`Total Coffee Bags (2500 gr units) across all months: ${totalCoffeeQty}`);
  console.log(`======================================================`);

  // Show top 20 sedes by coffee demand
  const sortedSedes = Array.from(sedesMap.values()).sort((a, b) => b.totalCoffeeBags - a.totalCoffeeBags);

  console.log(`\nTOP 25 LOCATIONS (SEDES) BY COFFEE BAG DEMAND:`);
  console.log(`${'SEDE'.padEnd(38)} ${'CECO'.padEnd(12)} ${'TOTAL'.padEnd(8)} ${'MONTHS'}`);
  console.log('-'.repeat(80));
  sortedSedes.slice(0, 25).forEach(s => {
    const monthsSummary = Object.entries(s.months).map(([m, q]) => `${m}:${q}`).join(', ');
    console.log(`${s.sede.padEnd(38).slice(0, 38)} ${s.ceco.padEnd(12)} ${(s.totalCoffeeBags + ' bags').padEnd(8)} ${monthsSummary}`);
  });

  // Check months distribution
  const monthTotals: Record<string, number> = {};
  sortedSedes.forEach(s => {
    Object.entries(s.months).forEach(([m, q]) => {
      monthTotals[m] = (monthTotals[m] || 0) + q;
    });
  });
  console.log(`\nDEMAND BY MONTH:`, monthTotals);
}

analyzeRFI().catch(console.error);
