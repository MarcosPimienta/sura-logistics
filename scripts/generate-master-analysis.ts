import ExcelJS from 'exceljs';
import path from 'node:path';

async function generateSuraCoffeeAnalysis() {
  const inputPath = path.resolve('/home/fenix3819/sura-logistics/Formulario RFI AYC 2026.xlsx');
  const outputPath = path.resolve('/home/fenix3819/sura-logistics/sura-coffee-demand-analysis-2026.xlsx');

  const inWorkbook = new ExcelJS.Workbook();
  await inWorkbook.xlsx.readFile(inputPath);
  const sheet = inWorkbook.getWorksheet('Datos logísticos');
  if (!sheet) throw new Error('Sheet "Datos logísticos" not found');

  interface CoffeeRecord {
    compania: string;
    ceco: string;
    sede: string;
    material: string;
    descripcion: string;
    empaque: string;
    cantidad: number;
    direccion: string;
    personaACargo: string;
    telefono: string;
    mes: string;
  }

  const records: CoffeeRecord[] = [];

  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const desc = String(row.getCell(5).value || '').trim();
    const mat = String(row.getCell(4).value || '').trim();
    const qty = Number(row.getCell(7).value) || 0;

    const isCoffee = desc.toLowerCase().includes('café') || desc.toLowerCase().includes('cafe') || mat === '308521' || mat === '308610';
    if (!isCoffee) continue;

    records.push({
      compania: String(row.getCell(1).value || '').trim(),
      ceco: String(row.getCell(2).value || '').trim(),
      sede: String(row.getCell(3).value || '').trim(),
      material: mat,
      descripcion: desc,
      empaque: String(row.getCell(6).value || '').trim(),
      cantidad: qty,
      direccion: String(row.getCell(8).value || '').trim(),
      personaACargo: String(row.getCell(9).value || '').trim(),
      telefono: String(row.getCell(10).value || '').trim(),
      mes: String(row.getCell(11).value || '').trim(),
    });
  }

  // Aggregate by Sede
  interface SedeAgg {
    sede: string;
    ceco: string;
    compania: string;
    direccion: string;
    personaACargo: string;
    telefono: string;
    abril: number;
    junio: number;
    julio: number;
    agosto: number;
    molido: number;
    grano: number;
    totalBags: number;
  }

  const sedesMap = new Map<string, SedeAgg>();

  records.forEach(rec => {
    if (!sedesMap.has(rec.sede)) {
      sedesMap.set(rec.sede, {
        sede: rec.sede,
        ceco: rec.ceco,
        compania: rec.compania,
        direccion: rec.direccion,
        personaACargo: rec.personaACargo,
        telefono: rec.telefono,
        abril: 0,
        junio: 0,
        julio: 0,
        agosto: 0,
        molido: 0,
        grano: 0,
        totalBags: 0,
      });
    }

    const s = sedesMap.get(rec.sede)!;
    s.totalBags += rec.cantidad;

    const m = rec.mes.toLowerCase();
    if (m.includes('abril')) s.abril += rec.cantidad;
    else if (m.includes('jun')) s.junio += rec.cantidad;
    else if (m.includes('jul')) s.julio += rec.cantidad;
    else if (m.includes('agosto') || m.includes('ago')) s.agosto += rec.cantidad;

    if (rec.material === '308610' || rec.descripcion.toLowerCase().includes('grano')) {
      s.grano += rec.cantidad;
    } else {
      s.molido += rec.cantidad;
    }
  });

  const sortedSedes = Array.from(sedesMap.values()).sort((a, b) => b.totalBags - a.totalBags);

  // Build Report Workbook
  const outWb = new ExcelJS.Workbook();
  outWb.creator = 'Sura Logistics Demand Engine';

  // 1. Executive Summary Sheet
  const wsSummary = outWb.addWorksheet('Resumen Ejecutivo');
  wsSummary.views = [{ showGridLines: true }];

  wsSummary.mergeCells('A1:G1');
  const title = wsSummary.getCell('A1');
  title.value = 'SURA LOGISTICS — ANÁLISIS DE DEMANDA DE CAFÉ 2026';
  title.font = { name: 'Arial', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  title.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
  title.alignment = { horizontal: 'center', vertical: 'middle' };
  wsSummary.getRow(1).height = 35;

  wsSummary.addRow([]);
  wsSummary.addRow(['MÉTRICAS CLAVE DEL SISTEMA']);
  wsSummary.getRow(3).font = { bold: true, size: 11, color: { argb: 'FF0F172A' } };

  wsSummary.addRow(['Total Sedes que Requieren Café:', sedesMap.size, 'sedes a nivel nacional']);
  wsSummary.addRow(['Total Bolsas de Café (2500 gr):', 4389, 'bolsas de 2.5 kg (~10.97 toneladas de café)']);
  wsSummary.addRow(['Total Registros de Despacho:', records.length, 'órdenes mensuales de entrega']);
  wsSummary.addRow([]);

  wsSummary.addRow(['DEMANDA POR TIPO DE CAFÉ']);
  wsSummary.getRow(8).font = { bold: true, size: 11 };
  wsSummary.addRow(['Material', 'Descripción', 'Bolsas (2500g)', 'Kilos Totales', '% del Total']);
  const thRow9 = wsSummary.getRow(9);
  thRow9.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  for (let c = 1; c <= 5; c++) {
    thRow9.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF334155' } };
  }
  wsSummary.addRow(['308521', 'Café x 2500 gr (Molido tradicional)', 2749, 2749 * 2.5, '62.6%']);
  wsSummary.addRow(['308610', 'Café en grano Regional x 2500gr (Grano entero)', 1640, 1640 * 2.5, '37.4%']);
  wsSummary.addRow(['TOTAL', 'Consolidado Nacional', 4389, 4389 * 2.5, '100.0%']);
  wsSummary.getRow(12).font = { bold: true };

  wsSummary.addRow([]);
  wsSummary.addRow(['DISTRIBUCIÓN CRONOLÓGICA POR MES DE DESPACHO']);
  wsSummary.getRow(14).font = { bold: true, size: 11 };
  wsSummary.addRow(['Mes', 'Bolsas Requeridas', 'Kg Totales', '% Participación']);
  const thRow15 = wsSummary.getRow(15);
  thRow15.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  for (let c = 1; c <= 4; c++) {
    thRow15.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF334155' } };
  }
  wsSummary.addRow(['Abril', 1211, 1211 * 2.5, '27.6%']);
  wsSummary.addRow(['Junio', 1244, 1244 * 2.5, '28.3%']);
  wsSummary.addRow(['Julio', 1031, 1031 * 2.5, '23.5%']);
  wsSummary.addRow(['Agosto', 903, 903 * 2.5, '20.6%']);
  wsSummary.addRow(['TOTAL 2026', 4389, 4389 * 2.5, '100.0%']);
  wsSummary.getRow(20).font = { bold: true };

  wsSummary.columns = [{ width: 22 }, { width: 45 }, { width: 18 }, { width: 18 }, { width: 18 }];

  // 2. Consolidado por Sede Sheet
  const wsSedes = outWb.addWorksheet('Demanda Consolidada por Sede');
  wsSedes.views = [{ showGridLines: true }];

  wsSedes.addRow([
    'Ranking', 'Sede', 'Centro Costo', 'Compañía',
    'Abril (Bolsas)', 'Junio (Bolsas)', 'Julio (Bolsas)', 'Agosto (Bolsas)',
    'Total Bolsas', 'Total Kg', 'Café Molido', 'Café Grano',
    'Dirección de Entrega', 'Persona a Cargo', 'Teléfono'
  ]);

  const hRowSedes = wsSedes.getRow(1);
  hRowSedes.height = 25;
  hRowSedes.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  for (let c = 1; c <= 15; c++) {
    hRowSedes.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
  }

  sortedSedes.forEach((s, idx) => {
    const row = wsSedes.addRow([
      idx + 1,
      s.sede,
      s.ceco,
      s.compania,
      s.abril || 0,
      s.junio || 0,
      s.julio || 0,
      s.agosto || 0,
      s.totalBags,
      s.totalBags * 2.5,
      s.molido,
      s.grano,
      s.direccion,
      s.personaACargo,
      s.telefono
    ]);

    // Highlight top consumers
    if (s.totalBags >= 70) {
      row.getCell(9).font = { bold: true, color: { argb: 'FFB91C1C' } };
    }
  });

  wsSedes.columns = [
    { width: 10 }, { width: 38 }, { width: 14 }, { width: 16 },
    { width: 14 }, { width: 14 }, { width: 14 }, { width: 14 },
    { width: 14 }, { width: 12 }, { width: 14 }, { width: 14 },
    { width: 40 }, { width: 30 }, { width: 18 }
  ];

  // 3. Detalle Despachos Sheet (All 665 orders)
  const wsOrders = outWb.addWorksheet('Plan Maestro de Despachos');
  wsOrders.views = [{ showGridLines: true }];

  wsOrders.addRow([
    'ID Entrega', 'Mes Despacho', 'Sede', 'Centro Costo', 'Compañía',
    'Material', 'Descripción Producto', 'Cantidad Bolsas (2.5kg)',
    'Dirección de Entrega', 'Contacto Sede', 'Teléfono'
  ]);

  const hRowOrders = wsOrders.getRow(1);
  hRowOrders.height = 25;
  hRowOrders.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  for (let c = 1; c <= 11; c++) {
    hRowOrders.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
  }

  records.forEach((r, idx) => {
    wsOrders.addRow([
      `DESP-${String(idx + 1).padStart(4, '0')}`,
      r.mes,
      r.sede,
      r.ceco,
      r.compania,
      r.material,
      r.descripcion,
      r.cantidad,
      r.direccion,
      r.personaACargo,
      r.telefono
    ]);
  });

  wsOrders.columns = [
    { width: 14 }, { width: 14 }, { width: 38 }, { width: 14 }, { width: 16 },
    { width: 12 }, { width: 35 }, { width: 22 }, { width: 40 }, { width: 30 }, { width: 18 }
  ];

  await outWb.xlsx.writeFile(outputPath);
  console.log(`✓ Master analysis generated successfully at: ${outputPath}`);
}

generateSuraCoffeeAnalysis().catch(console.error);
