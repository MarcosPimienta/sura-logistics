import fs from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';

async function generateReports() {
  const inputPath = path.resolve('/home/fenix3819/sura-logistics/Formulario RFI AYC 2026.xlsx');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputPath);
  const sheet = workbook.getWorksheet('Datos logísticos')!;

  interface RowData {
    cia: string;
    ceco: string;
    sede: string;
    mat: string;
    desc: string;
    qty: number;
    dir: string;
    resp: string;
    tel: string;
    mes: string;
  }

  const rows: RowData[] = [];
  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const desc = String(row.getCell(5).value || '').trim();
    const mat = String(row.getCell(4).value || '').trim();
    const isCoffee = desc.toLowerCase().includes('café') || desc.toLowerCase().includes('cafe') || mat === '308521' || mat === '308610';
    if (!isCoffee) continue;

    rows.push({
      cia: String(row.getCell(1).value || '').trim(),
      ceco: String(row.getCell(2).value || '').trim(),
      sede: String(row.getCell(3).value || '').trim(),
      mat,
      desc,
      qty: Number(row.getCell(7).value) || 0,
      dir: String(row.getCell(8).value || '').trim(),
      resp: String(row.getCell(9).value || '').trim(),
      tel: String(row.getCell(10).value || '').trim(),
      mes: String(row.getCell(11).value || '').trim(),
    });
  }

  // Aggregations by Sede
  const sedesMap = new Map<string, {
    sede: string;
    ceco: string;
    cia: string;
    dir: string;
    resp: string;
    tel: string;
    abril: number;
    junio: number;
    julio: number;
    agosto: number;
    molido: number;
    grano: number;
    total: number;
  }>();

  const ciasMap = new Map<string, number>();

  rows.forEach(r => {
    ciasMap.set(r.cia, (ciasMap.get(r.cia) || 0) + r.qty);

    if (!sedesMap.has(r.sede)) {
      sedesMap.set(r.sede, {
        sede: r.sede,
        ceco: r.ceco,
        cia: r.cia,
        dir: r.dir,
        resp: r.resp,
        tel: r.tel,
        abril: 0,
        junio: 0,
        julio: 0,
        agosto: 0,
        molido: 0,
        grano: 0,
        total: 0,
      });
    }

    const s = sedesMap.get(r.sede)!;
    s.total += r.qty;

    const m = r.mes.toLowerCase();
    if (m.includes('abril')) s.abril += r.qty;
    else if (m.includes('jun')) s.junio += r.qty;
    else if (m.includes('jul')) s.julio += r.qty;
    else if (m.includes('agosto') || m.includes('ago')) s.agosto += r.qty;

    if (r.mat === '308610' || r.desc.toLowerCase().includes('grano')) {
      s.grano += r.qty;
    } else {
      s.molido += r.qty;
    }
  });

  const sortedSedes = Array.from(sedesMap.values()).sort((a, b) => b.total - a.total);

  // 1. Generate Markdown Report with ALL 203 SEDES
  const mdContent = `# INFORME EJECUTIVO: DEMANDA Y DISTRIBUCIÓN DE CAFÉ SURA 2026

**Fecha de Emisión:** 4 de Septiembre de 2026  
**Origen de Datos:** Formulario RFI AYC 2026  
**Alcance:** Planificación Logística y Despacho de Café a Nivel Nacional (203 Sedes)  

---

## 1. RESUMEN EJECUTIVO

El presente informe consolida los requerimientos de suministro de café para las sedes de **SURA** a nivel nacional durante el ciclo operativo 2026. A partir del análisis del Formulario RFI AYC, se identifican las sedes que requieren producto, los volúmenes exactos en unidades (bolsas de 2.5 kg) y kilogramos, la distribución por tipo de café (molido vs. grano) y el cronograma de despachos por mes.

### Indicadores Clave (KPIs)
* **Demanda Total de Café:** **4,389 bolsas de 2,500 gramos** (Equivalente a **10,972.5 kg** / ~11 toneladas).
* **Total de Sedes con Requerimiento:** **203 sedes operativas** a nivel nacional.
* **Total de Despachos Programados:** **665 órdenes de entrega**.
* **Frecuencia Operativa:** 4 olas de despacho en los meses de **Abril, Junio, Julio y Agosto**.

---

## 2. DESGLOSE POR TIPO DE PRODUCTO

| Código Material | Descripción del Producto | Unidad de Empaque | Bolsas Solicitadas | Kilos Totales | Participación |
| :--- | :--- | :---: | :---: | :---: | :---: |
| **308521** | **Café x 2500 gr** (Molido tradicional) | Bolsa 2.5 kg | **2,749** | 6,872.5 kg | **62.6%** |
| **308610** | **Café en grano Regional x 2500gr** (Grano entero) | Bolsa 2.5 kg | **1,640** | 4,100.0 kg | **37.4%** |
| **TOTAL** | | | **4,389** | **10,972.5 kg** | **100.0%** |

---

## 3. CRONOGRAMA MENSUAL DE DESPACHOS

| Mes de Entrega | Bolsas Requeridas | Kilogramos Totales | Órdenes de Entrega | % del Volumen |
| :--- | :---: | :---: | :---: | :---: |
| **Abril** | **1,211** | 3,027.5 kg | 179 | 27.6% |
| **Junio** | **1,244** | 3,110.0 kg | 184 | 28.3% |
| **Julio** | **1,031** | 2,577.5 kg | 158 | 23.5% |
| **Agosto** | **903** | 2,257.5 kg | 144 | 20.6% |
| **TOTAL** | **4,389** | **10,972.5 kg** | **665** | **100.0%** |

---

## 4. DISTRIBUCIÓN POR COMPAÑÍA

| Compañía / Entidad | Bolsas Totales | Kilogramos | % Participación |
| :--- | :---: | :---: | :---: |
${Array.from(ciasMap.entries())
  .sort((a, b) => b[1] - a[1])
  .map(([cia, qty]) => `| **${cia}** | ${qty} | ${(qty * 2.5).toLocaleString()} kg | ${((qty / 4389) * 100).toFixed(1)}% |`)
  .join('\n')}

---

## 5. CONSOLIDADO NACIONAL: TODAS LAS SEDES QUE ORDENARON CAFÉ (203 SEDES)

Listado completo de las 203 sedes ordenadas por volumen total de demanda:

| # | Sede | Centro de Costo | Abril | Junio | Julio | Agosto | Total Bolsas | Café Molido | Café Grano |
| :-: | :--- | :---: | :-: | :-: | :-: | :-: | :-: | :-: | :-: |
${sortedSedes.map((s, idx) => 
  `| ${idx + 1} | **${s.sede}** | \`${s.ceco}\` | ${s.abril || 0} | ${s.junio || 0} | ${s.julio || 0} | ${s.agosto || 0} | **${s.total}** | ${s.molido || 0} | ${s.grano || 0} |`
).join('\n')}

---

## 6. RECOMENDACIONES OPERATIVAS Y LOGÍSTICAS

1. **Gestión de Lotes y Frescura:** Programar la tostión y molienda máximo 15 días antes de cada ventana mensual (Abril, Junio, Julio, Agosto).
2. **Embalaje Secundario:** Agrupar en cajas de 4 o 5 bolsas (10 a 12.5 kg por bulto) con rotulación del ID de entrega y sede de destino.
3. **Validación de Entrega en Destino:** Exigir firma y sello de remisión (POD) con verificación física de unidades por la persona a cargo.
4. **Plan Maestro Detallado:** Para consultar direcciones de entrega, nombres de los responsables y teléfonos de contacto de cada una de las 203 sedes, remitirse al archivo \`Informe_Logistico_Cafe_Sura_2026.xlsx\` (pestaña *Plan Maestro de Despachos*).
`;

  fs.writeFileSync('/home/fenix3819/sura-logistics/INFORME_EJECUTIVO_DEMANDA_CAFE_SURA_2026.md', mdContent);
  console.log('✓ Updated Markdown executive report with ALL 203 sedes: INFORME_EJECUTIVO_DEMANDA_CAFE_SURA_2026.md');

  // 2. Generate Standalone HTML Report with ALL 203 SEDES + Live Filter
  const htmlContent = `<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Informe Ejecutivo — Demanda de Café Sura 2026 (Consolidado Nacional 203 Sedes)</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700;800&family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
  <style>
    :root {
      --primary: #0f172a;
      --secondary: #1e293b;
      --accent: #d97706;
      --accent-light: #fef3c7;
      --accent-warm: #f59e0b;
      --text-dark: #0f172a;
      --text-muted: #64748b;
      --bg-page: #f8fafc;
      --bg-card: #ffffff;
      --border: #e2e8f0;
      --danger: #dc2626;
      --success: #16a34a;
    }

    * { box-sizing: border-box; margin: 0; padding: 0; }

    body {
      font-family: 'Plus Jakarta Sans', sans-serif;
      background: var(--bg-page);
      color: var(--text-dark);
      line-height: 1.6;
      padding: 2.5rem 1rem;
    }

    .report-wrapper {
      max-width: 1280px;
      margin: 0 auto;
      background: var(--bg-card);
      border-radius: 20px;
      box-shadow: 0 10px 30px rgba(0, 0, 0, 0.06);
      border: 1px solid var(--border);
      overflow: hidden;
    }

    .report-header {
      background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
      color: #ffffff;
      padding: 3rem 3.5rem;
      position: relative;
    }

    .badge-header {
      display: inline-block;
      background: rgba(245, 158, 11, 0.2);
      border: 1px solid rgba(245, 158, 11, 0.4);
      color: #fcd34d;
      font-size: 0.8rem;
      font-weight: 700;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      padding: 0.35rem 0.85rem;
      border-radius: 30px;
      margin-bottom: 1.25rem;
    }

    .report-title {
      font-family: 'Outfit', sans-serif;
      font-size: 2.4rem;
      font-weight: 800;
      line-height: 1.2;
      margin-bottom: 0.75rem;
    }

    .report-subtitle {
      color: #94a3b8;
      font-size: 1.05rem;
      max-width: 750px;
    }

    .report-meta-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
      gap: 1.5rem;
      margin-top: 2.5rem;
      padding-top: 2rem;
      border-top: 1px solid rgba(255, 255, 255, 0.1);
    }

    .meta-item .meta-label {
      font-size: 0.75rem;
      text-transform: uppercase;
      letter-spacing: 0.05em;
      color: #94a3b8;
      margin-bottom: 0.25rem;
    }

    .meta-item .meta-val {
      font-family: 'Outfit', sans-serif;
      font-size: 1.1rem;
      font-weight: 600;
      color: #ffffff;
    }

    .report-body {
      padding: 3rem 3.5rem;
    }

    .section-heading {
      font-family: 'Outfit', sans-serif;
      font-size: 1.5rem;
      font-weight: 700;
      color: var(--primary);
      margin: 2.5rem 0 1.25rem;
      display: flex;
      align-items: center;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 1rem;
    }

    .section-heading:first-child { margin-top: 0; }

    /* KPI Highlights Grid */
    .kpi-row {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 1.25rem;
      margin-bottom: 2.5rem;
    }

    .kpi-card {
      background: #f8fafc;
      border: 1px solid var(--border);
      border-radius: 14px;
      padding: 1.5rem;
      transition: transform 0.2s;
    }

    .kpi-card-highlight {
      background: linear-gradient(145deg, #fffbeb, #fef3c7);
      border-color: #fde68a;
    }

    .kpi-card-val {
      font-family: 'Outfit', sans-serif;
      font-size: 2.3rem;
      font-weight: 800;
      color: var(--primary);
      margin-bottom: 0.25rem;
    }

    .kpi-card-highlight .kpi-card-val {
      color: #b45309;
    }

    .kpi-card-label {
      font-size: 0.85rem;
      font-weight: 600;
      color: var(--text-muted);
    }

    /* Search & Filter Bar */
    .table-controls {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 1rem;
      margin-bottom: 1rem;
      flex-wrap: wrap;
    }

    .search-input {
      padding: 0.65rem 1rem;
      font-family: var(--font-body);
      font-size: 0.9rem;
      border: 1px solid var(--border);
      border-radius: 8px;
      width: 320px;
      outline: none;
      background: #ffffff;
      transition: border-color 0.2s;
    }

    .search-input:focus {
      border-color: var(--accent-warm);
    }

    .counter-badge {
      font-size: 0.85rem;
      color: var(--text-muted);
      font-weight: 600;
    }

    /* Tables */
    .data-table-wrapper {
      overflow-x: auto;
      border: 1px solid var(--border);
      border-radius: 12px;
      margin-bottom: 2rem;
      max-height: 680px;
    }

    table {
      width: 100%;
      border-collapse: collapse;
      text-align: left;
      font-size: 0.88rem;
    }

    th {
      background: #f1f5f9;
      color: #334155;
      font-weight: 700;
      padding: 0.85rem 1rem;
      border-bottom: 1px solid var(--border);
      font-size: 0.78rem;
      text-transform: uppercase;
      letter-spacing: 0.04em;
      position: sticky;
      top: 0;
      z-index: 2;
    }

    td {
      padding: 0.85rem 1rem;
      border-bottom: 1px solid #f1f5f9;
      color: #1e293b;
    }

    tr:last-child td { border-bottom: none; }
    tr:hover td { background: #f8fafc; }

    .tag-product {
      display: inline-block;
      padding: 0.25rem 0.6rem;
      border-radius: 6px;
      font-size: 0.75rem;
      font-weight: 700;
    }

    .tag-molido { background: #e0f2fe; color: #0369a1; }
    .tag-grano { background: #fef3c7; color: #92400e; }

    .print-btn {
      background: var(--accent-warm);
      color: #000;
      font-weight: 700;
      padding: 0.75rem 1.5rem;
      border-radius: 10px;
      border: none;
      cursor: pointer;
      font-size: 0.9rem;
      display: inline-flex;
      align-items: center;
      gap: 0.5rem;
      box-shadow: 0 4px 12px rgba(245, 158, 11, 0.25);
    }

    .print-btn:hover { background: #d97706; }

    @media print {
      body { padding: 0; background: #fff; }
      .report-wrapper { box-shadow: none; border: none; max-width: 100%; }
      .print-btn, .table-controls { display: none; }
      .data-table-wrapper { max-height: none; overflow: visible; }
    }
  </style>
</head>
<body>
  <div class="report-wrapper">
    <div class="report-header">
      <div style="display: flex; justify-content: space-between; align-items: flex-start;">
        <div>
          <span class="badge-header">Informe Operativo Oficial</span>
          <h1 class="report-title">Demanda y Abastecimiento de Café SURA</h1>
          <p class="report-subtitle">Consolidado Nacional Completo (203 Sedes) y Plan Maestro de Despacho 2026 — Formulario RFI AYC</p>
        </div>
        <button class="print-btn" onclick="window.print()">🖨 Imprimir / Guardar PDF</button>
      </div>

      <div class="report-meta-grid">
        <div class="meta-item">
          <div class="meta-label">Fecha del Informe</div>
          <div class="meta-val">Septiembre 2026</div>
        </div>
        <div class="meta-item">
          <div class="meta-label">Sedes con Requerimiento</div>
          <div class="meta-val">203 Sedes</div>
        </div>
        <div class="meta-item">
          <div class="meta-label">Órdenes de Despacho</div>
          <div class="meta-val">665 Entregas</div>
        </div>
        <div class="meta-item">
          <div class="meta-label">Volumen Total</div>
          <div class="meta-val">10.97 Toneladas</div>
        </div>
      </div>
    </div>

    <div class="report-body">
      <div class="section-heading">
        <span>📌 1. Resumen de Indicadores Clave</span>
      </div>
      <div class="kpi-row">
        <div class="kpi-card kpi-card-highlight">
          <div class="kpi-card-val">4,389</div>
          <div class="kpi-card-label">Bolsas de Café (2.5 kg) Solicitadas</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-card-val">2,749</div>
          <div class="kpi-card-label">Bolsas Café Molido (62.6%)</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-card-val">1,640</div>
          <div class="kpi-card-label">Bolsas Café en Grano (37.4%)</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-card-val">203</div>
          <div class="kpi-card-label">Total Sedes con Pedido</div>
        </div>
      </div>

      <div class="section-heading">
        <span>☕ 2. Desglose por Especificación de Producto</span>
      </div>
      <div class="data-table-wrapper" style="max-height: none;">
        <table>
          <thead>
            <tr>
              <th>Código Material</th>
              <th>Descripción del Producto</th>
              <th>Unidad de Empaque</th>
              <th>Bolsas Totales</th>
              <th>Kilogramos</th>
              <th>% Participación</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td><strong>308521</strong></td>
              <td>Café x 2500 gr (Molido tradicional)</td>
              <td>1 Bolsa X 2.5 kg</td>
              <td><strong>2,749</strong></td>
              <td>6,872.5 kg</td>
              <td><span class="tag-product tag-molido">62.6%</span></td>
            </tr>
            <tr>
              <td><strong>308610</strong></td>
              <td>Café en grano Regional x 2500gr (Grano entero)</td>
              <td>1 Bolsa X 2.5 kg</td>
              <td><strong>1,640</strong></td>
              <td>4,100.0 kg</td>
              <td><span class="tag-product tag-grano">37.4%</span></td>
            </tr>
            <tr style="background: #f8fafc; font-weight: 700;">
              <td colspan="3">TOTAL CONSOLIDADO NACIONAL</td>
              <td>4,389 bolsas</td>
              <td>10,972.5 kg</td>
              <td>100.0%</td>
            </tr>
          </tbody>
        </table>
      </div>

      <div class="section-heading">
        <span>📅 3. Cronograma Mensual de Despachos</span>
      </div>
      <div class="data-table-wrapper" style="max-height: none;">
        <table>
          <thead>
            <tr>
              <th>Mes de Entrega</th>
              <th>Bolsas Requeridas</th>
              <th>Kilos Totales</th>
              <th>Órdenes de Entrega</th>
              <th>% del Volumen Anual</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td><strong>Abril 2026</strong></td>
              <td>1,211 bolsas</td>
              <td>3,027.5 kg</td>
              <td>179 entregas</td>
              <td>27.6%</td>
            </tr>
            <tr>
              <td><strong>Junio 2026</strong></td>
              <td>1,244 bolsas</td>
              <td>3,110.0 kg</td>
              <td>184 entregas</td>
              <td>28.3%</td>
            </tr>
            <tr>
              <td><strong>Julio 2026</strong></td>
              <td>1,031 bolsas</td>
              <td>2,577.5 kg</td>
              <td>158 entregas</td>
              <td>23.5%</td>
            </tr>
            <tr>
              <td><strong>Agosto 2026</strong></td>
              <td>903 bolsas</td>
              <td>2,257.5 kg</td>
              <td>144 entregas</td>
              <td>20.6%</td>
            </tr>
            <tr style="background: #f8fafc; font-weight: 700;">
              <td>TOTAL 2026</td>
              <td>4,389 bolsas</td>
              <td>10,972.5 kg</td>
              <td>665 órdenes</td>
              <td>100.0%</td>
            </tr>
          </tbody>
        </table>
      </div>

      <div class="section-heading">
        <span>🏢 4. Consolidado Nacional: Todas las Sedes que Ordenaron Café (203 Sedes)</span>
      </div>

      <div class="table-controls">
        <input type="text" id="search-input" class="search-input" placeholder="🔍 Buscar por nombre de sede, centro de costo o dirección..." oninput="filterSedes()">
        <span class="counter-badge" id="counter-badge">Mostrando 203 de 203 sedes</span>
      </div>

      <div class="data-table-wrapper">
        <table id="sedes-table">
          <thead>
            <tr>
              <th>#</th>
              <th>Sede</th>
              <th>Centro Costo</th>
              <th>Compañía</th>
              <th>Abr</th>
              <th>Jun</th>
              <th>Jul</th>
              <th>Ago</th>
              <th>Total Bolsas</th>
              <th>Molido</th>
              <th>Grano</th>
            </tr>
          </thead>
          <tbody>
            ${sortedSedes.map((s, idx) => `
              <tr data-sede="${s.sede.toLowerCase()}" data-ceco="${s.ceco.toLowerCase()}" data-dir="${s.dir.toLowerCase()}">
                <td style="color: #94a3b8; font-weight: 600;">${idx + 1}</td>
                <td><strong>${escapeHtml(s.sede)}</strong><div style="font-size: 0.75rem; color: #64748b;">${escapeHtml(s.dir || '—')}</div></td>
                <td style="font-family: monospace; font-size: 0.8rem;">${escapeHtml(s.ceco)}</td>
                <td style="font-size: 0.8rem; color: #475569;">${escapeHtml(s.cia)}</td>
                <td>${s.abril || '—'}</td>
                <td>${s.junio || '—'}</td>
                <td>${s.julio || '—'}</td>
                <td>${s.agosto || '—'}</td>
                <td><strong style="color: #b45309; font-size: 1rem;">${s.total}</strong></td>
                <td>${s.molido || 0}</td>
                <td>${s.grano || 0}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>

      <div class="section-heading">
        <span>📦 5. Archivos Complementarios para Operaciones</span>
      </div>
      <div style="background: #f8fafc; border: 1px solid var(--border); border-radius: 12px; padding: 1.25rem; font-size: 0.9rem;">
        <p style="margin-bottom: 0.5rem;">Para la programación detallada de rutas, etiquetas y órdenes de compra con contactos y teléfonos de cada una de las 203 sedes, consulte:</p>
        <ul style="padding-left: 1.25rem; color: #334155;">
          <li><strong>Plan Maestro Excel:</strong> <code>Informe_Logistico_Cafe_Sura_2026.xlsx</code> (Contiene las 665 órdenes desglosadas por dirección, responsable y teléfono).</li>
          <li><strong>Formulario Original:</strong> <code>Formulario RFI AYC 2026.xlsx</code> (Fuente primaria de datos).</li>
        </ul>
      </div>
    </div>
  </div>

  <script>
    function filterSedes() {
      const query = document.getElementById('search-input').value.toLowerCase().trim();
      const rows = document.querySelectorAll('#sedes-table tbody tr');
      let visible = 0;

      rows.forEach(row => {
        const sede = row.getAttribute('data-sede') || '';
        const ceco = row.getAttribute('data-ceco') || '';
        const dir = row.getAttribute('data-dir') || '';
        
        if (!query || sede.includes(query) || ceco.includes(query) || dir.includes(query)) {
          row.style.display = '';
          visible++;
        } else {
          row.style.display = 'none';
        }
      });

      document.getElementById('counter-badge').textContent = 'Mostrando ' + visible + ' de ' + rows.length + ' sedes';
    }
  </script>
</body>
</html>`;

  fs.writeFileSync('/home/fenix3819/sura-logistics/Informe_Ejecutivo_Demanda_Cafe_Sura_2026.html', htmlContent);
  console.log('✓ Updated Standalone HTML executive report with ALL 203 sedes + search bar: Informe_Ejecutivo_Demanda_Cafe_Sura_2026.html');
}

function escapeHtml(str: string) {
  return String(str).replace(/[&<>"']/g, (m) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  })[m] || m);
}

generateReports().catch(console.error);
