import ExcelJS from 'exceljs';
import { LocationRecord, IngestionResult, RowError, LocationRecordSchema } from './types.js';

interface HeaderMapping {
  locationId?: number;
  locationName?: number;
  currentStock?: number;
  minimumThreshold?: number;
  dailyConsumption?: number;
  packagingUnit?: number;
  region?: number;
  contact?: number;
}

const HEADER_ALIASES: Record<keyof HeaderMapping, string[]> = {
  locationId: [
    'locationid', 'location_id', 'location id', 'loc id', 'store id', 'store_id',
    'store #', 'store number', 'id', 'code', 'location code', 'codigo', 'cod', 'sucursal id', 'sede id'
  ],
  locationName: [
    'locationname', 'location_name', 'location name', 'location', 'store name', 'store',
    'branch', 'branch name', 'sucursal', 'sede', 'nombre', 'sitio', 'point of sale'
  ],
  currentStock: [
    'currentstock', 'current_stock', 'current stock', 'stock', 'coffee bags', 'coffeebags',
    'bags', 'bags on hand', 'inventory', 'existencias', 'bolsas', 'bolsas de cafe', 'on hand', 'qty'
  ],
  minimumThreshold: [
    'minimumthreshold', 'minimum_threshold', 'minimum threshold', 'threshold', 'min threshold',
    'min stock', 'minimum stock', 'min', 'stock minimo', 'minimo', 'reorder point', 'safety stock'
  ],
  dailyConsumption: [
    'dailyconsumption', 'daily_consumption', 'daily consumption', 'consumption', 'avg daily consumption',
    'consumo', 'consumo diario', 'daily burn', 'burn rate'
  ],
  packagingUnit: [
    'packagingunit', 'packaging_unit', 'pack size', 'packsize', 'case size', 'unit size',
    'unidades por paquete', 'caja', 'paquete'
  ],
  region: ['region', 'city', 'zone', 'area', 'ciudad', 'zona', 'territory'],
  contact: ['contact', 'manager', 'contact person', 'email', 'phone', 'contacto', 'responsable'],
};

function normalizeHeaderString(val: unknown): string {
  if (val === null || val === undefined) return '';
  return String(val)
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '') // remove accents
    .replace(/[^a-z0-9]/g, '');
}

export function detectHeaders(row: ExcelJS.Row): { mapping: HeaderMapping; headerRowNumber: number } | null {
  const mapping: HeaderMapping = {};
  let matchedKeyCount = 0;

  row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    const rawVal = cell.value;
    const normalized = normalizeHeaderString(rawVal);
    if (!normalized) return;

    for (const [key, aliases] of Object.entries(HEADER_ALIASES) as [keyof HeaderMapping, string[]][]) {
      if (mapping[key] === undefined) {
        const matches = aliases.some(alias => {
          const normAlias = normalizeHeaderString(alias);
          return normalized === normAlias || normalized.includes(normAlias);
        });
        if (matches) {
          mapping[key] = colNumber;
          matchedKeyCount++;
          break;
        }
      }
    }
  });

  // A row is considered a valid header if it matches at least location and stock indicators
  if ((mapping.locationId || mapping.locationName) && (mapping.currentStock || mapping.minimumThreshold)) {
    return { mapping, headerRowNumber: row.number };
  }

  return null;
}

function parseNumericValue(val: unknown): number | null {
  if (val === null || val === undefined || val === '') return null;
  if (typeof val === 'number') return isNaN(val) ? null : val;
  if (typeof val === 'object' && 'result' in (val as Record<string, unknown>)) {
    // Excel formula result
    const result = (val as { result: unknown }).result;
    return typeof result === 'number' ? result : Number(result);
  }
  const cleanStr = String(val).replace(/[^0-9.-]/g, '').trim();
  if (cleanStr === '' || cleanStr === '-' || cleanStr === '.') return null;
  const parsed = Number(cleanStr);
  return isNaN(parsed) ? null : parsed;
}

function parseStringValue(val: unknown): string {
  if (val === null || val === undefined) return '';
  if (typeof val === 'object' && 'text' in (val as Record<string, unknown>)) {
    return String((val as { text: unknown }).text).trim();
  }
  return String(val).trim();
}

export async function parseExcelWorkbook(source: string | Buffer): Promise<IngestionResult> {
  const workbook = new ExcelJS.Workbook();

  if (typeof source === 'string') {
    await workbook.xlsx.readFile(source);
  } else {
    await workbook.xlsx.load(source as any);
  }

  const worksheet = workbook.worksheets[0];
  if (!worksheet || worksheet.rowCount === 0) {
    throw new Error('The workbook contains no sheets or is empty');
  }

  let headerDetection: { mapping: HeaderMapping; headerRowNumber: number } | null = null;

  // Scan the first 10 rows to locate the header row
  for (let r = 1; r <= Math.min(10, worksheet.rowCount); r++) {
    const row = worksheet.getRow(r);
    const detected = detectHeaders(row);
    if (detected) {
      headerDetection = detected;
      break;
    }
  }

  if (!headerDetection) {
    throw new Error(
      'Could not detect valid column headers. Please ensure columns for Location (Name or ID) and Coffee Bag Stock/Threshold exist.'
    );
  }

  const { mapping, headerRowNumber } = headerDetection;
  const records: LocationRecord[] = [];
  const errors: RowError[] = [];
  let totalRowsRead = 0;

  for (let r = headerRowNumber + 1; r <= worksheet.rowCount; r++) {
    const row = worksheet.getRow(r);
    // Skip completely empty rows
    let hasValue = false;
    row.eachCell(() => { hasValue = true; });
    if (!hasValue) continue;

    totalRowsRead++;

    const rawLocationId = mapping.locationId ? parseStringValue(row.getCell(mapping.locationId).value) : '';
    const rawLocationName = mapping.locationName ? parseStringValue(row.getCell(mapping.locationName).value) : '';
    const rawStock = mapping.currentStock ? parseNumericValue(row.getCell(mapping.currentStock).value) : null;
    const rawThreshold = mapping.minimumThreshold ? parseNumericValue(row.getCell(mapping.minimumThreshold).value) : null;
    const rawDailyConsumption = mapping.dailyConsumption ? parseNumericValue(row.getCell(mapping.dailyConsumption).value) : null;
    const rawPackSize = mapping.packagingUnit ? parseNumericValue(row.getCell(mapping.packagingUnit).value) : 1;
    const rawRegion = mapping.region ? parseStringValue(row.getCell(mapping.region).value) : undefined;
    const rawContact = mapping.contact ? parseStringValue(row.getCell(mapping.contact).value) : undefined;

    const locationId = rawLocationId || rawLocationName || `LOC-${r}`;
    const locationName = rawLocationName || rawLocationId || `Location ${r}`;

    if (!rawLocationId && !rawLocationName) {
      errors.push({
        rowNumber: r,
        message: 'Row has no Location ID or Location Name',
        rawValues: { row: r }
      });
      continue;
    }

    if (rawStock === null) {
      errors.push({
        rowNumber: r,
        locationId,
        locationName,
        field: 'currentStock',
        message: 'Current stock value is missing or not a valid number'
      });
      continue;
    }

    if (rawStock < 0) {
      errors.push({
        rowNumber: r,
        locationId,
        locationName,
        field: 'currentStock',
        message: `Current stock cannot be negative (found: ${rawStock})`
      });
      continue;
    }

    const minimumThreshold = rawThreshold !== null ? Math.max(0, rawThreshold) : 10; // Default threshold fallback
    const packagingUnit = rawPackSize && rawPackSize > 0 ? Math.floor(rawPackSize) : 1;

    const recordData = {
      locationId,
      locationName,
      currentStock: rawStock,
      minimumThreshold,
      dailyConsumption: rawDailyConsumption !== null ? rawDailyConsumption : undefined,
      packagingUnit,
      region: rawRegion || undefined,
      contact: rawContact || undefined,
    };

    const validated = LocationRecordSchema.safeParse(recordData);
    if (!validated.success) {
      errors.push({
        rowNumber: r,
        locationId,
        locationName,
        message: validated.error.issues.map(i => i.message).join(', ')
      });
    } else {
      records.push(validated.data);
    }
  }

  return {
    records,
    errors,
    totalRowsRead,
    validRowsCount: records.length,
    sheetName: worksheet.name,
  };
}
