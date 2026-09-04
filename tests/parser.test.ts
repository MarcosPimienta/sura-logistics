import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import { parseExcelWorkbook } from '../src/core/parser.js';

describe('Excel Parser & Schema Normalization', () => {
  it('parses valid spreadsheet with standard columns and extracts locations', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Locations');
    sheet.addRow(['Location ID', 'Location Name', 'Current Stock', 'Min Threshold', 'Pack Size', 'Region']);
    sheet.addRow(['LOC-101', 'Downtown Branch', 5, 20, 5, 'North']);
    sheet.addRow(['LOC-102', 'Airport Hub', 18, 15, 1, 'West']);

    const buffer = await workbook.xlsx.writeBuffer();
    const result = await parseExcelWorkbook(Buffer.from(buffer));

    expect(result.validRowsCount).toBe(2);
    expect(result.errors.length).toBe(0);
    expect(result.records[0].locationId).toBe('LOC-101');
    expect(result.records[0].currentStock).toBe(5);
    expect(result.records[0].minimumThreshold).toBe(20);
    expect(result.records[0].packagingUnit).toBe(5);
    expect(result.records[0].region).toBe('North');
  });

  it('handles spanish and alternative header aliases', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Inventario');
    sheet.addRow(['Sucursal', 'Bolsas de Cafe', 'Stock Minimo', 'Caja']);
    sheet.addRow(['Sede Central', 8, 25, 10]);

    const buffer = await workbook.xlsx.writeBuffer();
    const result = await parseExcelWorkbook(Buffer.from(buffer));

    expect(result.validRowsCount).toBe(1);
    expect(result.records[0].locationName).toBe('Sede Central');
    expect(result.records[0].currentStock).toBe(8);
    expect(result.records[0].minimumThreshold).toBe(25);
    expect(result.records[0].packagingUnit).toBe(10);
  });

  it('flags negative stock and non-numeric entries as row errors', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Data');
    sheet.addRow(['Store ID', 'Store Name', 'Bags', 'Min']);
    sheet.addRow(['LOC-201', 'Valid Store', 10, 15]);
    sheet.addRow(['LOC-202', 'Negative Store', -3, 15]);
    sheet.addRow(['LOC-203', 'Corrupted Store', 'N/A', 15]);

    const buffer = await workbook.xlsx.writeBuffer();
    const result = await parseExcelWorkbook(Buffer.from(buffer));

    expect(result.validRowsCount).toBe(1);
    expect(result.errors.length).toBe(2);
    expect(result.errors[0].locationId).toBe('LOC-202');
    expect(result.errors[1].locationId).toBe('LOC-203');
  });
});
