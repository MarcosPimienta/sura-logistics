import { describe, it, expect } from 'vitest';
import { analyzeLocationStock, analyzeInventory } from '../src/core/analyzer.js';
import { LocationRecord } from '../src/core/types.js';

describe('Demand Analysis Engine', () => {
  it('identifies location with zero stock as CRITICAL with full deficit', () => {
    const record: LocationRecord = {
      locationId: 'LOC-001',
      locationName: 'Central Station Cafe',
      currentStock: 0,
      minimumThreshold: 20,
      packagingUnit: 1,
    };

    const result = analyzeLocationStock(record);
    expect(result.urgency).toBe('CRITICAL');
    expect(result.deficit).toBe(20);
    expect(result.recommendedReplenishment).toBe(20);
  });

  it('identifies location below critical ratio (<= 30%) as CRITICAL', () => {
    const record: LocationRecord = {
      locationId: 'LOC-002',
      locationName: 'Airport Terminal 2',
      currentStock: 4,
      minimumThreshold: 20,
      packagingUnit: 1,
    };

    const result = analyzeLocationStock(record);
    expect(result.urgency).toBe('CRITICAL');
    expect(result.deficit).toBe(16);
  });

  it('identifies location between 30% and 100% of threshold as WARNING', () => {
    const record: LocationRecord = {
      locationId: 'LOC-003',
      locationName: 'Uptown Corner',
      currentStock: 12,
      minimumThreshold: 20,
      packagingUnit: 1,
    };

    const result = analyzeLocationStock(record);
    expect(result.urgency).toBe('WARNING');
    expect(result.deficit).toBe(8);
  });

  it('rounds up replenishment to packaging unit multiples', () => {
    const record: LocationRecord = {
      locationId: 'LOC-004',
      locationName: 'Midtown Hub',
      currentStock: 6,
      minimumThreshold: 20, // deficit = 14
      packagingUnit: 5, // round up 14 to next multiple of 5 => 15
    };

    const result = analyzeLocationStock(record);
    expect(result.deficit).toBe(14);
    expect(result.recommendedReplenishment).toBe(15);
  });

  it('marks location with exact minimum threshold as SUFFICIENT with 0 replenishment', () => {
    const record: LocationRecord = {
      locationId: 'LOC-005',
      locationName: 'North Gate',
      currentStock: 20,
      minimumThreshold: 20,
      packagingUnit: 5,
    };

    const result = analyzeLocationStock(record);
    expect(result.urgency).toBe('SUFFICIENT');
    expect(result.deficit).toBe(0);
    expect(result.recommendedReplenishment).toBe(0);
  });

  it('marks location with high surplus (>150% threshold) as SURPLUS', () => {
    const record: LocationRecord = {
      locationId: 'LOC-006',
      locationName: 'West Plaza',
      currentStock: 40,
      minimumThreshold: 20,
      packagingUnit: 1,
    };

    const result = analyzeLocationStock(record);
    expect(result.urgency).toBe('SURPLUS');
    expect(result.deficit).toBe(0);
  });

  it('calculates days of supply remaining when daily consumption is present', () => {
    const record: LocationRecord = {
      locationId: 'LOC-007',
      locationName: 'Express Kiosk',
      currentStock: 15,
      minimumThreshold: 25,
      dailyConsumption: 5,
      packagingUnit: 1,
    };

    const result = analyzeLocationStock(record);
    expect(result.daysOfSupplyRemaining).toBe(3);
  });

  it('correctly compiles overall inventory summary and prioritizes critical locations first', () => {
    const records: LocationRecord[] = [
      { locationId: 'A', locationName: 'Sufficient Store', currentStock: 25, minimumThreshold: 20, packagingUnit: 1 },
      { locationId: 'B', locationName: 'Critical Store', currentStock: 2, minimumThreshold: 20, packagingUnit: 5 }, // deficit 18 -> 20
      { locationId: 'C', locationName: 'Warning Store', currentStock: 14, minimumThreshold: 20, packagingUnit: 5 }, // deficit 6 -> 10
    ];

    const summary = analyzeInventory(records);
    expect(summary.totalLocations).toBe(3);
    expect(summary.locationsNeedingStock).toBe(2);
    expect(summary.criticalCount).toBe(1);
    expect(summary.warningCount).toBe(1);
    expect(summary.sufficientCount).toBe(1);
    expect(summary.totalBagsNeeded).toBe(30); // 20 + 10

    // Check sort order: Critical store first
    expect(summary.results[0].locationId).toBe('B');
    expect(summary.results[1].locationId).toBe('C');
    expect(summary.results[2].locationId).toBe('A');
  });
});
