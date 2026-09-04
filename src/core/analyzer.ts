import {
  LocationRecord,
  LocationAnalysisResult,
  AnalysisSummary,
  AnalysisOptions,
  ReplenishmentUrgency,
  RowError
} from './types.js';

export function analyzeLocationStock(
  record: LocationRecord,
  options: AnalysisOptions = {}
): LocationAnalysisResult {
  const minThreshold = record.minimumThreshold;
  const current = record.currentStock;
  const packSize = record.packagingUnit && record.packagingUnit > 0 ? record.packagingUnit : (options.defaultPackSize || 1);
  const criticalRatio = options.criticalThresholdPercentage ?? 0.3; // <= 30% of threshold is critical

  let deficit = 0;
  let recommendedReplenishment = 0;
  let urgency: ReplenishmentUrgency = 'SUFFICIENT';

  if (current < minThreshold) {
    deficit = minThreshold - current;
    // Round up replenishment to multiple of packaging unit
    recommendedReplenishment = Math.ceil(deficit / packSize) * packSize;

    if (current === 0 || (minThreshold > 0 && (current / minThreshold) <= criticalRatio)) {
      urgency = 'CRITICAL';
    } else {
      urgency = 'WARNING';
    }
  } else if (current > minThreshold * 1.5) {
    urgency = 'SURPLUS';
  } else {
    urgency = 'SUFFICIENT';
  }

  let daysOfSupplyRemaining: number | undefined;
  if (record.dailyConsumption && record.dailyConsumption > 0) {
    daysOfSupplyRemaining = Number((current / record.dailyConsumption).toFixed(1));
  }

  return {
    locationId: record.locationId,
    locationName: record.locationName,
    currentStock: current,
    minimumThreshold: minThreshold,
    deficit,
    recommendedReplenishment,
    packagingUnit: packSize,
    urgency,
    daysOfSupplyRemaining,
    region: record.region,
  };
}

export function analyzeInventory(
  records: LocationRecord[],
  ingestionErrors: RowError[] = [],
  options: AnalysisOptions = {}
): AnalysisSummary {
  const results = records.map(record => analyzeLocationStock(record, options));

  // Sort results: CRITICAL first, then WARNING, then SUFFICIENT, then SURPLUS
  // Within same urgency, higher deficit first
  const urgencyWeight: Record<ReplenishmentUrgency, number> = {
    CRITICAL: 0,
    WARNING: 1,
    SUFFICIENT: 2,
    SURPLUS: 3,
  };

  results.sort((a, b) => {
    const weightDiff = urgencyWeight[a.urgency] - urgencyWeight[b.urgency];
    if (weightDiff !== 0) return weightDiff;
    return b.deficit - a.deficit;
  });

  const locationsNeedingStock = results.filter(r => r.deficit > 0).length;
  const totalBagsNeeded = results.reduce((acc, curr) => acc + curr.recommendedReplenishment, 0);
  const criticalCount = results.filter(r => r.urgency === 'CRITICAL').length;
  const warningCount = results.filter(r => r.urgency === 'WARNING').length;
  const sufficientCount = results.filter(r => r.urgency === 'SUFFICIENT' || r.urgency === 'SURPLUS').length;

  return {
    totalLocations: records.length,
    locationsNeedingStock,
    totalBagsNeeded,
    criticalCount,
    warningCount,
    sufficientCount,
    generatedAt: new Date().toISOString(),
    results,
    ingestionErrors,
  };
}
