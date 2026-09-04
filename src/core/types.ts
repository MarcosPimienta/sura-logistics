import { z } from 'zod';

export const LocationRecordSchema = z.object({
  locationId: z.string().min(1, 'Location ID is required'),
  locationName: z.string().min(1, 'Location Name is required'),
  currentStock: z.number().min(0, 'Current stock must be zero or positive'),
  minimumThreshold: z.number().min(0, 'Minimum threshold must be zero or positive'),
  dailyConsumption: z.number().min(0).optional(),
  packagingUnit: z.number().int().min(1).default(1),
  region: z.string().optional(),
  contact: z.string().optional(),
});

export type LocationRecord = z.infer<typeof LocationRecordSchema>;

export interface RowError {
  rowNumber: number;
  locationId?: string;
  locationName?: string;
  field?: string;
  message: string;
  rawValues?: Record<string, unknown>;
}

export interface IngestionResult {
  records: LocationRecord[];
  errors: RowError[];
  totalRowsRead: number;
  validRowsCount: number;
  sheetName: string;
}

export type ReplenishmentUrgency = 'CRITICAL' | 'WARNING' | 'SUFFICIENT' | 'SURPLUS';

export interface LocationAnalysisResult {
  locationId: string;
  locationName: string;
  currentStock: number;
  minimumThreshold: number;
  deficit: number;
  recommendedReplenishment: number;
  packagingUnit: number;
  urgency: ReplenishmentUrgency;
  daysOfSupplyRemaining?: number;
  region?: string;
}

export interface AnalysisSummary {
  totalLocations: number;
  locationsNeedingStock: number;
  totalBagsNeeded: number;
  criticalCount: number;
  warningCount: number;
  sufficientCount: number;
  generatedAt: string;
  results: LocationAnalysisResult[];
  ingestionErrors: RowError[];
}

export interface AnalysisOptions {
  defaultThreshold?: number;
  defaultPackSize?: number;
  criticalThresholdPercentage?: number; // e.g. stock <= 25% of min is critical
}
