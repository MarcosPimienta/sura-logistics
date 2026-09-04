import express, { Request, Response } from 'express';
import multer from 'multer';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseExcelWorkbook } from '../core/parser.js';
import { analyzeInventory } from '../core/analyzer.js';
import { generateExcelReport, generateCsvReport } from '../core/exporter.js';
import { AnalysisSummary } from '../core/types.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 25 * 1024 * 1024 }, // 25MB limit
});

app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.resolve(__dirname, '../../public')));

// Health check
app.get('/api/health', (_req: Request, res: Response) => {
  res.json({ status: 'ok', time: new Date().toISOString() });
});

// Upload and analyze spreadsheet
app.post('/api/analyze', upload.single('file'), async (req: Request, res: Response) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No Excel file provided in request' });
    }

    const defaultThreshold = req.body.defaultThreshold ? Number(req.body.defaultThreshold) : undefined;
    const defaultPackSize = req.body.defaultPackSize ? Number(req.body.defaultPackSize) : undefined;

    const ingestion = await parseExcelWorkbook(req.file.buffer);
    const summary = analyzeInventory(ingestion.records, ingestion.errors, {
      defaultThreshold,
      defaultPackSize,
    });

    return res.json({
      success: true,
      filename: req.file.originalname,
      summary,
    });
  } catch (error: unknown) {
    const message = error instanceof Error ? error.message : 'Unknown error during analysis';
    return res.status(422).json({ error: message });
  }
});

// Export Excel Manifest
app.post('/api/export/excel', async (req: Request, res: Response) => {
  try {
    const summary: AnalysisSummary = req.body.summary;
    if (!summary || !summary.results) {
      return res.status(400).json({ error: 'Valid summary payload is required for export' });
    }

    const buffer = await generateExcelReport(summary);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="sura-replenishment-manifest.xlsx"');
    return res.send(buffer);
  } catch (error: unknown) {
    const message = error instanceof Error ? error.message : 'Failed to generate Excel export';
    return res.status(500).json({ error: message });
  }
});

// Export CSV Manifest
app.post('/api/export/csv', (req: Request, res: Response) => {
  try {
    const summary: AnalysisSummary = req.body.summary;
    if (!summary || !summary.results) {
      return res.status(400).json({ error: 'Valid summary payload is required for export' });
    }

    const csv = generateCsvReport(summary);
    res.setHeader('Content-Type', 'text/csv');
    res.setHeader('Content-Disposition', 'attachment; filename="sura-replenishment-manifest.csv"');
    return res.send(csv);
  } catch (error: unknown) {
    const message = error instanceof Error ? error.message : 'Failed to generate CSV export';
    return res.status(500).json({ error: message });
  }
});

// Catch-all route to serve the SPA
app.get('*', (_req: Request, res: Response) => {
  res.sendFile(path.resolve(__dirname, '../../public/index.html'));
});

if (process.env.NODE_ENV !== 'test') {
  app.listen(PORT, () => {
    console.log(`\x1b[1;32m✓ Sura Logistics Web Server running at http://localhost:${PORT}\x1b[0m`);
  });
}

export default app;
