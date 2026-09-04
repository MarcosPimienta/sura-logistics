# Sura Logistics — Coffee Bag Demand Analyzer

Automated spreadsheet ingestion and location demand analysis engine to determine which operating locations require coffee bag replenishments.

---

## 🚀 Quick Start

### 1. Web Dashboard
Start the local server and open the web dashboard:
```bash
npm start
```
Visit **[http://localhost:3000](http://localhost:3000)** in your browser:
- **Drag & Drop** any Excel workbook (`.xlsx`, `.xls`).
- **Interactive KPI Cards** displaying critical shortages, warning deficits, and total bags needed.
- **Search & Filter** locations by status (`CRITICAL`, `WARNING`, `SUFFICIENT`), location name, or region.
- **Export Replenishment Manifest** directly to Excel (`.xlsx`) or CSV.
- Click **"🧪 Load Sample Data"** in the top bar to preview with built-in data.

### 2. Command-Line Interface (CLI)
You can also run batch analysis directly from the terminal:
```bash
# Basic terminal summary table
npm run analyze -- sample-inventory.xlsx

# Export formatted Excel manifest alongside file
npm run analyze -- sample-inventory.xlsx --export-excel

# Export CSV manifest
npm run analyze -- sample-inventory.xlsx --export-csv
```

### 3. Run Automated Tests
```bash
npm test
```

---

## 📊 Spreadsheet Format & Column Mapping

The system automatically detects and maps column headers, supporting English and Spanish variations:

| Field | Required | Recognized Header Aliases |
| :--- | :---: | :--- |
| **Location ID** | Optional* | `Location ID`, `ID`, `Code`, `Store #`, `Codigo`, `Sucursal ID` |
| **Location Name** | Required* | `Location Name`, `Location`, `Store Name`, `Branch`, `Sucursal`, `Sede`, `Nombre` |
| **Current Stock** | Required | `Current Stock`, `Stock`, `Coffee Bags`, `Bags`, `Existencias`, `Bolsas`, `On Hand` |
| **Min Threshold** | Optional | `Min Threshold`, `Threshold`, `Minimum Stock`, `Min`, `Stock Minimo`, `Reorder Level` *(default: 10)* |
| **Pack Size** | Optional | `Pack Size`, `Package Size`, `Caja`, `Paquete`, `Packaging Unit` *(default: 1)* |
| **Daily Burn** | Optional | `Daily Consumption`, `Avg Daily Consumption`, `Consumo Diario` |
| **Region** | Optional | `Region`, `City`, `Zone`, `Ciudad`, `Area` |

*\*Note: If Location ID is omitted, Location Name is used as the unique key, and vice versa.*

---

## ⚙️ Demand & Replenishment Logic

1. **Deficit Calculation**:
   $$\text{Deficit} = \max(0, \text{Minimum Threshold} - \text{Current Stock})$$

2. **Packaging Unit Rounding**:
   If a packaging unit / pack size is specified (e.g. 5 bags per case), replenishment is rounded up to complete packs:
   $$\text{Recommended Replenishment} = \left\lceil \frac{\text{Deficit}}{\text{Pack Size}} \right\rceil \times \text{Pack Size}$$

3. **Urgency Classification**:
   - 🔴 **CRITICAL**: Current stock is $0$ or $\le 30\%$ of the minimum threshold.
   - 🟡 **WARNING**: Current stock is below the minimum threshold ($30\% - 99\%$).
   - 🟢 **SUFFICIENT**: Current stock meets or exceeds minimum threshold.
   - 🔵 **SURPLUS**: Current stock exceeds $150\%$ of the minimum threshold.

---

## 🛠 Spec-Driven Development Workflow (`/opsx`)

This project follows the OpenSpec development lifecycle:
- `/opsx:explore`: Map requirements and investigate problems.
- `/opsx:propose`: Create change proposals, designs, and task checklists.
- `/opsx:apply`: Implement change tasks.
- `/opsx:sync`: Sync delta specifications to main specs.
- `/opsx:archive`: Archive completed change records.
