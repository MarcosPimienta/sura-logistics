## ADDED Requirements

### Requirement: Excel File Upload and Ingestion
The system SHALL provide an interface and API endpoint to accept Excel files (.xlsx and .xls) containing location inventory data.

#### Scenario: Successful file upload
- **WHEN** a valid Excel file is uploaded via the ingestion interface
- **THEN** the system parses the workbook, validates structure, and confirms successful upload

#### Scenario: Unsupported file type rejected
- **WHEN** a non-Excel file (e.g., .txt or .pdf) is uploaded
- **THEN** the system rejects the file with a 400 Bad Request error and an explanatory message

### Requirement: Location Inventory Schema Validation
The system SHALL validate that each record in the ingested spreadsheet includes required columns: Location Identifier, Location Name, Current Coffee Bag Stock, and Minimum Threshold (or Average Daily Consumption).

#### Scenario: Missing required columns
- **WHEN** an uploaded spreadsheet is missing required columns
- **THEN** the system rejects the processing job and returns a list of missing required column names

#### Scenario: Malformed row data
- **WHEN** a row contains non-numeric values for stock or negative quantities
- **THEN** the system flags the row with a validation error indicating the row number and column issue
