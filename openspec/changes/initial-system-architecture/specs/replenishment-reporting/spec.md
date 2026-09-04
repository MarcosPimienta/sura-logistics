## ADDED Requirements

### Requirement: Location Replenishment Summary View
The system SHALL present a structured summary listing all analyzed locations, highlighting those requiring coffee bag replenishment with urgency indicators based on deficit severity.

#### Scenario: Viewing replenishment list
- **WHEN** analysis processing completes
- **THEN** the system displays a table of locations sorted by replenishment priority (critical deficit first)

### Requirement: Replenishment Manifest Export
The system SHALL allow users to export the replenishment calculation results as an Excel or CSV file ready for warehouse dispatch.

#### Scenario: Exporting replenishment manifest
- **WHEN** the user triggers export of the analysis result
- **THEN** the system downloads a spreadsheet file containing location details, current stock, threshold, and recommended coffee bag shipment quantity
