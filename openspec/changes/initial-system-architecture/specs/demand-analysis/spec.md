## ADDED Requirements

### Requirement: Location Coffee Bag Shortage Identification
The system SHALL compare current coffee bag stock against the minimum threshold for each location to identify deficit locations.

#### Scenario: Location with stock below threshold
- **WHEN** a location has 5 coffee bags in stock and a minimum threshold of 20
- **THEN** the system marks the location as requiring replenishment with a deficit of 15 bags

#### Scenario: Location with adequate stock
- **WHEN** a location has 25 coffee bags in stock and a minimum threshold of 20
- **THEN** the system marks the location status as sufficient with zero replenishment needed

### Requirement: Replenishment Quantity Calculation
The system SHALL calculate the recommended quantity of coffee bags to dispatch, rounding up to whole pack/case sizes if a packaging unit size is configured.

#### Scenario: Standard replenishment calculation
- **WHEN** deficit calculation completes for a location with pack size of 1
- **THEN** the recommended dispatch quantity equals the exact deficit amount

#### Scenario: Pack-size rounded replenishment calculation
- **WHEN** a deficit is 14 bags and the packaging unit rule specifies packs of 5
- **THEN** the system recommends dispatching 15 bags (3 packs) to meet or exceed minimum threshold
