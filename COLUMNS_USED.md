# Columns Used in Underwriting Analysis

This document lists all columns used by the analysis scripts (`underwriting_batch.py`, `app.py`, `underwriting_analysis.ipynb`).

## Required Input Columns (from CSV/Excel)

### Core Columns (Required)
1. **`statusDescription`** - Status of the case (e.g., "Completed", "On Hold", "In Progress")
   - Used for: Status count analysis, filtering

2. **`createDateTime`** - Date and time when case was created
   - Used for: Aging calculation, TAT calculation, seasonality analysis

3. **`completedDateTime`** - Date and time when case was completed (can be empty/null)
   - Used for: TAT calculation, completion rate, aging (non-completed cases)

4. **`onHoldReasonDescriptionsHistory`** - History of hold reasons (separated by ', ' or '|')
   - Used for: Case type classification (Straight Through, One Touch, Multi Hold), hold reason analysis

5. **`onHoldDatesHistory`** - History of on-hold dates (separated by ', ' or '|')
   - Used for: Holding time calculation

6. **`offHoldDatesHistory`** - History of off-hold dates (separated by ', ' or '|')
   - Used for: Holding time calculation

### Optional but Important Columns
7. **`requestTypeDescription`** - Type of request
   - Used for: Top 10 request types analysis, drill-down analysis

8. **`writeOutReasonDescriptionsHistory`** or **`writeOutReasonDescription`** - Write-out reasons
   - Used for: Top write-out reasons per TAT bucket analysis
   - Falls back to `writeOutReasonDescription` if history column not available

9. **`numberOfLocations`** - Number of locations
   - Used for: Location analysis (median TAT, touches, holding time by location count)

10. **`NumberOfVehicles`** - Number of vehicles
    - Used for: Vehicle analysis (median TAT, touches, holding time by vehicle count), correlation analysis

11. **`requestId`** - Unique identifier for each case
    - Used for: Counting cases in aggregations

12. **`onHoldReasonDescription`** - Current hold reason (optional)
    - Used for: Additional hold reason analysis

## Calculated Columns (Created by Script)

These columns are **generated** during processing and added to the dataframe:

1. **`Aging_Days`** - Number of days since creation for non-completed cases
   - Formula: `Current Date - createDateTime` (only for cases where `completedDateTime` is null)

2. **`CaseType`** - Classification: "Straight Through", "One Touch", or "Multi Hold (N touches)"
   - Based on: Count of items in `onHoldReasonDescriptionsHistory`

3. **`HoldingTimes`** - List of individual holding periods (in days) for each case
   - Calculated from: `onHoldDatesHistory` and `offHoldDatesHistory`

4. **`TotalHoldingTime`** - Sum of all holding periods for a case (in days)
   - Formula: `sum(HoldingTimes)`

5. **`NumberOfTouches`** - Number of times case was placed on hold
   - Formula: `len(HoldingTimes)`

6. **`TAT_Days`** - Turnaround Time in days (only for completed cases)
   - Formula: `completedDateTime - createDateTime`

7. **`TAT_Bucket`** - TAT performance category
   - Values: "0-5 days", "5-7 days", "7+ days"
   - Based on: `TAT_Days`

8. **`Month`** - Period object for month/year
   - Extracted from: `createDateTime`

9. **`Month_Str`** - String representation of month (e.g., "2025-01")
   - Extracted from: `createDateTime`

10. **`HoldCategory`** - Simplified hold category
    - Values: "Straight Through", "Single Hold", "Multi Hold"
    - Derived from: `CaseType`

## Column Usage Summary

### For Status Analysis
- `statusDescription`

### For Aging Analysis
- `createDateTime`
- `completedDateTime`
- → Calculates: `Aging_Days`

### For Case Type Classification
- `onHoldReasonDescriptionsHistory`
- → Calculates: `CaseType`, `HoldCategory`, `NumberOfTouches`

### For Holding Time Analysis
- `onHoldDatesHistory`
- `offHoldDatesHistory`
- → Calculates: `HoldingTimes`, `TotalHoldingTime`

### For TAT Analysis
- `createDateTime`
- `completedDateTime`
- → Calculates: `TAT_Days`, `TAT_Bucket`

### For Seasonality Analysis
- `createDateTime` (for month extraction)
- `completedDateTime` (for completion rate)
- `TotalHoldingTime` (for holding rate)
- `CaseType` (for straight-through vs multi-hold seasonality)

### For Location/Vehicle Analysis
- `numberOfLocations`
- `NumberOfVehicles`
- `TAT_Days`
- `NumberOfTouches`
- `TotalHoldingTime`

### For Drill-Down Analysis
- `requestTypeDescription` (top 10 request types)
- `onHoldReasonDescriptionsHistory` (top 10 hold reasons)
- `writeOutReasonDescriptionsHistory` or `writeOutReasonDescription` (top write-out reasons)
- `TAT_Bucket` (for filtering)

## Notes

- **Separator Handling**: Columns with history data (`onHoldReasonDescriptionsHistory`, `onHoldDatesHistory`, `offHoldDatesHistory`, `writeOutReasonDescriptionsHistory`) can contain values separated by:
  - `, ` (comma-space)
  - `|` (pipe)
  - The script automatically detects and handles both formats

- **Missing Columns**: The script handles missing columns gracefully:
  - If a column doesn't exist, related analyses are skipped
  - Error messages are printed for missing critical columns

- **Date Parsing**: All date columns are parsed with error handling:
  - Invalid dates become `NaT` (Not a Time)
  - Only valid dates are used in calculations

