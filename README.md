# Underwriting Issuance Dashboard

A comprehensive Streamlit-based dashboard for analyzing underwriting issuance data with advanced analytics and reporting capabilities.

## Features

This dashboard provides detailed analysis of underwriting issuance data including:

1. **Status Count Analysis** - Count of cases by status description
2. **Aging Analysis** - Days since creation for non-completed cases
3. **Case Type Classification** - Straight-through, one-touch, and multi-hold classification
4. **Holding Time Analysis** - Time spent on hold for each case
5. **Turnaround Time (TAT) Analysis** - Time to complete cases
6. **TAT Buckets** - Categorization of TAT into 0-5, 5-7, and 7+ days
7. **Drill-Down Analysis** - Detailed analysis by TAT bucket or number of touches
8. **Seasonality Analysis** - Monthly trends in completion and holding rates
9. **Location and Vehicle Analysis** - Impact of locations and vehicles on processing time
10. **Excel Export** - Comprehensive report generation

## Installation

1. Install required packages:
```bash
pip install -r requirements.txt
```

2. Run the Streamlit app:
```bash
streamlit run app.py
```

## Data Requirements

The dashboard accepts CSV or Excel files with the following key columns:

- `statusDescription` - Status of the case
- `createDateTime` - Date and time when case was created
- `completedDateTime` - Date and time when case was completed (can be empty)
- `onHoldReasonDescriptionsHistory` - History of hold reasons (separated by ', ' or '|')
- `onHoldDatesHistory` - History of on-hold dates (separated by ', ' or '|')
- `offHoldDatesHistory` - History of off-hold dates (separated by ', ' or '|')
- `requestTypeDescription` - Type of request
- `onHoldReasonDescription` - Current hold reason
- `numberOfLocations` - Number of locations
- `NumberOfVehicles` - Number of vehicles

## Field Calculations

### 1. Status Count (`statusDescription`)

**Calculation:**
- Counts the number of cases for each unique value in the `statusDescription` column
- Uses `value_counts()` to aggregate and display distribution

**Purpose:** Understand the distribution of cases across different statuses (e.g., "Completed", "On Hold", "In Progress")

---

### 2. Aging Days (`Aging_Days`)

**Calculation:**
```python
Aging_Days = Current Date - createDateTime
```
- Only calculated for cases where `completedDateTime` is null/empty
- Represents the number of days a case has been open without completion
- Uses `datetime.now()` as the reference point for current date

**Purpose:** Identify cases that have been pending for extended periods and may need attention

---

### 3. Case Type Classification (`CaseType`)

**Calculation:**
The classification is based on parsing `onHoldReasonDescriptionsHistory`:

- **Straight Through**: If `onHoldReasonDescriptionsHistory` is blank or empty
  - No holds were placed on the case
  - Case processed without interruption
  - Number of touches = 0

- **One Touch**: If `onHoldReasonDescriptionsHistory` contains exactly one hold reason
  - Case was placed on hold once
  - Single touch point in the process
  - Number of touches = 1

- **Multi Hold (N touches)**: If `onHoldReasonDescriptionsHistory` contains multiple hold reasons
  - Number of touches = count of separated values
  - Values can be separated by ', ' (comma-space) or '|' (pipe)
  - Example: "Reason1, Reason2" or "Reason1|Reason2" = 2 touches
  - Example: "Reason1, Reason2, Reason3" = 3 touches

**Parsing Logic:**
- First checks for '|' separator
- If not found, checks for ', ' separator
- Splits the string and counts non-empty values
- If blank/empty → Straight Through (0 touches)

**Purpose:** Categorize cases based on process complexity and identify bottlenecks

---

### 4. Holding Time (`HoldingTimes`, `TotalHoldingTime`)

**Calculation:**
For each case, holding times are calculated by pairing `onHoldDatesHistory` and `offHoldDatesHistory`:

```python
For each hold period:
    HoldingTime = offHoldDate - onHoldDate (in days)
    
If offHoldDate is missing:
    HoldingTime = Current Date - onHoldDate (in days)

TotalHoldingTime = Sum of all individual holding times
```

**Details:**
- Parses both date columns using the same separator logic (', ' or '|')
- Matches each on-hold date with corresponding off-hold date by index
- If an off-hold date is missing, calculates from on-hold date to current date
- Stores individual holding times in `HoldingTimes` list
- `TotalHoldingTime` is the sum of all holding periods

**Purpose:** Measure the total time cases spend on hold, identifying process delays

---

### 5. Number of Touches (`NumberOfTouches`)

**Calculation:**
```python
NumberOfTouches = Count of items in HoldingTimes list
```

- Directly corresponds to the number of hold periods
- Equals the number of entries in `onHoldReasonDescriptionsHistory` (after parsing)

**Purpose:** Quantify process complexity and identify cases requiring multiple interventions

---

### 6. Turnaround Time - TAT (`TAT_Days`)

**Calculation:**
```python
TAT_Days = completedDateTime - createDateTime (in days)
```

- Only calculated for cases where `completedDateTime` is not null
- Represents the total time from case creation to completion
- Measured in days (can include fractional days)

**Purpose:** Measure overall case processing efficiency and identify performance trends

---

### 7. TAT Buckets (`TAT_Bucket`)

**Calculation:**
```python
If TAT_Days <= 5:
    TAT_Bucket = "0-5 days"
Else if TAT_Days <= 7:
    TAT_Bucket = "5-7 days"
Else:
    TAT_Bucket = "7+ days"
```

**Purpose:** Categorize cases into performance tiers for easier analysis and reporting

---

### 8. Median TAT

**Calculation:**
```python
Median_TAT = Median of all TAT_Days values (excluding nulls)
```

- Uses `median()` function on non-null TAT values
- Provides a robust measure of central tendency (less affected by outliers than mean)

**Purpose:** Understand typical case processing time

---

### 9. Straight Through and Multi-Hold Counts

**Calculation:**
```python
Straight_Through_Count = Count of cases where CaseType == "Straight Through"
Multi_Hold_Count = Count of cases where CaseType contains "Multi Hold"
One_Touch_Count = Count of cases where CaseType == "One Touch"
```

**Purpose:** Understand the distribution of case complexity

---

### 10. Median TAT by Case Type

**Calculation:**
```python
For each CaseType:
    Median_TAT = Median of TAT_Days for that CaseType
```

- Groups cases by `CaseType`
- Calculates median TAT for each group
- Helps identify if case complexity affects processing time

**Purpose:** Compare processing efficiency across different case types

---

### 10b. Median TAT by Single Hold vs Multi-Hold (by TAT Buckets)

**Calculation:**
```python
# Create HoldCategory from CaseType
HoldCategory = 'Straight Through' if CaseType == 'Straight Through'
             = 'Single Hold' if CaseType == 'One Touch'
             = 'Multi Hold' if CaseType contains 'Multi Hold'

# Calculate median TAT by HoldCategory and TAT_Bucket
For each combination of (HoldCategory, TAT_Bucket):
    Median_TAT = Median of TAT_Days for that combination

# Also calculate overall median by HoldCategory
For each HoldCategory:
    Median_TAT = Median of TAT_Days for that HoldCategory
```

**Details:**
- Simplifies case classification into three categories: Straight Through, Single Hold, and Multi Hold
- Analyzes median TAT across both hold categories and TAT performance buckets
- Helps identify if hold complexity affects TAT performance within different time ranges

**Purpose:** Understand how case complexity (number of holds) relates to TAT performance across different time buckets

---

### 11. Top 10 Request Types and Hold Reasons (Drill-Down Analysis)

**Calculation:**
```python
# Filter data based on drill-down selection
If drill-down by "Straight Through":
    Drill_Data = All cases where CaseType == "Straight Through"
Else if drill-down by "Multi Hold":
    Drill_Data = All cases where CaseType contains "Multi Hold"
Else if drill-down by "TAT Bucket":
    Drill_Data = All cases where TAT_Bucket == selected_bucket
Else if drill-down by "Number of Touches":
    Drill_Data = All cases where NumberOfTouches == selected_touches

# Calculate top 10 for filtered data
Top_10_Request_Types = Drill_Data['requestTypeDescription'].value_counts().head(10)

# For Hold Reasons:
# Parse and flatten all hold reasons from filtered data
All_Hold_Reasons = []
For each case in Drill_Data:
    Reasons = Parse(onHoldReasonDescriptionsHistory)
    All_Hold_Reasons.extend(Reasons)
Top_10_Hold_Reasons = pd.Series(All_Hold_Reasons).value_counts().head(10)
```

**Drill-Down Options:**
1. **Straight Through**: Analyze only cases that went through without any holds
2. **Multi Hold**: Analyze only cases that had multiple holds
3. **By TAT Bucket**: Analyze cases within a specific TAT performance range (0-5, 5-7, or 7+ days)
4. **By Number of Touches**: Analyze cases with a specific number of touch points

**For Hold Reasons:**
- Parses all `onHoldReasonDescriptionsHistory` values in the filtered dataset
- Flattens separated values into individual reasons
- Counts frequency of each reason
- Selects top 10 most common

**Purpose:** Identify most common request types and hold reasons for targeted process improvement within specific case categories

---

### 12. Median of Highest and Lowest Holding Times

**Calculation:**
```python
# For each case, extract all individual holding times
All_Holding_Times = Flatten all individual holding times from all cases

# Overall statistics
Highest_Holding_Time_Overall = Max(All_Holding_Times)
Lowest_Holding_Time_Overall = Min(All_Holding_Times)
Median_Holding_Time_All = Median(All_Holding_Times)

# Per-case statistics
For each case:
    Highest_Per_Case = Max(HoldingTimes for that case)
    Lowest_Per_Case = Min(HoldingTimes for that case)

Median_of_Highest_Holding_Time = Median(Highest_Per_Case across all cases)
Median_of_Lowest_Holding_Time = Median(Lowest_Per_Case across all cases)
Average_Touches = Mean(NumberOfTouches)
```

**Details:**
- **Median Holding Time (All)**: Median of all individual holding periods across all cases
- **Median of Highest Holding Time**: For each case, finds the longest holding period, then calculates the median of these maximum values
- **Median of Lowest Holding Time**: For each case, finds the shortest holding period, then calculates the median of these minimum values
- Provides insight into both overall holding time distribution and per-case holding patterns

**Purpose:** Understand the range and distribution of holding times across all cases, and identify patterns in per-case holding behavior

---

### 13. Seasonality Analysis (Monthly)

**Calculations:**

**Monthly Completion Rate:**
```python
For each month:
    Total_Cases = Count of cases created in that month
    Completed_Cases = Count of cases with non-null completedDateTime
    Completion_Rate = (Completed_Cases / Total_Cases) * 100
```

**Monthly Holding Rate:**
```python
For each month:
    Total_Cases = Count of cases created in that month
    Cases_With_Holds = Count of cases with TotalHoldingTime > 0
    Holding_Rate = (Cases_With_Holds / Total_Cases) * 100
```

**Additional Monthly Metrics:**
- Median TAT per month
- Median Holding Time per month
- Average Number of Touches per month

**Purpose:** Identify seasonal patterns, trends, and performance variations throughout the year

---

### 14. Analysis by Location and Number of Vehicles

**By Number of Locations:**
```python
For each unique numberOfLocations:
    Median_TAT = Median TAT_Days
    Avg_Touches = Mean NumberOfTouches
    Median_Holding_Time = Median TotalHoldingTime
    Count = Number of cases
```

**By Number of Vehicles:**
```python
For each unique NumberOfVehicles:
    Median_TAT = Median TAT_Days
    Avg_Touches = Mean NumberOfTouches
    Median_Holding_Time = Median TotalHoldingTime
    Count = Number of cases
```

**Correlation Analysis:**
- Scatter plots with trend lines (OLS regression)
- Analyzes relationship between NumberOfVehicles and:
  - NumberOfTouches
  - TotalHoldingTime

**Purpose:** Understand if case complexity (locations/vehicles) impacts processing time and number of touches

---

## Data Parsing Logic

The dashboard handles data that may contain separated values in single cells:

### Separator Detection:
1. First checks for '|' (pipe) separator
2. If not found, checks for ', ' (comma-space) separator
3. If neither found, treats entire value as single item

### Example Parsing:
- Input: `"Reason1, Reason2, Reason3"` → Output: `["Reason1", "Reason2", "Reason3"]`
- Input: `"Reason1|Reason2"` → Output: `["Reason1", "Reason2"]`
- Input: `"Reason1"` → Output: `["Reason1"]`
- Input: `""` or `NaN` → Output: `[]`

### Date Parsing:
- Uses `pd.to_datetime()` with error handling
- Invalid dates are coerced to `NaT` (Not a Time)
- Only valid dates are used in calculations

---

## Excel Export

The dashboard generates a comprehensive Excel report with the following sheets:

1. **Summary** - Key metrics and statistics including:
   - Total Cases, Completed Cases
   - Median TAT and Median Aging
   - Straight Through, One Touch, and Multi Hold case counts
   - Median Holding Time statistics (all, highest, lowest)
   - Average Number of Touches
   - Highest and Lowest Holding Times overall

2. **Status_Counts** - Distribution of cases by status description

3. **Case_Type_Classification** - Breakdown by case type (Straight Through, One Touch, Multi Hold)

4. **TAT_Buckets** - Distribution across TAT buckets (0-5, 5-7, 7+ days)

5. **Median_TAT_by_CaseType** - Median TAT for each case type

6. **Median_TAT_by_Hold_Bucket** - Median TAT by Hold Category (Straight Through, Single Hold, Multi Hold) and TAT Bucket (pivot table)

7. **Median_TAT_by_HoldCategory** - Overall median TAT by Hold Category

8. **Monthly_Statistics** - Seasonality analysis including:
   - Total Cases and Completed Cases per month
   - Completion Rate and Holding Rate per month
   - Median TAT, Median Holding Time, and Average Touches per month

9. **Holding_Time_Stats** - Detailed holding time statistics:
   - Median Holding Time (All)
   - Median of Highest Holding Time
   - Median of Lowest Holding Time
   - Highest and Lowest Holding Times overall
   - Average Number of Touches

10. **Location_Analysis** - Analysis by number of locations (if available):
    - Median TAT, Average Touches, Median Holding Time, and Count for each location count

11. **Vehicle_Analysis** - Analysis by number of vehicles (if available):
    - Median TAT, Average Touches, Median Holding Time, and Count for each vehicle count

12. **Original_Data** - Complete original dataset with all original columns before any processing or calculations

13. **Processed_Data** - Complete dataset with all calculated fields:
    - Original columns plus: CaseType, HoldCategory, Aging_Days, TAT_Days, TAT_Bucket, TotalHoldingTime, NumberOfTouches, Month_Str

---

## Usage

1. **Upload Data:**
   - Choose "Upload File" to upload a CSV or Excel file
   - Or choose "Data Folder" to select from existing files in the `data/` folder

2. **View Analysis:**
   - Navigate through different sections using the scrollable dashboard
   - Interactive charts allow zooming and filtering

3. **Export Results:**
   - Click "Generate Excel Report" to create a comprehensive report
   - Download the generated Excel file

---

## Notes

- All date calculations use the current system date/time as reference
- Missing or invalid dates are handled gracefully (excluded from calculations)
- The dashboard automatically detects and handles different data formats
- Large datasets are processed efficiently using pandas operations

---

## Technical Stack

- **Streamlit** - Web application framework
- **Pandas** - Data manipulation and analysis
- **NumPy** - Numerical computations
- **Plotly** - Interactive visualizations
- **OpenPyXL** - Excel file generation

---

## Author

Underwriting Issuance Dashboard - Analytics and Reporting Tool
