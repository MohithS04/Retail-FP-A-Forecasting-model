# Power BI Implementation Guide: Retail FP&A Forecasting

This guide details the steps to build the 3-page Executive Dashboard requested for this model using the generated `Retail_FPA_Model.xlsx`.

## 1. Data Connection & Power Query Steps

1. **Load the Excel File**: Open Power BI Desktop -> Get Data -> Excel Workbook -> select `Retail_FPA_Model.xlsx`.
2. **Select Sheets**: Import the `P&L Summary`, `Historical Data`, and `Revenue Forecast` sheets.

### Data Modeling (Star Schema)
To build the required visuals efficiently, we need to unpivot the data and create a Date table.
*   **Calendar Table**: Create a new calculated table for dates spanning 2021-2025.
    ```dax
    Calendar = CALENDAR(DATE(2021,1,1), DATE(2025,12,31))
    ```
*   **Fact_Actuals**: In Power Query, take the `Historical Data` table.
    *   Promote headers, select the "Category" and "Unit" columns, right-click and **Unpivot Other Columns**.
    *   Rename the resulting columns "Date" and "Amount". Ensure Date is cast as a Date type and Amount as Currency.
*   **Fact_Forecast**: Do the same unpivot operation for `Revenue Forecast`, `COGS & Gross Margin`, and `OpEx` sheets, then append them into a single `Fact_Forecast` table.
*   **Dimension Tables**:
    *   Create a `Dim_Unit` table with unique values from the Unit column.
    *   Create a `Dim_P&L_Category` table with a custom sort order (Revenue = 1, COGS = 2, Margin = 3, Payroll = 4... EBITDA = 10) so matrices sort correctly.

Associate the `Date`, `Unit`, and `Category` columns between Facts and Dims to create a valid Star Schema.

## 2. Key DAX Measures

```dax
Total Revenue Actual = CALCULATE(SUM(Fact_Actuals[Amount]), Fact_Actuals[Category] = "Revenue")
Total Revenue Forecast = CALCULATE(SUM(Fact_Forecast[Amount]), Fact_Forecast[Category] = "Revenue")

Total COGS Actual = CALCULATE(SUM(Fact_Actuals[Amount]), Fact_Actuals[Category] = "COGS")
Gross Margin Actual = [Total Revenue Actual] - [Total COGS Actual]
Gross Margin % = DIVIDE([Gross Margin Actual], [Total Revenue Actual], 0)

Total OpEx Actual = CALCULATE(SUM(Fact_Actuals[Amount]), Fact_Actuals[Category] IN {"Payroll", "Rent", "Marketing", "Utilities", "D&A", "Other OpEx"})
EBITDA Actual = [Gross Margin Actual] - [Total OpEx Actual]
EBITDA Margin % = DIVIDE([EBITDA Actual], [Total Revenue Actual], 0)

// Variance Analysis
Revenue Variance $ = [Total Revenue Actual] - [Total Revenue Forecast]
Revenue Variance % = DIVIDE([Revenue Variance $], [Total Revenue Forecast])
```

## 3. Dashboard Design (3 Pages)

### Page 1: Executive Summary
*   **Color Palette**: Navy blue `#112244`, White, Gold `#D4AF37`, and Teal for positive variance.
*   **KPI Cards (Top row)**: `Total Revenue Actual`, `Gross Margin %`, `EBITDA Actual`. Add smaller text below the value for "% vs Budget".
*   **Budget vs. Actual Waterfall Chart**:
    *   Values: P&L Categories (Revenue (+), COGS (-), OpEx (-), EBITDA (Total))
    *   Breakdown: Measure showing the `$ Variance` between Actual and Budget.
*   **Revenue Trend Line (24-month rolling)**:
    *   Line Chart: X-Axis = `Calendar[Year-Month]`.
    *   Lines: `Total Revenue Actual` (Solid Blue), `Total Revenue Forecast` (Dashed Gold).

### Page 2: Unit-Level Performance
*   **Matrix Visual**:
    *   Rows: `Dim_Unit[Unit Name]`.
    *   Columns: `Dim_P&L_Category[Category]`.
    *   Values: `Actual Amount` measure.
    *   Apply conditional formatting (Color Scales) on the Gross Margin % and EBITDA Margin % columns across units.
*   **Scatter Plot**:
    *   X-Axis: `Total Revenue Actual`.
    *   Y-Axis: `EBITDA Margin %`.
    *   Details/Bubble: `Dim_Unit[Unit Name]`.
    *   Size: `Total Revenue Actual`.
*   **Slicer Panel**: Insert a slicer strip on the left-hand side for `Month/Year`.

### Page 3: Variance Deep Dive
*   **Tornado Chart (Clustered Bar)**:
    *   Y-Axis: `Dim_P&L_Category`.
    *   X-Axis: `$ Variance (Actual - Budget)`.
    *   Sort absolute greatest variance to the top.
*   **Trend of Variance**:
    *   Line Chart showing the Net Income Variance month-over-month.
*   **Drill-through capability**: Turn on "Drill-through" on Page 2 and 3 so clicking a specific unit's variance transports the user to a filtered view of just that unit's P&L stack.
