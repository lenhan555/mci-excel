---
title: Excel & Power BI Complete Reference Knowledge
description: Comprehensive reference extracted from 5 PDF course files covering Excel Formulas, PivotTables, Data Visualization, Pro Tips, Power Query, Power Pivot, and DAX.
tags: [excel, formulas, pivottable, data-visualization, power-query, power-pivot, dax, powerbi]
source_files:
  - Excel_Formulas_&_Functions.pdf
  - Excel_PivotTable_Slides.pdf
  - Excel_Data_Visualization.pdf
  - Excel+Pro+Tips+for+Power+Users.pdf
  - PowerQuery_PowerPivot_DAX.pdf
created: 2026-09-21
---

# Excel & Power BI Complete Reference Knowledge

---

## FILE 1: Excel Formulas & Functions

### Course Outline (9 Sections)

1. Excel Formulas 101 — Syntax, reference types, errors, auditing, shortcuts
2. Conditionals & Logical Operators
3. Basic Statistical Functions
4. Lookup/Reference Functions
5. Text Formulas
6. Date & Time Functions
7. Formula-Based Formatting
8. Array Formulas (expanded to Dynamic Arrays)
9. Extra Bonus Functions

---

### Formulas 101

- **Formula Syntax:** equals sign, arguments, parentheses, commas
- **Reference Types:**
  - Relative: `A1`
  - Mixed: `$A1`, `A$1`
  - Absolute: `$A$1`
- **Common Errors:** `######`, `#NAME?`, `#VALUE!`, `#DIV/0!`, `#REF!`, `#N/A`
- **Auditing Tools:** Trace Precedents, Trace Dependents, Show Formulas, Evaluate Formula, Error Checking
- **Keyboard Shortcuts:** `Ctrl+Arrow`, `Ctrl+Shift+Arrow`, `Ctrl+Home/End`, `Ctrl+PgUp/PgDn`, `F1`, `F2`, `F4`, `F9`
- **Mac Shortcuts:** `Command-T`, `Command-Y`, `Control-U`, `Command-Arrow`, `Command-Shift-Arrow`
- **Alt Key Tips:** `Alt-H-V-V`, `Alt-A-T`, `Alt-M-V`
- **Data Validation Types:** Number Type, Value, List, Date/Time, Text Length, Custom/Formula-Driven

---

### Conditional Statements

| Function | Notes |
|---|---|
| `IF` | Syntax, nested IF |
| `AND` / `OR` | |
| `NOT` / `<>` | Not-equal operator |
| `IFERROR` | |
| `ISBLANK` | IS Statements |
| `ISNUMBER` | |
| `ISTEXT` | |
| `ISERROR` | |
| `ISEVEN` / `ISODD` | |
| `ISLOGICAL` | |
| `ISFORMULA` | |

---

### Basic Statistical Functions

| Function | Notes |
|---|---|
| `COUNT`, `AVERAGE`, `MEDIAN`, `MODE`, `MAX`, `MIN` | Core stats |
| `PERCENTILE` | 25th and 75th |
| `STDEV`, `VAR` | |
| `RANK`, `SMALL`, `LARGE` | |
| `PERCENTRANK` | |
| `RAND()`, `RANDBETWEEN` | |
| `SUMPRODUCT` | With and without filters |
| `COUNTIF`, `SUMIF`, `AVERAGEIF` | Single-criteria |
| `COUNTIFS`, `SUMIFS`, `AVERAGEIFS` | Multi-criteria |

---

### Lookup & Reference Functions

| Function | Notes |
|---|---|
| Named Ranges | |
| `VLOOKUP` | Full syntax, Laws of Lookups |
| `HLOOKUP` | |
| `ROW`, `ROWS` | |
| `COLUMN`, `COLUMNS` | |
| `INDEX` | |
| `MATCH` | |
| `INDEX/MATCH` | Combined 2-way lookup |
| `XLOOKUP` | Office 365 only; comparison vs VLOOKUP |
| `CHOOSE` | Including array constant use |
| `OFFSET` | |

---

### Text Functions

| Function | Notes |
|---|---|
| `TRIM`, `UPPER`, `LOWER`, `PROPER` | |
| `CONCATENATE` | And `&` operator |
| `LEFT`, `MID`, `RIGHT`, `LEN` | |
| `TEXT`, `VALUE` | |
| `SEARCH` | Case-insensitive |
| `FIND` | Case-sensitive |
| `IF(ISNUMBER(SEARCH(...)))` | Pattern for text classification |

---

### Date & Time Functions

- **Date Values:** Serial numbers from 1/1/1900
- **Date Formatting Codes:** `d`, `dd`, `ddd`, `dddd`, `m`, `mm`, `mmm`, `mmmm`, `yy`, `yyyy`
- **Fill Series Options:** Copy Cells, Fill Days, Fill Weekdays, Fill Months, Fill Years

| Function | Notes |
|---|---|
| `TODAY()`, `NOW()` | |
| `YEAR`, `MONTH`, `DAY` | Serialization |
| `HOUR`, `MINUTE`, `SECOND` | |
| `EOMONTH` | |
| `YEARFRAC` | Basis: 0=US30/360, 1=actual/actual, 2=actual/360, 3=actual/365, 4=European30/360 |
| `WEEKDAY` | |
| `WORKDAY`, `NETWORKDAYS` | |
| `DATEDIF` | Units: D, M, Y, MD, YD, YM |
| `DATEVALUE`, `TIMEVALUE` | |

---

### Formula-Based Formatting

- Creating, editing, and managing formula-driven conditional formatting rules

---

### Dynamic Array Formulas (Office 365)

| Function | Notes |
|---|---|
| `SORT()` | Sort by one or multiple columns |
| `SORTBY()` | Sort by external array |
| `FILTER()` | AND conditions (multiply); OR conditions (sum) |
| `UNIQUE()` | `by_col`, `exactly_once` options |
| `SEQUENCE()` | rows, columns, start, step |
| `RANDARRAY()` | rows, columns, min, max, integer |
| `FREQUENCY` | Now dynamic |
| `TRANSPOSE` | Now dynamic |
| `CHOOSE({1,2},...)` | For joining arrays |
| `LET()` | Variable declaration in formulas |

- **Spill Ranges:** Dynamic Excel / `#SPILL!` errors

---

### Extra Bonus Functions

- `INDIRECT`
- `HYPERLINK`
- `WEBSERVICE`
- `FILTERXML`

---

## FILE 2: Excel PivotTable Slides

### Pivot Table Basics

- **Why PivotTables?** Powerful, Beautiful, Fast, Accurate, Flexible
- **Data Structure:**
  - Good: rectangular, no extra formatting
  - Bad: transposed, calculated fields in source
- **Inserting:** Insert PivotTable, Recommended PivotTables
- **Field List Areas:** Filters, Rows, Columns, Values; Layout options
- **Tabs:** Analyze & Design
- **Operations:** Selecting, Clearing, Moving, Refreshing, Change Data Source
- **How Pivots Work:** Detect coordinates → Apply arithmetic → Display result

---

### Formatting

- **Number Formatting:** Right-click column header
- **Table Styles:** Right-click to set default, customize
- **Table Layouts:**
  - Compact Form
  - Outline Form (recommended)
  - Tabular Form (non-repeating and repeating)
- **Conditional Formatting:** Text/Value-based, Data Bars, Color Scales, Icon Sets, Formula-Based Rules

---

### Sorting, Filtering & Grouping

| Feature | Notes |
|---|---|
| Sorting & Filtering | More Sort Options, Label Filters, Value Filters, Manual Selections |
| Grouping Data | Right-click → Group → creates new grouped field |
| Slicers | Pretty filter |
| Timelines | Date-specific filter |
| Report Filter Pages | Show Report Filter Pages → creates new tabs per filter value |

---

### Calculated Values & Fields

| Feature | Notes |
|---|---|
| Summarize Values By | SUM, COUNT, AVERAGE, MAX, MIN |
| Show Values As — % of Total Column | |
| Show Values As — % of Parent | |
| Show Values As — % Difference (prev year) | |
| Show Values As — Running Total (by year) | |
| Show Values As — Rank (Large to Small) | |
| Show Values As — Index | `(Cell Value × Grand Total) / (Row Total × Column Total)` |
| Calculated Fields | New measures from existing numerical fields |
| Calculated Items | New dimensions from existing dimensions |
| Solve Order | For multiple calculated items |
| List Formulas | New tab summarizing all calculated fields/items |

---

### Pivot Charts

- PivotChart 101 — dynamically tied to PivotTable
- Tabs: Analyze, Design, Format
- Chart Layouts & Styles
- Field Buttons — hide/show via PivotChart Tools → Analyze → Field Buttons
- Adding Slicers to PivotCharts
- Adding Timelines to PivotCharts

---

## FILE 3: Excel Data Visualization

### Key Principles

- Clarity & Simplicity (maximize impact, minimize noise)
- Narrative (tell a story, communicate insights)
- Balance Design & Function
- The Good, The Bad, The Ugly
- **The 3 Key Questions:** data type, what to communicate, who is the end user

---

### Chart Types Reference (16 Types)

| Chart Type | Common Uses | Key Pro Tips |
|---|---|---|
| Bar & Column Charts | Comparing numerical data across categories | Stacked/clustered; custom color-code rules |
| Histograms & Pareto Charts | Distribution of continuous data | Adjust bin size; Pareto for cumulative impact |
| Line Charts | Trends over time | Trendlines (linear/polynomial); combine with column charts |
| Area Charts | Changes in data composition over time | Keep categories <6; data validation for dynamic highlighting |
| Pie & Donut Charts | Proportions totaling 100% | Keep slices <6; donut for multiple series / "race track" |
| Scatter Plots | Correlations/relationships between series | Trendline/line of best fit; correlation ≠ causation |
| Bubble Charts | Three-dimensional scatter (X, Y, size) | Color as 4th dimension; form controls for animation |
| Box & Whisker Charts | Statistical characteristics, outliers | Quartile calculation includes/excludes median option |
| Tree Maps & Sunburst Charts | Hierarchical data with groups/sub-groups | Tree Maps for 1–2 levels; Sunburst for multiple levels |
| Waterfall Charts | Net value after series of +/- contributions | Use sub-totals as checkpoints |
| Funnel Charts | Progress through stages of a funnel | "Percent of total" calculations |
| Radar Charts | 3+ quantitative variables on 2D chart | Normalize to same scale; limit categories |
| Surface & Contour Charts | Optimum combinations across 3D | Avoid wireframe; don't use when heat map suffices |
| Stock Charts | Stock market data (volume, high, low, open, close) | Set axis min/max; switch date to text axis |
| Heat Maps | Trends/relationships using color scales | Custom format to hide cell values `(;;;)`; dynamic via COUNTIF |
| Geospatial/Choropleth Maps | Location-based data | Power Map plug-in; color + size for multiple attributes |

---

## FILE 4: Excel Pro Tips for Power Users

> Note: PDF text was largely non-extractable (image-based slides). Recoverable content below.

- **Sections visible:** `COUNTIFS`, `SUMIFS`, `AVERAGEIFS`
- **Formula snippets recovered:**
  - `=MATCH($B4,$I$2:$I$11,0)`
  - `=MATCH(D$1,$H$1:$M$1,0)`
  - `=MATCH(B6,$G$2:$G$13,0)+1`
- Inferred topics: `COUNTIFS`, `SUMIFS`, `AVERAGEIFS`, `INDEX/MATCH` patterns

---

## FILE 5: Power Query, Power Pivot & DAX

### Power Excel Workflow

```
RAW DATA → POWER QUERY (Get & Transform / ETL) → DATA MODEL → POWER PIVOT & DAX
```

---

### Power Query

#### Data Connection Types
- From File
- From Database
- From Azure
- From Online Services
- From Other Sources

#### Query Editor Components
- Formula Bar (M code)
- Applied Steps
- Data Preview
- Tabs: HOME (general/transformation), TRANSFORM (modify columns), ADD COLUMN (new columns)

#### Data Loading Options
- Table (worksheet)
- Connection Only
- Add to Data Model

#### Basic Table Transformations
- Sort, Change data types, Promote header row
- Keep/remove columns, Keep/remove rows, Remove Duplicates
- Duplicate/move/rename columns

#### Text-Specific Tools
- Split column (delimiter or character count)
- Extract characters
- Merge/concatenate fields
- Format: upper/lower/proper/prefix/suffix, Trim, Clean

#### Number-Specific Tools
- Statistics: sum, min/max, average, count, countdistinct
- Standard, Scientific, Trigonometry
- Information (binary TRUE/FALSE flags: even, odd, positive, negative)

#### Date-Specific Tools
- Age, Date Only, Year/Month/Quarter/Week/Day, Hour/Minute/Second, Earliest/Latest

#### Calendar Tables
- Creating a Basic Calendar Table
- Creating a Rolling Calendar: blank query → literal date → list → To Table → Add Column date parts

#### Query Operations
- Adding Index Column (sequential IDs from 0 or 1)
- Adding Conditional Column (IF/THEN rules)
- Grouping & Aggregating (Group By)
- Pivoting & Unpivoting (vs Transpose)
- Edit, Delete, Rename, Refresh, Duplicate, Merge, Append
- **Merging Queries:** JOIN based on common column (like VLOOKUP; adds columns)
- **Appending Queries:** Stack tables with same structure (adds rows; From Folder option)

#### Power Query Best Practices
- Clear naming (no spaces)
- Shape at source
- Only load needed data

---

### Data Modeling 101

| Concept | Description |
|---|---|
| Data View vs Diagram View | Two views of the Data Model |
| Database Normalization | Reduce redundancy, preserve integrity |
| Data Tables | "Fact" tables |
| Lookup Tables | "Dimension" tables |
| Primary Keys | Unique per row |
| Foreign Keys | Many instances |
| Star Schema | One lookup per dimension |
| Snowflake Schema | Lookup-to-lookup chains |

#### Relationships
- Creating: drag in Diagram View or Create Relationship in Design tab
- **Cardinality:** One-to-Many (valid), Many-to-Many (invalid), One-to-One (inefficient)
- Active vs Inactive (only one active at a time per table pair)
- **Filter Direction:** Always from "one"/lookup side to "many"/data side; filters flow downstream
- Never connect data tables directly — use shared lookup tables only

#### Hierarchies
- Example: Country > State > City
- Drill up/down in PivotTables

---

### Power Pivot & DAX Fundamentals

| Concept | Notes |
|---|---|
| Measures vs Calculated Fields | Terminology varies by Excel version |
| Calculated Columns | Row context, static values, stored in model (uses memory) — good for row properties; bad for aggregation |
| DAX Measures | Filter context, dynamic, placed in Values area only |
| Implicit Measures | Drag field to Values |
| Explicit Measures | Via AutoSum or Measure Dialog Box (with IntelliSense, format options) |
| Filter Context | Pivot coordinates determine each cell's calculation |

**Measure Calculation Steps:** Detect coordinates → Carry filters downstream → Evaluate formula

---

### DAX Syntax

```
TableName[ColumnName]          -- fully qualified column reference
[MeasureName]                  -- measure reference
'Transactions Table'[quantity] -- tables with spaces use single quotes
```

#### DAX Operators

| Type | Operators |
|---|---|
| Arithmetic | `+`, `-`, `*`, `/`, `^` |
| Comparison | `=`, `>`, `<`, `>=`, `<=`, `<>` |
| Text/Logical | `&` (concatenate), `&&` (AND), `\|\|` (OR), `IN` (list OR with curly brackets) |

---

### Common DAX Functions

#### Math & Stats

| Function | Syntax |
|---|---|
| `SUM` | `SUM(<column>)` |
| `AVERAGE` | `AVERAGE(<column>)` |
| `MAX` | `MAX(<column>)` or `MAX(<exp1>, <exp2>)` |
| `MIN` | `MIN(<column>)` or `MIN(<exp1>, <exp2>)` |
| `DIVIDE` | `DIVIDE(<numerator>, <denominator>, <other>)` |
| `COUNT` | `COUNT(<column>)` — numeric cells only |
| `COUNTA` | `COUNTA(<column>)` — all non-empty cells |
| `DISTINCTCOUNT` | `DISTINCTCOUNT(<column>)` |
| `COUNTROWS` | `COUNTROWS(<table>)` |

#### Iterator ("X") Functions

| Function | Notes |
|---|---|
| `SUMX` | `SUMX(<table>, <expression>)` |
| `AVERAGEX` | |
| `MAXX` / `MINX` | |
| `RANKX` | |
| `COUNTX` | |

#### Logical Functions

| Function | Syntax |
|---|---|
| `IF` | `IF(<logical test>, <value_if_true>, <value_if_false>)` |
| `IFERROR` | `IFERROR(value, value_if_error)` |
| `AND` | `AND(<logical1>, <logical2>)` |
| `OR` | `OR(<logical1>, <logical2>)` |
| `NOT` | |
| `SWITCH` | `SWITCH(<expression>, <value1>, <result1>, ..., <else>)` |
| `SWITCH(TRUE(), ...)` | Pattern for Boolean-range conditions |
| `TRUE`, `FALSE` | |

#### Text Functions (DAX)

| Function | Syntax |
|---|---|
| `LEN` | `LEN(<text>)` |
| `CONCATENATE` | `CONCATENATE(<text1>, <text2>)` and `&` |
| `LEFT` / `RIGHT` | `LEFT(<text>, <num_chars>)` |
| `MID` | `MID(<text>, <start_num>, <num_chars>)` |
| `UPPER` / `LOWER` / `PROPER` | |
| `SUBSTITUTE` | `SUBSTITUTE(<text>, <old_text>, <new_text>, <instance>)` |
| `SEARCH` | `SEARCH(<find_text>, <within_text>, <start_num>, <NotFoundValue>)` |
| `FIND`, `REPLACE`, `REPT`, `TRIM`, `UNICHAR`, `FORMAT` | |

#### Filter Functions (DAX)

| Function | Notes |
|---|---|
| `CALCULATE` | `CALCULATE(<expression>, <filter1>, <filter2>, ...)` — overrides PivotTable filter context |
| `FILTER` | `FILTER(<table>, <filter expression>)` — returns filtered table subset |
| `ALL` | Removes filter context; opposite of FILTER |
| `ALLEXCEPT` | |
| `RELATED` | Like VLOOKUP using model relationships; row context only |
| `RELATEDTABLE` | |
| `DISTINCT`, `VALUES` | |
| `EARLIER` / `EARLIEST` | |
| `HASONEVALUE`, `HASONEFILTER`, `ISFILTERED` | |
| `USERELATIONSHIP` | |

#### Date & Time Functions (DAX)

| Function | Notes |
|---|---|
| `DAY` / `MONTH` / `YEAR` | |
| `HOUR` / `MINUTE` / `SECOND` | |
| `TODAY()` / `NOW()` | |
| `WEEKDAY` / `WEEKNUM` | `WEEKDAY(<date>, <type>)` |
| `EOMONTH` | `EOMONTH(<start_date>, <months>)` |
| `DATEDIFF` | `DATEDIFF(<start_date>, <end_date>, <interval>)` |
| `YEARFRAC` | |

#### Time Intelligence Functions (DAX)

| Function | Notes |
|---|---|
| `DATESYTD(Calendar[Date])` | Year-to-Date |
| `DATESQTD(Calendar[Date])` | Quarter-to-Date |
| `DATESMTD(Calendar[Date])` | Month-to-Date |
| `DATEADD(Calendar[Date], -1, MONTH)` | Previous Period |
| `DATESINPERIOD(Calendar[Date], MAX(Calendar[Date]), -10, DAY)` | Rolling Period |

---

### Pro Tips — Disconnected Slicers

1. Create parameter table
2. Add to Data Model (not connected to other tables)
3. Use as slicer
4. Capture selection in a `FILTER`/`CALCULATE` measure

---

### DAX Best Practices

- Write explicit measures for even simple calculations
- Break measures into simple component parts (measure trees)
- Use fully qualified column references; bare measure names for measures

---

### Data Visualization Options from Data Model

- PivotCharts & Conditional Formatting
- CUBE functions (pull values from data model without pivots)
- Power View, Power Map
- Microsoft Power BI

**Power Excel vs Power BI:** Both share the same engine; Power BI adds drag-and-drop reports, custom visuals, and cloud collaboration tools.

---

## Summary: All Excel Functions & Features Across All 5 PDFs

### Logical / Conditional
`IF`, Nested IF, `AND`, `OR`, `NOT`, `IFERROR`, `ISERROR`, `ISBLANK`, `ISNUMBER`, `ISTEXT`, `ISEVEN`, `ISODD`, `ISLOGICAL`, `ISFORMULA`, `SWITCH`, `SWITCH(TRUE())`, `TRUE`, `FALSE`

### Math & Stats
`SUM`, `AVERAGE`, `MEDIAN`, `MODE`, `MAX`, `MIN`, `PERCENTILE`, `STDEV`, `VAR`, `RANK`, `SMALL`, `LARGE`, `PERCENTRANK`, `RAND`, `RANDBETWEEN`, `SUMPRODUCT`, `COUNTIF`, `SUMIF`, `AVERAGEIF`, `COUNTIFS`, `SUMIFS`, `AVERAGEIFS`, `COUNT`, `COUNTA`, `COUNTROWS`, `DISTINCTCOUNT`, `DIVIDE`, `FREQUENCY`

### Lookup & Reference
`VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `INDEX`, `MATCH`, `INDEX/MATCH`, `CHOOSE`, `OFFSET`, `INDIRECT`, `ROW`, `ROWS`, `COLUMN`, `COLUMNS`

### Text
`CONCATENATE`, `LEFT`, `MID`, `RIGHT`, `LEN`, `TEXT`, `VALUE`, `SEARCH`, `FIND`, `TRIM`, `UPPER`, `LOWER`, `PROPER`, `REPLACE`, `REPT`, `SUBSTITUTE`, `UNICHAR`, `FORMAT`, `IF(ISNUMBER(SEARCH()))` pattern

### Date & Time
`TODAY`, `NOW`, `DATEVALUE`, `TIMEVALUE`, `YEAR`, `MONTH`, `DAY`, `HOUR`, `MINUTE`, `SECOND`, `EOMONTH`, `YEARFRAC`, `WEEKDAY`, `WORKDAY`, `NETWORKDAYS`, `DATEDIF`/`DATEDIFF`

### Dynamic Arrays (Office 365)
`SORT`, `SORTBY`, `FILTER`, `UNIQUE`, `SEQUENCE`, `RANDARRAY`, `TRANSPOSE`, `LET`

### Other / Bonus
`HYPERLINK`, `WEBSERVICE`, `FILTERXML`

### DAX-Specific
`CALCULATE`, `FILTER`, `ALL`, `ALLEXCEPT`, `RELATED`, `RELATEDTABLE`, `DISTINCT`, `VALUES`, `EARLIER`, `EARLIEST`, `HASONEVALUE`, `HASONEFILTER`, `ISFILTERED`, `USERELATIONSHIP`, `SUMX`, `AVERAGEX`, `MAXX`, `MINX`, `RANKX`, `COUNTX`, `DATESYTD`, `DATESQTD`, `DATESMTD`, `DATEADD`, `DATESINPERIOD`

### Features (Non-Formula)
PivotTables, PivotCharts, Slicers, Timelines, Calculated Fields, Calculated Items, Show Values As, Conditional Formatting, Data Validation, Named Ranges, Auditing Tools (Trace Precedents/Dependents, Evaluate Formula), Power Query (Get & Transform), Data Model, Power Pivot, DAX, CUBE functions, Power Map, Power BI
