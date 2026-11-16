# 🏦 Banking Transaction Data Quality & Business Intelligence Analysis

**Microsoft Excel · Power Query · Power BI · DAX**

> A beginner-friendly walkthrough of the documented Excel cleaning checks and the three-page Power BI report. The DAX formulas below are **educational reconstructions**, not verified exports of the original measure expressions.

## 📑 Contents

- [Excel Data Cleaning](#-excel-data-cleaning)
- [DAX Measures](#-dax-measures-meaning-creation-and-examples)
- [Dashboard Design and Formatting](#-dashboard-design-and-common-formatting)
- [Page 1 — Executive Overview](#-page-1--executive-overview)
- [Page 2 — Fraud & Risk Analysis](#-page-2--fraud--risk-analysis)
- [Page 3 — Customer & Channel Insights](#-page-3--customer--channel-insights)
- [Testing Checklist](#-testing-interactivity-and-checking-results)

---

## 🧹 Excel Data Cleaning

### Standardize column headings

**Where:** Edit the header cells in the cleaned Excel worksheet, or use Power Query → Transform → Rename.

| Original | Readable heading |
| --- | --- |
| transaction_id | Transaction ID |
| customer_id | Customer ID |
| transaction_date | Transaction Date |
| transaction_amount | Transaction Amount |
| transaction_status | Transaction Status |

**Purpose:** Make fields understandable when dragging them into a Power BI visual and when writing DAX.

**Final consistency check:** The documented cleaned worksheet retains is fraud and a trailing space in Loan Type ` `. A final model may rename them to Is Fraud and Loan Type, but then adjust DAX expressions accordingly. Do not claim those final renames already happened.

### Step 4. Standardize categorical values

**Where:** In Excel, select the appropriate column and use Ctrl + H (Find and Replace), taking care to match whole cells. In Power Query, select a column → Transform → Replace Values. Review the result before saving.

| Column | Before | After | Rows affected |
| --- | --- | --- | --- |
| State | UP | Uttar Pradesh | 42 |
| Transaction Type | ATM_Withdrawal | ATM Withdrawal | 67 |
| Transaction Type | Net_Banking | Net Banking | 54 |
| Transaction Type | Auto_Debit | Auto Debit | 32 |
| Transaction Type | Credit_Card | Credit Card | 25 |
| Channel | Mobile_App | Mobile App | 244 |
| Channel | POS_Terminal | POS Terminal | 74 |

**Purpose:** Without standardization, one real-world category may appear under multiple labels. Consistent names make charts and slicers easier to read. These are label changes; they are not changes to financial amounts.

### Step 5. Check intended column data types

**Where:** Excel → Home → Number Format for basic inspection. Set authoritative reporting types in Power Query later.

| Text | Date / Time | Decimal number | Whole number |
| --- | --- | --- | --- |
| Transaction ID, Customer ID, State, Channel, Status, KYC Status, Transaction Type | Transaction Date (Date), Transaction Time (Time) | Transaction Amount, Account Balance, EMI Amount | Credit Score, Transaction Hour, Has Loan, is fraud |

**Purpose:** Using the wrong type can break date grouping, make arithmetic incorrect, or cause identifiers to lose formatting. Setting Excel's cell display format is not always equivalent to converting underlying values to the right data type.

### Step 6. Handle ambiguous loan details without inventing data

Observed change: 21 Loan Type values changed from None to Needs Review when original and cleaned worksheets were compared.

**Purpose:** A missing or conflicting loan category is a data-quality issue, not permission to guess whether the customer has a particular loan. Keeping Needs Review makes unresolved records visible and preserves transaction totals.

### Step 7. Flag unusual transaction amounts

A new Amount Outlier field classifies 109 records as Outlier and 541 as Normal.

**Purpose:** Make unusually valued transactions easy to investigate on the risk page, without deleting them. A high amount is not automatically fraudulent.

**Reproducibility warning:** The actual rule/threshold used for these flags was not available in the documented source material. Do not state that the existing 109 flags were produced by IQR, Z-score, or any fixed rupee threshold unless the original formula or M expression confirms it.

### Step 8. Add three validation-result fields

The cleaned worksheet additionally contains:
- Date Validation
- Time Validation
- Status Validation
All 650 rows are labeled Valid in each field. These are reporting labels; the original validation conditions should be retrieved before claiming a reproducible rule. Purpose: expose checks rather than hiding them inside a manual review.

### Step 9. Audit the cleaned worksheet

| Check | Documented result |
| --- | --- |
| Blank cells | 0 |
| Duplicate transaction IDs | 0 |
| Duplicate complete rows | 0 |
| Zero/negative transaction amounts | 0 |
| Transaction-hour mismatches | 0 |
| is fraud = 1 | 6 |
| Successful transactions | 599 |
| Failed transactions | 21 |
| Pending transactions | 20 |
| Reversed transactions | 10 |

Check that transaction statuses sum to 650 and compare the total Transaction Amount before and after cleaning. The recorded total is ₹18,754,341.09 across all statuses, not only successful or settled items.

### Step 10. Compare before and after

| Property | Original | Cleaned |
| --- | --- | --- |
| Rows | 650 | 650 |
| Columns | 20 | 24 |
| Column names | Snake case | Mostly readable headings |
| Some categorical values | Abbreviated / underscore-delimited | Standardized |
| Loan-type ambiguity | None in identified cases | 21 marked Needs Review |
| Amount outlier label | Not present | Added |
| Date / time / status validation labels | Not present | Added |

**Outcome:** An analysis-ready worksheet was prepared while retaining all 650 transaction rows.

---

## 🧮 DAX Measures: Meaning, Creation, and Examples

### What is DAX?

DAX (Data Analysis Expressions) is the formula language used by Power BI to calculate totals, counts, percentages and other metrics. A measure is a calculation evaluated in the current filter context. For example, when the user selects one State, the same measure recalculates for that state.

### How to create a measure — click by click

1. In Power BI Desktop, open Report view.
2. In the Data/Fields pane, right-click the table Indian_Banking_Transactions_Cleaned_Verion.
3. Click New measure. Alternatively use Modeling → New measure or Home → New measure, depending on the UI version.
4. In the formula bar, type a name, =, and the DAX expression.
5. Press Enter or the check mark. A calculator icon appears beside the measure.
6. Select the measure and use Measure tools / Modeling to set display format, such as Whole number, Currency or Percentage.
7. Drag the measure to a Card, or to a chart's Y-axis / Values field well.

### Actual measure names referenced in the PBIX report visuals

The PBIX report configuration references the following names (including original spelling):

**Measures:** Total Transactions, Total Transaction Amount, Unique Customers, Successful Transactions, Success Rate, Fraud Tansaction, Fraud Rate, Outlier Transactions, Outlier Rate, Failed Transactions, Average Transaction Value.

**Very important distinction:** The names and their use in visuals were directly verified from Bank Transactions.pbix. The original DAX expression text was not reliably extractable, so the formulas below are educational reconstructions. They are not presented as exact formulas recovered from the uploaded PBIX. Compare them in Power BI Desktop by clicking each existing measure and reading the formula bar before treating them as the project's definitive code.

> For readability, all examples below assume the actual model table is called Indian_Banking_Transactions_Cleaned_Verion. DAX uses the exact stored names: is fraud, Transaction Status and Amount Outlier.

### Example DAX measure reference

#### 1. Total Transactions — How many transaction records exist?

```dax
Total Transactions =
COUNTROWS('Indian_Banking_Transactions_Cleaned_Verion')
```

- Function: COUNTROWS counts transaction rows in the current filter context.
- Why needed: Overall activity KPI; also used in channel, status and state comparisons.
- Unfiltered expected result: 650, provided the model includes the same 650 cleaned records.

#### 2. Total Transaction Amount — What is the recorded transaction value?

```dax
Total Transaction Amount =
SUM('Indian_Banking_Transactions_Cleaned_Verion'[Transaction Amount])
```

- Function: SUM adds the amounts in currently visible rows.
- Format: Currency (₹) or decimal with Indian rupee label as appropriate.
- Unfiltered expected result: ₹18,754,341.09.
- Important: This includes all statuses unless you explicitly filter it; it is not revenue.

#### 3. Unique Customers — How many distinct customers appear?

```dax
Unique Customers =
DISTINCTCOUNT('Indian_Banking_Transactions_Cleaned_Verion'[Customer ID])
```

- Function: DISTINCTCOUNT counts each unique ID once.
- Why: One customer may make multiple transactions.
- Unfiltered expected result: 91.

#### 4. Successful Transactions — How many transactions succeeded?

```dax
Successful Transactions =
CALCULATE(
    [Total Transactions],
    'Indian_Banking_Transactions_Cleaned_Verion'[Transaction Status] = "Success"
)
```

- Function: CALCULATE evaluates a calculation using an additional filter.
- Check first: Confirm the model's exact successful-status label; the documented status category is Success, but if your real model says Successful, edit the string accordingly.
- Unfiltered documented count: 599.

#### 5. Success Rate — What percentage of transactions succeeded?

```dax
Success Rate =
DIVIDE([Successful Transactions], [Total Transactions], 0)
```

- Function: DIVIDE(numerator, denominator, fallback) avoids division-by-zero errors.
- Format: Percentage; do not multiply by 100 if using Percentage formatting.
- Unfiltered value implied by the documented counts: 599 ÷ 650 ≈ 92.15%.

#### 6. Fraud Tansaction — How many records are fraud-flagged?

```dax
Fraud Tansaction =
CALCULATE(
    [Total Transactions],
    'Indian_Banking_Transactions_Cleaned_Verion'[is fraud] = 1
)
```

- Name note: Fraud Tansaction is intentionally spelled as in the actual PBIX. Renaming it in Power BI is possible, but then check dependent measures and visual labels.
- Unfiltered documented count: 6.
- Interpretation: Flagged records, not confirmed fraud losses.

#### 7. Fraud Rate — What share of transactions are flagged?

```dax
Fraud Rate =
DIVIDE([Fraud Tansaction], [Total Transactions], 0)
```

- Format: Percentage.
- Unfiltered value implied by documented counts: 6 ÷ 650 ≈ 0.92%.
- Purpose: Compare descriptive flag rates by KYC status or other categories; sample sizes matter.

#### 8. Outlier Transactions — How many records have an outlier label?

```dax
Outlier Transactions =
CALCULATE(
    [Total Transactions],
    'Indian_Banking_Transactions_Cleaned_Verion'[Amount Outlier] = "Outlier"
)
```

- Unfiltered documented count: 109.
- Purpose: Compare already-flagged transactions; this measure does not calculate the outlier threshold.

#### 9. Outlier Rate — What fraction of transactions are flagged as outliers?

```dax
Outlier Rate =
DIVIDE([Outlier Transactions], [Total Transactions], 0)
```

- Format: Percentage.
- Unfiltered value implied by counts: 109 ÷ 650 ≈ 16.77%.

#### 10. Failed Transactions — How many failed?

```dax
Failed Transactions =
CALCULATE(
    [Total Transactions],
    'Indian_Banking_Transactions_Cleaned_Verion'[Transaction Status] = "Failed"
)
```

- Unfiltered documented count: 21.
- Purpose: Monitor operationally unsuccessful transactions.

#### 11. Average Transaction Value — What is the average recorded amount?

```dax
Average Transaction Value =
DIVIDE([Total Transaction Amount], [Total Transactions], 0)
```

- Format: ₹ currency with 2 decimal places.
- Purpose: Understand typical value as a mathematical mean, not the median.
- Expected unfiltered value using documented totals: approximately ₹28,852.83.

### Where to find the exact original DAX in the PBIX

Open the supplied Bank Transactions.pbix in Power BI Desktop → find a measure (calculator icon) in the Data pane → select the measure → copy its expression from the formula bar. Repeat for all 11 names above and, if desired, store them in a dax/measures.dax documentation file. Do not overwrite correct existing measures with example formulas without comparing their behavior first.

---

## 🎨 Dashboard Design and Common Formatting

### Understanding visual components

| Element | What it is | Why use it? |
| --- | --- | --- |
| Card | One prominent number or percentage | Quickly display a KPI |
| Clustered bar chart | Horizontal bars comparing groups | Compare channel or state performance |
| Donut chart | Parts of a whole | Compare category composition |
| Line chart | Values across a time axis | See changes over time |
| Slicer | Interactive filter on the report canvas | Explore one date, state or channel |
| Tooltip | Information shown on hover | Add context without cluttering the chart |
| Text box | Title, subtitle or explanation | Help users navigate the dashboard |

### Create a visual — reusable beginner workflow

1. Open the desired page in Report view.
2. Click a blank area on the page.
3. In Visualizations / Build visual, choose Card, Clustered bar chart, Donut chart, Line chart, or Slicer. In Power BI versions with on-object interaction, select the visual first and use Build a visual to add fields.
4. Drag the required field/measure from the Data pane into its Data, X-axis, Y-axis, Legend, Values, or Tooltips field well (labels vary by visual/version).
5. Resize and position the visual on the canvas.
6. Select the visual → Format visual (paintbrush) → customize the visual and its general properties.
7. Confirm its title, data labels, axis labels, currency/percentage format and tooltip values.
8. Test slicer and chart interactions.

### How to format a Card

1. Insert a Card and add one measure to Data/Values.
2. In Format visual → Visual, use the corresponding Callout value / Value settings to control font size, color and display units.
3. Set Category label / Label to show what the number means.
4. Under General → Effects, adjust background, border and corner radius if available.
5. For rates, select the measure itself and set Percentage in measure formatting; this is different from changing the card's font.
6. Set a readable title or category label such as Success Rate.

**Formatting evidence from the PBIX:** The report uses a dark blue canvas (stored page background #0B1425) with light text. Some Executive cards have a dark #182840 fill, white main values, muted labels, and a recorded value font size of 32. These are observed examples, not a claim that every visual has identical formatting.

### How to format a clustered bar chart

1. Insert Clustered bar chart.
2. Put a category such as Channel or State in the Y-axis / Category well.
3. Put a numeric measure such as [Total Transactions] in the X-axis / Values well (the saved report schema calls this the Y measure projection; UI labels can differ).
4. Optionally add other measures to Tooltips.
5. Go to Format visual → adjust Bars / Data colors, X-axis, Y-axis, Data labels, Gridlines, Title and Background.
6. For rankings, use the visual menu More options (…) → Sort axis → measure → Descending.
7. For a Top N chart, apply a visual-level Top N filter (detailed on Page 3).

**Purpose:** Compare the size of a metric across categories instead of showing only one total.

### How to format a donut chart

1. Insert Donut chart.
2. Place a category in Legend / Category.
3. Place [Total Transactions] in Values.
4. Under Format visual, set segment colors, detail labels, legend position and title.
5. Check that the category total agrees with the overall transaction count under the same filters.

**Purpose:** Show how an overall total is distributed among categories, not a time trend.

### How to format a line chart

1. Insert Line chart.
2. Place Transaction Date (Date hierarchy, if appropriate) on the X-axis.
3. Place [Total Transaction Amount] on the Y-axis.
4. Choose Year and Month hierarchy levels / drill depth to show monthly data.
5. Sort dates chronologically; check that months from different years are not mistakenly combined.
6. Under Format visual, adjust lines, markers, axis labels, currency display and chart title.

**Purpose:** See how recorded value changes across time.

### How to add and format a slicer

1. Insert a Slicer.
2. Add one column such as Transaction Date, State, or Channel.
3. Set the appropriate style in Format visual → Slicer settings (for example, a date range, dropdown, or list depending on version and field type).
4. Adjust title/header, text size, background, outline, search and selection behavior.
5. Choose a state/channel and check that other visuals change.
6. If needed, select slicer → Format → Edit interactions → choose Filter or None over individual target visuals.
7. Optional: use View → Sync slicers to coordinate filters across pages. The PBIX confirms repeated slicers, not that synchronization was configured, so test synchronization before claiming it.

**Purpose:** Users can compare subsets (for example, one state or one channel) without rebuilding charts or using Excel filters.

### Titles, theme and page layout

- Text: Use Insert → Text box to add clear page headings and descriptions.
- Theme: The saved report includes a Fluent2 theme resource, dark page backgrounds and light/high-contrast text.
- Executive Overview page: Saved dimensions 1280 × 900; Fit to Page display.
- Fraud & Risk Analysis page: Saved dimensions 1920 × 1080; Fit to Width display.
- Customer & Channel Insights page: Saved dimensions 1920 × 1080; Fit to Page display.
- Accessibility: Ensure sufficient contrast and meaningful labels; avoid using color alone to communicate risk.

---

## 📊 Page 1 — Executive Overview

**Actual PBIX page name:** Executive Overview
**Main purpose:** Give an at-a-glance summary of transaction activity, recorded value, success, customers, fraud flags, categories and monthly trends.

### Page 1 visual inventory verified from the PBIX

| Visual | Actual field / measure | Purpose |
| --- | --- | --- |
| Card — Total Transactions | [Total Transactions] | Number of transaction records |
| Card — Total Transaction Amount | [Total Transaction Amount] | Sum of recorded amounts |
| Card — Unique Customers | [Unique Customers] | Distinct customer count |
| Card — Successful Transactions | [Successful Transactions] | Count of successful records |
| Card — Success Rate | [Success Rate] | Proportion successful |
| Card — Fraud Tansaction | [Fraud Tansaction] | Number flagged for fraud |
| Donut — Transaction Status | Transaction Status + [Total Transactions] | Status mix |
| Donut — Debit vs Credit | Transaction Direction + [Total Transactions] | Transaction direction mix |
| Bar — Transactions by Channel | Channel + [Total Transactions]; tooltip [Total Transaction Amount] | Compare channel counts, with value on hover |
| Line — Monthly Transaction Value Trend | Transaction Date hierarchy (two levels) + [Total Transaction Amount] | Time trend of recorded amount |
| Slicer — Date | Transaction Date | Time filtering |
| Slicer — State | State | Regional filtering |
| Slicer — Channel | Channel | Channel filtering |

**Note:** The saved page also contains extra card instances referencing [Success Rate] and [Fraud Tansaction], plus text boxes. These are repeated visual instances, not new measures.

### Rebuild Page 1 — each component

#### 1A. Add the page title

1. Click + New page and rename it Executive Overview (right-click page tab → Rename).
2. Choose Insert → Text box and type Banking Transaction Analytics.
3. Add a smaller text box: EXECUTIVE OVERVIEW.
4. Under Format, make the title larger, use light text on a dark background, and position it above the KPIs.
Why: Readers immediately know which page they are viewing.

#### 1B. Create KPI Cards

For each row below, repeat: click Card → drag the named measure to Data/Values → Format visual → show a readable label, set value size/color and arrange cards in a row or grid.

| Card | Drop this measure | Display guidance | Business question |
| --- | --- | --- | --- |
| Total Transactions | [Total Transactions] | Whole number, no unnecessary decimals | How many transactions? |
| Total Transaction Amount | [Total Transaction Amount] | Rupee currency or compact units clearly labeled | How much recorded value? |
| Unique Customers | [Unique Customers] | Whole number | How many different customers? |
| Successful Transactions | [Successful Transactions] | Whole number | How many succeeded? |
| Success Rate | [Success Rate] | Percentage (e.g., 2 decimals) | What proportion succeeded? |
| Fraud Tansaction | [Fraud Tansaction] | Whole number; label as fraud-flagged | How many were flagged? |

**Formatting tip:** Apply the numeric format at Measure tools first; then use card formatting for font/size/background. The actual PBIX uses light card values on dark fills.

#### 1C. Create “Transactions by Channel” bar chart

1. Click Clustered bar chart.
2. Add Channel to the category / Y-axis.
3. Add [Total Transactions] to Values / numeric axis.
4. Add [Total Transaction Amount] to Tooltips — this additional tooltip is confirmed in the PBIX.
5. Set the chart title to Transactions by Channel.
6. Format bar color, category labels, numeric labels and background.
7. Optionally sort descending by total transactions.

**Purpose:** Identify which transaction channel recorded more activity. Hover to also see recorded amount for that channel.

#### 1D. Create “Transaction Status” donut

1. Click Donut chart.
2. Set Legend/Category = Transaction Status.
3. Set Values = [Total Transactions].
4. Set title to Transaction Status.
5. Enable legible labels/legend and adjust segment colors.
6. Cross-check the status counts: 599 Success, 21 Failed, 20 Pending, 10 Reversed (unfiltered documented counts).

**Purpose:** See the share of records at each processing status.

#### 1E. Create “Debit vs Credit” donut

1. Click Donut chart.
2. Set Legend/Category = Transaction Direction.
3. Set Values = [Total Transactions].
4. Set title to Debit vs Credit.
5. Format category labels and the legend.

**Purpose:** Compare the transaction count distribution by direction. The chart uses counts, not amounts.

#### 1F. Create “Monthly Transaction Value Trend” line chart

1. Click Line chart.
2. Set X-axis = Transaction Date date hierarchy, with the required Year/Month drill levels.
3. Set Y-axis = [Total Transaction Amount].
4. Choose the Year/Month level needed to show an ordered monthly trend.
5. Set the title to Monthly Transaction Value Trend.
6. Format axis labels, line and currency units.

**Purpose:** Track how the total recorded amount changes over time. The PBIX confirms two hierarchy-level projections, but the precise drill state should be verified in Desktop.

#### 1G. Create Date, State and Channel slicers

Repeat Insert Slicer → add field → format slicer settings for:

| Slicer | Field | Example use |
| --- | --- | --- |
| Date | Transaction Date | Focus on a period |
| State | State | View transactions from one region |
| Channel | Channel | Compare a specific channel |

Test a channel selection: Total Transactions, the chart bars and other visuals should reflect the filtered subset unless deliberately excluded by Edit interactions.

**Page 1 expected takeaway:** High-level transaction performance. The unfiltered dataset contains 650 transactions, 91 customers, ₹18,754,341.09 recorded amount, 599 successful transactions, and 6 fraud-flagged transactions.

---

## 🛡️ Page 2 — Fraud & Risk Analysis

**Actual PBIX page name:** Fraud & Risk analysis (capitalization as saved)
**Main purpose:** Monitor records flagged for fraud, unusual amounts and failures. These indicators support review; they do not identify proven crimes or forecast future fraud.

### Page 2 visual inventory verified from the PBIX

| Visual | Actual field / measure | Purpose |
| --- | --- | --- |
| Card — Fraud Tansaction | [Fraud Tansaction] | Fraud-flagged count |
| Card — Fraud Rate | [Fraud Rate] | Fraud-flagged percentage |
| Card — Outlier Transactions | [Outlier Transactions] | Outlier-flagged count |
| Card — Outlier Rate | [Outlier Rate] | Outlier-flagged percentage |
| Card — Failed Transactions | [Failed Transactions] | Count of failed records |
| Bar — Fraud by Channel | Channel + [Fraud Tansaction] | Fraud-flag counts by channel |
| Bar — Outliers by Channel | Channel + [Outlier Transactions] | Outlier-flag counts by channel |
| Bar — Failed Transactions by Channel | Channel + [Failed Transactions] | Failure counts by channel; title not explicitly stored in inspected configuration |
| Bar — Fraud Rate by KYC Status | KYC Status + [Fraud Rate]; tooltips [Fraud Tansaction], [Total Transactions] | Compare rates with supporting counts |
| Slicer — Date | Transaction Date | Filter by time |
| Slicer — State | State | Filter by region |
| Slicer — Channel | Channel | Filter by channel |

### Rebuild Page 2 — each component

#### 2A. Create page heading

1. Add another report page and rename it Fraud & Risk analysis.
2. Insert a text box with Fraud & Risk Analysis.
3. Insert a subtitle text box: Transaction Monitoring and Risk Indicators.
4. Apply the same dark-blue/light-text design as Page 1.

#### 2B. Add the five risk KPI Cards

Create five Card visuals, one at a time:

| Card | Measure | Display format | Why it matters |
| --- | --- | --- | --- |
| Fraud Tansaction | [Fraud Tansaction] | Whole number | Number of flagged records |
| Fraud Rate | [Fraud Rate] | Percentage | Share of records flagged |
| Outlier Transactions | [Outlier Transactions] | Whole number | Unusual-amount labels |
| Outlier Rate | [Outlier Rate] | Percentage | Share with outlier labels |
| Failed Transactions | [Failed Transactions] | Whole number | Operational failures |

Exact procedure for every card: Visualizations → Card → drag measure into Data/Values → Measure tools for number/percentage → Format visual for value font, label, background and corners.

**Documented unfiltered benchmarks:** 6 flagged fraud transactions, 109 outliers and 21 failures. The reconstructed rate examples produce about 0.92% fraud-flag rate and 16.77% outlier rate.

#### 2C. Create “Fraud by Channel” bar chart

1. Insert Clustered bar chart.
2. Category/Y-axis = Channel.
3. Values/numeric axis = [Fraud Tansaction].
4. Title = Fraud by Channel.
5. Format bars, data labels and axes; optionally sort descending.

**Purpose:** See which channels contain more fraud-flagged records. For comparison of rates, use a denominator such as total transactions per channel; counts alone can be affected by channel size.

#### 2D. Create “Outliers by Channel” bar chart

1. Insert Clustered bar chart.
2. Category = Channel.
3. Values = [Outlier Transactions].
4. Title = Outliers by Channel.
5. Apply clear labels and consistent formatting.

**Purpose:** See where flagged unusual amounts occur. A record can be an outlier without being fraud.

#### 2E. Create the failed-transactions-by-channel bar chart

1. Insert Clustered bar chart.
2. Category = Channel.
3. Values = [Failed Transactions].
4. Add an explanatory title such as Failed Transactions by Channel when rebuilding (the PBIX confirms fields but does not store this title explicitly in the examined chart).
5. Format axes and labels.

**Purpose:** Compare failed transaction counts across channels, which can suggest areas for operational review.

#### 2F. Create “Fraud Rate by KYC Status” bar chart

1. Insert Clustered bar chart.
2. Category = KYC Status.
3. Values = [Fraud Rate].
4. Tooltips = [Fraud Tansaction] and [Total Transactions] — both confirmed in PBIX.
5. Title = Fraud Rate by KYC Status.
6. Set fraud rate to Percentage and make axis labels legible.

**Purpose:** Compare flagged share by KYC category while allowing viewers to see both the flagged numerator and the transaction denominator on hover. With just six fraud-flagged records overall, treat differences as descriptive and potentially unstable.

#### 2G. Add three slicers

Create slicers for Transaction Date, State, and Channel using the common workflow. The saved PBIX uses listSlicer visuals on this page. Use a list style when reconstructing it, with readable titles and selection controls.

Try this: Select a particular channel and observe how fraud counts, failed counts, outliers and rates respond. Clear the filter before comparing with whole-dataset totals.

**Page 2 expected takeaway:** This page supports risk monitoring and follow-up questions, not a conclusion that one category or group causes fraud.

---

## 👥 Page 3 — Customer & Channel Insights

**Actual PBIX page name:** Customer & Channel Insights  (saved with a trailing space)
**Main purpose:** Understand customers, average transaction value, account mix, regional activity and differences between channels.

### Page 3 visual inventory verified from the PBIX

| Visual | Actual field / measure | Purpose |
| --- | --- | --- |
| Card — Unique Customers | [Unique Customers] | Distinct customers |
| Card — Total Transactions | [Total Transactions] | Transaction volume |
| Card — Total Transaction Amount | [Total Transaction Amount] | Recorded amount |
| Card — Average Transaction Value | [Average Transaction Value] | Average amount per transaction |
| Bar — Transactions by Channel | Channel + [Total Transactions] | Activity by channel; title not explicitly stored |
| Bar — Top 7 States by Transactions | State + [Total Transactions]; tooltip [Total Transaction Amount]; Top 7 filter | Highest-activity seven states |
| Donut — Transaction Distribution by Account Type | Account Type + [Total Transactions] | Share of transactions by account category |
| Bar — Total Transaction Amount by Channel | Channel + [Total Transaction Amount]; tooltips [Total Transactions], [Unique Customers], [Average Transaction Value] | Compare recorded amounts with extra context |
| Slicer — Date | Transaction Date | Filter by date |
| Slicer — State | State | Filter by state |
| Slicer — Channel | Channel | Filter by channel |

### Rebuild Page 3 — each component

#### 3A. Create the page heading

1. Add a report page and rename it Customer & Channel Insights.
2. Use Insert → Text box for the page title.
3. Add subtitle: Customer Behavior | Account Segmentation | Regional & Channel Performance.
4. Keep the dark theme, text contrast and spacing consistent with the earlier pages.

#### 3B. Create four KPI Cards

| Card | Measure | Format | Purpose |
| --- | --- | --- | --- |
| Unique Customers | [Unique Customers] | Whole number | Distinct customer reach |
| Total Transactions | [Total Transactions] | Whole number | Transaction activity |
| Total Transaction Amount | [Total Transaction Amount] | ₹ / currency | Recorded value |
| Average Transaction Value | [Average Transaction Value] | ₹, 2 decimals | Mean value per transaction |

For each: Card → Data/Values → choose measure → Format visual. Keep each card's label visible so that ₹ amounts are never mistaken for counts.

#### 3C. Create the channel transaction-count chart

1. Insert Clustered bar chart.
2. Category = Channel.
3. Values = [Total Transactions].
4. Give it a descriptive title such as Transactions by Channel when rebuilding.
5. Adjust bar colors and sort/labels as needed.

**Purpose:** Compare volume across channels; this is different from comparing rupee amounts.

#### 3D. Create “Top 7 States by Transactions” — including the real Top N filter

1. Insert Clustered bar chart.
2. Category = State.
3. Values = [Total Transactions].
4. Tooltips = [Total Transaction Amount] — verified in PBIX.
5. Open the Filters pane and expand Filters on this visual.
6. Add/select State in this visual's filters.
7. Switch its filter type to Top N.
8. Choose Top and enter 7 in Show items.
9. Drag [Total Transactions] into By value.
10. Click Apply filter, then sort the bars by [Total Transactions] descending.
11. Title = Top 7 States by Transactions; format axes, labels and background.

**Verified configuration:** The inspected PBIX stores a State TopN visual filter with Top: 7, ordered by the [Total Transactions] measure. This is a concrete report setting, not a hypothetical example.

**Purpose:** Focus attention on the seven states with the most transactions in the current relevant context, while preserving other states in the underlying dataset.

#### 3E. Create “Transaction Distribution by Account Type” donut

1. Insert Donut chart.
2. Legend/Category = Account Type.
3. Values = [Total Transactions].
4. Title = Transaction Distribution by Account Type.
5. Format legend, labels and segment colors.

**Purpose:** See which account types contribute to transaction counts.

#### 3F. Create “Total Transaction Amount by Channel” bar chart

1. Insert Clustered bar chart.
2. Category = Channel.
3. Values = [Total Transaction Amount].
4. Add all three saved tooltip measures:
   - [Total Transactions]
   - [Unique Customers]
   - [Average Transaction Value]
5. Set title = Total Transaction Amount by Channel.
6. Format the value axis as rupee amount (or visibly labeled compact units) and make labels readable.
7. Optionally sort by [Total Transaction Amount] descending.

**Purpose:** A channel may have many transactions but a smaller overall monetary amount. The tooltip gives additional context: volume, customers and average size.

#### 3G. Add Date, State and Channel slicers

1. Insert a Slicer using Transaction Date.
2. Repeat for State.
3. Repeat for Channel.
4. Use list styling consistent with the saved page's listSlicer visuals.
5. Test whether selecting a state recalculates the Top 7 visual, other channel charts and KPI Cards as expected.

**Page 3 expected takeaway:** Customer count, transaction volume and total value are related but different measures. The report provides views into account types, channels and state-level distribution.

---

## ✅ Testing Interactivity and Checking Results

Follow this checklist after building the three report pages:
- [ ] Record preservation: 650 cleaned worksheet records remain available under the unfiltered model context.
- [ ] Total amount: The total recorded amount matches ₹18,754,341.09 when no filters are active.
- [ ] Customer count: [Unique Customers] returns 91 in the unfiltered context.
- [ ] Status reconciliation: Success 599 + Failed 21 + Pending 20 + Reversed 10 = 650.
- [ ] Fraud/outliers: Fraud-flagged = 6 and outlier-flagged = 109 unfiltered.
- [ ] Percentages: Measures are formatted as percentages, not incorrectly scaled twice.
- [ ] Date: Monthly chart is sorted in actual chronological order.
- [ ] Charts: Axis measure is correct — count versus amount must not be confused.
- [ ] Top N: Page 3's State visual uses Top 7 by Total Transactions.
- [ ] Slicers: Date, State and Channel selections update intended visuals.
- [ ] Clear selections: Clearing slicers restores unfiltered KPIs.
- [ ] Tooltips: Page 1 channel count chart shows amount; Page 2 KYC rate chart shows flagged/total counts; Page 3 amount-by-channel chart shows transaction/customer/average measures.
- [ ] Interpretation: Fraud flags and outliers are clearly described as indicators, not confirmed fraud losses.
