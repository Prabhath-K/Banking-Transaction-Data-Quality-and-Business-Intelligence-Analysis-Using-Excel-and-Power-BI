# Banking-Transaction-Data-Quality-and-Business-Intelligence-Analysis-Using-Excel-and-Power-BI

Project overview
This project prepares an Indian banking transaction dataset for analysis in Microsoft Power BI. The workflow uses Microsoft Excel to review and clean the raw data and Power Query to transform, validate, and load it into the reporting model. The cleaned data supports a three-page Power BI report covering transaction performance, fraud/risk indicators, and customer/channel insights.

Objective: Produce a consistent, analysis-ready transaction table without changing legitimate financial amounts or inventing missing business information.
Tools
- Microsoft Excel: Initial data inspection, cleaning, and comparison with the source worksheet.
- Power Query (Power BI Desktop): Applied transformations, type conversion, custom columns, filtering, and loading the query.
- Power BI / DAX: Downstream measures and dashboard visuals (documented separately from data cleaning).
Dataset overview
Property	Result
Original dataset	650 records, 20 columns
Cleaned worksheet	650 records, 24 columns
Unique transaction IDs	650
Unique customer IDs	91
Missing/blank values in the cleaned worksheet	0
Duplicate transaction IDs in the cleaned worksheet	0
Duplicate complete rows in the cleaned worksheet	0
Total recorded transaction amount	₹18,754,341.09


Source worksheets: Indian Banking Transactions (raw) and Indian Banking Transactions Cle (cleaned), from Indian Banking Transactions(1).xlsx. The cleaned worksheet includes four additional fields for outlier and validation reporting. Counts above were checked against the uploaded cleaned workbook; they are not claims about other datasets.
Data-cleaning workflow
1. Inspect the raw dataset
- Reviewed transaction identifiers, customer identifiers, dates, times, financial values, categorical attributes, fraud flags, and loan details.
- Compared the original worksheet with the cleaned worksheet to identify transformations.
- Checked for blanks and duplicate records in the cleaned result.
2. Standardize column names
The original worksheet used snake_case headings such as transaction_id, customer_id, and transaction_amount. The cleaned worksheet uses readable headings including Transaction ID, Customer ID, Transaction Date, Transaction Amount, and Transaction Status.
Minor cleanup to consider: The uploaded cleaned sheet retains a trailing space in the Loan Type  header and uses is fraud rather than Is Fraud. Rename these for consistent final documentation if they still exist in the final model.

3. Standardize categorical values
The original and cleaned worksheets confirm the following replacements:
Field	Original value	Cleaned value	Records updated
State	UP	Uttar Pradesh	42
Transaction Type	ATM_Withdrawal	ATM Withdrawal	67
Transaction Type	Net_Banking	Net Banking	54
Transaction Type	Auto_Debit	Auto Debit	32
Transaction Type	Credit_Card	Credit Card	25
Channel	Mobile_App	Mobile App	244
Channel	POS_Terminal	POS Terminal	74


These changes improve category grouping, slicer usability, and chart readability while preserving the underlying transaction records.
4. Review data types
In Power Query, the transformation history includes Changed Type steps. The key target types for reporting are:
Column	Intended type
Transaction ID, Customer ID	Text
Transaction Date	Date
Transaction Time	Time
Transaction Amount, Account Balance, EMI Amount	Decimal number
Credit Score, Transaction Hour, Has Loan, is fraud	Whole number
State, Transaction Type, Channel, Transaction Status, KYC Status	Text


Power Query's column-quality preview displayed 100% valid values for the visible ID, date, time, account-type, and transaction-type columns in the uploaded screenshot. That preview alone does not prove every business rule is correct.
5. Handle inconsistent loan information without guessing
The raw-versus-cleaned comparison identifies 21 records where Loan Type changed from None to Needs Review. These represent unresolved loan-type information, not confirmed corrections to a specific loan category.

Decision: Flag ambiguous records for further review instead of assigning an unsupported loan type. Retain the transaction so that financial totals are not silently distorted.
6. Flag unusual transaction amounts
A derived Amount Outlier column identifies:
Label	Records
Outlier	109
Normal	541


Outliers were flagged, not automatically removed. A high transaction value is not, by itself, evidence of fraud.
Methodology note: The exact outlier detection formula/threshold (for example, IQR or a business-defined cutoff) was not recoverable from the available workbook and screenshots. Add the actual Excel formula or Power Query custom-column expression here before claiming a specific statistical method.

7. Add validation fields
The cleaned worksheet contains these additional columns:
- Date Validation
- Time Validation
- Status Validation
All 650 records are labeled Valid in each of these three columns. They are retained as explicit data-quality indicators. The precise formulas/conditions used to assign the labels should be recorded separately if the transformation needs to be fully reproducible.
8. Check data integrity
The cleaned worksheet audit shows:
Check	Result
Missing/blank cells	0
Duplicate transaction IDs	0
Duplicate complete rows	0
Zero or negative transaction amounts	0
Transaction-hour mismatches	0
Fraud-flagged transactions (is fraud = 1)	6
Successful transactions	599
Failed transactions	21
Pending transactions	20
Reversed transactions	10


The original and cleaned sheets retained 650 transactions; cleaning did not require deleting transaction rows. The total transaction amount represents recorded values across statuses, not bank revenue or confirmed settled value.
Power Query — Applied Steps
The Power Query Editor screenshot shows that the transformation pipeline contains these applied-step types (some recur multiple times):
1. Renamed Columns — standardized field names for reporting.
2. Changed Type / Changed Type1 / Changed Type2 — set column types, including a visible final conversion of is fraud to Int64.Type.
3. Replaced Value / Replaced Value1 … Replaced Value10 — standardized or adjusted values across multiple steps.
4. Added Custom / Added Custom1 … Added Custom4 — introduced calculated/custom fields.
5. Filtered Rows / Filtered Rows1 — applied filtering steps within the query.
6. Reordered Columns — organized the final column order.
7. Close & Apply — loaded the transformed query into the Power BI model.
Important distinction: These are the step names visible in Power Query, not a complete export of their M expressions. The screenshots do not establish the precise predicate of every Filtered Rows step or the formula inside every Added Custom step. For a completely reproducible pipeline, copy the full M script from Power Query Editor → Home → Advanced Editor into a separate file such as power_query/transformations.m and describe those formulas here.
Example of a verified Power Query type-conversion step
The screenshot of the final applied step shows this M expression:
= Table.TransformColumnTypes(#"Replaced Value10", {{"is fraud", Int64.Type}})
This is shown as evidence of an actual applied step; it is not intended to replace the complete query.
Final dataset and downstream analysis
The cleaned transaction table was used to build Power BI measures and three report pages:
1. Executive Overview — transaction volume, total recorded value, success rate, channel breakdown, monthly value trend, and transaction status.
2. Fraud & Risk Analysis — fraud flags, outliers, failed transactions, channel comparisons, and fraud-flag rates by KYC status.
3. Customer & Channel Insights — distinct customers, transaction value, average transaction value, account-type breakdowns, and state/channel comparisons.
The report uses interactive Date, State, and Channel slicers. Fraud flags and outlier flags are descriptive indicators, not confirmed fraud losses or predictive risk scores.
Data quality and project limitations
- The cleaned table is a single transaction-level dataset; it is not a complete enterprise banking data warehouse.
- Needs Review loan types are intentionally unresolved.
- The outlier classification rule needs to be explicitly documented from the original formula or M step.
- The validation columns report Valid, but the underlying validation formulas should be supplied for full reproducibility.
- A source Excel file may be linked through an absolute local Windows path in Power BI. Users opening the PBIX on another machine should update the file path in Power Query → Data source settings before refreshing.
- With only six flagged fraud transactions, apparent differences in fraud-flag rates should not be treated as reliable predictions or causal effects.
