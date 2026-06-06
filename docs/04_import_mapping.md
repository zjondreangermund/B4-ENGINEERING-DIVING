# Excel Import and Export Mapping

## Strategy
Use a staging import system. Never write uploaded Excel data directly into live production tables.

## Import stages

### Stage 1: Upload
Save original workbook file.

### Stage 2: Detect
Detect sheet names and headers.

### Stage 3: Map
Map Excel columns to database fields.

### Stage 4: Validate
Check required fields, number formats, dates, duplicates and invalid values.

### Stage 5: Staging
Load valid rows into import_staging_jobs.

### Stage 6: Approval
Admin reviews import summary and approves.

### Stage 7: Commit
Create/update production records.

## Main data mapping

| Excel Column | Database Field | Notes |
|---|---|---|
| Month | reporting_month | Can be calculated from date |
| Year | reporting_year | Can be calculated from date |
| FY | financial_year | Must match company FY rules |
| Date | job_date | Required |
| Job Nr | job_number | Required unique per company |
| Invoice No. | invoice_number | Optional |
| Revenue Incl VAT | revenue_incl_vat | Decimal |
| Revenue Excl VAT | revenue_excl_vat | Decimal |
| Start Time | start_time | Time/decimal handling required |
| End Time | end_time | Time/decimal handling required |
| Total Hours | total_hours | Can be calculated or imported |
| Labour_Costs | labour_cost | Decimal |
| Equipment_Costs | equipment_cost | Decimal |
| Workshop Cost | workshop_cost | Decimal |
| Total_Costs | total_cost | Should equal labour + equipment + workshop + other |
| OPS Manager | ops_manager_name | Resolve to user/person record |
| Division | division_name | Resolve to division record |

## Validation rules

- Job number cannot be blank.
- Job date cannot be blank.
- Revenue and cost fields must be numeric.
- Total cost must be checked against component costs.
- Profit must be recalculated by the system, not trusted blindly from Excel.
- Duplicate job numbers must be flagged before import.
- Unknown divisions and OPS managers must be staged for mapping.

## Export rules

The export must support:
- Current app data to Excel format
- Job register export
- Data table export
- Dashboard validation export
- Original Excel-compatible column order
