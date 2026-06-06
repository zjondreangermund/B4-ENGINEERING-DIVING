# Product Blueprint

## Working product name
Nautilus Enterprise

## Goal
Create a web-first operational platform for diving, marine, pollution and civil work that captures jobs, calculates costs, produces job cards, and gives management live profitability visibility.

## Core modules

### 1. Executive Dashboard
- Revenue excl. VAT
- Gross profit
- Gross profit percentage
- Total hours
- Total jobs
- Cost split: labour, equipment, workshop
- Profit by division
- Profit by OPS manager
- Month/year filters

### 2. Job Register
The job register must feel familiar to Excel but work like a proper app.

Key fields:
- Job number
- Job date
- Division head / OPS manager
- Division
- Client
- Job/project description
- Completion date
- Quote number
- PO number
- Report number/date
- Invoice number
- Revenue incl. VAT
- Revenue excl. VAT
- Status

### 3. Job Card Conversion
This module converts field work into office-ready records.

Inputs:
- Date
- OPS manager
- Description
- Start time
- End time
- Hours
- Labour cost
- Equipment cost
- Workshop cost
- Division

Outputs:
- Job card PDF
- Costing record
- Invoice-ready line
- Excel export row

### 4. Costing
- Labour cost per job
- Equipment cost per job
- Workshop cost per job
- Other expenses
- Total cost
- Gross profit
- Gross profit %

### 5. Master Data
- Users
- Clients
- Divisions
- Vessels
- Employees
- Equipment
- Cost categories
- Salary records

### 6. Audit Trail
Every change must be recorded because this is a serious company system.

Audit events:
- Job created
- Job edited
- Cost changed
- Invoice number added
- Job card approved
- Export generated
- Import completed

## Non-negotiables
- No data loss.
- Excel import/export must be exact.
- Every financial calculation must be traceable.
- Admins must be able to see who changed what.
- The app must work cleanly on laptop, tablet and mobile browser.
