# Workbook Analysis

## Detected workbook sheets

### Job Register
Rows: 41 | Columns: 10
Sample first rows:
-  |  |  |  |  | B4 ENG & DIVING JOB REGISTER |  |  |  | 
-  |  |  |  |  |  |  |  |  | 
-  | Search By |  |  |  |  |  |  |  | 
### Analysis
Rows: 48 | Columns: 21
Sample first rows:
-  |  |  |  |  |  |  |  |  | 
-  |  |  |  |  |  |  |  |  | 
-  |  |  |  |  |  |  |  |  | 
### Profit per Job
Rows: 958 | Columns: 12
Sample first rows:
- ***Note: | Select Year,Division and OPS Manager to Analyse seperatly |  |  |  |  |  |  |  | 
-  |  |  |  | FY24 | 1 Nov 2023 - 31 Oct 2024 |  |  |  | 
- Year | 2024 |  |  | FY25 | 1 Nov 2024 - 31 Oct 2025 |  |  |  | 
### Eugen
Rows: 28 | Columns: 11
Sample first rows:
- Start Date | OPS Manager | Job Nr |  Revenue Excl VAT |  Total Hours |  Labour_Costs |  Equipment_Costs |  Workshop Cost |  Total_Costs |  Profit
- 2024-10-14 00:00:00 | Eugen | 1089 | 0 | 3.333333333333334 | 1451.1888111888113 | 350 | 670.7558528150216 | =SUM(Table1[[#This Row],[ Labour_Costs]:[ Workshop Cost]]) | =Table1[[#This Row],[ Revenue Excl VAT]]-Table1[[#This Row],[ Total_Costs]]
- 2024-11-05 00:00:00 | Eugen | 1132 | 0 | 3 | 1392.5838461538467 | 350 | 641.1064999007347 | =SUM(Table1[[#This Row],[ Labour_Costs]:[ Workshop Cost]]) | =Table1[[#This Row],[ Revenue Excl VAT]]-Table1[[#This Row],[ Total_Costs]]
### Data
Rows: 757 | Columns: 43
Sample first rows:
- Month | Year | FY | Date | Job Nr | Invoice No. | Revenue Incl VAT | Revenue Excl VAT | Start Time  | End Time
- =TEXT(D2,"mmmm") | =YEAR(D2) |  | =_xlfn.XLOOKUP(E2,'Job Register '!$A:$A,'Job Register '!$B:$B,0) | 1 | =_xlfn.XLOOKUP(E2,'Job Register '!$A:$A,'Job Register '!$K:$K,0) | =SUMIFS('Job Register '!$M:$M,'Job Register '!$A:$A,Data!E2) | =SUMIFS('Job Register '!$N:$N,'Job Register '!$A:$A,E2) | =SUMIFS('Job Card Conversion'!$H:$H,'Job Card Conversion'!$E:$E,E2)*24 | =SUMIFS('Job Card Conversion'!$I:$I,'Job Card Conversion'!$E:$E,E2)*24
- =TEXT(D3,"mmmm") | =YEAR(D3) |  | =_xlfn.XLOOKUP(E3,'Job Register '!$A:$A,'Job Register '!$B:$B,0) | 2 | =_xlfn.XLOOKUP(E3,'Job Register '!$A:$A,'Job Register '!$K:$K,0) | =SUMIFS('Job Register '!$M:$M,'Job Register '!$A:$A,Data!E3) | =SUMIFS('Job Register '!$N:$N,'Job Register '!$A:$A,E3) | =SUMIFS('Job Card Conversion'!$H:$H,'Job Card Conversion'!$E:$E,E3)*24 | =SUMIFS('Job Card Conversion'!$I:$I,'Job Card Conversion'!$E:$E,E3)*24
### Jobs
Rows: 56 | Columns: 2
Sample first rows:
- Month | December
- Year | 2025
-  | 
### Job Card Conversion
Rows: 2340 | Columns: 17
Sample first rows:
- Date | Month | Year | FY | Job Nr | OPS manager | Description | Start Time  | End Time | Hours
- 2024-01-01 00:00:00 | =TEXT(Job_Card_April_202421[[#This Row],[Date]],"mmmm") | =YEAR(Job_Card_April_202421[[#This Row],[Date]]) | 2024 | 1 | Jolanda | Anuket Coral STS Colombo Express | 04:30:00 | 19:30:00 | =(Job_Card_April_202421[[#This Row],[End Time]]-Job_Card_April_202421[[#This Row],[Start Time ]])*24
- 2024-01-02 00:00:00 | =TEXT(Job_Card_April_202421[[#This Row],[Date]],"mmmm") | =YEAR(Job_Card_April_202421[[#This Row],[Date]]) | 2024 | 2 | Jolanda | Docking Bourbon Monsoon | 07:00:00 | 09:00:00 | =(Job_Card_April_202421[[#This Row],[End Time]]-Job_Card_April_202421[[#This Row],[Start Time ]])*24
### Settings
Rows: 6 | Columns: 2
Sample first rows:
- Divison Head | Division 
- Eugen | Diving
- Bresler | Diving
### Job Register 
Rows: 766 | Columns: 20
Sample first rows:
-  |  |  |  |  | Job Register |  |  |  | 
- Job_No | Date | Divison Head | Division  | Client | Job/Project_Description | Completion_Date | Quote_Nr | PO_No | Report No. and date issued
- 1 | 2024-01-01 00:00:00 | Jolanda | Pollution | SGM | Anuket Coral STS Colombo Express |  |  |  | 
### Salaries
Rows: 47 | Columns: 6
Sample first rows:
-  |  |  |  |  | 
-  | 2024 |  |  | 2025 | 
-  | Month | Total Salaries |  | Month | Total Salaries


## Main data table headers

Month, Year, FY, Date, Job Nr, Invoice No., Revenue Incl VAT, Revenue Excl VAT, Start Time , End Time, Total Hours, Labour_Costs, Equipment_Costs, Workshop Cost, Total_Costs, OPS Manager, Division

## Important finding

The workbook is already operating like a business management system inside Excel. The app should not start as a blank system. It should digitize the existing workflow and keep the Excel logic available for import/export validation.

## Core business calculation

Revenue Excl VAT - Total Costs = Gross Profit

Total Costs currently consists mainly of:

- Labour_Costs
- Equipment_Costs
- Workshop Cost

## Divisions found in dummy data

- Pollution: 493 jobs | Revenue N$ 52,332,161.02 | Profit N$ 36,527,018.24
- Diving: 365 jobs | Revenue N$ 23,204,451.68 | Profit N$ 14,346,691.65
- Civil: 28 jobs | Revenue N$ 6,441,445.19 | Profit N$ 3,671,713.60

## OPS Managers found

- Jolanda: 493 jobs | Revenue N$ 52,332,161.02 | Profit N$ 36,527,018.24
- Eugen: 348 jobs | Revenue N$ 20,275,851.97 | Profit N$ 13,176,293.81
- Carmon: 28 jobs | Revenue N$ 6,441,445.19 | Profit N$ 3,671,713.60
- Bresler: 17 jobs | Revenue N$ 2,928,599.71 | Profit N$ 1,170,397.85