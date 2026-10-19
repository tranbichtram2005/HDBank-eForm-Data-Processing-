# HDBank eForm Data Processing 

## Overview

This repository contains Python/Pandas scripts used to validate, standardize and transform eForm survey data on over-the-counter transactions, serving as the input for a Power BI dashboard.

The raw survey data (Excel) is in a wide format, where each business process has several answer columns. The scripts convert the data into **Long Format** (Unpivot) to make data modeling and visualization easier.

### Scripts

| File | Description |
|------|-------------|
| `buildstarschema.py` | Reads survey data from Excel, automatically detects business processes, and unpivots them into a `Fact_NghiepVu` table (support level, processing time, error type). Cross-checks related columns to flag contradictory responses with `ĐVKD_Flag`, and adds sort-order columns (`Thu_Tu_NV`, `Thu_Tu_SLA`) so items display in the correct order in Power BI. Also extracts the open-ended feedback block (TH1, TH2) into `Fact_GopY`.Standardizes column names and unpivots the feedback data (TH1, TH2) into Long Format. Classifies each scenario as error / no error (`Trang_Thai_TinhHuong`) and flags cases that may have been missed (`Co_The_Bo_Sot`) |

### Outputs



### Tools

Microsoft Excel, Python (Pandas, NumPy), Microsoft Power BI.
