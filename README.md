# Enterprise-Advanced-Microsoft-Excel-Portfolio: Finance, Operations & Supply Chain Analytics

![Microsoft Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Power Query](https://img.shields.io/badge/Power_Query-M_Language-F2C811?style=for-the-badge&logo=powerbi&logoColor=black)
![Power Pivot](https://img.shields.io/badge/Power_Pivot-DAX-0078D4?style=for-the-badge&logo=microsoft&logoColor=white)
![VBA](https://img.shields.io/badge/VBA-Macros_%26_Automation-3665A4?style=for-the-badge&logo=visual-studio-code&logoColor=white)

## Overview
Enterprise Advanced Microsoft Excel Portfolio: Solutions for complex operational, financial, and supply chain environments. Demonstrate skills in handling hundreds of thousands of rows, automation of multi-source ETL pipelines, building relational Data Models, and deploying VBA-driven automation for BI, Manufacturing, Logistics, Corporate Finance.


> [!NOTE] NOTE
> Since finding open-source datasets for real-world companies is rare (due to the confidential nature of the information), these projects are done using a combination of open-source sample databases from select sources (e.g. Microsoft, Kaggle) and realistic synthetic datasets that reflect real-world environments.



## Core Stack & Skills
* **ETL & Data Engineering (Power Query / M):** 
	* Automated folder ingestion, API connections, PDF table extraction, unpivoting messy ERP exports, custom M functions, and fuzzy matching.
* **Data Modeling & BI (Power Pivot / DAX):** 
	* Star-schema architecture, handling 500k+ rows bypassing the 1M worksheet limit via the Data Model, time-intelligence measures (`SAMEPERIODLASTYEAR`, `DATESYTD`), and KPI disconnected slicers.
* **Advanced Formula Engine:** 
	* Dynamic Array functions (`LET`, `LAMBDA`, `MAP`, `BYROW`, `XLOOKUP`, `FILTER`, `UNIQUE`, `SORTBY`) for real-time calculation engines.
* **Financial Modeling:** 
	* 3-Statement modeling, Monte Carlo simulations, sensitivity analysis (Data Tables), WACC, DCF, and working capital optimization.
* **Process Automation (VBA & Office Scripts):** 
	* Automated PDF report bursting, ERP data sanitization, Outlook email integration, and custom UserForms.



## Projects 

### Project 1. Corporate Finance & FP&A: Rolling 3-Statement Model & Working Capital Simulator
- **Domain:**
	- Corporate Finance / FP&A / Treasury  
- **Data Volume:**
	- ~350,000 rows | **Refresh Type:** General Ledger (GL) Pipeline + Monte Carlo VBA

#### Business Problem
To remove reliance on static, fragile monthly budget spreadsheets that take days to close and  reconcile. Add the ability to stress-test cash flow covenants and working capital requirements against raw material price inflation, FX volatility, and customer payment delays.

#### Data Sources & Architecture

Leveraged the industry-standard Microsoft Contoso Data Warehouse to simulate a massive, enterprise-scale ERP environment. Engineered a Power Query pipeline to bypass the limitations of traditional VLOOKUPs, demonstrating how modern Excel handles big-data financial modeling.

* `SAP_GL_Transactions.csv`:
	* 300,000 rows of raw General Ledger journal entries across 4 fiscal years.
* `AR_AP_Subledger.xlsx`:
	* 50,000 rows of open and closed customer invoices and vendor bills (Aging analysis).
* `Chart_of_Accounts_Mapping.xlsx`:
	* Hierarchical mapping table linking GL account numbers to Income Statement, Balance Sheet, and Cash Flow lines.

#### Technical Implementation
* **Power Query GL Engine:** 
	* Automated trial balance generation by ingesting 300k raw journal entries, stripping formatting errors, and mapping accounts to financial statement lines without a single `VLOOKUP`.
* **3-Statement Financial Model:** 
	* Fully integrated **Income Statement, Balance Sheet, and Cash Flow Statement** (Direct & Indirect methods) driven by a centralized assumptions tab (Base, Upside, Downside cases).
* **Working Capital Cash Conversion Cycle (CCC):** 
	* Granular cohort analysis calculating **DSO** (Days Sales Outstanding), **DIO** (Days Inventory Outstanding), and **DPO** (Days Payable Outstanding) directly from subledger invoice timestamps.
* **VBA Monte Carlo Simulation:**
	* Custom VBA routine (`modMonteCarlo.bas`) running 5,000 iterations across stochastic variables (COGS inflation, interest rates, volume growth) to output a probability distribution of **Free Cash Flow (FCF)** and **EBITDA**.

#### Results
* Executive **P&L Variance Bridge (Waterfall Chart)** decomposing Budget vs. Actuals into Price, Volume, and Mix effects.
* Automated liquidity stress-testing dashboard monitoring debt service coverage ratios (DSCR).


---

### Project 2. Supply Chain & Logistics: Multi-Echelon Inventory & Freight Optimization
- **Domain:** 
	- Supply Chain / Global Logistics / Procurement  
- **Data Volume:**
	- ~265,000 rows | **Refresh Type:** Multi-Source Relational Model

#### Business Problem
- Supply chain planners faced simultaneously high stockout rates on fast-moving SKUs and excess working capital tied up in slow-moving inventory across 6 regional distribution centers (DCs), alongside unpredictable freight surcharges from 3PL carriers.

#### Data Sources & Architecture
* `WMS_Inventory_Ledger.accdb`:
	* Microsoft Access database table (150,000 rows) of daily stock movements, receipts, and cycle counts.
* `/Carrier_Invoices_PDF/`: 
	* 50+ multi-page carrier freight invoices in PDF format containing lane rates, fuel surcharges, and accessorial fees.
* `ERP_Purchase_Orders.csv`:
	* 115,000 rows of historical POs, supplier lead times, and promised vs. actual delivery dates.
* `Live FX Rates (Web API)`:
	* Real-time currency conversion for international freight billing.

#### Technical Implementation
* **Power Query (PDF & DB Ingestion):**
	* Extracted tabular invoice data across dozens of PDF files, appended Access DB ledger entries, and connected to a live REST API to normalize USD, EUR, and KES/GBP freight costs.
* **Dynamic Arrays & Statistical Modeling:** 
	* Used `LET` and `LAMBDA` functions to calculate dynamic **Safety Stock**, **Reorder Points (ROP)**, and **Economic Order Quantity (EOQ)** factoring in both demand volatility ($\sigma_D$) and supplier lead-time variability ($\sigma_{LT}$).
* **ABC-XYZ Segmentation Matrix:** 
	* Automated a 9-box inventory classification matrix combining revenue volume (ABC) with coefficient of variation in demand (XYZ).
* **Solver & VBA:**
	- Configured an Excel Solver routine wrapped in VBA to determine optimal stock transfer routes between DCs to minimize freight cost per ton-kilometer.

#### Results
* **OTIF (On-Time In-Full) Vendor Scorecard** ranking 80+ suppliers on lead-time reliability and fill rates.
* Freight audit module flagging carrier invoice discrepancies against contracted lane rates.

---
### Project 3. Manufacturing & Production: OEE, Downtime & Scrap Cost Intelligence
- **Domain:**
	- Manufacturing / Production Engineering / Business Intelligence  
- **Data Volume:**
	- ~420,000 rows | **Refresh Type:** Automated Folder ETL

#### Business Problem
- A multi-plant manufacturing operation suffered from fragmented reporting. Machine telemetry logs, ERP production schedules, and quality assurance (QA) scrap logs were stored in separate systems, delaying root-cause analysis on machine downtime and inflating unit production costs.

#### Data Sources & Architecture
* `/Telemetry_Logs/`:
	* 24 monthly CSV dumps (~350,000 rows) of machine cycle times and PLC fault codes.
* `ERP_Production_Orders.xlsx`:
	* 45,000 rows of planned vs. actual production runs, shift schedules, and standard BOM costs.
* `QA_Defect_Logs.txt`: 
	* Tab-delimited text file (25,000 rows) tracking material scrap, rework hours, and defect categories.
* `Machine_Master.json`: 
	* Equipment hierarchy, ideal cycle times, and maintenance schedules.

#### Technical Implementation
* **Power Query (ETL):** 
	* Consolidated 24 CSV files from a local directory using binary combine; parsed `.json` machine specs; cleaned shift timestamps and merged QA logs with production order IDs.
* **Power Pivot & DAX:**
	* Built a Star Schema connecting 3 Fact tables (Telemetry, Orders, Defects) to 4 Dimension tables (Calendar, Machines, Shifts, Products). Authored DAX measures for **Availability %**, **Performance %**, **Quality %**, and composite **OEE (Overall Equipment Effectiveness)**.
* **VBA Automation:** 
	* Built a one-click macro (`modExportPlantReports.bas`) that loops through each Plant Manager, filters the dashboard, exports a high-resolution PDF summary, and drafts an automated shift-handover email via Outlook.

#### Results
* Interactive Pareto analysis identifying the top 20% of fault codes causing 80% of unscheduled line stops.
* Real-time variance tracking of **Standard vs. Actual Cost per Unit**, isolating material scrap impact.


---

## Repo Structure

```text
├── 01_Finance_FPA_3Statement_Model/
│   ├── Data_Sources/
│   │   ├── SAP_GL_Transactions.csv
│   │   ├── AR_AP_Subledger.xlsx
│   │   └── Chart_of_Accounts_Mapping.xlsx
│   ├── VBA_Modules/
│   │   └── modMonteCarlo.bas
│   ├── FPA_WorkingCapital_Model.xlsm
│   └── README.md
│
├── 02_SupplyChain_Inventory_Freight/
│   ├── Data_Sources/
│   │   ├── Carrier_Invoices_PDF/    # Raw 3PL PDF invoices
│   │   ├── WMS_Inventory_Ledger.accdb
│   │   └── ERP_Purchase_Orders.csv
│   ├── VBA_Modules/
│   │   └── modRouteOptimizer.bas
│   ├── SupplyChain_Control_Tower.xlsm
│   └── README.md
│
├── 03_Manufacturing_OEE_Analytics/
│   ├── Data_Sources/
│   │   ├── Telemetry_Logs/          # Monthly CSV machine logs
│   │   ├── ERP_Production_Orders.xlsx
│   │   ├── QA_Defect_Logs.txt
│   │   └── Machine_Master.json
│   ├── VBA_Modules/
│   │   └── modExportPlantReports.bas
│   ├── Manufacturing_OEE_Suite.xlsm
│   └── README.md
└── 
