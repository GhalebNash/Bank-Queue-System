# Bank Queue Management System (Excel + VBA)

## Overview
An Excel-based discrete-event simulation of a bank queue. Automates customer check-in, service assignment, and KPI reporting. Uses **VBA macros** and **Solver** to test staffing scenarios and reduce wait times.

## Features
- One-click **Start/Reset** buttons (VBA)
- Real-time queue and teller status
- KPI dashboard: avg wait, service time, throughput, utilization
- Scenario testing with **Solver** (e.g., optimal # of tellers)
- Exportable daily performance report

## File(s)
- `BankQueueSystem.xlsm` – main workbook with macros and dashboard  
- `screenshots/` – UI and dashboard previews

## How to Run
1. Download `BankQueueSystem.xlsm`.
2. Open in **Microsoft Excel (Desktop)** and **Enable Macros**.
3. On the **Dashboard** sheet, click **Start Simulation**.
4. Use the **Scenario** sheet to change arrivals/service rates, then run **Solver** to find optimal staffing.


## Methods & Tools
- **Excel, VBA**, **Solver**
- Queuing concepts (arrival rate λ, service rate μ, utilization ρ)
- KPI design and data visualization

## What I Learned
- Building a small simulation engine in Excel/VBA  
- Designing decision dashboards for business ops  
- Formulating optimization problems with Solver
