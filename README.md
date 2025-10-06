# PULSE — Platform for Unified Learning & Streaming Evidence
> Educational analytics project for the fictional company **PULSE** — an international online retailer of digital music operating across several countries.
---

## Project Overview
This repository contains the complete analytics environment for **PULSE Analytics**, built on top of the **Chinook PostgreSQL dataset** adapted for learning and streaming data.  
It demonstrates an end-to-end analytical pipeline:

- Importing structured datasets (Artists, Albums, Tracks, Genres, Customers, Invoices, etc.)  
- Ensuring data integrity and schema validation  
- Writing advanced SQL queries with **multiple JOINs** for real insights  
- Generating visual analytics using **matplotlib** and **Plotly**  
- Exporting styled reports to Excel via **openpyxl**  

---

## 📊 Main Analytics (Screenshots)
![Pie Chart](charts/01_pie_revenue_by_category.png)  
![Bar Chart](charts/02_bar_top_sellers_by_revenue.png)  
![Horizontal Bar](charts/06_duration_by_genre_strip.png)  

---


## 📦 Dataset
Public **Chinook Database** — a realistic model of a digital media store used here to simulate PULSE’s operational data.  
It includes:
- Artists, Albums, Tracks, Genres  
- Customers, Employees, Invoices, and InvoiceLines  
- Relational links between content creators, clients, and sales operations  

> 📚 Source: [Chinook Database (official sample)](https://github.com/lerocha/chinook-database)

---

## 📈 Key Analytics & Deliverables

### SQL Business Queries
All analytics rely on **2+ table JOINs**, designed to answer meaningful business questions such as:
- Revenue share by genre  
- Top-performing artists by total revenue  
- Average order value by country  
- Monthly revenue growth trend  
- Track duration distribution  
- Correlation between price and duration (top genres)

### Static Reporting (matplotlib)
- Six chart types: **pie, bar, horizontal bar, line, histogram, scatter**  
- Automatically saved in `/charts/`  
- Each chart includes titles, axis labels, and legends (if needed)  
- Console report includes number of rows and chart purpose  

### Interactive Visualization (Plotly)
- Dynamic **bar chart with a time slider** (`animation_frame="month"`) to visualize monthly revenue by country  
- Output saved as `/charts/timeslider_revenue_by_country.html`  
- Used during demo to show interactivity  

### Excel Export (openpyxl)
- All query results exported to `/exports/`  
- Features:
  - Frozen header + first column  
  - Filters enabled on all columns  
  - 3-color gradient for numeric columns  
  - Conditional formatting for min/max  
- Example console log:
  ```text
  Created file excel_20251006_112530.xlsx, 6 sheets, 1234 rows
  Full path: /exports/excel_20251006_112530.xlsx

