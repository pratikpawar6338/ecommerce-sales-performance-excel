# Technical Documentation: Data Cleaning, Formulas & Business Q&A

This document provides the granular methodology behind the **Vrinda Store Annual Report 2022** Excel project for technical reviewers and hiring managers.

---

## 1. Dataset Overview & Data Dictionary
* **Total Records:** 31,047 transactional rows (after cleaning)
* **Time Period:** January 2022 – December 2022
* **Granularity:** One row per order line item

| Column Name | Data Type | Description / Cleaning Applied |
| :--- | :--- | :--- |
| `Order ID` | Text | Unique transaction identifier (used for `Count of Order ID`) |
| `Cust ID` | Integer | Unique customer identifier |
| `Gender` | Categorical | Standardized from `M/W/Men/Women` to `Men` and `Women` |
| `Age` | Integer | Customer age at time of purchase |
| `Age Group` | Categorical *(Engineered)* | `Young Adult` (<30), `Adult` (30–49), `Senior` (50+) |
| `Date` | Date | Order placement date |
| `Month` | Text *(Engineered)* | 3-letter month (`Jan`–`Dec`) extracted via `=TEXT(Date, "mmm")` |
| `Status` | Categorical | `Delivered`, `Cancelled`, `Returned`, `Refunded` |
| `Channel` | Categorical | Sales platform (`Amazon`, `Myntra`, `Flipkart`, `Ajio`, `Meesho`, `Nalli`, `Others`) |
| `Category` | Categorical | Standardized uppercase apparel category (`SET`, `KURTA`, `SAREE`, `TOP`, etc.) |
| `Qty` | Integer | Standardized numeric quantity (`One`/`Two` converted to `1`/`2`) |
| `Amount` | Currency (₹) | Transaction value in Indian Rupees |
| `ship-state` | Categorical | Destination Indian state (standardized uppercase) |

---

## 2. Step-by-Step Data Cleaning & Logic Log

### Step A: Data Quality Audit
* Applied AutoFilters (`Ctrl + Shift + L`) across all columns to inspect unique values for spelling errors, inconsistent casing, and blanks.
* Checked numeric vs. text alignment (identified text strings `"One"` and `"Two"` inside `Qty`).
* Identified trailing space in the `Channel ` header and mixed casing in `Category` (`kurta` vs `Set`).

### Step B: Feature Engineering Formulas
1. **Age Group Categorization:**
   ```excel
   =IF(E2>=50, "Senior", IF(E2>=30, "Adult", "Young Adult"))
   ```
   *Analyst Note:* Replaced generic `"Teenager"` catch-all logic with `"Young Adult"` so customers aged 20–29 are accurately represented in demographic reporting.
2. **Month Extraction:**
   ```excel
   =TEXT(G2, "mmm")
   ```
3. **Category Casing Standardization:**
   Applied `=UPPER(...)` (or `PROPER(...)`) and locked values via **Paste Values (`Alt + H + V + V`)** before refreshing Pivot caches.

### Step C: Dynamic KPI Card Architecture
To prevent `#REF!` or `GETPIVOTDATA` errors when linking dashboard shapes to Pivot Tables:
1. Created a dedicated helper Pivot Table (`Orders KPI` sheet) summarizing `Amount` and `Count of Order ID` by `Status`.
2. Built dynamic summary cells referencing the Pivot output:
   * **Realised Revenue:** `=SUMIF(A4:A7, "Delivered", B4:B7)` → **₹ 1,97,10,544**
   * **Lost Revenue:** `=IFERROR(B8 - Realised_Revenue_Cell, 0)` → **₹ 14,65,833**
   * **Total Orders:** Linked to Pivot Grand Total of `Count of Order ID` → **31,047**
   * **Average Order Value (AOV):** `=IFERROR(Gross_Sales / Total_Orders, 0)` → **₹ 682**
3. Linked rounded-rectangle shapes directly to these helper cells via the Formula Bar and connected all 3 dashboard slicers (`Months`, `Channel`, `Category`) to the `Orders KPI` Pivot Table using **Report Connections**.

---

## 3. Detailed Answers to Business Questions

| # | Business Question | Finding & Supporting Data |
| :-: | :--- | :--- |
| **1** | **Compare sales and orders in a single chart.** | Built a Combo Chart (Clustered Column for Total Sales on Primary Axis + Line for Total Orders on Secondary Axis). Orders and sales move in lockstep throughout the year, confirming stable AOV (~₹682) across months. |
| **2** | **Which month got the highest sales and orders?** | **March** peaked at **~₹1.93M** in sales and **~2,819 orders**, followed by February and January. Sales softened steadily from September through December. |
| **3** | **Who purchased more—Men or Women in 2022?** | **Women** generated **64%** of total sales vs. **36%** for **Men**. |
| **4** | **What are the different order statuses in 2022?** | **Delivered: 92%** (₹1.97 Cr realised), **Cancelled: 3%**, **Returned: 3%**, and **Refunded: 2%** (combined ₹14.66 Lakh in lost revenue). |
| **5** | **Which Top 5 states contribute most to sales?** | 1. **Maharashtra** (₹2.99M) · 2. **Karnataka** (₹2.65M) · 3. **Uttar Pradesh** (₹2.10M) · 4. **Telangana** (₹1.71M) · 5. **Tamil Nadu** (₹1.68M). The top 3 alone account for ~35% of total sales. |
| **6** | **What is the relationship between Age and Gender?** | **Adult Women (30–49)** lead at **34.59%** of all orders, followed by **Young Adult Women (<30)** at **21.13%**, **Adult Men** at **15.47%**, **Senior Women (50+)** at **13.70%**, **Young Adult Men** at **9.20%**, and **Senior Men** at **5.91%**. |
| **7** | **Which channels contribute maximum orders?** | **Amazon (35%)**, **Myntra (23%)**, and **Flipkart (22%)** account for **80%** of total order volume. |
| **8** | **What is the highest-selling category & AOV driver?** | **`SET`** drives ~50% of revenue with the highest AOV (~₹848), while **`KURTA`** generates high order volume at a lower AOV (~₹475)—creating a clear upsell/bundling opportunity. |