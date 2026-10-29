# Sales Data Analyzer (Excel Project)

A beginner-friendly Excel project that analyzes **200 sales orders** to answer simple business questions: *Which region sells the most? Which product is the best seller? Who are the top customers? How does profit relate to sales? Are monthly sales going up or down?*

**Main file:** `Project_2_Analyzer.xlsx`

---

## Table of Contents

1. [Project Goal](#1-project-goal)
2. [About the Data](#2-about-the-data)
3. [What Is Inside the Workbook](#3-what-is-inside-the-workbook)
4. [Sheet-by-Sheet Explanation](#4-sheet-by-sheet-explanation)
5. [Dashboard](#5-dashboard)
6. [Key Findings](#6-key-findings)
7. [Excel Features and Formulas Used](#7-excel-features-and-formulas-used)
8. [How to Use This Workbook](#8-how-to-use-this-workbook)
9. [Requirements](#9-requirements)
10. [Notes and Limitations](#10-notes-and-limitations)

---

## 1. Project Goal

Turn a raw table of sales orders into useful insights by using Excel tools:

- Highlight important values automatically (conditional formatting)
- Summarize data quickly (pivot tables)
- Understand the data with statistics and a regression line
- Track monthly growth with arrows
- Find high-value customers with a filter formula
- Show everything on one visual dashboard

---

## 2. About the Data

- **Rows:** 200 orders
- **Time period:** 10 April 2024 to 7 April 2025
- **Regions:** Central, East, North, South, West
- **Product categories:** Books, Clothing, Electronics, Furniture, Office Supplies

| Column | Meaning | Example |
|---|---|---|
| `Customer_ID` | Unique code for the customer | CUST019 |
| `Customer_Name` | Name of the customer | James Davis |
| `Region` | Area where the order was placed | West |
| `Product_Category` | Type of product sold | Books |
| `Sales` | Money earned from the order | 1036.09 |
| `Quantity` | Number of units sold | 4 |
| `Discount` | Discount given (0.20 = 20%) | 0.2 |
| `Order_Date` | Date of the order | 10-Apr-2024 |
| `Profit` | Money left after costs | 331.55 |

**Quick numbers:** total sales = **195,217.76**, total profit = **68,287.01**, average order = **976.09**, overall profit margin = about **35%**.

---

## 3. What Is Inside the Workbook

The workbook has **10 sheets**. Each sheet focuses on one task.

| # | Sheet name | One-line purpose |
|---|---|---|
| 1 | `Conditional_formatting` | The original data, with the biggest sales highlighted |
| 2 | `Compare_Discount` | A small calculator showing how a discount changes the final price |
| 3 | `Linear_regression` | Shows how profit changes as sales change |
| 4 | `Descriptive_statistics` | Statistical summary of the sales data |
| 5 | `Arrow_for_Monthly_sale` | Monthly sales and growth, shown with up/down arrows |
| 6 | `TimeStamp` | Adds the current date and time to records |
| 7 | `High_vslue_Cus` | Finds high-value customers using a filter |
| 8 | `Pivot_table` | Pivot tables for sales by region and by product |
| 9 | `Pie_Line_Bar chart` | The charts built from the pivot tables |
| 10 | `Dashbord` | One page with all four charts together |

> Sheet names are written exactly as in the file (including spelling such as `Dashbord` and `High_vslue_Cus`).

---

## 4. Sheet-by-Sheet Explanation

### 1) `Conditional_formatting`
**What it is:** The main dataset (all 9 columns, 200 rows).
**What it does:** Applies **Top 10** rules to the `Sales` column so the highest values stand out in color automatically. No manual searching is needed.
**How to read it:** Look for the colored cells in the Sales column. These are your biggest orders.

### 2) `Compare_Discount`
**What it is:** A "what-if" calculator for a manager.
**Example used:** 10 units of a PS5 at 50,000 each with a 40% discount.

| Item | Formula | Result |
|---|---|---|
| Total Amount | `Quantity x Price` | 500,000 |
| Discount Amount | `Total Amount x Discount` | 200,000 |
| Final Amount | `Total Amount - Discount Amount` | 300,000 |

**Tip:** Change the quantity, price, or discount and the results update instantly.

### 3) `Linear_regression`
**What it is:** A way to predict `Profit` from `Sales`.
**The idea:** Draw the best straight line through the data: `y = m x + c`
- `m` (slope) = `SLOPE(Profit, Sales)`. It tells you how much profit rises for each extra unit of sales.
- `c` (intercept) = `INTERCEPT(Profit, Sales)`. It is the starting value when sales are zero.
- Column `K` calculates the **predicted profit** for every order.

**Result:** The slope is about **0.35**, so on average **every 100 of extra sales adds roughly 35 of profit**. The relationship is strong and positive (correlation about 0.77).

### 4) `Descriptive_statistics`
**What it is:** A statistical summary of the sales data (column `J`).
**Why it matters:** Statistics such as the average, spread, minimum, and maximum give a quick feel for the data before deeper analysis.

For reference, these are the key figures for the dataset:

| Measure | Sales | Profit |
|---|---:|---:|
| Average | 976.09 | 341.44 |
| Middle value (median) | 976.72 | 332.26 |
| Smallest | 258.51 | 64.63 |
| Largest | 1,923.66 | 769.46 |

### 5) `Arrow_for_Monthly_sale`
**What it is:** A monthly sales tracker.
**Columns:**
- `month`: Apr to Dec
- `Monthly_Sales`: total sales for that month using `SUM` over that month's rows
- `Monthly_Sales_Growth`: `(Next month - This month) / This month`, the percentage change

**Arrows:** Icon sets show a quick visual signal:
- Green up arrow = sales grew
- Yellow sideways arrow = about the same
- Red down arrow = sales dropped

### 6) `TimeStamp`
**What it is:** The data with an extra `TimeStamp` column using `=NOW()`.
**What it does:** Shows the current date and time, and updates each time the workbook recalculates. Useful for showing when a report was last refreshed.

### 7) `High_vslue_Cus`
**What it is:** A list of important customers (this sheet uses its own sample list of 20 customers).
**What it does:**
- Uses the `FILTER` function to show only customers with **Sales above 50,000 AND Profit above 10,000**. If nobody matches, it shows "No customers found".
- Uses `INDEX` + `MATCH` to look up a customer's sales value by name.
- Uses `INDEX` + `MATCH` + `MAX` to find the **name of the customer with the highest sales**.

### 8) `Pivot_table`
**What it is:** Two pivot tables that add up sales automatically.
- **Total sales by Region**
- **Total sales by Product**

| Region | Total Sales |
|---|---:|
| West | 44,083.58 |
| Central | 43,468.33 |
| South | 40,248.79 |
| North | 34,975.42 |
| East | 32,441.64 |
| **Grand Total** | **195,217.76** |

| Product | Total Sales |
|---|---:|
| Books | 47,002.46 |
| Clothing | 41,927.59 |
| Electronics | 41,044.00 |
| Furniture | 36,612.89 |
| Office Supplies | 28,630.82 |

### 9) `Pie_Line_Bar chart`
**What it is:** The pivot tables turned into four charts: Pie, Bar, Line, and 3D Bar.

### 10) `Dashbord`
**What it is:** The final summary page that places all four charts together. See the next section.

---

## 5. Dashboard

The `Dashbord` sheet is the quickest way to see the results.

| Chart | Type | What it shows |
|---|---|---|
| Total sales by Region | Pie | Each region's share of total sales |
| Total sales by product | Bar | Which product categories sell the most |
| Total sales by region | Line | Region totals compared side by side |
| Top 10 Highest sales customers | 3D Bar | The biggest sales values |

---

## 6. Key Findings

- **Best region:** West (44,083.58), followed closely by Central. **East** is the lowest.
- **Best product:** Books (47,002.46). **Office Supplies** is the lowest.
- **Profit link:** Higher sales usually mean higher profit (about 35 profit per 100 sales).
- **Discounts:** Range from 0% to 20%, averaging about 10%.
- **Monthly trend:** Sales go up and down through the year, so the arrows help spot the strong and weak months.

---

## 7. Excel Features and Formulas Used

| Feature | Used for |
|---|---|
| `SUM` | Monthly sales totals |
| `SLOPE`, `INTERCEPT` | Regression line |
| `NOW` | Timestamp |
| `FILTER` | Pulling out high-value customers |
| `INDEX` + `MATCH` | Looking up values and names |
| `MAX` | Finding the largest sales value |
| Conditional formatting (Top 10) | Highlighting biggest sales |
| Icon sets (3 arrows) | Showing growth direction |
| Pivot tables | Summarizing by region and product |
| Charts (Pie, Bar, Line, 3D Bar) | Visual dashboard |

---

## 8. How to Use This Workbook

1. Download and open `Project_2_Analyzer.xlsx` in Excel.
2. Start with the **`Dashbord`** sheet for the big picture.
3. Open the other sheets to see how each result was calculated. Click any cell with a formula to see how it works.
4. If you change the data, go to **Data > Refresh All** so the pivot tables and charts update.
5. To try the discount calculator, change the numbers in `Compare_Discount`.

---

## 9. Requirements

- Microsoft Excel 2021 or Microsoft 365 (needed for the `FILTER` function)
- No macros or add-ins required

---

## 10. Notes and Limitations

- `Arrow_for_Monthly_sale` adds up **fixed row ranges** for each month. If you sort the data or add rows, update those ranges.
- `High_vslue_Cus` uses a separate sample list of customers, not the main 200 orders.
- `TimeStamp` uses `NOW()`, so the value changes every time Excel recalculates.
- Results are based on this sample dataset only and should not be treated as real-world forecasts.

---

## Author

_Add your name here_
