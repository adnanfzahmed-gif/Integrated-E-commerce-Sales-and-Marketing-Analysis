# 🛒 E-Commerce Sales Dashboard & Business Insights

![Excel](https://img.shields.io/badge/Microsoft%20Excel-2019%20%7C%202021%20%7C%20365-217346?logo=microsoftexcel&logoColor=white)
![Pivot Tables](https://img.shields.io/badge/Pivot%20Tables-Slicers%20%26%20Charts-blue)
![Focus](https://img.shields.io/badge/Focus-Data%20Analysis-success)
![License](https://img.shields.io/badge/Use-Educational-lightgrey)

> Sales data analysis and an interactive dashboard built entirely in **Microsoft Excel**, using data cleaning, descriptive statistics, Pivot Tables, Pivot Charts and Slicers to turn 2,000 retail transactions into business insights.

---

## 📌 Project Overview

This project analyses retail sales data to show how a business is performing and where the opportunities are. The dataset covers customer details, product categories, store information, sales amounts, discounts and payment methods.

**What it delivers**
- 🧹 A cleaned dataset (duplicates removed, missing values handled, formats checked)
- 📐 Descriptive statistics calculated with Excel formulas
- 📊 Pivot Tables and Pivot Charts for the key sales views
- 🎛️ An interactive dashboard with Slicers and filters to explore sales by category, region and customer segment

---

## 🗂️ Dataset

A star-schema workbook (`Ecommerce_Sales_Dataset_modified.xlsx`) with one fact table and three dimension tables, joined into a single analysis table.

| Sheet | Rows | Contents |
|---|---|---|
| `Sales_Fact` | 2,000 | Order date, customer, product, store, quantity, unit price, discount, payment type, total amount |
| `Customer_Dim` | 500 | Name, age, gender, city, state, loyalty level |
| `Product_Dim` | 100 | Product name, category, sub-category, brand, cost, stock |
| `Store_Dim` | 20 | Store name, region, city, store type |
| `Sales Analysis Table 1` | 2,000 | Joined table with Month, Year and Age Range added |

Other sheets: `Sales Pivot`, `All pivots`, `Statistical Analysis` and `Ecommerce Sales Dashboard`.

**Period covered:** 14 Oct 2023 to 13 Oct 2025 (2 years), 492 distinct customers.

**Calculation:** `Total_Amount = Quantity × Unit_Price × (1 − Discount)`.

---

## 🧰 Tools & Requirements

| Requirement | Details |
|---|---|
| Software | Microsoft Excel 2019, Excel 2021 or Microsoft 365 |
| Operating system | Windows 10/11 or macOS |
| Skills | Basic Excel formulas and Pivot Tables |

---

## 🚀 Getting Started

1. **Download** the Excel project file (`.xlsx`).
2. **Open** it in Microsoft Excel.
3. **Enable editing** if prompted.
4. **Refresh** the Pivot Tables if the data has been updated (right-click a Pivot Table, then Refresh).

---

## 🔧 Steps Performed

### Step 1: Open the workbook
Open the Excel workbook and review the raw data.

### Step 2: Clean the data
- Remove duplicates (none found in the source tables)
- Handle missing values
- Check data formats

| Table | Missing values found |
|---|---|
| `Sales_Fact` | Quantity 5, Unit_Price 3, Discount 4, Total_Amount 4 |
| `Customer_Dim` | Loyalty_Level 4, plus a stray unnamed numeric column |
| `Product_Dim` | Cost 6, Stock 3 |

The analysis table has no blanks, with the missing `Total_Amount` values recalculated from the formula above.

### Step 3: Descriptive statistics
Calculated with Excel formulas (2,000 orders):

| Measure | Excel function | Total amount | Unit price | Discount | Customer age | Quantity |
|---|---|---|---|---|---|---|
| Mean | `AVERAGE` | 1,087.20 | 513.58 | 14.8% | 41.6 | 2.48 |
| Median | `MEDIAN` | 868.72 | 522.69 | 14.0% | 42 | 2 |
| Mode | `MODE.SNGL` | n/a | 213.69 | 10% | 29 | 2 |
| Minimum | `MIN` | 18.46 | 20.09 | 0% | 18 | 1 |
| Maximum | `MAX` | 3,908.24 | 999.89 | 30% | 64 | 4 |
| Variance (sample) | `VAR.S` | 671,775 | 79,721 | 0.0077 | 184.3 | 1.24 |
| Variance (population) | `VAR.P` | 671,439 | 79,681 | 0.0077 | 184.2 | 1.23 |
| Std deviation (sample) | `STDEV.S` | 819.62 | 282.35 | 0.088 | 13.58 | 1.11 |
| Std deviation (population) | `STDEV.P` | 819.41 | 282.28 | 0.088 | 13.57 | 1.11 |
| Skewness | `SKEW` | 0.88 | −0.03 | 0.06 | −0.04 | 0.02 |
| Kurtosis | `KURT` | 0.05 | −1.19 | −1.21 | −1.19 | −1.34 |

`Count` is 2,000 for every column.

### Step 4: Pivot Tables and Pivot Charts
| View | Question answered |
|---|---|
| Total sales | How much did the business sell? |
| Sales by product category | Which categories earn the most? |
| Sales by region | Where is demand strongest? |
| Sales by gender | How do customer groups differ? |
| Sales by payment method | How do customers prefer to pay? |
| Top products | What sells best? |
| Top stores | Which stores perform best? |

### Step 5: Interact with the dashboard
Use **Slicers** and **Filters** to slice the metrics above by any dimension.

---

## 📈 Key Insights

### Headline numbers
| Metric | Value |
|---|---|
| **Total sales** | **2,174,395** |
| Orders | 2,000 |
| Average order value | 1,087 (median 869) |
| Units sold | 4,964 |
| Average discount | 14.8% |

Order values are right-skewed (skewness 0.88): a minority of large orders pulls the mean above the median.

### Sales by product category
| Category | Sales | Share |
|---|---|---|
| Sports | 542,010 | 24.9% |
| Electronics | 504,311 | 23.2% |
| Home | 427,296 | 19.7% |
| Clothing | 379,216 | 17.4% |
| Beauty | 321,561 | 14.8% |

### Sales by region
| Region | Sales | Share | Stores | Sales per store |
|---|---|---|---|---|
| North | 646,101 | 29.7% | 6 | 107,683 |
| South | 635,213 | 29.2% | 6 | 105,869 |
| East | 622,404 | 28.6% | 6 | 103,734 |
| West | 270,676 | 12.4% | 2 | 135,338 |

West has the lowest total only because it has just 2 stores; per store it is the **strongest region**.

### Store type
| Type | Sales | Share | Stores |
|---|---|---|---|
| Online | 914,254 | 42.0% | 8 |
| Flagship | 848,577 | 39.0% | 8 |
| Outlet | 411,565 | 18.9% | 4 |

### Customers
| Segment | Finding |
|---|---|
| **Gender** | Women generate 56.6% of sales from 54.8% of orders, with a higher average order (1,124 vs 1,043 for men) |
| **Age** | The 46-55 group leads (499,637, 23.0%), followed by 36-45 (22.0%); 18-25 is smallest (16.4%) but spends the most per order (1,149) |
| **Loyalty** | Platinum is the top tier at 30.2% of sales; Silver is lowest at 21.2% |
| **Payment** | PayPal 34.8%, Credit Card 32.9%, Cash on delivery 32.3%: almost an even three-way split |
| **Top customer** | CUST1219, with 17,148 in purchases |

### Top performers
| | Name | Sales |
|---|---|---|
| Top product | Portable Power Bank | 41,853 |
| 2nd | Shin Guards (Adult) | 38,281 |
| 3rd | Men's Casual Sneakers | 35,490 |
| Top store | Bean, Jones and Park | 158,971 (7.3% of sales) |
| 2nd | Johnson-Wright | 125,012 |

### Time patterns
- **Best sales month:** March 2025 (107,589), then March 2024 (103,116) and November 2024 (102,983).
- **Peak weekdays:** Tuesday (334,889) and Monday (333,980); Friday is the slowest (287,860).
- **Full-year comparison:** 2024 totalled 1,093,912. 2023 (from 14 Oct) and 2025 (to 13 Oct) are partial years, so compare them month by month rather than by total.
- **Dips:** May 2024 (57,825) and June 2025 (65,636) were weak months before recovering.

### 🎯 Recommendations
1. **Expand in the West.** Its 2 stores outsell the others by roughly 25% each, so a third store is worth testing.
2. **Grow the Outlet format cautiously.** Outlets are 19% of sales from 20% of stores, and sell less per store than Online stores.
3. **Promote to the 46-55 and 36-45 groups,** which make up 45% of sales, and use offers to lift the 18-25 group's order count since it spends the most per order.
4. **Support every payment method.** No single option dominates, so removing any would risk about a third of sales.
5. **Plan inventory around March and November,** the recurring peaks, and look into the May and June dips.

---

## ⚠️ Data Quality Notes

- **Product, category, region and store columns in `Sales Analysis Table 1` are wrong.** Every one of the 2,000 rows shows the same product (Cricket Bat, PROD97), store (STORE3), region (North) and store type (Outlet); only about 1% match the real product in `Sales_Fact`. The lookup formulas seem to have been filled with fixed references. As a result, the category, region and top product/store pivots in the workbook show a single value at 100%. The figures in this README are recalculated by joining `Sales_Fact` to the product and store tables on their IDs. To fix the workbook, rebuild those columns with `XLOOKUP` or `INDEX/MATCH` using each row's own Product_ID and Store_ID.
- **Unaffected measures:** total sales, gender, age range, loyalty, payment type, dates and the statistics table all agree with the source data.
- **16 orders (0.8% of sales) have no loyalty level** because 4 customers have it missing.
- **Average discount is 14.8%,** and 18.7% of orders carry a discount of 25% or more; margin impact cannot be assessed without the cost column, which is incomplete.

---

## 🆘 Help

| Problem | Fix |
|---|---|
| Pivot Table is not updating | Right-click the Pivot Table and select **Refresh** |
| Charts do not reflect new data | Refresh the Pivot Table, or update the data source range |
| Formula errors | Check the cell references and make sure the data is numeric |
| Every row shows the same product or store | Check that lookup formulas use relative references to each row's own ID |

---

## 🕓 Version History

**Version 1.0**
- Initial dataset analysis
- Data cleaning
- Statistical analysis
- Interactive Excel dashboard creation

---

## 👤 Author

**Adnan Ahmed**
Data Analyst Project

---

## 📄 License

This project is created for educational purposes.

## 🙏 Acknowledgments

- Microsoft Excel documentation
- Data analysis and visualisation concepts
- Excel Pivot Table and dashboard techniques

---

⭐ If you found this useful, consider starring the repo.
