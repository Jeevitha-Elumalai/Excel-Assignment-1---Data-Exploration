# 📊 Product Data Analysis Using Excel

## 📌 Project Overview

This project focuses on **Product Data Analysis using Microsoft Excel**. It involves organizing product information, exploring product prices and quantities, categorizing products, and analyzing data using essential Excel functions and text functions.

The project demonstrates how Excel can be used to summarize data, identify price ranges, and extract useful information from product IDs.

## 🎯 Project Objectives

- Organize and analyze product data.
- Calculate the sum, count, average, minimum, and maximum values.
- Analyze data using conditional functions such as `SUMIF` and `COUNTIF`.
- Categorize products based on their price ranges.
- Extract specific information using text functions.
- Improve data analysis and Excel skills.

## 📁 Dataset Description

The dataset contains product information with the following columns:

| Column Name | Description |
|---|---|
| Product ID | Unique identifier for each product |
| Product Name | Name of the product |
| Brand Name | Brand associated with the product |
| Price ($) | Price of the product in dollars |
| Quantity | Quantity of products |
| Category | Category to which the product belongs |
| Price Range | Classification of the product by price |
| Day | Day associated with the product record |
| Country Code | Country identifier |
| Month | Month associated with the product record |

## 🗂️ Sample Dataset

| Product ID | Product Name | Brand Name | Price ($) | Quantity | Category | Price Range | Day | Country Code | Month |
|---|---|---|---:|---:|---|---|---:|---|---|
| 28-JAN-US | Laptop | Dell | 1000 | 30 | Electronics | High Price | 28 | US | JAN |
| 15-FEB-US | Sneakers | Nike | 80 | 15 | Fashion | Standard Price | 15 | US | FEB |
| 03-MAR-US | Coffee Maker | Keurig | 130 | 40 | Kitchen | Standard Price | 3 | US | MAR |

## 🧮 Excel Functions Used

The following Excel functions were used to perform calculations and explore the product dataset.

| Excel Function | Result | Purpose |
|---|---:|---|
| SUM | 10,100 | Calculates the total of numeric values |
| COUNT | 34 | Counts cells containing numbers |
| AVERAGE | 297.06 | Calculates the average of numeric values |
| MIN | 30 | Identifies the minimum numeric value |
| MAX | 1,000 | Identifies the maximum numeric value |
| SUMIF | 8,050 | Adds values that meet a specified condition |
| COUNTIF | 11 | Counts cells that meet a specified condition |

**Note:** The results above are based on the calculations performed in the Excel workbook. The exact meaning of each result depends on the selected cell range and formula criteria.

### 1. SUM Function

The `SUM` function adds numbers in a selected range.

**Syntax:**
```excel
=SUM(range)

Example:

=SUM(D2:D35)

Purpose: Calculates the total of product prices in the selected range.

2. COUNT Function

The COUNT function counts the number of cells containing numeric values.

Syntax:

=COUNT(range)

Example:

=COUNT(D2:D35)

Purpose: Counts the numeric entries in the selected product price range.

3. AVERAGE Function

The AVERAGE function calculates the arithmetic mean of numeric values.

Syntax:

=AVERAGE(range)

Example:

=AVERAGE(D2:D35)

Purpose: Calculates the average product price.

4. MIN Function

The MIN function returns the smallest numeric value in a range.

Syntax:

=MIN(range)

Example:

=MIN(D2:D35)

Purpose: Identifies the lowest product price in the selected range.

5. MAX Function

The MAX function returns the largest numeric value in a range.

Syntax:

=MAX(range)

Example:

=MAX(D2:D35)

Purpose: Identifies the highest product price in the selected range.

6. SUMIF Function

The SUMIF function adds values that satisfy a specified condition.

Syntax:

=SUMIF(range, criteria, sum_range)

Example:

=SUMIF(F2:F35,"Electronics",D2:D35)

Purpose: Calculates the total price of products belonging to the Electronics category.

7. COUNTIF Function

The COUNTIF function counts the number of cells that satisfy a specified condition.

Syntax:

=COUNTIF(range, criteria)

Example:

=COUNTIF(F2:F35,"Electronics")

Purpose: Counts the products belonging to the Electronics category.

🔤 Text Functions Used

Text functions help extract specific information from product IDs and other text values.

1. LEFT Function

The LEFT function extracts characters from the beginning of a text string.

Syntax:

=LEFT(text, num_chars)

Example:

=LEFT(A2,2)

Purpose: Extracts the first two characters from a product ID.

2. RIGHT Function

The RIGHT function extracts characters from the end of a text string.

Syntax:

=RIGHT(text, num_chars)

Example:

=RIGHT(A2,2)

Purpose: Extracts the last two characters, such as the country code US from 28-JAN-US.

3. MID Function

The MID function extracts characters from the middle of a text string.

Syntax:

=MID(text, start_num, num_chars)

Example:

=MID(A2,4,3)

Purpose: Extracts the month abbreviation JAN from the product ID 28-JAN-US.

📈 Key Analysis Areas

Product price analysis

Product quantity analysis

Brand and category exploration

Price range classification

Country code identification

Month and day extraction

Conditional calculations using SUMIF and COUNTIF

Summary statistics using basic Excel functions

🛠️ Tools and Technologies

Microsoft Excel – Data organization, calculations, and analysis

Excel Formulas – Numerical and conditional calculations

Text Functions – Extraction of information from product IDs

GitHub – Project documentation and version control

🚀 How to Use This Project

Download or open the Excel workbook.

Review the product dataset and its column headings.

Check the formulas used for calculations.

Apply functions such as SUM, COUNT, AVERAGE, MIN, and MAX.

Use SUMIF and COUNTIF to perform conditional analysis.

Apply LEFT, RIGHT, and MID to extract information from product IDs.

Compare the results and identify useful insights from the dataset.

💡 Learning Outcomes

Through this project, I practiced:

Understanding and organizing structured product data.

Applying basic Excel statistical functions.

Performing conditional calculations.

Extracting text information using Excel formulas.

Summarizing and interpreting product data.

Documenting a data analysis project using GitHub Markdown.

⭐ Conclusion

This project provides hands-on experience with Microsoft Excel for product data exploration and analysis. By combining numerical functions, conditional formulas, and text functions, the project demonstrates how raw product information can be organized and transformed into meaningful insights.
information can be organized and transformed into meaningful insights.
