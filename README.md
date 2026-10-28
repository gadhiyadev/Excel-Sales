# Excel Sales Data Analysis & Business Intelligence Dashboard

## 📌 Project Overview

This project focuses on analyzing sales transaction data using Microsoft Excel to identify sales trends, understand customer purchasing behavior, evaluate product performance, and generate meaningful business insights.

The workbook contains 249 customer transaction records covering different products, categories, regions, payment methods, and customer segments. I used Excel formulas, PivotTables, charts, statistical analysis, and what-if analysis to explore the data and organize the results into a structured analytical workbook.

## 🎯 Project Objectives

- Analyze sales transactions and calculate total revenue.
- Identify popular products and frequently occurring product names.
- Evaluate customer spending and identify high-value customers.
- Compare sales performance across product categories and regions.
- Apply Excel date and text functions to extract useful information.
- Perform scenario analysis to understand how changes in quantity and unit price affect revenue.
- Use linear regression to examine relationships between variables.
- Present analytical results using charts and summary tables.

## 🛠️ Tools and Technologies Used

- **Microsoft Excel** — Data analysis, calculations, reporting, and visualization.
- **Excel Formulas and Functions** — Revenue calculations, date calculations, text manipulation, counting, filtering, and matching.
- **PivotTables** — Summarizing sales by category, region, and customer.
- **Charts** — Comparing sales results and visualizing data.
- **Scenario Manager** — Comparing best-case, average-case, and worst-case scenarios.
- **Data Analysis ToolPak** — Performing linear regression analysis.


## 📊 Dataset Description

The dataset includes the following fields:

- `Transaction_ID` — Unique transaction identifier.
- `Date` — Date of the transaction.
- `Customer_ID` — Unique customer identifier.
- `Customer_Name` — Customer name.
- `Product_ID` — Product identifier.
- `Product_Name` — Name of the purchased product.
- `Category` — Product category.
- `Quantity` — Number of units purchased.
- `Unit_Price` — Price per unit.
- `Payment_Method` — Payment method used.
- `Region` — Geographical sales region.
- `Customer_Segment` — Customer classification.
- `Customer_Since` — Customer registration date.
- `Total_Amount` — Calculated transaction amount.

## 🔍 Data Analysis and Excel Techniques

### 1. Sales and Revenue Calculations

I calculated the total transaction amount by multiplying the quantity purchased by the unit price.

**Formula used:**

```excel
=H2*I2
```

This calculation was applied to the `Total_Amount` column to determine the monetary value of each transaction.

### 2. Date and Time Functions

I explored Excel date and time functions to work with transaction dates and customer information.

Functions used included:

- `TODAY()` — Returns the current date.
- `NOW()` — Returns the current date and time.
- `DATEDIF()` — Calculates the difference between two dates.
- `EOMONTH()` — Returns the last day of a month relative to a specified date.

These functions demonstrate how Excel can support date-based calculations and time-related analysis.

### 3. Product Frequency Analysis

I used `COUNTIF()` to count how frequently each product appeared in the transaction dataset.

```excel
=COUNTIF(F:F,F2)
```

I also used `INDEX()`, `MATCH()`, and `MAX()` together to identify the product associated with the highest frequency count.

```excel
=INDEX(F:F,MATCH(MAX(Q:Q),Q:Q,0))
```

This approach helps identify frequently purchased products based on transaction records.

### 4. Filtering and Matching Data

I used the `FILTER()` function to retrieve records that met a specified condition, such as transactions belonging to the Electronics category.

I also explored matching names between lists and extracting matching entries into another column.

These techniques are useful for isolating relevant records and comparing datasets.

### 5. Text Manipulation

I applied Excel text functions to manipulate customer names and extract useful information.

Functions included:

- `LEFT()` — Extracts characters from the beginning of a text string.
- `MID()` — Extracts characters from a specified position.
- `FIND()` — Locates a character or substring.
- `UPPER()` — Converts text to uppercase.
- `LOWER()` — Converts text to lowercase.

For example, the following formula extracts the first initials of a customer's first and last names:

```excel
=LEFT(D2,1)&MID(D2,FIND(" ",D2)+1,1)
```

### 6. PivotTable Analysis

I created PivotTables to summarize the transaction data and examine sales performance across different dimensions.

The analysis includes:

- Total sales by product category.
- Transaction counts by region.
- Total spending by customer.
- Customer-level sales comparisons.

PivotTables help convert individual transaction records into summaries that are easier to interpret.

### 7. High-Value Customer Analysis

I summarized transaction amounts by customer to identify customers contributing higher sales values.

This analysis helps demonstrate how customer-level aggregation can support customer segmentation, targeted marketing, and customer retention decisions.

### 8. What-If Analysis Using Scenario Manager

I used Excel's Scenario Manager to compare different combinations of quantity and unit price and observe their effect on the total transaction amount.

The workbook includes three scenarios:

- **Best-case scenario:** Quantity of 10 and unit price of 900, producing a calculated amount of 9,000.
- **Average-case scenario:** Quantity of 5 and unit price of 700, producing a calculated amount of 3,500.
- **Worst-case scenario:** Quantity of 2 and unit price of 59.99, producing a calculated amount of 119.98.

These scenarios illustrate how changes in input values affect the calculated transaction amount. They are hypothetical scenarios, not actual sales forecasts.

### 9. Linear Regression Analysis

I used Excel's Data Analysis ToolPak to generate a linear regression output containing:

- Regression statistics.
- R-squared and adjusted R-squared.
- ANOVA results.
- Regression coefficients.
- P-values and confidence intervals.
- Predicted values and residuals.

The saved regression output reports an R-squared value of approximately **0.210**, meaning the model explains around 21% of the variation in the dependent variable within the analyzed sample.

This indicates that the model has limited explanatory power and that other factors may influence the outcome. The result should be interpreted in the context of the variables selected for the regression.

## 📈 Key Findings

Based on the saved workbook summaries:

- **Total transaction value:** Approximately 224,692.52 in the dataset's currency units.
- **Electronics sales:** Approximately 167,256.10.
- **Furniture sales:** Approximately 49,097.68.
- **Appliances sales:** Approximately 8,338.74.
- **Highest-sales category:** Electronics, accounting for the largest share of the recorded transaction value.
- **Regional transaction analysis:** The workbook summarizes transactions across Central, East, North, South, and West regions.
- **Customer analysis:** Customer-level transaction values are summarized to help compare spending contributions.

The currency is not explicitly identified in the workbook, so the monetary figures are reported without assigning a currency symbol.

## 📊 Data Visualization

The workbook includes charts for visual exploration, including:

- Bar charts.
- Pie charts.
- A line chart.
- Pivot-based visualizations.

These charts provide a visual way to compare selected sales metrics and communicate patterns in the transaction data.

## 💡 Business Applications

The techniques demonstrated in this project can help businesses:

- Understand which product categories contribute most to revenue.
- Identify frequently purchased products.
- Recognize customers with higher spending levels.
- Compare transaction activity across regions.
- Evaluate how changes in quantity and price affect transaction values.
- Use statistical analysis to investigate relationships within sales data.
- Support data-informed sales and marketing decisions.

## 📚 Skills Demonstrated

- Microsoft Excel
- Sales Data Analysis
- Data Aggregation
- PivotTables and Charts
- Excel Formulas and Functions
- Date and Time Calculations
- Text Manipulation
- Data Filtering and Matching
- Customer Value Analysis
- What-If Analysis
- Linear Regression
- Data Visualization
- Business Insight Generation

## ⚠️ Project Scope and Limitations

This project is based on the transaction data provided in the workbook. The analysis describes the recorded data and should not be treated as a prediction of future business performance.

The regression results depend on the variables and observations selected for the analysis. The workbook also contains a dashboard worksheet with charts, while the written report worksheet remains unpopulated.

## 👤 Author

Created as a Microsoft Excel data analysis project to practice sales reporting, customer analysis, statistical techniques, and business intelligence concepts.
