# BOM ABC Analysis Automation

생산계획, BOM, 자재마스터, 재고 데이터를 활용해 자재별 12주 계획 사용량과 사용금액을 계산하고, 사용금액 기준으로 ABC 등급을 자동 분류하는 Python pandas 기반 Excel 자동화 프로젝트입니다.

---

## 1. Project Overview

This project automates BOM-based ABC analysis using production plans, BOM master data, material master data, and inventory data.

The main purpose of this project is to identify high-value and critical materials that have a significant impact on cost, inventory burden, and procurement priorities.

The script calculates 12-week material usage quantities, converts them into usage amounts using unit prices, and classifies materials into A, B, and C classes based on cumulative usage amount ratio.

The final result is exported as a formatted Excel report for practical SCM, procurement, inventory, and cost management use.

---

## 2. Business Background

As the number of BOM components increases, it becomes difficult to manually determine which materials have the largest impact on cost and inventory.

In real procurement and SCM operations, not every material should be managed with the same priority. High-value materials should receive more attention because they often have greater influence on cost reduction, supplier negotiation, inventory control, and delivery risk management.

This project solves that problem by automatically calculating material usage and classifying key materials through ABC analysis.

---

## 3. Key Features

* Read multiple Excel input files using pandas
* Validate file existence before execution
* Validate required columns for each input file
* Merge production plan data with BOM data using `product_code`
* Calculate material required quantity based on production quantity and BOM usage quantity
* Aggregate 12-week material usage by material
* Merge material master data including supplier, lead time, MOQ, and unit price
* Merge inventory data including current stock
* Calculate 12-week usage amount by material
* Sort materials by usage amount in descending order
* Calculate cumulative amount and cumulative ratio
* Automatically classify materials into A, B, and C classes
* Generate multiple Excel report sheets
* Apply Excel formatting for better readability
* Highlight ABC classes with different colors
* Print final summary results after execution

---

## 4. Input Files

The code is designed to run in a Kaggle Notebook environment.

All input files should be located in the following directory:

```python
/kaggle/working/
```

### 4.1 production_plan.xlsx

This file contains weekly production plan data.

| Column       | Description                 |
| ------------ | --------------------------- |
| plan_week    | Production plan week        |
| customer     | Customer name               |
| model        | Product model               |
| product_code | Product code                |
| qty          | Planned production quantity |

### 4.2 bom_master.xlsx

This file contains BOM information by product.

| Column        | Description                                 |
| ------------- | ------------------------------------------- |
| product_code  | Product code                                |
| material_code | Material code                               |
| material_name | Material name                               |
| usage_qty     | Required material quantity per product unit |

### 4.3 material_master.xlsx

This file contains material master information.

| Column         | Description            |
| -------------- | ---------------------- |
| material_code  | Material code          |
| supplier       | Supplier name          |
| lead_time_days | Lead time in days      |
| moq            | Minimum order quantity |
| unit_price     | Unit price of material |

### 4.4 inventory.xlsx

This file contains current stock information.

| Column        | Description            |
| ------------- | ---------------------- |
| material_code | Material code          |
| current_stock | Current stock quantity |

---

## 5. Output File

The final Excel report is saved to the following path:

```python
/kaggle/working/bom_abc_analysis.xlsx
```

The output file includes the following sheets:

| Sheet Name          | Purpose                              |
| ------------------- | ------------------------------------ |
| abc_analysis        | Main ABC analysis result by material |
| summary_by_abc      | Summary by ABC class                 |
| supplier_summary    | Usage amount summary by supplier     |
| weekly_usage_detail | Weekly material usage detail         |

---

## 6. Calculation Logic

### 6.1 Material Required Quantity

The production plan and BOM master are merged by `product_code`.

The required material quantity is calculated as follows:

```python
material_required_qty = qty * usage_qty
```

For example, if the planned production quantity is 300 units and the BOM usage quantity is 2 units per product, the required material quantity is:

```python
300 * 2 = 600
```

---

### 6.2 Total 12-Week Usage Quantity

The script aggregates material required quantity by material over the selected 12-week period.

```python
total_12week_usage_qty = sum(material_required_qty)
```

This shows how much of each material is required for the entire 12-week production plan.

---

### 6.3 Usage Amount

The 12-week usage amount is calculated by multiplying total usage quantity by unit price.

```python
usage_amount_12week = total_12week_usage_qty * unit_price
```

This value is used as the main basis for ABC classification.

---

### 6.4 Cumulative Amount and Cumulative Ratio

Materials are sorted by usage amount in descending order.

The cumulative amount is calculated as follows:

```python
cumulative_amount = cumulative sum of usage_amount_12week
```

The cumulative ratio is calculated as follows:

```python
cumulative_ratio = cumulative_amount / total_usage_amount
```

---

## 7. ABC Classification Criteria

ABC classification is based on cumulative usage amount ratio.

| Class | Criteria                      | Meaning                     |
| ----- | ----------------------------- | --------------------------- |
| A     | Cumulative ratio <= 80%       | High-priority key materials |
| B     | 80% < Cumulative ratio <= 95% | Medium-priority materials   |
| C     | Cumulative ratio > 95%        | Low-priority materials      |

A-class materials are considered the most important materials for cost reduction, inventory control, and supplier negotiation.

---

## 8. Excel Report Sheets

### 8.1 abc_analysis

This is the main analysis sheet.

It includes material-level ABC analysis results.

Main columns:

* material_code
* material_name
* supplier
* lead_time_days
* moq
* unit_price
* current_stock
* total_12week_usage_qty
* usage_amount_12week
* amount_ratio
* cumulative_amount
* cumulative_ratio
* abc_class

---

### 8.2 summary_by_abc

This sheet summarizes the analysis result by ABC class.

Main columns:

* abc_class
* material_count
* total_usage_qty
* total_usage_amount
* amount_ratio

This sheet helps users quickly understand how many materials belong to each class and how much each class contributes to total usage amount.

---

### 8.3 supplier_summary

This sheet summarizes usage amount by supplier.

Main columns:

* supplier
* material_count
* a_class_count
* total_usage_amount
* a_class_usage_amount

This sheet helps identify suppliers that are highly related to high-value materials.

---

### 8.4 weekly_usage_detail

This sheet provides weekly material usage details.

Main columns:

* plan_week
* material_code
* material_name
* required_qty
* unit_price
* usage_amount
* supplier
* abc_class

This sheet is useful for checking weekly demand trends by material.

---

## 9. Validation Logic

The script includes validation logic to reduce data errors.

### 9.1 File Existence Check

If an input file does not exist, the script prints which file is missing.

### 9.2 Required Column Check

If a required column is missing, the script prints the missing column name.

### 9.3 BOM Matching Check

If a `product_code` exists in the production plan but not in the BOM master, the script prints a warning message.

### 9.4 Unit Price Check

If a material has missing or zero unit price, the script prints a warning message.

### 9.5 Calculation Check

The script checks whether the following calculations are correct:

```python
material_required_qty = qty * usage_qty
usage_amount_12week = total_12week_usage_qty * unit_price
cumulative_ratio = cumulative_amount / total_usage_amount
```

### 9.6 ABC Classification Check

The script validates whether A, B, and C classes are assigned based on the correct cumulative ratio criteria.

---

## 10. Excel Formatting

The output Excel file includes the following formatting:

* Bold header row
* Filter applied to the header row
* Freeze top row
* Auto-adjusted column width
* Thousand separator for amount columns
* Percentage format for ratio columns
* Highlight A-class rows with orange background
* Highlight B-class rows with yellow background
* Highlight C-class rows with light gray background

These formatting rules make the report easier to read and use in practical business situations.

---

## 11. Final Report Output

After execution, the script prints the following summary:

* Output file path
* Total number of materials
* Number of A-class materials
* Number of B-class materials
* Number of C-class materials
* Total 12-week usage amount
* A-class usage amount
* A-class usage amount ratio
* Top material by usage amount
* Top supplier by usage amount

---

## 12. Tech Stack

* Python
* pandas
* NumPy
* openpyxl
* Excel

---

## 13. How to Run

### Step 1. Upload Input Files

Upload the following files to `/kaggle/working/`.

```python
production_plan.xlsx
bom_master.xlsx
material_master.xlsx
inventory.xlsx
```

### Step 2. Run the Python Code

Run the Python script in a Kaggle Notebook.

### Step 3. Download the Output File

The result file will be created at:

```python
/kaggle/working/bom_abc_analysis.xlsx
```

---

## 14. Expected Business Impact

This project can support procurement and SCM decision-making by helping users identify key materials that have the greatest cost impact.

Expected business use cases include:

* Selecting cost reduction targets
* Prioritizing supplier negotiation
* Identifying high-value materials
* Reviewing inventory risk
* Supporting procurement planning
* Improving material management efficiency

---

## 15. Project Summary

This project automates BOM-based ABC analysis using Python pandas and Excel.

It connects production plan, BOM, material master, and inventory data to calculate 12-week material usage and usage amount. Based on cumulative usage amount ratio, materials are automatically classified into A, B, and C classes.

The final Excel report helps procurement and SCM users quickly identify high-value materials and set priorities for cost reduction, inventory control, and supplier management.
