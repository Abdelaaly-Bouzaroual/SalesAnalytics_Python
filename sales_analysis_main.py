# ==================================================
# Sales Analysis Project - Python Data Science
# Author: ChatGPT - Based on user project description
# ==================================================

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime

# Configure seaborn style for consistency
sns.set(style="whitegrid")

# Main function for running the analysis
def main():
    # === Step 1: Load and clean the data ===
    file_path = "sales.xlsx"  # <-- Ensure this file is in the same folder
    xls = pd.ExcelFile(file_path)
    sales_df = pd.read_excel(xls, sheet_name='Sales Orders')
    cust_df = pd.read_excel(xls, sheet_name='Customers')
    region_df = pd.read_excel(xls, sheet_name='Regions')
    product_df = pd.read_excel(xls, sheet_name='Products')

    # Convert dates
    sales_df['OrderDate'] = pd.to_datetime(sales_df['OrderDate'])
    sales_df['Ship Date'] = pd.to_datetime(sales_df['Ship Date'])
    sales_df['Year'] = sales_df['OrderDate'].dt.year
    sales_df['Month'] = sales_df['OrderDate'].dt.month

    # Calculate Sales, Profit and Cost
    sales_df['Sales'] = sales_df['Order Quantity'] * sales_df['Unit Selling Price']
    sales_df['Profit'] = sales_df['Order Quantity'] * (sales_df['Unit Selling Price'] - sales_df['Unit Cost'])
    sales_df['Cost'] = sales_df['Order Quantity'] * sales_df['Unit Cost']

    # Merge with additional information
    sales_df = sales_df.merge(product_df, left_on='Product Description Index', right_on='Index', how='left')
    sales_df = sales_df.merge(cust_df, left_on='Customer Name Index', right_on='Customer Index', how='left')
    sales_df = sales_df.merge(region_df, left_on='Delivery Region Index', right_on='Index', how='left')

    # Create date dimension table
    date_range = pd.date_range(sales_df['OrderDate'].min(), sales_df['OrderDate'].max())
    date_df = pd.DataFrame({
        'Date': date_range,
        'Year': date_range.year,
        'Month': date_range.month,
        'Quarter': date_range.quarter,
        'DayOfWeek': date_range.dayofweek
    })

    # === Step 2: KPI Aggregation ===
    kpi_df = sales_df.groupby(['Year']).agg({
        'Sales': 'sum',
        'Profit': 'sum',
        'Cost': 'sum',
        'Order Quantity': 'sum'
    }).reset_index()
    kpi_df['Profit Margin %'] = 100 * kpi_df['Profit'] / kpi_df['Sales']
    kpi_df['Sales PY'] = kpi_df['Sales'].shift(1)
    kpi_df['Sales Var'] = kpi_df['Sales'] - kpi_df['Sales PY']
    kpi_df['Sales Var %'] = 100 * kpi_df['Sales Var'] / kpi_df['Sales']

    # === Step 3: Visualizations ===
    plt.figure(figsize=(8, 5))
    sns.barplot(x='Year', y='Sales', data=kpi_df)
    plt.title('Total Sales per Year')
    plt.tight_layout()
    plt.show()

    plt.figure(figsize=(8, 5))
    sns.lineplot(x='Year', y='Profit', data=kpi_df, marker='o', color='green')
    plt.title('Profit per Year')
    plt.tight_layout()
    plt.show()

    plt.figure(figsize=(8, 5))
    sns.lineplot(x='Year', y='Profit Margin %', data=kpi_df, marker='s', color='orange')
    plt.title('Profit Margin % per Year')
    plt.tight_layout()
    plt.show()

    # More visualizations can be added as needed, following similar structure...

    # === Step 4: Print KPI Summary ===
    print("========= KPI SUMMARY =========")
    print(kpi_df[['Year', 'Sales', 'Sales PY', 'Sales Var', 'Sales Var %', 'Profit', 'Profit Margin %']])

# Entry point
if __name__ == "__main__":
    main()
