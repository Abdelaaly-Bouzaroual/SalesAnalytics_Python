# ==================================================
# Sales Analysis Project - Python Data Science
# ==================================================

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Configure seaborn style for consistency
sns.set(style="whitegrid")

# Main function for running the analysis
def main():
    # ===============================================
    # Sales Analysis Project - Python Data Science
    # ===============================================

    # Set global plot style
    sns.set(style="whitegrid")

    # Load Excel file
    file_path = "sales.xlsx"
    xls = pd.ExcelFile(file_path)
    sales_df = pd.read_excel(xls, sheet_name='Sales Orders')
    cust_df = pd.read_excel(xls, sheet_name='Customers')
    region_df = pd.read_excel(xls, sheet_name='Regions')
    product_df = pd.read_excel(xls, sheet_name='Products')

    # Step 1: Clean and transform sales data
    sales_df['OrderDate'] = pd.to_datetime(sales_df['OrderDate'])
    sales_df['Ship Date'] = pd.to_datetime(sales_df['Ship Date'])
    sales_df['Year'] = sales_df['OrderDate'].dt.year
    sales_df['Month'] = sales_df['OrderDate'].dt.month
    sales_df['Sales'] = sales_df['Order Quantity'] * sales_df['Unit Selling Price']
    sales_df['Profit'] = sales_df['Order Quantity'] * (sales_df['Unit Selling Price'] - sales_df['Unit Cost'])
    sales_df['Cost'] = sales_df['Order Quantity'] * sales_df['Unit Cost']

    # Merge auxiliary info
    sales_df = sales_df.merge(product_df, left_on='Product Description Index', right_on='Index', how='left')
    sales_df = sales_df.merge(cust_df, left_on='Customer Name Index', right_on='Customer Index', how='left')
    sales_df = sales_df.merge(region_df, left_on='Delivery Region Index', right_on='Index', how='left')

    # Create date table
    date_range = pd.date_range(sales_df['OrderDate'].min(), sales_df['OrderDate'].max())
    date_df = pd.DataFrame({
        'Date': date_range,
        'Year': date_range.year,
        'Month': date_range.month,
        'Quarter': date_range.quarter,
        'DayOfWeek': date_range.dayofweek
    })

    # Step 2: KPI Measures
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

    # --- Visualization Section ---

    # Total Sales per Year
    plt.figure(figsize=(8, 5))
    sns.barplot(x='Year', y='Sales', data=kpi_df)
    plt.title('Total Sales per Year')
    plt.ylabel('Sales')
    plt.tight_layout()
    plt.show()

    # Profit per Year
    plt.figure(figsize=(8, 5))
    sns.lineplot(x='Year', y='Profit', data=kpi_df, marker='o', color='green')
    plt.title('Profit per Year')
    plt.ylabel('Profit')
    plt.tight_layout()
    plt.show()

    # Margin %
    plt.figure(figsize=(8, 5))
    sns.lineplot(x='Year', y='Profit Margin %', data=kpi_df, marker='s', color='orange')
    plt.title('Profit Margin % per Year')
    plt.ylabel('Margin %')
    plt.tight_layout()
    plt.show()

    # Sales by Product (current year vs previous year)
    plt.figure(figsize=(10, 6))
    top_products = sales_df.groupby(['Product Name', 'Year'])['Sales'].sum().reset_index()
    sns.barplot(x='Product Name', y='Sales', hue='Year', data=top_products)
    plt.title('Sales by Product and Year')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

    # Sales by Month (across years)
    plt.figure(figsize=(10, 5))
    month_sales = sales_df.groupby(['Year', 'Month'])['Sales'].sum().reset_index()
    sns.lineplot(x='Month', y='Sales', hue='Year', data=month_sales, marker='o')
    plt.title('Monthly Sales per Year')
    plt.xticks(range(1, 13))
    plt.tight_layout()
    plt.show()

    # Top 5 cities by sales
    top_cities = sales_df.groupby('City')['Sales'].sum().sort_values(ascending=False).head(5)
    plt.figure(figsize=(8, 5))
    sns.barplot(x=top_cities.index, y=top_cities.values)
    plt.title('Top 5 Cities by Sales')
    plt.ylabel('Sales')
    plt.tight_layout()
    plt.show()

    # Profit by Channel
    channel_profit = sales_df.groupby(['Channel', 'Year'])['Profit'].sum().reset_index()
    plt.figure(figsize=(10, 5))
    sns.barplot(x='Channel', y='Profit', hue='Year', data=channel_profit)
    plt.title('Profit by Channel and Year')
    plt.tight_layout()
    plt.show()

    # Top 5 clients by sales
    client_sales = sales_df.groupby(['Customer Names', 'Year'])['Sales'].sum().reset_index()
    top_clients = client_sales.groupby('Customer Names')['Sales'].sum().nlargest(5).index
    plt.figure(figsize=(10, 6))
    sns.barplot(x='Customer Names', y='Sales', hue='Year',
                data=client_sales[client_sales['Customer Names'].isin(top_clients)])
    plt.title('Top 5 Clients by Sales per Year')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

    # Last 5 clients by sales
    last_clients = client_sales.groupby('Customer Names')['Sales'].sum().nsmallest(5).index
    plt.figure(figsize=(10, 6))
    sns.barplot(x='Customer Names', y='Sales', hue='Year',
                data=client_sales[client_sales['Customer Names'].isin(last_clients)])
    plt.title('Bottom 5 Clients by Sales per Year')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

    # Summary Printout
    print("========= KPI SUMMARY =========")
    print(kpi_df[['Year', 'Sales', 'Sales PY', 'Sales Var', 'Sales Var %', 'Profit', 'Profit Margin %']])
# Entry point
if __name__ == "__main__":
    main()
