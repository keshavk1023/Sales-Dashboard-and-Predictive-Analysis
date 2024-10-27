import pandas as pd
import pyodbc
import matplotlib.pyplot as plt
import seaborn as sns

# Connect to SQL Server
conn = pyodbc.connect('DRIVER={SQL Server};SERVER=your_server_name;DATABASE=your_database_name;UID=your_username;PWD=your_password')

# Data Extraction using SQL Queries
query = """
SELECT 
    OrderDate,
    ProductID,
    ProductName,
    SalesAmount,
    Quantity,
    CustomerID,
    Region
FROM 
    SalesData
"""
sales_data = pd.read_sql(query, conn)

# Data Transformation
# Convert OrderDate to datetime format
sales_data['OrderDate'] = pd.to_datetime(sales_data['OrderDate'])

# Extract year and month from OrderDate
sales_data['Year'] = sales_data['OrderDate'].dt.year
sales_data['Month'] = sales_data['OrderDate'].dt.month

# Calculate total sales and quantity per product
product_sales = sales_data.groupby(['ProductID', 'ProductName']).agg({
    'SalesAmount': 'sum',
    'Quantity': 'sum'
}).reset_index()

# Calculate total sales and quantity per region
region_sales = sales_data.groupby('Region').agg({
    'SalesAmount': 'sum',
    'Quantity': 'sum'
}).reset_index()

# Calculate monthly sales trends
monthly_sales = sales_data.groupby(['Year', 'Month']).agg({
    'SalesAmount': 'sum',
    'Quantity': 'sum'
}).reset_index()

# Visualize Sales Trends
plt.figure(figsize=(12, 6))
sns.lineplot(data=monthly_sales, x='Month', y='SalesAmount', hue='Year')
plt.title('Monthly Sales Trends')
plt.xlabel('Month')
plt.ylabel('Sales Amount')
plt.legend(title='Year')
plt.show()

# Visualize Top Performing Products
top_products = product_sales.sort_values(by='SalesAmount', ascending=False).head(10)
plt.figure(figsize=(12, 6))
sns.barplot(data=top_products, x='SalesAmount', y='ProductName')
plt.title('Top Performing Products')
plt.xlabel('Sales Amount')
plt.ylabel('Product Name')
plt.show()

# Visualize Sales by Region
plt.figure(figsize=(12, 6))
sns.barplot(data=region_sales, x='SalesAmount', y='Region')
plt.title('Sales by Region')
plt.xlabel('Sales Amount')
plt.ylabel('Region')
plt.show()

# Save the transformed data to CSV for Power BI
sales_data.to_csv('transformed_sales_data.csv', index=False)

print("Sales Dashboard and Predictive Analysis Completed. Data saved to 'transformed_sales_data.csv' for Power BI visualization.")
