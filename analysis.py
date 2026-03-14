import pandas as pd

print("Loading dataset...")

df = pd.read_excel("SuperStoreOrders.csv.xlsx")

print("Dataset loaded successfully!")

print(df.head())

# Total Sales
total_sales = df["sales"].sum()
print("Total Sales:", total_sales)

# Total Profit
total_profit = df["profit"].sum()
print("Total Profit:", total_profit)

# Total Orders
total_orders = df["order_id"].nunique()
print("Total Orders:", total_orders)

# Sales by Region
sales_by_region = df.groupby("region")["sales"].sum()
print("\nSales by Region:")
print(sales_by_region)

# Sales by Category
sales_by_category = df.groupby("category")["sales"].sum()

print("\nSales by Category:")
print(sales_by_category)

# Top 10 Products by Sales
top_products = df.groupby("product_name")["sales"].sum().sort_values(ascending=False).head(10)

print("\nTop 10 Products by Sales:")
print(top_products)

# Monthly Sales Trend
monthly_sales = df.groupby("Months")["sales"].sum()

print("\nMonthly Sales Trend:")
print(monthly_sales)

# Profit by Category
profit_by_category = df.groupby("category")["profit"].sum()

print("\nProfit by Category:")
print(profit_by_category)

# Save cleaned dataset
df.to_csv("superstore_cleaned.csv", index=False)

print("\nCleaned dataset saved successfully!")

import pandas as pd
import matplotlib.pyplot as plt

print("Loading dataset...")

# load excel file
df = pd.read_excel("SuperStoreOrders.csv.xlsx")

print("Dataset loaded successfully!")

# -------------------------
# ANALYSIS
# -------------------------

# Sales by Region
sales_by_region = df.groupby("region")["sales"].sum()

# Profit by Category
profit_by_category = df.groupby("category")["profit"].sum()

# Top 10 Products
top_products = df.groupby("product_name")["sales"].sum().sort_values(ascending=False).head(10)

# Monthly Sales Trend
monthly_sales = df.groupby("Months")["sales"].sum()

# -------------------------
# DASHBOARD STYLE CHARTS
# -------------------------

fig, axes = plt.subplots(2,2, figsize=(14,8))

sales_by_region.plot(kind="bar", ax=axes[0,0])
axes[0,0].set_title("Sales by Region")

profit_by_category.plot(kind="bar", ax=axes[0,1])
axes[0,1].set_title("Profit by Category")

top_products.plot(kind="bar", ax=axes[1,0])
axes[1,0].set_title("Top 10 Products")

monthly_sales.plot(kind="line", marker="o", ax=axes[1,1])
axes[1,1].set_title("Monthly Sales Trend")

plt.tight_layout()
plt.show()

print("Analysis Completed!")