import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

df = pd.read_excel("Global_Ecommerce_Sales_Data.xlsx")
#file_path
file_path="Global_Ecommerce_Sales_Data.xlsx"
df = pd.read_excel(file_path)

# Show basic info
print("Data Loaded Successfully.")
print(df.columns)
print(df.head())



#Step 2: Clean the data 
# Strip extra spaces from column names
df.columns = df.columns.str.strip()

# Drop rows with missing important values
df.dropna(subset=['Order ID', 'Product', 'Quantity Ordered', 'Price Each', 'Revenue'], inplace=True)

# Convert types
df['Quantity Ordered'] = pd.to_numeric(df['Quantity Ordered'], errors='coerce')
df['Price Each'] = pd.to_numeric(df['Price Each'], errors='coerce')
df['Revenue'] = pd.to_numeric(df['Revenue'], errors='coerce')
df['Order Date'] = pd.to_datetime(df['Order Date'], errors='coerce')

# Drop rows with invalid data after conversion
df.dropna(inplace=True)

# Final check
print("Data after cleaning:")
print(df.info())
  


#Step 3: Exploratory Data Analysis (EDA) 
# 1. Total Revenue
total_revenue = df['Revenue'].sum()
print("Total Revenue:", total_revenue)

# 2. Total Unique Customers
unique_customers = df['Order ID'].nunique()
print("Unique Orders:", unique_customers)

# 3. Revenue by Country
revenue_by_country = df.groupby('Country')['Revenue'].sum().sort_values(ascending=False)
print("Revenue by Country:")
print(revenue_by_country)

# 4. Top 5 Products by Revenue
top_products = df.groupby('Product')['Revenue'].sum().sort_values(ascending=False).head(5)
print("Top 5 Products by Revenue:")
print(top_products)

# 5. Monthly Sales
df['Month'] = df['Order Date'].dt.month
monthly_sales = df.groupby('Month')['Revenue'].sum()
print("Monthly Sales:")
print(monthly_sales)




#Step 4: Visualize the data 

# Calculate order counts by country if not already defined
order_counts_by_country = df['Country'].value_counts()
# Set plot style for better visuals

sns.set_style=("whitegrid")



# Create 2x2 subplot layout
fig, axs = plt.subplots(2, 2, figsize=(14, 10), constrained_layout= True)

# 1. Revenue by Country
axs[0, 0].bar(revenue_by_country.index, revenue_by_country.values, color='skyblue')
axs[0, 0].set_title("Revenue by Country")
axs[0, 0].set_xlabel("Country")
axs[0, 0].set_ylabel("Total Revenue")
axs[0, 0].tick_params(axis='x', rotation=45)

# 2. Top 5 Products by Revenue
axs[0, 1].bar(top_products.index, top_products.values, color='orange')
axs[0, 1].set_title("Top 5 Products by Revenue")
axs[0, 1].set_ylabel("Revenue")
axs[0, 1].tick_params(axis='x', rotation=45)

# 3. Monthly Revenue
axs[1, 0].plot(monthly_sales.index, monthly_sales.values, marker='o', linestyle='--', color='green')
axs[1, 0].set_title("Monthly Revenue")
axs[1, 0].set_xlabel("Month")
axs[1, 0].set_ylabel("Revenue")
axs[1, 0].grid(True)

# 4. Optional: Orders by Country (example)
axs[1, 1].bar(order_counts_by_country.index, order_counts_by_country.values, color='purple')
axs[1, 1].set_title("Number of Orders by Country")
axs[1, 1].set_ylabel("Order Count")
axs[1, 1].tick_params(axis='x', rotation=45)

# Show all together
plt.show()


#Step 5: Export Results
# Export summary results to Excel
summary = {
    'Total Revenue': [total_revenue],
    'Unique Orders': [unique_customers],
    'Top Product': [top_products.idxmax()],
    'Top Product Revenue': [top_products.max()]
}

summary_df = pd.DataFrame(summary)
summary_df.to_excel("ecommerce_summary.xlsx", index=False)
print("Summary exported to 'ecommerce_summary.xlsx'")