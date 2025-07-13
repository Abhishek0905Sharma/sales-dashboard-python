import pandas as pd
import matplotlib.pyplot as plt
import os

# Load data
data_path = "../data/sales_data.xlsx"
df = pd.read_excel(data_path)

# Basic summary
summary = df.groupby("Region")["Revenue"].sum().sort_values(ascending=False)
print("Revenue by Region:")
print(summary)

# Plot sales trend
df['Date'] = pd.to_datetime(df['Date'])
monthly_sales = df.groupby(df['Date'].dt.to_period('M'))["Revenue"].sum()

plt.figure(figsize=(10, 5))
monthly_sales.plot(kind='line', marker='o')
plt.title("Monthly Sales Trend")
plt.xlabel("Month")
plt.ylabel("Revenue")
plt.grid(True)
plt.tight_layout()

# Save figure
os.makedirs("../images", exist_ok=True)
plt.savefig("../images/monthly_sales_trend.png")

# Export summary
os.makedirs("../output", exist_ok=True)
summary.to_csv("../output/revenue_by_region.csv")

print("Analysis completed. Check 'images/' and 'output/' folders.")
