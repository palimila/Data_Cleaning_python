import pandas as pd
import numpy as np
from fuzzywuzzy import process
import logging

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Load the Excel file
excel_data = pd.ExcelFile('Sales_Data_.xlsx')
sales_df = excel_data.parse('Sales')
product_master_df = excel_data.parse('Product Master')
region_master_df = excel_data.parse('Region Master')


sales_df.columns = sales_df.iloc[2]
sales_df = sales_df[3:]
sales_df = sales_df.iloc[:, :10]  
sales_df.columns = ['City', 'Region Code', 'No.of Cars', 'Price per car', 'Total Amount',
                    'Order Date', 'Month', 'Year', 'Product', 'Sales Person']
sales_df.dropna(how='all', inplace=True)

product_master_df.columns = product_master_df.iloc[0]
product_master_df = product_master_df[1:]
product_master_df = product_master_df.iloc[:, 3:]
product_master_df.columns = ['Car Model', 'Car Make', 'Category', 'Manufacturing cost']
product_master_df.dropna(how='all', inplace=True)


# Remove rows with null in important columns
sales_df.dropna(subset=['City', 'Region Code', 'No.of Cars', 'Price per car',
                        'Total Amount', 'Order Date', 'Product'], inplace=True)

# Remove rows with amount = 0 and empty key fields
sales_df = sales_df[~(
    (sales_df['Total Amount'].astype(float) == 0) &
    (sales_df[['City', 'Product', 'Order Date']].isnull().all(axis=1))
)]

# Clean City names
sales_df['City'] = sales_df['City'].str.strip().str.title()

# Split Product into Car Make and Car Model
product_split = sales_df['Product'].str.split('|', expand=True)
sales_df['Car Model'] = product_split[0]
sales_df['Car Make'] = product_split[1]

# Convert Order Date
sales_df['Order Date'] = pd.to_datetime(sales_df['Order Date'], dayfirst=True, errors='coerce')

# Flag inconsistent data
sales_df['Inconsistent_Flag'] = sales_df.isnull().any(axis=1)


# Merge Region Name and Country
merged_df = pd.merge(sales_df, region_master_df, how='left', on='Region Code')

# Fuzzy match missing region codes
missing_codes = merged_df[merged_df['Region Name'].isna()]['Region Code'].unique()
fuzzy_map = {}
for code in missing_codes:
    match, score = process.extractOne(code, region_master_df['Region Code'].astype(str))
    if score > 80:
        fuzzy_map[code] = match

# Replace using fuzzy map and remerge
sales_df['Region Code'] = sales_df['Region Code'].replace(fuzzy_map)
merged_df = pd.merge(sales_df, region_master_df, how='left', on='Region Code')

# Merge Category from product master
merged_df = pd.merge(merged_df, product_master_df, how='left', on=['Car Model', 'Car Make'])


# Total Sales by Country
merged_df['Total Amount'] = merged_df['Total Amount'].astype(float)
country_sales = merged_df.groupby('Country')['Total Amount'].sum().reset_index()

# Invalid Dates
invalid_dates = merged_df[merged_df['Order Date'].isna()]

# Remove duplicates keeping latest
merged_df = merged_df.sort_values('Order Date').drop_duplicates(subset=['City', 'Product'], keep='last')

# Product contribution %
total_sales = merged_df['Total Amount'].sum()
merged_df['Product Contribution %'] = (merged_df['Total Amount'] / total_sales) * 100

# Top-performing Car Make in each Region
top_car_make = merged_df.groupby(['Region Name', 'Car Make'])['Total Amount'].sum().reset_index()
top_car_make = top_car_make.sort_values(['Region Name', 'Total Amount'], ascending=[True, False])
top_performers = top_car_make.drop_duplicates('Region Name')

# Quarterly trends
merged_df['Quarter'] = pd.PeriodIndex(merged_df['Order Date'], freq='Q').astype(str)
quarterly_sales = merged_df.groupby(['Car Make', 'Quarter'])['Total Amount'].sum().reset_index()


logging.info("Data cleaning and transformation complete.")
logging.info(f"Total countries: {country_sales.shape[0]}")
logging.info(f"Invalid dates found: {invalid_dates.shape[0]}")
logging.info("Top performers and quarterly trends computed.")

# Optional: Save to Excel
with pd.ExcelWriter('Cleaned_Sales_Report.xlsx') as writer:
    merged_df.to_excel(writer, sheet_name='Cleaned Sales', index=False)
    invalid_dates.to_excel(writer, sheet_name='Invalid Dates', index=False)
    country_sales.to_excel(writer, sheet_name='Country Sales', index=False)
    top_performers.to_excel(writer, sheet_name='Top Car Make by Region', index=False)
    quarterly_sales.to_excel(writer, sheet_name='Quarterly Trends', index=False)
