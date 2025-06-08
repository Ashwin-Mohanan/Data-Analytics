import pandas as pd
import xlsxwriter
import numpy as np
from datetime import datetime

def convert(dt):
    try:
        return datetime.strptime(dt, '%d/%m/%Y').strftime("%Y-%m-%d")
    except ValueError:
        return datetime.strptime(dt, '%m/%d/%Y').strftime("%Y-%m-%d")


fileName = 'D:/ashwin/devScreen/Excesise/New folder (2)/SampleSales.csv'
rawdf = pd.read_csv(fileName)
rawdf = rawdf.dropna()
rawdf = rawdf.drop_duplicates()
# checking raw data type
print(rawdf.info())

# converting date in yyy-mm-dd
rawdf['Order Date'] = rawdf['Order Date'].apply(convert)
rawdf['Ship Date'] = rawdf['Ship Date'].apply(convert)

# converting in datetimeformat in py
rawdf['Order Date'] = pd.to_datetime(rawdf['Order Date'])
rawdf['Ship Date'] = pd.to_datetime(rawdf['Ship Date'])

rawdf['Year Month'] = rawdf['Order Date'].dt.to_period('M')
rawdf['Year'] = rawdf['Order Date'].dt.to_period('Y')
rawdf['Month'] = rawdf['Order Date'].dt.month
df_grouped = rawdf.groupby(['Year','Month']).agg({'Sales':'sum'}).reset_index()

# df_grouped['Growth_Rate'] = (df_grouped['Sales'] - df_grouped['Sales'].shift(1)) / df_grouped['Sales'].shift(1)
meanvals = df_grouped["Sales"].mean()
stdvals = df_grouped["Sales"].std()

 # Get last available year and month
last_year = df_grouped['Year'].max()
last_month = df_grouped[df_grouped['Year'] == last_year]['Month'].max()
df_grouped['Lower_Bound'] = df_grouped["Sales"]
df_grouped['Upper_Bound'] = df_grouped["Sales"]

# Generate future periods
periods = 12 #12months
future_dates = []
for i in range(1, periods + 1):
    future_month = (last_month + i - 1) % 12 + 1
    future_year = last_year + (last_month + i - 1) // 12
    future_dates.append((future_year, future_month))

# Forecast sales using mean and std
future_sales = []
for year, month in future_dates:
    forecast = np.random.normal(loc=meanvals, scale= stdvals)  # Generate forecasted value
    lower_bound = forecast - (1 * stdvals)
    upper_bound = forecast + (1 * stdvals)
    future_sales.append([year, month, forecast, lower_bound, upper_bound])

forecast_df = pd.DataFrame(future_sales, columns=['Year', 'Month', 'Forecast_Sales', 'Lower_Bound', 'Upper_Bound'])
final_df = pd.concat([df_grouped, forecast_df.rename(columns={'Forecast_Sales': 'Sales'})], ignore_index=True)
final_df['Year Month'] = final_df['Year'].astype(str)+'-'+final_df['Month'].astype(str)
print(final_df) 
# Fill NaN values in the first row with the average growth rate
# avg_growth = df_grouped['Growth_Rate'].mean()
# df_grouped['Growth_Rate'].fillna(avg_growth, inplace=True)
