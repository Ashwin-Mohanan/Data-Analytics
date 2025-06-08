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

rawdf['Sales'] = pd.to_numeric(rawdf['Sales'])
rawdf['Profit'] = pd.to_numeric(rawdf['Profit'])
# rawdf['Sales'] = rawdf['Sales'].fillna(0)
# rawdf['Profit'] = rawdf['Profit'].fillna(0)
# Groupby data on category
rawdf['Year Month'] = rawdf['Order Date'].dt.to_period('M')
rawdf['Year'] = rawdf['Order Date'].dt.to_period('Y')
df = rawdf.groupby(['Country/Region','Category','Region','Year Month','City']).agg({'Sales':'sum','Profit':'sum'}).reset_index().sort_values('Sales')
regionwise = rawdf.groupby(['Region']).agg({'Sales':'sum','Profit':'sum'}).reset_index().sort_values('Sales')
categorywise = rawdf.groupby(['Category']).agg({'Sales':'sum','Profit':'sum'}).reset_index().sort_values('Sales')
# writer = pd.ExcelWriter('exploratory.xlsx')
# regionwise.to_excel(writer, sheet_name='regionwise', index=False)
# categorywise.to_excel(writer, sheet_name='categorywise', index=False)

# workbook1 = writer.book
# worksheet1 = writer.sheets["regionwise"]

# chart1 = workbook1.add_chart({"type": "line"})
# chart1.add_series({"values": "=regionwise!$C$2:$C$4"})
# worksheet1.insert_chart("E2", chart1)

# workbook2 = writer.book
# worksheet2 = writer.sheets["categorywise"]
# chart2 = workbook2.add_chart({"type": "bar"})
# chart2.add_series({"values": "=categorywise!$C$2:$C$4"})
# worksheet2.insert_chart("E2", chart2)

# writer.close()


# Outliner using IQR rule.

# Q1 = df['Sales'].quantile(0.25) # finding values which are below 25%
# Q3 = df['Sales'].quantile(0.75) # finfing values which are below 75%
# IQR = Q3-Q1
# outliners = df[(df['Sales']<(Q1 - 1.5 *IQR)) | (df['Sales']>(Q3 +1.5*IQR))]

# print(outliners)
# print(rawdf.shape)


# Outliner using Controll limits
test=[]
# tempDf = rawdf.groupby(['Category','Region','Year Month']).agg({'Sales':'mean','Sales':'std'}).reset_index()
tempDf = rawdf.groupby(['Category', 'Region', 'Year Month']).agg({'Sales':'sum'}).reset_index()
# tempDf['Year Month'] = tempDf['Year Month'].dt.strftime('%y-%b-%d')
# for i in ['mean','std']:
#     grp = grouped.rename(columns={"Sales":'Sales_'+i})
#     xx = grp.groupby(['Category', 'Region', 'Year Month']).agg({'Sales_'+i:i}).reset_index()
#     test.append(xx)
# Calculate mean and standard deviation
# tempDf = grouped.agg(['mean', 'std']).reset_index()
# tempDf = tempDf.groupby(['Region', 'Category','Year Month'])['Sales'].agg(['mean', 'std']).reset_index()
tempDf['mean'] = tempDf["Sales"].mean()
tempDf['std'] = tempDf["Sales"].std()


tempDf.rename(columns={'mean': 'Avg_Sales', 'std': 'Std_Sales'}, inplace=True)
# salesMean = tempDf['mean']
# salesStd = tempDf['std']

tempDf['OutlierLower1'] = tempDf['Avg_Sales'] - (1 * tempDf['Std_Sales'])
tempDf['OutlierUpper1'] = tempDf['Avg_Sales'] + (1 * tempDf['Std_Sales'])

tempDf['OutlierLower2'] = tempDf['Avg_Sales'] - (2 * tempDf['Std_Sales'])
tempDf['OutlierUpper2'] = tempDf['Avg_Sales'] + (2 * tempDf['Std_Sales'])

tempDf['OutlierLower3'] = tempDf['Avg_Sales'] - (3 * tempDf['Std_Sales'])
tempDf['OutlierUpper3'] = tempDf['Avg_Sales'] + (3 * tempDf['Std_Sales'])
# finaloutlierdf = rawdf.merge(tempDf, on=['Category', 'Region', 'Year Month'], how='left')
# finaloutlierdf['outliers1'] = (finaloutlierdf['Sales'] < finaloutlierdf['OutlierLower1']) | (finaloutlierdf['Sales'] > finaloutlierdf['OutlierUpper1'])

# finaloutlierdf['outliers2'] = (finaloutlierdf['Sales'] < finaloutlierdf['OutlierLower2']) | (finaloutlierdf['Sales'] > finaloutlierdf['OutlierUpper2'])

# finaloutlierdf['outliers3'] = (finaloutlierdf['Sales'] < finaloutlierdf['OutlierLower3']) | (finaloutlierdf['Sales'] > finaloutlierdf['OutlierUpper3'])

# finaloutlierdf = finaloutlierdf[['Category', 'Region', 'Sales', 'Year Month','Year', 
#     'OutlierLower1', 'OutlierUpper1', 
#     'OutlierLower2', 'OutlierUpper2', 
#     'OutlierLower3', 'OutlierUpper3']]
# tempDf = tempDf.sort_values(by='Year Month')
finaloutlierdf = tempDf.fillna(0)
finaloutlier = finaloutlierdf.to_excel('finaloutlierdf2.xlsx')
print('Done')