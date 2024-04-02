# SalesDashboard
Microsoft Excel Sales Dashboard

In this project, I created a Microsoft Excel sales dashboard featuring pivot charts and slicers. I also looked at Excel's three-dimensional animated map. Here is a look at the end result:

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/Dashboard1.png>

## Table of Contents ##

1. Data Cleaning
2. Pivot Tables
3. Pivot Charts
4. Dashboard Creation
5. 3D Mapping

First, I used data cleaning techniques to prepare the data with power query. Next, I created pivot tables to pull relevant data from the worksheet and subsequently used them to create pivot charts. Then, the pivot charts were applied onto the dashboard. Lastly, I wanted to see how Excel's 3D mapping could be used to visualize the data through animation.

### 1. Data Cleaning ###

To start, the first thing I needed was a dataset with sales data. The dataset that I used is at https://data.world/dataman-udit/us-regional-sales-data (credit: Udit Kumar Chatterjee, @dataman-udit).

After downloading the dataset, I imported it into Microsoft Excel using *Get Data* > *From File* > *From Excel Workbook*.

<img width=300px length=300px src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning1.png>

Looking at the dataset, there are six tables, each on a different sheet. I could have chosen either the tables or worksheets, but I chose the worksheets to bring into power query.

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning2.png>

After loading in the data, I needed to consider data cleaning methods, including standardizing table column names, removing duplicates, filtering unnecessary data, converting data types, and more. Luckily, there was no duplicates or missing data to worry about.

I started by standardizing the column names to a single format. For example, I changed *OrderNumber* to *Order Number*. I also removed underscores when they were present.

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning3.png>

I realized on the regions sheet that the column names were occupying the first row, so I fixed this by selecting *Use First Row as Headers*.

<img width=55% src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning4.png>

Next, I wanted to remove unnecessary data. Before doing this, though, it was important for me to consider what questions I wanted to be answered by the dashboard. Considering this was a personal project, there were nobody to report to or any specific questions to answer. Therefore, I wanted to focus on sales, products, and customers. Also, I considered creating separate warehouse and sales team dashboards, so I wanted to keep that relevant data as well.

In Microsoft Word, I highlighted columns that existed in multiple worksheets to get a better understanding of the structure of the dataset.

<img width=55% src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/Table1.png>

I deleted columns I did not need by right-clicking on them and selecting *Remove*. The columns that I removed were: <br/>
Sales Orders Sheet - Order Number, Procured Date, Currency Code, and Discount Applied <br/>
Store Locations Sheet - County, State Code, Type, Latitude, Longitude, Location, Area Code, Population, Household Income, Median Income, Land Area, Water Area, and Time Zone <br/>
Regions Sheet - State Code

This data could be valuable when doing other analysis like regression, but I did not need it for the dashboard. Also, some data I kept would possibly not be used, but when it came to data that I was unsure about, it was best to keep it.

<img width=25% src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning5.png>

<img width=55% src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/Table2.png>

Next, I wanted to merge data from the other five sheets into the sales orders sheet so it would be simple to create pivot tables and pivot charts. 

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning6.png>

<img width=50% src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning7.png>

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning8.png>

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning9.png>

<img width=50% src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning10.png>

<img width=20% src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning11.png>

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning12.png>

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning13.png>

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning14.png>

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning15.png>

### 2. Pivot Tables ###

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/PivotTable1.png>

<img width=25% src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/PivotTable2.png>

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/PivotTable3.png>

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/PivotTable4.png>

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/PivotTable5.png>

<img width=30% src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/PivotTable6.png>

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/PivotTable7.png>

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/PivotTable8.png>

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/PivotTable9.png>

### 3. Pivot Charts ###

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/PivotChart1.png>

<img width=40% src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/PivotChart2.png>

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/PivotChart3.png>

### 4. Dashboard Creation ###

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/Dashboard1.png>

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/Dashboard2.png>

### 5. 3D Mapping ###

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/World1.png>

<img width=80% src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/World2.png>

<img width=25% src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/World3.png>
