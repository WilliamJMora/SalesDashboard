# SalesDashboard
Microsoft Excel Sales Dashboard

In this project, I created a Microsoft Excel sales dashboard featuring pivot charts and slicers. I also looked at Excel's three-dimensional animated map. Here is a look at the end result:

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/Dashboard1.png>

## Table of Contents ##

1. Data Cleaning
2. Pivot Tables
3. Pivot Charts and Slicers
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

Next, I wanted to merge data from the other five sheets into the sales orders sheet so it would be simple to create pivot tables and pivot charts. I did this by selecting *Merge Queries*.

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning6.png>

There were five merges that I wanted to complete. The first merge was with the customers sheet. I selected the customer ID columns from both tables as the matching columns and I made the sales orders table the primary table. I selected left outer join as the join type since I wanted all of the rows from the left table (sales orders sheet) and the matching rows from the right table (customers sheet).

<img width=50% src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning7.png>

I expanded the customers sheet that was now present in the sales orders sheet and selected the columns I wanted to retain, which was only the customer names. Then, I removed the customer ID column from the sales orders sheet since I no longer needed it.

The four other merges I completed were: <br/>
Sales Orders Sheet (Sales Team ID) > Sales Teams Sheet (Sales Team ID) => kept sales team and region from sales teams sheet, removed sales team ID from sales orders sheet <br/>
Sales Orders Sheet (Product ID) > Products Sheet (Product ID) => kept product name from products sheet, removed product ID from sales orders sheet <br/>
Sales Orders Sheet (Store ID) > Store Locations (Store ID) => kept city and state from store locations, removed store ID from sales orders sheet <br/>
Sales Orders Sheet (State) > Regions Sheet (State) => kept region from regions sheet, but did not remove state from sales orders sheet

I changed region from the sales teams sheet to sales team region since there were two region columns (the other referring to state regions).

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning8.png>

After the merges, I wanted to create some new columns using calculations. I started by selecting *Custom Column*.

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning9.png>

I created price by multiplying order quantity by unit price. I also made a cost column by multiplying order quantity by unit cost, then a profit solumn by subtracting cost from price. I removed the unit price and unit cost solumns thereafter.

<img width=50% src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning10.png>

I converted these columns to currency by selecting the data type icon on the column header and choosing *Currency*. Another option was to change the data types after loading the data into Excel, which is what the screenshot below is from.

<img width=20% src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning11.png>

Then, I created columns that found the difference between order date and ship date (named order to ship) and ship date and delivery date (named ship to delivery). This was done by selecting the two columns, then choosing *Date* > *Subtract Days*.

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning12.png>

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning13.png>

Lastly, I wanted to get the month and day from both the order date and ship date. I did this by selecting the column and just like before, went to *Date* > *Name of Month*. The same method applied for the day.

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning14.png>

I seleted *Close & Load* to bring in the changes to Microsoft Excel. Now, I was ready to create pivot tables and slicers.

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

### 3. Pivot Charts and Slicers ###

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
