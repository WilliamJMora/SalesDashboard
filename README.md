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

I created sale by multiplying order quantity by unit price. I also made a cost column by multiplying order quantity by unit cost, then a profit solumn by subtracting cost from price. I removed the unit price and unit cost solumns thereafter.

<img width=50% src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning10.png>

I converted these columns to currency by selecting the data type icon on the column header and choosing *Currency*. Another option was to change the data types after loading the data into Excel, which is what the screenshot below is from.

<img width=20% src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning11.png>

Then, I created columns that found the difference between order date and ship date (named order to ship) and ship date and delivery date (named ship to delivery). This was done by selecting the two columns, then choosing *Date* > *Subtract Days*.

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning12.png>

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning13.png>

Lastly, I wanted to get the month and day from both the order date and ship date. I did this by selecting the column and just like before, went to *Date* > *Name of Month*. The same method applied for the day.

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning14.png>

I selected *Close & Load* to bring in the changes to Microsoft Excel. Now, I was ready to create pivot tables.

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/DataCleaning15.png>


### 2. Pivot Tables ###

To start creating pivot tables, I selected *Pivot Table* > *From Table/Range* and then highlighted the data in the sales orders sheet.

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/PivotTable1.png>

The first table I created displayed sales by sales channel. I put sales channel in the rows field and sale in the values field.

<img width=25% src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/PivotTable2.png>

I changed the sales data to the currency data type.

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/PivotTable3.png>

The next table I wanted to create was the average amount of time it took for each warehouse to ship packages. This is why I created the order to ship column earlier. First, I copied and pasted the first pivot table next to the first one. Then, I followed the same process as above with dragging the warehouse into the filters field and order to ship into the values field. I had to change the aggregation of the order to ship column from sum to average. I did this by selecting the drop-down arrow, *Value Field Settings...* > *Average*. 

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/PivotTable4.png>

I rounded the values to two decimal places using the numbers tab.

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/PivotTable5.png>

<img width=30% src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/PivotTable6.png>

The next pivot table I made was simply a sales aggregate. I dragged sales into the values field and made sure the aggregation was sum.

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/PivotTable7.png>

Going back to the sales by sales channel pivot table, I sorted the data from largest to smallest so the pivot chart goes from largest to smallest.

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/PivotTable8.png>

This shows all of the pivot tables I created. Now, I could use these tables to create pivot charts.

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/PivotTable9.png>


### 3. Pivot Charts and Slicers ###

To make pivot charts, I put my cursor on the relevant pivot chart and selected *Pivot Chart*.

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/PivotChart1.png>

For the sales by sales channel chart, I went with a pie chart to show how much money in sales were made by each sales channel as a whole. If I wanted the raw number of sales or the money made by each sales channel, I would have went with a bar chart. This article (credit: Janis Gulbis) gives an overview of which chart to use given what is wanted to be visualized: https://eazybi.com/blog/data-visualization-and-chart-types.

<img width=40% src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/PivotChart2.png>

I wanted green, blue, and purple colors on the dashboard, so I went with a built-in blue style. Styling can be discovered and adjusted with by exploring Excel. As a side note, the only chart that was not created with a pivot chart was the regions map, in which the regions sheet was used instead.

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/PivotChart3.png>

Once I finished making the pivot charts, I made the slicers. To do this, I selected *Slicer* while my cursor was in a pivot table. Since the slicers are connected to all of the charts (with the exception of the regions map), it did not matter which chart the cursor was in. 

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/Slicer1.png>

For the first slicer, I wanted to be able to filter the charts by order month, so I checked that category.

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/Slicer2.png>

To make all the charts update with the slicer, I selected *Report Connections...* and checked each pivot table with the exceptions of the average time pivot tables that I created.

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/Slicer3.png>

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/Slicer4.png>

Then, I made other slicers that filtered data by order year, city, state, and region. Now, I could finally put all of the data on the dashboard.

### 4. Dashboard Creation ###

Choosing the style of the dashboard depends on personal choice, but I made the gray background so the charts can stand out. This was done by inserting a rounded rectangle shape and removing the border. For each chart, I also inserted a rounded rectangle and then applied a shadow. The icons that are on each chart come with Microsoft Excel. I created a custom slicer style. All of these options can be found in the tabs on the top ribbon in Excel. The charts referring to products and customers are filtered by the top 5 and top 10 data values. This is done by selecting the relevant pivot table and then choosing *Value Filters* > *Top 10*.

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/Dashboard1.png>

This is what the chart looks like when filtering by the year 2019. If I were to filter by city, the state slicer would automatically choose the state that the city is in.

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/Dashboard2.png>

The warehouse pivot tables that I created were not used in this dashboard, but I can create a separate dashboard visualizing the warehouse data. The same can be said for the sales team data.


### 5. 3D Mapping ###

One other thing that I wanted to explore was 3D mapping. To create a three-dimensional map, I selected *3D Map* in the tours section.

<img src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/World1.png>

To see sales by city, I selected *Location* > *City* and *Height* > *Sales (Sum)*.

<img width=25% src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/World3.png>

<img width=80% src=https://github.com/WilliamJMora/SalesDashboard/blob/main/Pictures/World2.png>

One last thing I wanted to see was sales over time, so in the time field, I selected *Order Date (None)*. After that, I edited the time syntax that appeared in the top left corner of the map.

![3DSales](https://github.com/WilliamJMora/SalesDashboard/assets/116101032/5c338190-aa64-44ec-b193-9eddf0386321)

