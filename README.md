# Customer-Churn-Analysis-For-a-Bank
An Excel project to determine why customer churn happens, when customers stop doing business with a company or leave a service.

## Table of contents

[Introduction](https://github.com/GideonAdon/Customer-Churn-Analysis-For-a-Bank/edit/main/README.md#introduction)

[Aim of the Analysis](https://github.com/GideonAdon/Customer-Churn-Analysis-For-a-Bank/edit/main/README.md#aim-of-the-analysis)

[Objective of the Analysis](https://github.com/GideonAdon/Customer-Churn-Analysis-For-a-Bank/edit/main/README.md#objectives-of-the-analysis)

[Data Source](https://github.com/GideonAdon/Customer-Churn-Analysis-For-a-Bank/edit/main/README.md#data-source)

[Data Cleaning and Processing](https://github.com/GideonAdon/Customer-Churn-Analysis-For-a-Bank/edit/main/README.md#data-cleaning-and-processing)

[Data Cleaning Objectives](https://github.com/GideonAdon/Customer-Churn-Analysis-For-a-Bank/edit/main/README.md#-data-cleaning-formulas)

[Data Cleaning Formulas](https://github.com/GideonAdon/Customer-Churn-Analysis-For-a-Bank/edit/main/README.md#-data-cleaning-formulas)

[Business Value of the Analysis](https://github.com/GideonAdon/Customer-Churn-Analysis-For-a-Bank/edit/main/README.md#-final-business-value-of-the-analysis)

[Conclusion](https://github.com/GideonAdon/Customer-Churn-Analysis-For-a-Bank/edit/main/README.md#-conclusion)

## Introduction

Customer churn happens when customers stop doing business with a company or leave a service.
In this dataset, the bank wants to understand:

-	Why customers are leaving
-	Which customers are likely to churn
-	What factors influence churn
-	How the bank can improve customer retention

The dataset contains customer information such as:

-	Credit score
-	Geography
-	Gender
-	Age
-	Balance
-	Products used
-	Activity status
-	Salary
-	Churn status (Exited)
  ### AIM OF THE ANALYSIS
The main aim of this analysis is:
To identify the key factors causing customer churn and provide data-driven insights that help the bank reduce customer loss and improve retention.
________________________________________
### OBJECTIVES OF THE ANALYSIS
________________________________________
### 1 Determine the Overall Churn Rate
Objective:
To know the percentage of customers that left the bank.
Why?
This helps management understand the severity of customer loss.
________________________________________
### Excel Formula:
=COUNTIF(M:M,1)/COUNTA(A:A)
### Explanation:
-	COUNTIF(M:M,1) → counts customers who exited
-	COUNTA(A:A) → counts total customers
-	Result = churn percentage
### 2️ Identify Which Geography Has the Highest Churn
Objective:
To determine which country loses the most customers.
### Why?
The bank can focus retention strategies in high-risk regions.
________________________________________
### How to Achieve It:
Use a Pivot Table.
### Pivot Table Setup:
### Area	Field
Rows	Geography
Values	Exited (Sum)
________________________________________
### Insight Example:
### Country	Churn

Germany	Highest

France	Medium

Spain	Lowest
________________________________________
### Business Meaning:
Customers in Germany may be dissatisfied with services, pricing, or customer experience.
________________________________________
### 3️ Analyze Churn by Gender
### Objective:
To identify whether males or females churn more.
Why?
Helps the bank design targeted customer retention campaigns.
________________________________________
### Pivot Table:
Area	Field
Rows	Gender
Values	Exited (Sum)
________________________________________
### Optional Formula:
=COUNTIFS(E:E,"Female",M:M,1)
### Explanation:
Counts females who exited.
________________________________________
## 4️ Analyze Churn by Age Group
### Objective:
To understand which age category churns the most.
Why?
Different age groups have different banking behaviors.
________________________________________
### Step 1: Create Age Groups
Formula:
=IF(F2<30,"Young",IF(F2<50,"Middle Age","Senior"))
________________________________________
### Step 2: Use Pivot Table
Rows	Age Group
Values	Exited (Sum)
________________________________________
### Expected Insight:
Older customers may churn more than younger customers.
________________________________________
### 5️ Determine Whether Active Members Churn Less
### Objective:
To see if active customers are more loyal.
### Why?
Helps measure customer engagement effectiveness.
________________________________________
## Formula:
=COUNTIFS(K:K,1,M:M,1)
### Explanation:
Counts active members who churned.
________________________________________
### Compare With:
=COUNTIFS(K:K,0,M:M,1)

👉 This shows inactive customers who churned.
________________________________________
### Expected Business Insight:
Inactive members usually churn more.
________________________________________
### 6️ Analyze Relationship Between Number of Products and Churn
### Objective:
To determine whether customers with more products stay longer.
### Why?
Cross-selling products may improve retention.
________________________________________
### Pivot Table:
Rows	NumOfProducts
Values	Exited (Sum)
________________________________________
### KPI Formula:
 Average products used:
=AVERAGE(I:I)
________________________________________
### 7️ Analyze Credit Score Impact on Churn
### Objective:
To see whether customers with poor credit scores churn more.
### Why?
Financial risk can influence customer behavior.
________________________________________
### Formula:
Average Credit Score:
=AVERAGE(C:C)
________________________________________
### Optional Analysis:
Create score bands:
=IF(C2<500,"Poor",IF(C2<700,"Good","Excellent"))
Then build Pivot Table.
________________________________________
### 8️ Analyze Customer Balance and Salary
### Objective:
To identify whether high-value customers are leaving.
Why?
Losing wealthy customers is costly for the bank.
________________________________________
### Average Balance:
=AVERAGE(H:H)
________________________________________
### Average Salary:
=AVERAGE(L:L)
________________________________________
### High Balance Category:
=IF(H2>100000,"High Balance","Normal Balance")
## Data Source
This dataset is gotten from [Maven Analytics](https://app.mavenanalytics.io/datasets?search=ch) website, design specifically for practice purpose. It is presented in Excel file with 1 table, comprising of 15,863,431,470 rows and total of 13 columns respectively. The dataset includes key attributes essential for a comprehensive analysis, such as:

				
- Customer Id	A unique identifier for each customer
- Surname	The customer's last name	
- Credit Score	A numerical value representing the customer's credit score
- Geography	The country where the customer resides (France, Spain or Germany)
- Gender	The customer's gender (Male or Female)
- Age	The customer's age		
- Tenure	The number of years the customer has been with the bank
- Balance	The customer's account balance
- NumOfProducts	The number of bank products the customer uses (e.g., savings account, credit card)
- Has CrCard	Whether the customer has a credit card (1 = yes, 0 = no)
- IsActiveMember	Whether the customer is an active member (1 = yes, 0 = no)
- EstimatedSalary	The estimated salary of the customer
- Exited	Whether the customer has churned (1 = yes, 0 = no)

### Data Cleaning and Processing
Having clean data for analysis is essential because it ensures that the insights are accurate, reliable, and free from errors. Clean data makes the analysis process smoother, helping researchers, analysts, and decision-makers to draw meaningful conclusions and make informed decisions based on trustworthy information. After thoroughly checking the data for quality and suitability, including looking for errors, inconsistencies, missing values, and duplicates, I found that the dataset is well-organized and consistent except the order id column which requires the PROPER function to ensure that each word begin with an upper case. The data follows a standard format and uses consistent names. All the necessary fields are present and filled out correctly. The data values are accurate, and each column has the correct data type. Importantly, there are no duplicate records.
Based on this detailed review, the dataset is well-prepared and ready for analysis, ensuring that the insights I gain will be accurate, reliable, and trustworthy.
Additional Data Processing Steps:
### 1. Addition of New Columns:
Several new columns were added to enhance the dataset and enable more detailed analysis. These new columns included:
### 1️ Churn Flag Column
Formula
=IF(M2=1,"Churned","Retained")
Where:
-	1 = Customer left the bank 
-	0 = Customer stayed 
________________________________________
### Why Was It Created?
The original Exited column contains only numbers:
## Exited
1
0
1
While computers understand 1 and 0, business stakeholders often don't.
The Churn Flag converts technical data into business language:
## Exited	Churn Flag
1	Churned
0	Retained
1	Churned
________________________________________
### Business Purpose
Makes reports and dashboards easier to understand.
Instead of showing:
Exited = 1
You can show:
Customer Status = Churned
This is especially useful for:
-	Pivot Tables 
-	Charts 
-	Executive dashboards 
-	Presentations to management 
________________________________________
## Business Question It Helps Answer
-	How many customers churned? 
-	How many customers were retained? 
-	What percentage of customers were retained? 
## 2️ Age Group Column
## Formula
=IF(F2<30,"Young",
IF(F2<50,"Middle Age",
"Senior"))
________________________________________
## Why Was It Created?
The Age column contains many unique values:
## Age
22
23
24
25
26
58
67
Analyzing every age individually is difficult.
Instead, we group customers into meaningful categories.
________________________________________
## Result
## Age	Age Group
25	Young
42	Middle Age
65	Senior
________________________________________
## Business Purpose
Allows us to compare customer behavior by age segment.
Banks don't usually target:
"42-year-old customers"
They target:
"Young adults"
"Middle-aged professionals"
"Senior customers"
________________________________________
## Business Questions It Helps Answer
-	Which age group churns the most? 
-	Which age group has the highest balance? 
-	Which age group uses the most products? 
-	Which age group should receive retention campaigns? 
________________________________________
## Example Insight
After analysis you might discover:
## Age Group	Churn Rate
Young	12%
Middle Age	18%
Senior	35%
Recommendation
The bank should investigate why senior customers are leaving at a higher rate.
## 3️ Balance Category Column
## Formula
=IF(H2=0,"Zero Balance",
IF(H2<50000,"Low Balance",
"High Balance"))
________________________________________
## Why Was It Created?
The Balance column contains thousands of different values.
Example:
Balance
0
23,450
55,230
89,000
120,000
Comparing thousands of unique balances isn't practical.
Grouping balances creates meaningful customer segments.
________________________________________
## Result
Balance	Balance Category
0	Zero Balance
20,000	Low Balance
80,000	High Balance
________________________________________
## Business Purpose
Helps identify which type of customers are leaving.
The bank wants to know:
-	Are low-value customers leaving? 
-	Are high-value customers leaving? 
-	Are dormant customers leaving? 
________________________________________
## Business Questions It Helps Answer
-	Which balance segment has the highest churn? 
-	Are wealthy customers leaving? 
-	Are inactive customers maintaining zero balances? 
-	Which balance group generates the most revenue? 
________________________________________
## Example Insight
## Balance Category	Churn Rate
Zero Balance	10%
Low Balance	18%
High Balance	42%
Recommendation
High-balance customers are leaving at a much higher rate.
This is a major risk because these customers hold more money and are typically more profitable.


These additional columns are crucial for examining sales trends over time, identifying seasonal patterns, and determining peak sales periods.


________________________________________
### 🧹 DATA CLEANING OBJECTIVE
Before analysis, data must be cleaned to ensure accuracy.
________________________________________
### Cleaning Goals:
### Task	Purpose

Remove duplicates	Prevent double counting

Handle missing values	 Improve accuracy

Standardize text	 Ensure consistency

Validate data types	 Avoid calculation errors

Detect outliers	 Improve reliability
________________________________________
## 🔧 DATA CLEANING FORMULAS
________________________________________
### Detect Missing Values
=IF(COUNTBLANK(A2:M2)>0,"Missing","OK")
________________________________________
### Remove Extra Spaces
=TRIM(D2)
________________________________________
### Convert Text to Uppercase
=UPPER(E2)
________________________________________
### Detect Duplicate Customer IDs
=IF(COUNTIF(A:A,A2)>1,"Duplicate","Unique")
________________________________________
### 📊 KPI DASHBOARD OBJECTIVE
The dashboard helps management:
-	Monitor churn performance
-	Track customer behavior
-	Make strategic decisions quickly
________________________________________
### 📈 Recommended Dashboard KPIs
### KPI	Formula

Total Customers	=COUNTA(A:A)

Churned Customers	=COUNTIF(M:M,1)

Churn Rate	=COUNTIF(M:M,1)/COUNTA(A:A)

Active Members	=COUNTIF(K:K,1)

Avg Balance	=AVERAGE(H:H)

Avg Credit Score	=AVERAGE(C:C)
________________________________________
### 🧠 FINAL BUSINESS VALUE OF THE ANALYSIS
### This analysis helps the bank:
- Reduce customer churn
- Improve customer satisfaction
- Identify high-risk customers
- Improve customer engagement
- Increase profitability
- Make data-driven decisions
________________________________________
### 🏁 Conclusion
As a professional data analyst, the goal is not just to create charts, but to:
-	Solve business problems
-	Discover meaningful insights
-	Help the company retain customers
-	Support strategic decision-making using data

This customer churn analysis transforms raw customer data into actionable business intelligence.

