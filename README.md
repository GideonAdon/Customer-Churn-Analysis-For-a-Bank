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

[Data Analysis](https://github.com/GideonAdon/Customer-Churn-Analysis-For-a-Bank/edit/main/README.md#data-anlysis)

[Business Value of the Analysis](https://github.com/GideonAdon/Customer-Churn-Analysis-For-a-Bank/edit/main/README.md#-final-business-value-of-the-analysis)

[Recommendations](https://github.com/GideonAdon/Customer-Churn-Analysis-For-a-Bank/edit/main/README.md#recommendations)

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
**1 Determine the Overall Churn Rate**
Objective:
To know the percentage of customers that left the bank.

**Why?**
This helps management understand the severity of customer loss.
________________________________________
**Excel Formula:**
```=COUNTIF(M:M,1)/COUNTA(A:A)```

 **Explanation:**
-	```COUNTIF(M:M,1)``` → counts customers who exited
-	```COUNTA(A:A)``` → counts total customers
-	Result = churn percentage

 	**2️ Identify Which Geography Has the Highest Churn**

**Objective:**
To determine which country loses the most customers.

**Why?**
The bank can focus retention strategies in high-risk regions.
________________________________________
**How to Achieve It:**
Use a Pivot Table.

**Pivot Table Setup:**

**Area**	**Field**
Rows	Geography
Values	Exited (Sum)
________________________________________

**Insight Example:**

**Country	Churn**

Germany	Highest

France	Medium

Spain	Lowest
________________________________________
**Business Meaning:**
Customers in Germany may be dissatisfied with services, pricing, or customer experience.
________________________________________
**3️ Analyze Churn by Gender**

**Objective:**

To identify whether males or females churn more.

**Why?**
Helps the bank design targeted customer retention campaigns.
________________________________________
**Pivot Table:**
Area	Field
Rows	Gender
Values	Exited (Sum)
________________________________________
**Optional Formula:**

```=COUNTIFS(E:E,"Female",M:M,1)```

**Explanation:**
Counts females who exited.
________________________________________
**Analyze Churn by Age Group**

**Objective:**
To understand which age category churns the most.
Why?
Different age groups have different banking behaviors.
________________________________________
**Step 1: Create Age Groups**
Formula:
```=IF(F2<30,"Young",IF(F2<50,"Middle Age","Senior"))```
________________________________________
**Step 2: Use Pivot Table**
Rows	Age Group
Values	Exited (Sum)
________________________________________
**Expected Insight:**
Older customers may churn more than younger customers.
________________________________________
**5️ Determine Whether Active Members Churn Less**

**Objective**:

To see if active customers are more loyal.

**Why?**

Helps measure customer engagement effectiveness.
________________________________________

**Formula:**
```=COUNTIFS(K:K,1,M:M,1)```

**Explanation:**
Counts active members who churned.
________________________________________
**Compare With:**
```=COUNTIFS(K:K,0,M:M,1)```

👉 This shows inactive customers who churned.
________________________________________
**Expected Business Insight:**
Inactive members usually churn more.
________________________________________
**Analyze Relationship Between Number of Products and Churn**

**Objective:**
To determine whether customers with more products stay longer.

**Why?**
Cross-selling products may improve retention.
________________________________________
**Pivot Table:**
Rows	NumOfProducts
Values	Exited (Sum)
________________________________________
**KPI Formula**
 Average products used:
=AVERAGE(I:I)
________________________________________
**7️ Analyze Credit Score Impact on Churn**

**Objective:**
To see whether customers with poor credit scores churn more.

**Why?**
Financial risk can influence customer behavior.
________________________________________
**Formula:**
Average Credit Score:
```=AVERAGE(C:C)```
________________________________________
**Optional Analysis:**
Create score bands:
```=IF(C2<500,"Poor",IF(C2<700,"Good","Excellent"))```
Then build Pivot Table.
________________________________________
**8️ Analyze Customer Balance and Salary**
**Objective:**
To identify whether high-value customers are leaving.
Why?
Losing wealthy customers is costly for the bank.
________________________________________
**Average Balance:**
=AVERAGE(H:H)
________________________________________
**Average Salary:**
=AVERAGE(L:L)
________________________________________
**High Balance Category:**
```=IF(H2>100000,"High Balance","Normal Balance")```
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

**1. Addition of New Columns:**
Several new columns were added to enhance the dataset and enable more detailed analysis. These new columns included:

**1️ Churn Flag Column**
Formula
```=IF(M2=1,"Churned","Retained")```
Where:
-	1 = Customer left the bank 
-	0 = Customer stayed 
________________________________________
**Why Was It Created?**
The original Exited column contains only numbers:

**Exited**
1
0
1
While computers understand 1 and 0, business stakeholders often don't.
The Churn Flag converts technical data into business language:

**Exited	Churn Flag**
1	Churned
0	Retained
1	Churned
________________________________________
**Business Purpose**
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
**Business Question It Helps Answer**
-	How many customers churned? 
-	How many customers were retained? 
-	What percentage of customers were retained? 

**2️ Age Group Column**

**Formula**

=IF(F2<30,"Young",
IF(F2<50,"Middle Age",
"Senior"))
________________________________________
**Why Was It Created?**
The Age column contains many unique values:

**Age**

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
**Result**

**Age	Age Group**
25	Young
42	Middle Age
65	Senior
________________________________________

**Business Purpose**

Allows us to compare customer behavior by age segment.
Banks don't usually target:
"42-year-old customers"
They target:
"Young adults"
"Middle-aged professionals"
"Senior customers"
________________________________________
**Business Questions It Helps Answer**
-	Which age group churns the most? 
-	Which age group has the highest balance? 
-	Which age group uses the most products? 
-	Which age group should receive retention campaigns? 
________________________________________
**Example Insight**

After analysis you might discover:

**Age Group	Churn Rate**
Young	12%
Middle Age	18%
Senior	35%
Recommendation
The bank should investigate why senior customers are leaving at a higher rate.

**3️ Balance Category Column**

**Formula**

=IF(H2=0,"Zero Balance",
IF(H2<50000,"Low Balance",
"High Balance"))
________________________________________
**Why Was It Created?**

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
**Result**
Balance	Balance Category
0	Zero Balance
20,000	Low Balance
80,000	High Balance
________________________________________
**Business Purpose**
Helps identify which type of customers are leaving.
The bank wants to know:
-	Are low-value customers leaving? 
-	Are high-value customers leaving? 
-	Are dormant customers leaving? 
________________________________________
**Business Questions It Helps Answer**
-	Which balance segment has the highest churn? 
-	Are wealthy customers leaving? 
-	Are inactive customers maintaining zero balances? 
-	Which balance group generates the most revenue? 
________________________________________
**Example Insight**

**Balance Category	Churn Rate**

Zero Balance	10%
Low Balance	18%
High Balance	42%

**Recommendation**
High-balance customers are leaving at a much higher rate.
This is a major risk because these customers hold more money and are typically more profitable.


These additional columns are crucial for examining sales trends over time, identifying seasonal patterns, and determining peak sales periods.


________________________________________
### 🧹 DATA CLEANING OBJECTIVE
Before analysis, data must be cleaned to ensure accuracy.
________________________________________
**Cleaning Goals:**

| Task                | Purpose                      |
|---------------------|------------------------------|
| Remove Duplicates   | Prevent double counting      |
| Handle Missing Values | Improve accuracy           |
| Standardize Text    | Ensure consistency           |
| Validate Data Types | Avoid calculation errors     |
| Detect Outliers     | Improve reliability          |
________________________________________
## 🔧 DATA CLEANING FORMULAS
________________________________________
**Detect Missing Values**

```=IF(COUNTBLANK(A2:M2)>0,"Missing","OK")```
________________________________________
**Remove Extra Spaces**

```=TRIM(D2)```
________________________________________
**Convert Text to Uppercase**

```=UPPER(E2)```
________________________________________
**Detect Duplicate Customer IDs**

```=IF(COUNTIF(A:A,A2)>1,"Duplicate","Unique")```
________________________________________
**📊 KPI DASHBOARD OBJECTIVE**

The dashboard helps management:
-	Monitor churn performance
-	Track customer behavior
-	Make strategic decisions quickly
________________________________________
**📈 Recommended Dashboard KPIs**
| KPI                | Formula                               |
|--------------------|---------------------------------------|
| Total Customers    | =COUNTA(A:A)                          |
| Churned Customers  | =COUNTIF(M:M,1)                       |
| Churn Rate         | =COUNTIF(M:M,1)/COUNTA(A:A)           |
| Active Members     | =COUNTIF(K:K,1)                       |
| Avg Balance        | =AVERAGE(H:H)                         |
| Avg Credit Score   | =AVERAGE(C:C)                         |

________________________________________
### 🧠 FINAL BUSINESS VALUE OF THE ANALYSIS

**This analysis helps the bank:**
- Reduce customer churn
- Improve customer satisfaction
- Identify high-risk customers
- Improve customer engagement
- Increase profitability
- Make data-driven decisions
## Data Anlysis
**Customer Churn by Age Group**

![](https://github.com/GideonAdon/Customer-Churn-Analysis-For-a-Bank/blob/main/Churn%20by%20Age.png)

This chart shows the number of customers who have churned (left the bank) across three age segments:

| Age Group  | Churned Customers | % of Total Churn |
|------------|------------------:|-----------------:|
| Middle Age | 1,432             | 70.3%            |
| Young      | 309               | 15.2%            |
| Senior     | 297               | 14.6%            |
| **Total**  | **2,038**         | **100%**         |

**Key Insights for the Board**

**1. Middle-Aged Customers Drive the Majority of Churn**

The most significant finding is that **Middle-Aged customers account for over 70% of all customer losses.**
- 1,432 customers churned in this segment.
- This is almost **5 times higher** than either the Young or Senior segments.
- The bank's customer retention challenge is concentrated within this age category.

**Business Implication:**
Any reduction in churn among middle-aged customers will generate the greatest impact on customer retention and revenue preservation.

**2. Young and Senior Customers Show Similar Churn Levels**

The Young (309) and Senior (297) segments exhibit nearly identical churn volumes.

- Difference between the two groups is only 12 customers.
- Neither group appears to be a major contributor to overall churn.

**Business Implication:**
Retention efforts should remain in place, but they should not be the primary focus compared to the middle-aged segment.

**3. Potential Risk to Revenue**

Middle-aged customers are often:

- High-income earners
- Mortgage holders
- Credit card users
- Investors
- Multi-product customers

Losing customers from this segment may have a disproportionate effect on:

- Revenue
- Cross-selling opportunities
- Customer Lifetime Value (CLV)

**Business Implication:**
The financial impact of churn may be significantly higher than the customer count alone suggests.

**Possible Causes to Investigate**

The chart identifies **where churn is occurring,** but not**why.**

The next analysis should examine whether middle-aged churners have common characteristics such as:

- Low customer satisfaction
- Poor credit scores
- High account balances
- Limited product usage
- Long waiting times
- Lower engagement with digital channels
- Better offers from competitors

**Recommended Actions**

**Immediate Priorities**

**1. Launch a Middle-Age Retention Program**

Target customers aged approximately 35–55 with:

- Loyalty rewards
- Personalized offers
- Relationship management programs
- Mortgage and investment incentives

**2. Identify High-Value Customers at Risk**

Build a churn-risk dashboard that flags:

- Declining balances
- Reduced transactions
- Complaints
- Inactive accounts

**3. Conduct Root Cause Analysis**

Break down churn by:

- Geography
- Gender
- Credit Score
- Tenure
- Balance
- Number of Products

This will reveal the specific drivers behind the high churn rate.

**Board-Level Presentation Statement**

"Customer churn is heavily concentrated among middle-aged customers, who account for approximately 70% of all customer losses. This segment represents the bank's greatest retention risk and likely contains many high-value customers. Strategic retention initiatives focused on this demographic could significantly reduce overall churn and protect future revenue growth."

**Suggested Insight Title for Dashboard**

**"Middle-Aged Customers Represent 70% of Total Customer Churn – Primary Retention Risk Segment"**

This is a strong executive-level insight because it immediately highlights **where management should focus retention investments for maximum business impact.**
**Customer Churn by Gender**
![](https://github.com/GideonAdon/Customer-Churn-Analysis-For-a-Bank/blob/main/Churn%20by%20Gender.png)

**Overview**

This chart illustrates customer churn distribution by gender.

| Gender | Churned Customers | Percentage |
|--------|------------------:|-----------:|
| Male   | 1,124             | 55.2%      |
| Female | 914               | 44.8%      |
| Total  | 2,038             | 100.0%     |

**Key Findings**
**1. Male Customers Have the Highest Churn Rate**

Male customers account for **1,124** churned customers, representing approximately **55% of total churn.**

Female customers account for **914 churned customers**, representing approximately **45% of total churn.**

**Board Insight**

While churn affects both genders, male customers are leaving at a higher rate, contributing an additional **210 customer losses** compared to females.

**2. Gender Gap Exists but Is Not Extreme**

The difference between male and female churn is approximately 10 percentage points.

This suggests:

- Churn is not exclusively a male or female issue.
- Retention challenges exist across both customer groups.
- However, males appear to be slightly more vulnerable to churn.

**Business Implication**

A gender-specific retention strategy may help improve customer retention, particularly among male customers.

**3. Revenue Risk Assessment**

If male customers hold:

- Higher account balances
- More loans
- More investment products
- More premium banking services

Then the financial impact of male churn may be greater than the customer count indicates.

**Board Recommendation**

Management should assess:

- Average balance by gender
- Product ownership by gender
- Revenue contribution by gender
- Customer Lifetime Value (CLV) by gender

This will determine whether the higher male churn translates into higher revenue loss.

**Strategic Questions for Further Analysis**

The chart identifies who is leaving, but not why they are leaving.

The next phase should investigate:

**Male Customers**
- Are they concentrated in a specific age group?
- Do they have lower satisfaction scores?
- Are they using fewer banking products?
- Are they more sensitive to fees or interest rates?
**Female Customers**
- Are there specific customer segments with elevated churn?
- Are there service-related factors contributing to attrition?

**Recommended Actions**

**Priority 1: Focus on High-Risk Male Segments**

Identify male customers who exhibit:

- Falling balances
- Reduced transaction activity
- Declining product usage
- Customer complaints

Implement targeted retention campaigns before they exit.

**Priority 2: Segment by Age and Gender**

A powerful next analysis would be:
| Segment             | Churn Risk |
|---------------------|------------|
| Middle-Aged Male    | High       |
| Young Male          | Medium     |
| Senior Male         | Low        |
| Middle-Aged Female  | High       |
| Young Female        | Medium     |
| Senior Female       | Low        |


This segmentation often reveals the true churn drivers.

**Priority 3: Personalize Retention Strategies**

Instead of a generic retention program:

- Tailor offers to customer needs.
- Improve engagement through personalized communication.
- Offer loyalty incentives to at-risk customers.
**Board-Level Narrative**

"Customer churn is relatively balanced across genders; however, male customers represent 55% of all customer departures compared to 45% for females. While the difference is not substantial enough to indicate a systemic gender issue, it highlights an opportunity to investigate the drivers of churn among male customers. Future retention initiatives should focus on understanding the interaction between gender, age, product usage, and customer value to reduce attrition and protect revenue."

**Dashboard Insight Statement**

**"Male customers account for 55% of total churn, indicating a moderate but noteworthy retention risk compared to female customers."**

**Important Analyst Note**

Before concluding that men are more likely to churn, verify the overall customer base composition. For example:

- If 70% of all customers are male, then a 55% churn share actually indicates **better retention among males.**
- If customers are split 50/50 by gender, then males are genuinely churning at a higher rate.

**Customer Churn by Geography**

![](https://github.com/GideonAdon/Customer-Churn-Analysis-For-a-Bank/blob/main/Churn%20by%20Geography.png)

**Overview**

This chart shows the distribution of churned customers across four geographic categories.
| Geography | Churned Customers | Percentage of Total Churn |
| --------- | ----------------- | ------------------------- |
| France    | 686               | 33.7%                     |
| Spain     | 529               | 26.0%                     |
| Germany   | 505               | 24.8%                     |
| French    | 318               | 15.6%                     |
| **Total** | **2,038**         | **100%**                  |


**Key Findings**

**1. France Records the Highest Customer Churn**

France accounts for **686 churned customers**, representing approximately **34% of all customer losses.**

This makes France the largest contributor to overall churn among the geographic segments analyzed.

**Board Insight**

More than one-third of all customer departures originate from the French market, indicating that customer retention challenges are most pronounced in this region.

**2. Spain and Germany Show Similar Churn Patterns**

Spain and Germany contribute:

- Spain: 529 churned customers
- Germany: 505 churned customers

The difference between these two markets is minimal.

**Business Implication**

This suggests that customer attrition is not isolated to one market but may be driven by similar competitive, service, or product-related factors across these regions.

**3. Potential Data Quality Concern**

The chart contains both "France" and "French" as separate categories.

This is likely a data quality issue because:

France refers to a country.
French is a nationality or language, not a geographic location.
Analyst Recommendation

Before presenting final findings, validate the source data to determine whether:

"French" should be merged with "France", or
It represents a different category entered incorrectly.

If merged:

Geography	Churned Customers
France (Combined)	1,004
Spain	529
Germany	505

This would mean France contributes nearly 50% of total churn, significantly strengthening the business case for intervention in that market.

Strategic Business Implications
Customer Retention Risk

The concentration of churn in France suggests possible issues such as:

Competitive pressure
Customer dissatisfaction
Product-market fit challenges
Pricing concerns
Service quality differences

These factors should be investigated through deeper segmentation.

Revenue Impact

If French customers hold:

Larger balances
More investment products
More loans
Higher-value accounts

Then the financial impact of churn may exceed what customer counts alone indicate.

Board Recommendation

Analyze:

Average account balance by geography
Revenue contribution by geography
Product ownership by geography
Customer lifetime value by geography
Recommended Next-Level Analysis

To identify the root cause of churn, cross-analyze geography with:

Age Group
Which age segments are driving churn in France?
Gender
Are male or female customers more likely to leave in specific regions?
Credit Score
Is churn concentrated among high-risk customers?
Balance
Are high-balance customers leaving?
Number of Products
Do customers with fewer products churn more frequently?
Actions Recommended to Management
Immediate Priority

Conduct a focused churn investigation within the French market.

Medium-Term Priority

Develop region-specific retention campaigns, including:

Loyalty programs
Product bundling
Personalized offers
Customer engagement initiatives
Long-Term Priority

Build a geographic churn monitoring dashboard to identify emerging retention risks before they escalate.

Board-Level Narrative

"Customer churn is geographically concentrated, with France accounting for the highest number of customer departures at 686 customers, representing approximately one-third of total churn. Spain and Germany follow closely behind. However, a potential data quality issue exists with the separate classification of 'France' and 'French', which may understate the true churn concentration in the French market. Management should prioritize validation of geographic data and focus retention efforts on the regions contributing the highest customer losses."

Executive Dashboard Insight

"France is the largest contributor to customer churn, generating 34% of all customer departures and representing the bank's most significant geographic retention challenge."

Data Analyst Observation

Before presenting this chart to executives, I would recommend correcting the 'French' vs. 'France' categorization. Board-level decisions should be based on standardized geographic data to avoid misleading conclusions and ensure accurate targeting of retention strategies.

## Recommendations ##
Based on the findings from the analysis, the following recommendations are proposed:

**1. Prioritize Retention of Middle-Aged Customers**

Middle-aged customers account for the largest share of churn.

**Actions:**
Develop loyalty programs targeted at customers aged 35–55.
Offer personalized financial products and incentives.
Improve customer engagement through regular relationship management.

**Expected Outcome:**

Reduction in churn among the highest-risk customer segment.

**2. Increase Product Adoption**

Customers with only one product account for nearly 70% of total churn.

**Actions:**
Implement cross-selling campaigns.
Bundle products such as savings accounts, credit cards, loans, and investments.
Offer incentives for customers who adopt multiple products.
Expected Outcome:

Higher customer loyalty and lower churn rates.

**3. Focus Retention Efforts on High-Churn Geographic Markets**

France (and potentially the "French" category after data correction) contributes the largest number of churned customers.

**Actions:**
Conduct regional customer satisfaction surveys.
Analyze competitor activity in high-churn markets.
Develop location-specific retention campaigns.
Expected Outcome:

Improved customer retention in the most affected regions.

**4. Implement a Churn Prediction and Early Warning System**

Many churned customers show warning signs before leaving.

**Actions:**

Monitor:

Declining account balances
Reduced transaction frequency
Customer complaints
Product cancellations
Reduced digital engagement
Expected Outcome:

Early intervention before customers decide to leave.

**5. Improve Data Quality and Governance**

The presence of both "France" and "French" indicates potential data quality issues.

**Actions:**
Standardize geographic categories.
Establish data validation rules.
Regularly audit customer data.
Expected Outcome:

More accurate reporting and better decision-making.

**6. Focus on Churn Drivers Rather Than Gender**

Analysis shows that male and female churn rates are very similar.

**Actions:**
Avoid gender-based retention strategies.
Allocate resources to stronger churn predictors such as:
Age
Product ownership
Geography
Customer value
Expected Outcome:

More effective use of retention budgets.

**Final Executive Recommendation**

The bank should focus its retention strategy on middle-aged customers, single-product customers, and high-churn geographic markets. By increasing product adoption, improving customer engagement, and implementing a proactive churn monitoring system, the bank can significantly reduce customer attrition, protect revenue, and improve long-term customer lifetime value. These initiatives should be supported by strong data governance to ensure accurate and reliable decision-making.

### 🏁 Conclusion
As a professional data analyst, the goal is not just to create charts, but to:
-	Solve business problems
-	Discover meaningful insights
-	Help the company retain customers
-	Support strategic decision-making using data

This customer churn analysis transforms raw customer data into actionable business intelligence.

