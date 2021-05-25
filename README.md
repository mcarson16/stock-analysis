# stock-analysis
## **Overview of Project:** 
The objective was to compare stocks from a list of 12 Green Energy companies to advise a new investor's stock purchases. 

## **Results:**

In 2017, the top 3 stocks for return on investment were 
1. DQ - 199.4%
2. SEDG - 184.5%
3. ENPH - 129.5%

The 3 most traded stocks were
1. SPWR - 782,187,000
2. FSLR - 684,181,400
3. CSIQ - 310,592,800

![2017 Stock TDV and Return](https://user-images.githubusercontent.com/83254435/118413965-9bff0400-b667-11eb-9d84-fb4a61ad0627.PNG)

In 2018, only 2 stocks had a positive return on investment.
1. RUN - 84.0%
2. ENPH - 81.9%

The 3 most traded stocks were
1. ENPH - 607,473,500
2. SPWR - 538,024,300
3. RUN - 502,757,100

![2018 Stock TDV and Return](https://user-images.githubusercontent.com/83254435/118413967-9e615e00-b667-11eb-8971-0e47d0269943.PNG)

While ENPH took a downward turn between 2017 and 2018, RUN rose to become the highest performing stock in the array. Assuming the trend will hold, one could recommend investing primarily in RUN stock. Though ENPH's return fell, it remained positive, and it remained a highly traded stock. One could advise conservative investment in ENPH's stock.

## **Summary:**
Using VBA in Excel to analyze data from 2017 and 2018, we calculated the number of times each company's stock was traded in a given calendar year, and calculated the investment return if one had kept each company's stock for the given calendar year. The end user can add additional sheets for different years to analyze. They can also run the same analysis for companies not initially listed, if those companies' data is added to the spreadsheet for a given year, and their tickers are added to the list in the macro.

**Advantages:** Refactoring reduced redundancy, and sped up the analysis of data by taking fewer steps. 

**Disadvantages:** The process is very time consuming. The alteration of variables introduced numerous code breaks at first. These all had to be reconciled before the refactored code accomplished its purpose. Rather than editing the existing code from the original macro, it was easier to just reference the existing code to write a new macro for the intended purpose. 
