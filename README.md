# Stock Analysis
## Overview of Project
### The goal of this project is to review green energy stocks for our friend Steve's parents. Steve's parents are especially interested in the stock for the company DAQO (ticker: DQ). Stock data from 2017 and 2018 is analyzed to review total volume and percent return. The final results are displayed in one worksheet to provide a simplified review.
## Results
### As seen in the images below, the stocks performed stronger in 2017 than 2018. There are 3 columns in the analysis: Ticker, Total Daily Volume, and Return. A positive return is highlighted in green and a negative return is highlighted in red. In 2017, 11 of the 12 stocks had a positive return opposed to 2018 in which only 2 of the stocks had a positive return.

<img width="220" alt="2017" src="https://user-images.githubusercontent.com/67160240/166156222-8ed3c55d-8937-4136-b487-735da53d6b95.PNG">

<img width="224" alt="2018" src="https://user-images.githubusercontent.com/67160240/166156223-240cedbf-4168-4b57-9c62-6d2274294c7e.PNG">

### To arrive at the these results, VBA was utilized. A script was created that loops through the 2017 and 2018 daily stock data which is stored in separate worksheets separated by year.

<img width="425" alt="Loop" src="https://user-images.githubusercontent.com/67160240/166156232-340707b2-7a71-4714-a956-13e7e2cd8360.PNG">

<img width="513" alt="Loop2" src="https://user-images.githubusercontent.com/67160240/166156239-1732244f-90e1-4a95-9dc3-77f07a8a3be0.PNG">

### The code for the 2017 analysis ran in 0.46 seconds and the code for 2018 ran in 0.48 second.

![Screen Shot 2022-04-30 at 8 39 30 PM](https://user-images.githubusercontent.com/67160240/166156262-8bc3c202-ede2-44ee-9132-35fcc57607ab.png)

![Screen Shot 2022-04-30 at 8 43 44 PM](https://user-images.githubusercontent.com/67160240/166156272-6ba35169-b664-46c7-b1b1-c1961b2289f5.png)

##  Summary
### Refactoring was utilized to complete the final stock analysis. The advantage of refactoring is the result is having code that is more efficient, less complex, and easier to maintain. 
### In addition to a more simplified process, the analysis for each year ran quicker with the refactored code. Refactored 2017 ran in 0.46 seconds compared to 0.83 seconds and refactored 2018 ran in 0.48 seconds compared to 0.85 seconds.
