# Stock Analysis
## Overview of Project
### The goal of this project is to review green energy stocks for our friend Steve's parents. Steve's parents are especially interested in the stock for the company DAQO (ticker: DQ). Stock data from 2017 and 2018 is analyzed to review total volume and percent return. The final results are displayed in one worksheet to provide a simplified review.
## Results
### As seen in the images below, the stocks performed stronger in 2017 than 2018. There are 3 columns in the analysis: Ticker, Total Daily Volume, and Return. A positive return is highlighted in green and a negative return is highlighted in red. In 2017, 11 of the 12 stocks had a positive return opposed to 2018 in which only 2 of the stocks had a positive return.
<img width="220" alt="2017 (1)" src="https://user-images.githubusercontent.com/67160240/166110701-b993fe54-3e21-43e1-9ac3-253c19378c53.PNG">
<img width="224" alt="2018 (1)" src="https://user-images.githubusercontent.com/67160240/166110709-05f7aaf2-df70-4fc3-b12e-3a800ef64c68.PNG">
### To arrive at the these results, VBA was utilized. A script was created that loops through the 2017 and 2018 daily stock data which is stored in separate worksheets separated by year.
<img width="425" alt="Loop" src="https://user-images.githubusercontent.com/67160240/166110985-95d72df6-18ba-4e68-91b7-865f7151eacd.PNG">
<img width="513" alt="Loop2 (1)" src="https://user-images.githubusercontent.com/67160240/166111001-57efb091-243a-4d57-a6b3-ff3e74d91b11.PNG">
### The code for the 2017 analysis ran in 0.46 seconds and the code for 2018 ran in 0.48 seconds.
