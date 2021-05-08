# Stock-Analysis

## Overview of the Project
Analyzing a large data set should be always done on one of the coding programs like VBA. In this project, I've helped Steve to analyize 12 stocks by looking at their daily volumes and annual returns and refactoring a code that's already used in my previous tasks. In order to do it, I've used the Visual Basic for Applications on Excel which includes a large data set of the 12 stocks for the years 2017 and 2018 to compare and analyze their perfromance.

## Results
### Explanation of the Refactored Code and the Original Scripts
  Using the original script of the code, we've made some changes to make our work more efficient and to improve the logic of the code to make it easier for future users to read. We've started by defining the variables, adding an input box to specify the year, and start the timer of the code. Next we have added headers for the output data to be understandable and readable and created arrays for the ticker, daily stock volume, the starting price, and the ending price. The last three arrays were not available in the original script since we were expecting to use only these 12 stocks. The most important part here is the loops and nested loops that are used to go over each row of data and each stock to extract the convenient total volume, the starting price, and ending price in order to determine the return and start our analysis. Finally, we have formatted the output with the suitable colors and number formats to create a nice visualization of data. For each loop in each stock, we were starting the total volume by zero and ending it with a bigger number due to the sophisticated looping formula that adds the total volume of a ticker to the next total volume of the same ticker. For example:
- In the original script, we have only used the StartingPrice and the EndingPrice variables without creating them as arrays:
  - Dim StartingPrice As Double
  - Dim EndingPrice As Double
- Regarding the total volume formula, it looked like the following by taking into consideration that the total volume is equal to zero:
  - totalvolume = totalvolume + Cells(j, 8).Value
- However, in the refactored script, we've created a new variable (TickerIndex) to be used as an index in the three arrays (TickerVolume, EndingPrice, and StartingPrice):
  - Dim tickerIndex As Integer
    - tickerIndex = 0
  - Dim tickerVolume(12) As Long
  - Dim tickerStartingPrices(12) As Single
  - Dim tickerEndingPrices(12) As Single
- Here is how the tickervolume formula looks like in the refactored code:
  - tickerVolume(tickerIndex) = tickerVolume(tickerIndex) + Cells(i, 8).Value
        
### Results Based on the Code Output
  After finishing the code and creating the buttons to activate it by one click, we can see that the most two attractive stocks in terms of their high total daily volume which is explained as high liquid stocks and their high and positive return are ENPH and RUN stocks as we see in the following image. 

![VBA_Challenge_2018](https://user-images.githubusercontent.com/80184581/117552835-d5a68e00-b01b-11eb-99ae-21929900614d.PNG)

  On the other hand, DQ stock that was chosen by Tom's parents has a low total daily volume and a negative return in 2018. We can also look at SPWR and FSLR stocks that are liquid stocks because of their high daily total volume, but they have a low and negative return. In this case, we can only look at ENPH and RUN stocks. Both of them have a large increase in their total daily volume between 2017 and 2018 according to the output of 2017 data set in the folowing image. 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/80184581/117552840-ddfec900-b01b-11eb-8ade-e82870fd3e0f.PNG)

  Although they have similar returns in 2018, the only stock that increased was RUN. We can recommend Tom's parents to invest in RUN's stock since it's a very liquid stock with a very high increase in return when all the other stock were decreasing. However, we an also advise them to diversify their portfolio and invest part of their fund in ENPH. I always recommend to diversify in more than once stock especially when we have another opportunity to invest in a good stock like ENPH, but it's also the decision of the investor. 
### Comparing the Time Execution Between the Original and Refactored Scripts
  As we see in the images above, the time that's taken to execute the refactored code is less than 0.2 seconds in 2017 and 2018. However, if we do the same analysis on the original script, we see that the time increases to more than 0.9 seconds as we see in the following two images below about the original scripts in 2017 and 2018. 

![VBA_Original_2017](https://user-images.githubusercontent.com/80184581/117553309-d2f96800-b01e-11eb-9f10-a5f0e9eab5f9.PNG)

![VBA_Original_2018](https://user-images.githubusercontent.com/80184581/117553311-d68cef00-b01e-11eb-8168-b39f775b457b.PNG)

  The advantages of the code will be discussed later, but, in brief, one of the most important advantage of the code refactoring is the execution time that's taken to present the code output. The main reason behing this conclusion is the ease in understanding and reading the code thourgh the system by using simplier codes and felxible ways to analyze. 

## Pros and Cons of Refactoring a Script and Refactoring the Original VBA Script
### Pros and Cons of Refactoring a Script
#### Pros of Refactoring the Script
   - Code refactoring makes the code more extendible to a wider range of data (In our example, we are now able to do analysis to more than 12 stocks)
   - Code refactoring makes the code easier to understand and read and less complex to maintain because of the well organized set of new variables and loops
   - Code refactoring makes the code faster to be executed as we see in the timer in the images above
#### Cons of Refactoring the Script
   - Code refactoring takes too much time to be well edited and updated to a newer version since the flow of code shouldn't be changed
   - Code refactoring may lead to increase the chances of mistakes since most of the code is updated and edited  
### Advantages and Disadvantages of Refactoring the Original VBA Script
  In general, refactoring the original script has given the opportunity for the code user to use the same code for a much bigger number of variables (stocks) in a shorter period of time as we see in the above images. In addition, we are able now to use the code in other places without changing the whole code. We can just change the variable names, numbers, and output to be able to apply other analysis. Finally, the time which was spent on refactoring the code is spent only once, so we don't need too much time to change the code to create other projects and analysis. 
