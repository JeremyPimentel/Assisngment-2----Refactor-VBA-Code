# Assisngment-2----Refactor-VBA-Code
Refactoring VBA Code for sample stock data.
- by: Jeremy Pimentel

# Overview of Project
- This project aims to continue to help develop essential skills for VBA coding. This project asked for specific details of a subroutine including creating variables, for loops, If-then statments and finding time to run a script. 


## Results
- Using an array of all 12 tickers we are able to use For loops to loop through each ticker symbol and find total volume, starting prices and ending prices of each stock. 

![Array_1](https://user-images.githubusercontent.com/107723677/178126591-b2a40888-b15b-4a17-ab4b-722ea9583f3b.PNG)

- Using this setup we can run through each ticker using a for loop and i as a placeholder for the tickers 0 to 11. 
This for loop goes through each row, checks that the ticker symbol of the row before and after is the same, if it is, then it adds the volume. 


![For loop with array function](https://user-images.githubusercontent.com/107723677/178126801-529692ef-330f-4c1f-a056-5b5ebc221e7d.PNG)

- After going through each ticker, a final loop outputs each variable: tickers(i), tickerVolumes(i) and return (tickerEndingPrices(i) / tickerStartingPrices(i) - 1)


### 2017
![outputs the data](https://user-images.githubusercontent.com/107723677/178126819-b51a4dd7-eb57-47b0-81dc-58c8bba96b4d.PNG)

- From our outputs on the All Stocks Analysis worksheet we can see that in 2017, all stocks saw profits (Returns) except for stock ticker TERP which had a return of -7.2%. 
- The gains the other stocks showed ranged from 5.5% to 199.4% with RUN being the lowest and DQ being the highest. 
- The total daily volume for these stocks ranged from 35,796,200 to 782,187,000 with DQ having the lowest volume and SPWR having the highest volume. 


![VBA_Challenge_2017](https://user-images.githubusercontent.com/107723677/178127114-eb4440b8-95d1-473d-abb8-141b1856b743.PNG)

### 2018
- From 2017 to 2018 most stocks turned negative profits excpet for ENPH and RUN which saw returns of 81.9% and 84% respectively. The range of returns changed to -62.6% to 84.0% and the range of Total Daily Volume changed  83,079,900 to 607,473,500 with AY having the lowest volume and ENPH having the highest volume. 


![VBA_Challenge_2018](https://user-images.githubusercontent.com/107723677/178127115-8672fa30-004b-49e6-a3dc-25b056319cef.PNG)


## Summary
### Advantages & Disadvantages of Refactoring Code
- One advantage of refactoring code is that we can make the code simpler and easier to understand, which can help when we re-use the code for other projects and also help reduce run time. 
- A disadvantage of refactoring is that we are likely going to encounter errors when editing the original code and may not be able to get the code to work as it was intended. For example when I introduced a new variable in the refactored code, I had introduced an error. 
### Advantages & Disadvantages of Original and Refactored Code
- The original code had the advantage of being simpler in terms of theory. However this comes with a disadvantage because we are making the overall code messy and less easy to understand. 
- The advantage of the refactored code is that that it makes the overall finished code a lot more clean. The use of an array makes it so we dont need to repeat loops for each ticker symbol and can instead use a much more condensed code block to achieve the same result. Again this can be a disadvantage because it can be harder to code and can also introduce new bugs or errors. 
