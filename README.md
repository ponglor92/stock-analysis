# stock-analysis
#VBA 
#developer
#stocks

#VBA Stock Analysis
##Overview of the project
###Purpose
- the purpose of this project is, we are editing the Stock Market Dataset with VBA so we can loop through all the data. Then, after refactoring, we will determine if it worked by running it and see if it successfully ran faster than before refactoring it. 
##Analysis and Challenge
###Analysis of VBA Challenge
- While refactoring the all stocks analysis refactoring file, I was not able to run the file succcessfully and get a time on both 2017 and 2018. However, while working through the file, what I had to do was write in my own codes using either "for loop" or "nested loop" techniques within the numbered steps that were in the starter code file. With the knowledge of VBA and what I learned in module 2, I tried to the best of my ability to write codes that would work and try to run. 
- The first step of the VBA challenge was to create a tickerIndex variable and setting it equal to zero, then secondly, I had to create arrays for tickers, tickerVolumes, tickerstartingprices and tickerendingprices. To do that, I used the function Dim. Third, I had to create a for loop to initialize the tickerVOlume to zero, and to do that I lsited the tickers and equaled it to zero. After that, I cerated another loop and had to write a script inside of step 2b. For that, I used the formula that was given on the hint in the module 2 challenge instruction: tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value. I then created if- then statements, and created another loop for all the rows and ran the file. Unfortunately, I could not get a time.
###Challenges of VBA Challenge
- I encountered many challenges and errors that I was able to firgure out because most were small mistakes, such as indenting or minor mispelling. However, for the major mistakes such as "overflowing", I was not able to figure it out even after doubling the variable.
##Results
- In order to perform refactoring, you must do it in small steps. However those small changes can make your code successful and leave the file working.Disadvantages are: it is a long procedure and complex. The refactoring process can affect test outcomes. THe advantages are: the logical errors can pop up and let us know what eneds to be fixed. Also, VBA can reveal patterns for us.
- These pros and cons apply to refactoring the VBA script by improving the code without changing much of the entire file.
