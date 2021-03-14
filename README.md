# Stocks Analysis with VBA 
## Overview of Project

### Purpose

The purpose of this project is to help Steve analyze green energy stocks in order to help his parents decide whether or not invest their money into DAQO new energy corporation, which is a company that makes silicon wafers for solar panels. In this project, I used Excel VBA to analyze the green energy stock data with 12 stocks to determine the yearly return for 2017 and 2018 of each stock. Also, I refactored the original code in order to improve the code performance in time efficiency.

## Results

### Example of code

The below image is an example of my code before refactoring:

The below image is an example of my code after refactoring:




### Stock Performance between 2017 and 2018

Stock Performance of 2017:

image

Stock Performance of 2018:

image

Comparing the above images, I found that the overall green energy stock performance of 2017 is much better than 2018, which indicates there is bubble exist in green energy market. Looking at DQ stock particularly, the yearly return for 2017 is 199.4%, while for 2018 is -62.6%. The yearly return is not increasing consistently, therefore, DQ stock is not worth to invest.


### Comparison of execution times of the original script and refactored script

- Original Script

- Refactored Script

## Summary

- Advantages and disadvantages of refactoring code.

  ###### Advantages:
    1.	It can improve the logic of the code to make other people easier to read and understand.
    2.	It can use less memory and make the code run faster.
 
  ###### Disadvantages:
    1.	It is very time consuming to refactor code when the code contains thousands and thousands of lines.
    2.	It is easy to make a mistake during the process of refactoring code and cause the code to fail to run correctly. 

- Advantages and disadvantages of the original and refactored VBA script

  ###### Advantage of the original VBA script:

    We can construct the code faster because we are writing it according to our own logic without worrying about the efficiency of the code.

  ###### Advantages of the refactored VBA script:

    1.	VBA script is easier for other people to read.
    2.	The code runs more efficient. According to the images above, we can see that the run time for refactored code is X times faster than the original code.

  ###### Disadvantages of the original VBA script:

     It takes longer time to run the code.
  
  ###### Disadvantages of the refactored VBA script:

     The process of refactoring code is very time-consuming, and it is very easy to make a mistake to make the code cannot run correctly.
