# Stock Analysis Refactoring Challenge

## Overview of Project

### Purpose
Refactor an Excel/VBA Stock Analysis application to be scaleable to analyze larger data sets efficiently. Try to optimize
the code by reducing unneccessary steps, improving the logic and documenting it's operation thoroughly in the code.

## Analysis 

### Analysis of Stock performance between 2017 and 2018
I delivered charts showing stock performance in 2017 and in 2018.   The charts I delivered aren't as useful as they could be because they don't combine data over
the 2 years to easily display the trend in return per stock over two years, and instead are just a snapshot in time at the end of each year.   Nonetheless, it's easy to see even from the single chart per year that 2017 was a much "kinder" stock market for the stocks being evaluated than 2018. Most of the stocks
lost value in 2018.  
![2017 Return Performance](https://user-images.githubusercontent.com/107505166/176564320-28f86230-722d-416c-bc9b-24689c1b9f31.png)
![2018 Return Performance](https://user-images.githubusercontent.com/107505166/176564366-63db6a4a-ad04-4bd5-aa76-3f70d272e6b1.png)

<br> Using the example of DQ stock, it's easy to see why Steve's parents had interest in the DQ
stock based on it's performance in 2017, but it's performance in 2018 was abysmal.  

### Analysis of runtime performance of the Stock Analysis application after refactoring
<br> I delivered two charts showing the runtime performance of the application before and after refactoring.  In the first chart, all the trial runs are included and in the second chart (the one below) the "outlier" data - the very first run which takes susbstantially longer than subsequent runs is excluded.  Below is a chart showing performance for all subsequent runs after the initial run for both the original and the refactored Stock Analysis app.

![Run performance after refactoring - exlcluding initial run](https://user-images.githubusercontent.com/107505166/176564708-8d3502a1-cfc8-406b-b8e8-f4446069f1c9.png)

<br>The impact of the refactoring on the time performance of the Stock Analysis application looks iffy to me.  I don't know if that means I should have done many more trial runs to see a stronger trend or if I missed something obvious in
refactoring the code, or if given the small amount of data we're working with (roughly 3K rows of ticker data per year over 2 years) any performance improvements aren't very substantial.

<br>Regardless of the inconclusive results re: the performance run time of the refactored application, the refactored code should be easier to maintain as it's now
more logical and better documented.


## Results 

- What are advantages or disadvantages of refactoring code?
<br>There are mainly advantages to refactoring code.  Adding clear in-line documentation makes it easier for current and future programmers to understand what the code is doing and
to maintain and or modify the code.   Improving the flow of logic and use of memory will reduce the amount of time it takes the code to run and save time and
frustration for the users.<br>The only disadvantage I can think of in refactoring code is perhaps if one does it when the payoff won't be great, simply because it bugs them to see inefficient or inelegant programming style.  Refactoring code, even for the best purposes could easily introduce errors and takes time.  There should also be at least a rough estimate of Cost vs. Benefit ratio in undertaking the time to refactor code.

- How do these pros and cons apply to refactoring the original VBA script?
<br>There was a good payoff in refactoring the original VBA script.  It was a good way to learn about and practice using output arrays and made the program flow more logical and efficient and it helped me to understand the code better to go through and document it as part of the refactoring. By outputting the results to the screen only once at the end of program execution, it reduced annoying flickering that occurred in the original program which displayed results data row by row as the program ran.   It also certainly drove home the point that changing existing code
that works - even if maybe not ideally - can take time and introduce errors!  It was an interesting exercise to see if the refactoring
had a noticeable impact on performance (time per analysis run).

