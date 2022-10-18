# Stock Analysis - Refactored

## Overview of Project

This project aims to help Steve build on the success of our VBA Analysis spreadsheet and create an even better spreadsheet that will be able to look at the entire stock market.

### Purpose

The purpose of this analysis was to take a working set of code that we used in the DQ Analysis example and refactor it to make the code more efficient. We’ll attempt to use a more efficient “For Loop” that will cycle through the data one time to collect the info rather than making multiple loops through the sheet. Once we have found a more efficient “For Loop” we’ll be able to measure the level of success by utilizing the timer function that is in Visual Basic. We’ll also be able to take a more efficient set of code an apply it to an even larger data set in the future.

## Results: 


<img width="272" alt="2018" src="https://user-images.githubusercontent.com/110848660/191589739-efc010d1-e386-46d1-b6c4-97fd0e371a68.png">
In this image, we can see the run time with the initial set of code for the 2018 stock data. The run time shows that it took just over a second to complete the code.

<img width="259" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/110848660/191589756-9288eb60-2105-49cf-bbe6-ce608b217e15.png">
This image shows the drop in run time from the initial code compared to the refactored code. The refactored code ran in .11 seconds compared to a full second with the initial code.  This was due to a simplified and more efficient set of code that we were able to use. 


<img width="239" alt="2017" src="https://user-images.githubusercontent.com/110848660/191589768-80828b88-22fe-40f0-a3e2-88bde9d77936.png">
We saw the same thing when we ran the original "All Stocks Analysis" macro for the 2017 set of data. The timer function shows that this macro ran in 1 second. 

<img width="272" alt="VBA_Challenge_2017 png" src="https://user-images.githubusercontent.com/110848660/191589780-5c6e1d91-0912-4f3c-ab1b-ed5793c7b9d2.png">
This image shows the similar results that we saw with the refactored code for the 2018 data where the refactored code ran more efficient. This set of code ran in .11 seconds compared to the non-factored version that ran in a second. 


<img width="266" alt="Screenshot 2022-09-21 155054" src="https://user-images.githubusercontent.com/110848660/191597506-6a2dc2be-4f9a-449a-8b20-fcf351323463.png">

We did this by creating 12 output arrays to hold the data in as you can see from the code below. You can see the code below that we used to help create the rays. This allowed us to store data in the arrays and cut down on having to run multiple loops which you can see from the photos of the refactored code below.


<img width="494" alt="image" src="https://user-images.githubusercontent.com/110848660/191826389-0d461ffe-c167-479a-984c-c7f844ae66e8.png">

<img width="541" alt="image" src="https://user-images.githubusercontent.com/110848660/191826486-f45d9b8d-8ec0-44d8-a272-9787d219cba5.png">

Once completed, we used the same coding pasting and formatting our data on the "All Stocks Analysis" worksheet


## Summary

### Advantages and Disadvantages of refactoring code

The main advantage of using refactored code is that it should make it more efficient to run. This can be measured in the run time of the code which we were able to see in our Module. It can also clean up the code to make it easier to read for anyone coming behind you who is also interested in refactoring it. That is also beneficial for anyone who is new to coding so they are able to see how the code should be designed while eliminating redundant tasks within the code. This was seen in our module with doing the original "All Stocks Analysis" exercise and then finding a more efficient way of doing it in Module 2. As someone who is new to coding, it was beneficial to compare the code between the two exercises and see ways to make it more efficient.

One disadvantage of refactoring code is that it requires manpower so you are using resources (time and money) looking for a solution that may not exist. You may even find a solution, but the changes may be so subtle that it may not be enough to justify having someone spending the time or money to look at it. That would especially be the case at a smaller company where resources are limited. It would be similar to hiring a consultant to come in and look at a business to find improvements. You may end up spending more time or money on the refactoring process without seeing enough of a benefit to justify the investment. The last disadvantage of refactoring the code is that you are taking a working set of code and making changes that may result in it not working. That is why it is important to keep testing the code frequently to make sure your solution is still working as it should. It’s also important to make notes or comments on what you’re changing in case you need to go back to it later.

## How do these pros and cons apply to refactoring the original VBA script?

In this example, we saw a big difference in how the code ran in the Module Challenge compared to how it ran in the original All Stocks Analysis code. There was a noticeable difference in the amount of time that it took for the code to run between the two examples. Using the first All Stocks Analysis macro, our run times were showing around 1 second. Using the refactored code, we saw that run time dropped to around .1 seconds. This was almost a full second difference. This may not seem like a lot but this was only for 12 tickers and that number would make a big diffference if we were doing a large portion of the stock market. 

One of the Cons from refactoring the code was the amount of time that I spent on this exercise. Being very new to coding, it took a lot of time and thought to get the code refactored down to a more efficient version. I found one of the best things for me was to print it out and work through it that way as opposed to scrolling through it on my screen. I was also able to reach out to one of the learning assistants in BCS. He was able to walk me through and show me my mistakes in my code. Once he did that and I could see it, it made much more sense to me. This exercise was not in a real world setting, but it showed me how frustrating it can be to go from a working set of code and run into error after error while trying to improve it. This is an example of the time and money (if someone was actually paying me) that could be tied up while refactoring code. 
