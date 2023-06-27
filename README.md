# VBA-challenge
this is challenge for VBA

I started by declaring the variables that I would need in my code
Next, I allocated initial values to the variables.
To determine the yearly change, its percentage and stock volume for each ticker, I used a loop that would go through our data, adding up the volume values for one ticker and storing the opening value for the first day (this was done outside of the loop for the very first ticker) and once the loop determines the next ticker we take the close value for the last day for the ticker and from there we can calculate the yearly change and percntage change and allocate those values to cells. in our excel sheet (volume is calculated in the lioop and stored as well). Once the loop completes the calculations for one ticker, it moves on to the net, saving the opening value for the next ticker before starting again. We also of course re set the volume variable to 0 to start calculating the volume for the next ticker.

Next I declared additional variables to determine the greatest increase % , decrease % and greatest total volume. Using simple loops we go through the percentage change column and compare values to determine the percentages needed and the values are stored in the excel sheet as well.

Finally, to color code the yearly changes, I used a loop to color red for negative changes and green for positive changes. 

