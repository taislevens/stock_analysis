Overview of Project:
 To show the the volume and return of 11 stocks in the year of 2017 and 2018 using VBA commands.

Summary os results and comparison on codes
The data base used has 3012 rows which represents trnasactions by ticker and by day and 2 separate worksheet for each year (2017 and 2018). 
After analysing the data, it is noted that the volume of transactions in 2018 increase by 4% compared to 2017.
[Table comparison_collors.png](https://github.com/taislevens/Challenge_file/blob/main/Table%20comparison_collors.png)

It was also possible to calculate the return of the stock on that year by dividing the ending price to the closing price.
After running this to all the stocks, and even though the volume of transactions was higher in 2018, the return of most of the securities were negative, except for 2: ENPH and RUN.

In terms of the VBA code, the return was calculated using one loop:
[Return of stocks calculation.png](https://github.com/taislevens/Challenge_file/blob/main/Return%20of%20stocks%20calculation.png)
Based on the selected year choose at the start of the macro:
[Option at the start.png](https://github.com/taislevens/Challenge_file/blob/main/Option%20at%20the%20start.png)

The code was refactored which offered more organization and optimization when the code is run (more organize and simplified loops and less steps when running the code). This resulted on a less time used to run the code:
[Comparision Refactored and original.png](https://github.com/taislevens/Challenge_file/blob/main/Comparision%20Refactored%20and%20original.png)

Refactoring could also take a longer time to be fixed or adapted compared to creating the code from the start 
