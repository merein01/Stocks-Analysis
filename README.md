# Stocks-Analysis
This is an analysis on Green stocks. 

## Overview of Project

The purpose of this project is to analyze and help Steve choose the best green stock based on its previous performance. In Excel VBA, we created scripts\ for Steve to run where he can compare past performance on each of the different stocks. We accomplished this by extracting data from the total daily volume and percentage of return. 

## Results

<img width="347" alt="Stocks_Outputs_2017" src="https://user-images.githubusercontent.com/105119376/170176422-8661f8f1-7bc5-4015-b6a0-22d82eb2c1a6.png">
<img width="348" alt="Stocks_Outputs_2018" src="https://user-images.githubusercontent.com/105119376/170176448-9c4753fd-02c4-4b3b-8e6c-dde48ade2362.png">

## Stock Performance

After viewing the tables above, we can conclude that in 2018, most of the stocks had a negative return compared to 2017. All of the stocks with a negative return are in red and the code run in VBA to run this was: ElseIf Cells(i, 3) < 0 Then  'Change cell color to red,  Cells(i, 3).Interior.Color = vbRed. There were only two stocks (ENPH and RUN) with positive return and were formatted in green. The code run for this was: If Cells(i, 3) > 0 Then 'change cell color to green, Cells(i,3).Interior.Color = vbGreen. Out of all the stocks, DQ had the biggest difference in return from 2017 to 2018.

## Run Time Performance

<img width="346" alt="Original_Run_Time_2017" src="https://user-images.githubusercontent.com/105119376/170176917-15c68c09-ef57-4cc1-9e72-7c991d33a853.png">
<img width="345" alt="Orginal_Run_Time_2018" src="https://user-images.githubusercontent.com/105119376/170176979-e96ed8dd-5257-4a4c-88b9-f1125196e2f6.png">
<img width="352" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/105119376/170176994-6a04e8e9-25c3-4b6e-b5b6-777830057e75.png">
<img width="349" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/105119376/170177014-c3826de1-5b36-41cb-a089-fdfacae2cc19.png">

I included 4 screenshots: the original run times for 2017 and 2018, and the elapsed run times after refactoring the code in VBA. As you can see, the elapsed run times were 0.5 seconds quicker for both years than their original run times. 

## Summary 

The biggest advantage of refactoring code in general is that it improves the database you’re working with, making it clearer and more concise to better function. It also proved to run quicker. The biggest disadvantage is that refactoring can be risky because you can easily make a mistake with creating scripts in VBA and it can make it time consuming debugging and figuring out what went wrong. 

With the original script, the advantage was that the code could’ve been easily troubleshooted 
and wasn’t too difficult to work with. The disadvantage was that the code took a while to write as each step was stand alone, which could explain why it took longer to run the code. With the refactored script, the code was more efficient and reduced the run time from the original script. The disadvantage was that it’s more time consuming as you are redoing the work you already done in the original script. 

