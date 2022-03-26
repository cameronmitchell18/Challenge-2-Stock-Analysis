# VBA Challenge :dollar: :chart_with_upwards_trend:

## Overview of Project: 

This project was created to determine if a certain list of stocks were profitable by the end of the given year or not. Steve, a finacial advisor, wanted use to determine if certain stocks made or lost money, to advise his parents against putting all of their money into a single a company. By creating a Macro friendly Excel File, we were able to give Steve a dynamic file that lets him choose which year he wants the stock results from. We made this file easy to read to by giving him the stocks total volume for the year and a percentage of their loss or gain that is color cordinated.  

## How we were able to accomplish this: 

Below are explinations of how we were able to accomplish this and their corrisponding code. 

### Creating the headings: 

Steve needed three column headings; Ticker, Total Daily Volume, and Return. We are able to accomplish this right from the VBA editor with some pretty simple code. 

'  
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return" '

This allows us to have nice, easy to read headings at the top of our excel file so we know what everything represnets. 

### Defining our tickers: 

We needed to define our tickers and organize them into an array so we could later call on them to how well, or not so well, they performed at the end of the year. We defined our tickers ans their array using the block of code listed below: 

' Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"  '


### Getting Excel to work for us: 

instead of filtering each years work book to the ticker we wanted, then calculating the their Total Daily Volume, and then calculating if that stock made or lost money that year, we wanted to Excel to do the heavy lifiting for us. We were able to accomplish this by indexing the tickers and then looping through them. From there, we told excel that for each ticket, calculate that stocks tickers Total Daily Volume and yearly percemtage return. The below block of code shows you how by a few simple line's of code we can acchive that: 

'  Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    
    tickerIndex = 0 

    '1b) Create three output arrays   
    
    Dim tickerVolumes(12) As Long 
    Dim tickerStartingPrice(12) AS Single
    Dim tickerEndingPrice(12) As Single

    ''2a) Create a for loop to initialize the tickerVolumes to zero. 
    For i = 0 To 11

        tickerVolumes(i) = 0 
        tickerStartingPrice(i) = 0
        tickerEndingPrice(i) = 0 

    Next i
        
    ''2b) Loop over all the rows in the spreadsheet. 
    For i = 2 To RowCount

        '3a) Increase volume for current ticker

       tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i , 8).Value

        '3b) Check if the current row is the first row with the selected tickerIndex.
        
        If Cells(i , 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then

            tickerStartingPrice(tickerIndex) = Cells(i , 6).Value
         
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        
        If Cells(i , 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickersIndex) Then

            tickerEndingPrice(tickerIndex) = Cells(i , 6).Value
            
        End If 

            '3d Increase the tickerIndex. 
            If Cells(i , 1).Value = tickers(tickerIndex) And Cells(i + 1 , 1).Value <> tickers(tickerIndex) Then 
                
                tickerIndex = tickerIndex + 1 
            
            End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate

        Cells(4 + i , 1).Value = tickers(i)
        Cells(4 + i , 2).Value = tickerVolumes(i)
        Cells(4 + i , 3).Value = (tickerEndingPrice(i)) / (tickerStartingPrice(i)) - 1
        
        
    Next i 

        '


### The finishing touches 

We dont our spread sheet to look dull and boring. We want to add some color and life into it. By adding color not only does that give the Excel sheet some life, but it makes it much easier to read at just a glance. We marked the stocks who lost money for that given year in red and the ones who made money in green for easy and familiar readability. The block of code below highlights how that is accomplished: 

'  Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    '

## How did the stocks perform? 

By running the Macro, you can enter in which year of performances you would like to see. In this case, we will start with 2017. 2017 every stock but 1 was in the green. Indicating that they made money that year. In contrast, to 2018 where we can see that all stocks except 2 were in the red meaning they lost money. In this case, it is highly advisable for Steves parents to spread their investments out because if they invested all of their money in 1 company, they have huge potential to loose lots of money. 

## Refactoring a script:

I think that refactoring a script is so useful and beneficial for someone who is just learning how to code and still wrestling with the concepts. It gives an oversight of how the script should run but then wrestling with how to make it run quicker and with less lines of code. The down side to refactoring code would be that you may not work as hard on those ideas because you already know that the code works. This might lead to some hesitation in changing the code because you might break it and freak out because it doesnt run anymore. That though, in my opinion, is all in the journey of learning to code. Breaking code often times is the best teacher. Staying diligent and keeping with it until it runs again and hopefully even quicker, is where some might see the downside. 

The pros of refactoring the code we already wrote in the practice modules were that we were familiar with what the instructions were asking us to do. We could look back at our practice scripts and notes to get a reminder of what that directions code looked like. 

The cons though, the practice module scripts were all over the place though. The Modules have you create what seemed like 12 different scripts just to work on the same file. I wish more than anything the Modules would tell you to stop creating new modules and just build off the one you are already working on and starting with. It just becomes confusing when you have to go dig for code in a seperate module when really, it should all be in one place. 


