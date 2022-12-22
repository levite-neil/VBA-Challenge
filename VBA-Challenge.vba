Attribute VB_Name = "Module1"
Sub stockAnny():

'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.

'Declare variables
Dim worksheetname As String
Dim i, j As Integer

'Using the for loop to traverse the Worksheets
For Each ws In Worksheets


'Get the name of the work sheet.
worksheetname = ws.Name

'Get the last row of the work sheet
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set the formatting for the Percent Change and Greatest Values
ws.Range("K2:K" & LastRow).NumberFormat = "0.00 %"
ws.Range("Q2:Q3").NumberFormat = "0.00 %"

'Declare variables
Dim tickeranem  As String
Dim yearchange, percentchange, openingprice, closingprice As Double
Dim StockVolumeTotal, tickercount As Integer


'Reset values to zero
tickercount = 0
percentchange = 0
yearchange = 0
openingprice = 0
closingprice = 0
StockVolumeTotal = 0

'Set Headings
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"


    For i = 2 To LastRow
    
    
        'Begin the ticker counter and stock volume total
        StockVolumeTotal = StockVolumeTotal + ws.Cells(i, 7).Value
        tickercount = 1 + tickercount

        'Checking the Ticker name with the row that follows
        If (ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value) Then
            
            'See if the initial price has been collected
            If openingprice = 0 Then
                openingprice = ws.Cells(i, 3).Value
            End If

       'If the next row is not equal to current ticker and the above row is we can begin placing and calculating
        ElseIf (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) And (ws.Cells(i, 1).Value = ws.Cells(i - 1, 1).Value) Then
            
            'Get placement for rows
            RowCount = ws.Cells(Rows.Count, 9).End(xlUp).Row + 1
            
            'Get closing price and ticker name
            closingprice = ws.Cells(i, 6).Value
            tickername = ws.Cells(i, 1).Value
            
            'Assign values of ticker name approprate cells
            ws.Cells(RowCount, 9).Value = tickername
            yearlychange = (closingprice - openingprice)
            
            'Compare only the changes that are above zero and assigned index color
            If yearlychange > 0 Then
                ws.Cells(RowCount, 10).Value = yearlychange
                ws.Cells(RowCount, 10).Interior.ColorIndex = 4
            Else
                'Compare only the changes that are below zero and assign index color
                ws.Cells(RowCount, 10).Value = yearlychange
                ws.Cells(RowCount, 10).Interior.ColorIndex = 3
            End If
            
            'Calculate percentage change and assign percentage change and stock volume totals
            percentchange = (yearlychange / openingprice)
            ws.Cells(RowCount, 11).Value = percentchange
            ws.Cells(RowCount, 12).Value = StockVolumeTotal
            
            'Reset values to zero
            StockVolumeTotal = 0
            tickercount = 0
            openingprice = 0
            closingprice = 0
            tickername = ""
            percentagechange = 0
            
        End If
    Next i
    
    'Get row count for additional summary data and create headings for table
    DataRowCount = ws.Cells(Rows.Count, 9).End(xlUp).Row
    ws.Cells(1, 15).Value = ws.Name
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    'Declare variables
    Dim increasevalue, decreasevalue, volumevalue As Double
    Dim increasevaluename, decreasevaluename, volumename As String
    
    'Set variable to zero
    increasevalue = 0
    decreasevalue = 0
    volumevalue = 0
    
    'Using for loop to traverse the newly summerized data
    For j = 2 To DataRowCount
        
        'Only compare if it's greater than zero and the assigned increase value
        If ws.Cells(j, 10).Value > 0 And ws.Cells(j, 11).Value > increasevalue Then
            increasevaluename = ws.Cells(j, 9).Value
            increasevalue = ws.Cells(j, 11).Value
        'Only compare if it's less than zero and the assigned decrease value
        ElseIf ws.Cells(j, 10).Value < 0 And ws.Cells(j, 11).Value < decreasevalue Then
            decreasevaluename = ws.Cells(j, 9).Value
            decreasevalue = ws.Cells(j, 11).Value
        End If
        
        'Compare the assigned volume value to the current volume value
        If ws.Cells(j, 12).Value > volumevalue Then
            volumename = ws.Cells(j, 9).Value
            volumevalue = ws.Cells(j, 12).Value
        End If
        
            
    Next j
    
    'Fill in the table with the collected data
    ws.Cells(2, 16).Value = increasevaluename
    ws.Cells(3, 16).Value = decreasevaluename
    ws.Cells(4, 16).Value = volumename
    ws.Cells(2, 17).Value = increasevalue
    ws.Cells(3, 17).Value = decreasevalue
    ws.Cells(4, 17).Value = volumevalue
    
            
            

Next ws


End Sub

Sub reset_anny():

'This is used to rest the columns to quickly clear columns and re-run report

'Using the for loop to traverse the Worksheets
For Each ws In Worksheets
    
    ws.Columns("I:Q").Clear
    
Next ws
    

End Sub

