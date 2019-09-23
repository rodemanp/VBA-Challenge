'PVR - Create subroutine
Sub Stock()
    
'PVR - Create worksheet as a variable; in this case ws is Worksheet
Dim WS As Worksheet

    'PVR - Create worksheet loop
    For Each WS In ActiveWorkbook.Worksheets
    'PVR - Activate worksheets by entering "ws.Activate"
    WS.Activate
    
        'PVR - Create column labels for I-L; use cells
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
       
        'PVR - Create variable to hold/store Value; we will need start price and end price to get yearly change, ticker name, and percentage change
        Dim StartPrice As Double
        Dim EndPrice As Double
        Dim YearlyChange As Double
        Dim TickerName As String
        Dim PercentChange As Double
        
        'PVR - Create Total volume variable to get total volume
        Dim TotalVolume As Double
        TotalVolume = 0
        
        'PVR - Identify rows and columns as variable data points
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        
        'PVR - Identify and create variable for Ticker column
        Dim i As Long
        
        
        'PVR - Identify last row; 71226 is the last row between all the sheets but use sh.Cells(sht.Rows.Count, "A").End(xlUp).Row from https://www.thespreadsheetguru.com/blog/2014/7/7/5-different-ways-to-find-the-last-row-or-last-column-using-vba
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        
        'PVR - Test and verify
        MsgBox (LastRow)
        
        'PVR - Find initial start price (second row, and 3 column)  since column = 1 from above, we are adding two columns to the right
        StartPrice = Cells(2, Column + 2).Value
         
        'PVR -  Loop through all ticker names in Column "A"; 2 is start
        For i = 2 To LastRow
         
            'PVR - Check if we are still within the same ticker symbol, if it is not...
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                
                'PVR - Set Ticker name and where it will show
                TickerName = Cells(i, Column).Value
                Cells(Row, Column + 8).Value = TickerName
                
                'PVR - Set Close Price to store
                EndPrice = Cells(i, Column + 5).Value
                
                'PVR - make function to get Yearly Change (gettting start price our of for loop to find the looped end price)*important to know where you are writing this code
                YearlyChange = EndPrice - StartPrice
                Cells(Row, Column + 9).Value = YearlyChange
                
                'PVR - Now we need to add the Percentage Change by adding if...else if statement
                If (StartPrice = 0 And EndPrice = 0) Then
                    PercentChange = 0
                ElseIf (StartPrice = 0 And EndPrice <> 0) Then
                    PercentChange = 1
                Else
                    PercentChange = YearlyChange / StartPrice
                    Cells(Row, Column + 10).Value = PercentChange
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
                
                'PVR Add Total Volume for each ticker
                TotalVolume = TotalVolume + Cells(i, Column + 6).Value
                Cells(Row, Column + 11).Value = TotalVolume
                
                'PVR - locate where to put in summary table row
                Row = Row + 1
                
                'PVR - Must have reset statement to for Start Price
                StartPrice = Cells(i + 1, Column + 2)
                
                'PVR - Need to reset the Volume Total
                Volume = 0
            
            'PVR - Need to build elseif to create cells of the same ticker
            Else
                TotalVolume = TotalVolume + Cells(i, Column + 6).Value
            
            End If
            
        'PVR - finish looping through ticker column loop
        Next i
        
        'PVR - We need to find last row of yearly chnage per Worksheet; same as we did prior
        YearlyChangeLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
        
        'PVR - For loop "j" and Format and set the Cell Colors
        For j = 2 To YearlyChangeLastRow
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, Column + 9).Interior.ColorIndex = 10
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 3
            End If
        Next j
        
        'PVR - We need to create more labels for the summary table further to the right(01:Q3)
        Cells(2, Column + 14).Value = "Greatest % Increase"
        Cells(3, Column + 14).Value = "Greatest % Decrease"
        Cells(4, Column + 14).Value = "Greatest Total Volume"
        Cells(1, Column + 15).Value = "Ticker"
        Cells(1, Column + 16).Value = "Value"
        
        'PVR - Now we need to do a loop "z" to look through each row to find the greatest value and ticker
        For Z = 2 To YearlyChangeLastRow
            If Cells(Z, Column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YearlyChangeLastRow)) Then
                Cells(2, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(2, Column + 16).Value = Cells(Z, Column + 10).Value
                'PVR - Format to percentage
                Cells(2, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YearlyChangeLastRow)) Then
                Cells(3, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(3, Column + 16).Value = Cells(Z, Column + 10).Value
                'PVR - Format to percentage
                Cells(3, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YearlyChangeLastRow)) Then
                Cells(4, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(4, Column + 16).Value = Cells(Z, Column + 11).Value
         
            End If
        
        'PVR - End loop "z"
        Next Z
        
    'PVR - Loop through sheets
    Next WS
        
End Sub