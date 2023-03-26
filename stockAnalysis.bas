Attribute VB_Name = "Module1"
Sub stockAnalysis()

    ' loop through all worksheets in file
    '' declare variable to refer to a worksheet for the loop
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        
    
        ' add column titles to I, J, K, and L columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ' initialize counters for greatest % increase, decrease, and volume
        greatestPercIncrease = 0
        greatestPercDecrease = 0
        greatestTotalVol = 0
    
        ' loop through all stocks for one year (one sheet)
        ''find last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' initialize vol counter
        Dim volCounter As Double
        volCounter = 0
    
        ' initialize year
        Dim year As String
        
        ' initialize openPrice, closePrice, yearlyChg, perChg, and arrayRow counters
        Dim openPrice As Double
        Dim closePrice As Double
        Dim yearlyChg As Double
        'Dim perChg As Double
        Dim arrayRow As Integer
        arrayRow = 1
    
        ' pull year from title of sheet
        year = ws.Name
    
        ' assemble year and date cells to identify the opening and closing dates
        StartDate = year + "0102"
        EndDate = year + "1231"
        'MsgBox (StartDate + EndDate)
    
        
        ''loop through all rows
        For i = 2 To LastRow
        
            ' add to volume counter
            volCounter = volCounter + ws.Cells(i, 7).Value
        
            ' if it's the first date, save that opening price
            If ws.Cells(i, 2).Value = StartDate Then
            
                openPrice = ws.Cells(i, 3).Value
            
            End If
            
            ' if it's the last date, save the closing price
            If ws.Cells(i, 2).Value = EndDate Then
            
                closePrice = ws.Cells(i, 6).Value
                
                ' add to the arrayRow counter
                arrayRow = arrayRow + 1
            
                ' calculate yearlyChange and insert into column J
                yearlyChg = closePrice - openPrice
                ws.Cells(arrayRow, 10).Value = yearlyChg
                
                
                ' if yearly change is positive, color the yearlyChg cell green
                If yearlyChg > 0 Then
                    
                    ws.Cells(arrayRow, 10).Interior.ColorIndex = 4
                    
                End If
                
                ' if yearly change is negative, color the yearlyChg cell red
                If yearlyChg < 0 Then
                
                    ws.Cells(arrayRow, 10).Interior.ColorIndex = 3
                    
                End If
            
                ' save the ticker symbol - insert into column I
                ws.Cells(arrayRow, 9).Value = ws.Cells(i, 1).Value
                
                ' insert volCounter into column L
                ws.Cells(arrayRow, 12).Value = volCounter
                
                ' if volCounter is greatest than greatestTotalVol
                If volCounter > greatestTotalVol Then
                
                    ' set greatestTotalVol = to volCounter
                    greatestTotalVol = volCounter
                    
                    ' record ticker
                    ws.Range("P4").Value = ws.Cells(i, 1).Value
                    
                    ' record greatestTotalVol
                    ws.Range("Q4").Value = greatestTotalVol
                
                End If
                
                ' set volumn counter to 0
                volCounter = 0
                
                ' calculate % increase = yearlyChange ? Original Number ? 100.
                perChg = yearlyChg / openPrice '* 100 (got rid of this for now)
                
                ' if perChg is greater that greatestPercChange record it
                If perChg > greatestPercIncrease Then
                
                    ' set greatestPercChange = perChg
                    greatestPercIncrease = perChg
                    
                    ' record ticker
                    ws.Range("P2").Value = ws.Cells(i, 1).Value
                    
                    ' record greatest percentage increase
                    ws.Range("Q2").Value = FormatPercent(greatestPercIncrease)
                
                End If
                
                ' if perChg is less than greatestPercDecrease record it
                If perChg < greatestPercDecrease Then
                
                    ' set greatestPercDecrease = perChg
                    greatestPercDecrease = perChg
                    
                    ' record ticker
                    ws.Range("P3").Value = ws.Cells(i, 1).Value
                    
                    'record greatest % increase
                    ws.Range("Q3").Value = FormatPercent(greatestPercDecrease)
                
                End If
            
                
                ''format perChg as percentage
                ''and insert % increase to column K
                perChg = FormatPercent(perChg)
                ws.Cells(arrayRow, 11).Value = perChg
                
                
            End If
        ' move to next row
        Next i
        
        
    ' autofit all the columns that we've put information into
    ws.Columns("I:I").EntireColumn.AutoFit
    ws.Columns("J:J").EntireColumn.AutoFit
    ws.Columns("K:K").EntireColumn.AutoFit
    ws.Columns("L:L").EntireColumn.AutoFit
    ws.Columns("O:O").EntireColumn.AutoFit
    ws.Columns("P:P").EntireColumn.AutoFit
    ws.Columns("Q:Q").EntireColumn.AutoFit
    ' move to next worksheet
    Next ws
End Sub


