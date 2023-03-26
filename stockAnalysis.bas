Attribute VB_Name = "Module1"

Sub stockAnalysis()

    ' add column titles to I, J, K, and L columns
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

    ' loop through all stocks for one year (one sheet)
    ''find last row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
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
    year = ActiveSheet.Name

    ' assemble year and date cells to identify the opening and closing dates
    StartDate = year + "0102"
    EndDate = year + "1231"
    MsgBox (StartDate + EndDate)

    
    ''loop through all rows
    For i = 2 To LastRow
    
        ' add to volume counter
        volCounter = volCounter + Cells(i, 7).Value
    
        ' if it's the first date, save that opening price
        If Cells(i, 2).Value = 20200102 Then
        
            openPrice = Cells(i, 3).Value
        
        End If
        
        ' if it's the last date, save the closing price
        If Cells(i, 2).Value = 20201231 Then
        
            closePrice = Cells(i, 6).Value
            
            ' add to the arrayRow counter
            arrayRow = arrayRow + 1
        
            ' calculate yearlyChange and insert into column J
            yearlyChg = openPrice - closePrice
            Cells(arrayRow, 10).Value = yearlyChg
            
            ' if yearly change is positive, color the yearlyChg cell green
            If yearlyChg > 0 Then
                
                Cells(arrayRow, 10).Interior.ColorIndex = 4
                
            End If
            
            ' if yearly change is negative, color the yearlyChg cell red
            If yearlyChg < 0 Then
            
                Cells(arrayRow, 10).Interior.ColorIndex = 3
                
            End If
        
            ' save the ticker symbol - insert into column I
            Cells(arrayRow, 9).Value = Cells(i, 1).Value
            
            ' insert volCounter into column L
            Cells(arrayRow, 12).Value = volCounter
            
            ' set volumn counter to 0
            volCounter = 0
            
            ' calculate % increase = yearlyChange ? Original Number ? 100.
            ''format perChg as percentage
            ''and insert % increase to column K
            perChg = yearlyChg / openPrice '* 100 (got rid of this for now)
            perChg = FormatPercent(perChg)
            Cells(arrayRow, 11).Value = perChg
            
            
            
        End If
        
    Next i
End Sub

