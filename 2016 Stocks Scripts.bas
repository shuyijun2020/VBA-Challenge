Attribute VB_Name = "Module1"
Sub Multiple_Year_stock_summary()

 Dim WorksheetName As String

    
    
    'Headings of sum table
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "% Yearly Change"
    Range("L1").Value = "Total Volume"
    Range("O2").Value = "Greatst % Increase"
    Range("O3").Value = "Greatst % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    

    
    'Set initial variable for ticker symbol
    Dim Ticker As String
    
    Dim GIncreaseVTic As String
    Dim GDecreaseTic As String
    Dim GIncreaseTic As String
    
    
    'Set initial variables for yearly change
    Dim Yearopen As Double
    Dim Yearclose As Double
    Dim Percentchange As Double
    
    'Set initial variable for total volume
    Dim TotalV As Double
    TotalV = 0
    
    'Set variables for greatest increase in price and volume, decrease in price
    Dim GIncrease As Double
    Dim GDecrease As Double
    Dim GIncreaseV As Double
    
    GIncrease = 0
    GIncreaseV = 0
    GDecrease = 9999

    
    'Find lastrow count
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
        MsgBox ("Last row in column A is " & lastrow)
    
     
    
    'Keep track of the location for ticker count
    
    TickerRow = 2
    
        
    'Yr opening price
    Yearopen = Cells(2, 3).Value
    
    
    'Loop through all ticker rows
    For i = 2 To lastrow
    
        'Check if we still within the same ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
         
        Ticker = Cells(i, 1).Value
        
        
        'Yr close price
        Yearclose = Cells(i, 6).Value
        
        'Yr price change
        Yearchange = Cells(i, 6).Value - Yearopen
        
              
        'Exclude zero year open price
        If Yearopen > 0 Then
        
                
        'Percent yearly change in price
        Percentchange = Yearchange / Yearopen
        
        Else
        
        Percentchange = 0
        
        End If
        
        'Yr opening price
        Yearopen = Cells(i + 1, 3).Value
        
        
        
        'Add to ticker total volume
        TotalV = TotalV + Cells(i, 7).Value
        
        'Greatest increse V
        If TotalV > GIncreaseV Then
        GIncreaseV = TotalV
        
        GIncreaseVTic = Cells(i, 1)
               
        End If
        
        
        'Greatest decrease %
        If Percentchange < GDecrease Then
        GDecrease = Percentchange
        
        GDecreaseTic = Cells(i, 1)
        
        End If
        
        'Greatest increase %
        If Percentchange > GIncrease Then
        GIncrease = Percentchange
        
        GIncreaseTic = Cells(i, 1)
        
        End If
        
        
        'Format column yearly change of opening price
        Range("J2:J" & lastrow).NumberFormat = "0.00"
        
        'Format column percentage of yearly changes, greatest % increase or greatest % decrease to percentage
        Range("K2:K" & lastrow).NumberFormat = "0.00%"
        Range("Q2:Q3").NumberFormat = "0.00%"
        
        
        
        'Print the ticker and yearly change
        Range("I" & TickerRow).Value = Ticker
        
        'Print yearly change in opening price
        Range("J" & TickerRow).Value = Yearchange
        
        'Print percent yearly change in opening price
        Range("K" & TickerRow).Value = Percentchange
        
        'Print total volume
        Range("L" & TickerRow).Value = TotalV
        
        'Add row to the sum table
        TickerRow = TickerRow + 1
        
        'Reset total volume
        TotalV = 0
        
        'If same ticker
        Else
        
        'Add to ticker volume total
        TotalV = TotalV + Cells(i, 7).Value
    
               
        
        End If
    
    Next i
    
        Cells(2, 16).Value = GIncreaseTic
        Cells(2, 17).Value = GIncrease
        Cells(3, 16).Value = GDecreaseTic
        Cells(3, 17).Value = GDecrease
        Cells(4, 16).Value = GIncreaseVTic
        Cells(4, 17).Value = GIncreaseV

        'Format column yearly change of price
        Range("J2:J" & lastrow).NumberFormat = "0.00"
        
        Dim Color As Range
        Set Color = Worksheets("2016").Range("J2:J" & lastrow)
        
        For Each Cell In Color
        
        If Cell.Value < 0 Then
        Cell.Interior.ColorIndex = 3
        ElseIf Cell.Value > 0 Then
        Cell.Interior.ColorIndex = 4
        Else
        Cell.Interior.ColorIndex = xlNone
        
        End If
        
        Next



End Sub
