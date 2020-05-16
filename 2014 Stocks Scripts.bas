Attribute VB_Name = "Module1"
Sub Creditcardcharges()
    'Set an initial variable for holding the brand name
    Dim Brand_Name As String
    
    
    'Set an initial variable for holding the total per CC brand
     Dim Total As Double
     Total = 0
    
    'Keep track of the location for each CC brand in the summary table
    Dim Sum_Table_Row As Integer
    Sum_Table_Row = 2
 
    
    'Loop through all credit card purhases
    For i = 2 To 100
    
        'Check if we still within the same credit card brand, if we are not
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        'Set brand name
        Brand_Name = Cells(i, 1).Value
        
        'Add brand total
        Total = Total + Cells(i, 3).Value


     
        
        'Print the CC brand in table
        Range("G" & Sum_Table_Row).Value = Brand_Name
            
        'Print the brand amount in table
        Range("H" & Sum_Table_Row).Value = Total
            
        'Add one to the sum table row
        Sum_Table_Row = Sum_Table_Row + 1
            
        'Reset the brand total
        Total = 0
        
           'If same brand
        Else
        
                   
        'Add to brand total
        Total = Total + Cells(i, 3).Value
        
        End If
        
    Next i
    
    End Sub
        
        
