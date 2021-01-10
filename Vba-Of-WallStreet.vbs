Sub TckrFind()
    Dim ticker As String
    Dim TiSumTable As Integer
    Dim DiffTable As Integer
    Dim ChRateTable As Integer
    Dim TotSumTable As Integer
    Dim ClSumTable As Integer
    Dim OpSumTable As Integer
    Dim TotVol As Double
    Dim j As Integer
    Dim k As Integer
    
    '----Setting the column position to populate summary table----
    TiSumTable = Cells(5, Columns.Count).End(xlToLeft).Column + 2
    DiffTable = Cells(5, Columns.Count).End(xlToLeft).Column + 3
    ChRateTable = Cells(5, Columns.Count).End(xlToLeft).Column + 4
    TotSumTable = Cells(5, Columns.Count).End(xlToLeft).Column + 5
    
    'Not requested at the homework but included to compare results
    OpSumTable = Cells(5, Columns.Count).End(xlToLeft).Column + 6
    ClSumTable = Cells(5, Columns.Count).End(xlToLeft).Column + 7
    
    '----Setting the headers for Summary table----
    Cells(1, TiSumTable) = "Ticker"
    Cells(1, DiffTable) = "Yearly Change"
    Cells(1, ChRateTable) = "Percent Change"
    Cells(1, TotSumTable) = "Total Stock Volume"
    
    'Not requested at the homework but included to compare results
    Cells(1, OpSumTable) = "Open position BOY"
    Cells(1, ClSumTable) = "Close position EOY"
     
    '----Counters and sum variables----
    j = 2 'Counter to populate summary table
    TotVol = 0 'Variable where Volume per ticker is summed
    k = 1 'Conditional to check first Open position of each ticker
    
    '----Table exploration----
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    
        ticker = Cells(i, 1).Value
        TotVol = TotVol + Cells(i, 7).Value
                
        
        'Conditional to check for first Open position of each ticker
        If k = 1 Then
        
            OpenP = Cells(i, 3).Value
            
            k = k + 1
            
        End If
        
        'Finding the last same value
        If ticker <> (Cells(i + 1, 1).Value) Then
        
            'Populating Summary table
            If OpenP = 0 Then
                OpenP = Cells(i, 6).Value
            End If
            
            Cells(j, TiSumTable).Value = Cells(i, 1).Value
            Cells(j, DiffTable).Value = Cells(i, 6).Value - OpenP
            Cells(j, ChRateTable).Value = FormatPercent((Cells(i, 6).Value / OpenP) - 1)
            Cells(j, TotSumTable).Value = TotVol
            'Not requested at the homework but included to compare results
            Cells(j, OpSumTable).Value = OpenP
            Cells(j, ClSumTable).Value = Cells(i, 6).Value
            
            'Formatting "Yearly Change" column
            If Cells(i, 6).Value - OpenP < 0 Then
                Cells(j, DiffTable).Interior.ColorIndex = 3
            Else
                Cells(j, DiffTable).Interior.ColorIndex = 4
            End If
                   
            j = j + 1
            k = 1
            TotVol = 0
                              
        End If
        
    Next i
    
End Sub