Sub LoopOverSheets()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Select
        Call WallStreet
    Next
End Sub


Sub WallStreet()
    'Set Variable to hold ticker
    'Dim ticker As String
    
    'Set Variable to hold opening and closing price (column 3 for open and column 6 for close)
    Dim opening As Double
    Dim closing As Double
    
    'Set Variable to hold yearly change and percent change
    Dim yearly_change As Double
    Dim percentage_change As Double
    
    'Set Variable to hold total stock volume
    Dim total_volume As Double
    
    'Set opening price
    opening = Cells(2, 3).Value
    
    'Create Columns
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percentage Change"
    Range("L1").Value = "Total Stock Volume"
    
    'Declare iterators
    Dim i As Long
    Dim k As Integer
    
    k = 0
    
    
    'Loop over all rows
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    'For i = 2 To 1703
        'Check to see if ticker is different
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    
        
            'Set Ticker
            Range("I" & 2 + k).Value = Cells(i, 1).Value
            
            'Calculate and store total volume for new ticker
            total_volume = total_volume + Cells(i, 7).Value
            Range("L" & 2 + k).Value = total_volume
            
            'set closing
            closing = Cells(i, 6).Value
            
            'Calculates and stores the difference between open and close
            yearly_change = closing - opening
            Range("J" & 2 + k).Value = yearly_change
            
            'Calculates and stores the percentage change
            If opening <> 0 Then
                percentage_change = yearly_change / opening
                Range("K" & 2 + k).Value = percentage_change
            Else
                percentage_change = 0
                Range("K" & 2 + k).Value = percentage_change
            End If
            
            'Reset stock total volume
            total_volume = 0
            
            
            'Set opening value
            opening = Cells(i + 1, 3).Value
            j = i
            If opening = 0 Then
                j = j + 1
                opening = Cells(j, 3).Value
            Else
                j = 0
           End If
            
            
            
           k = k + 1
            
        Else
            total_volume = total_volume + Cells(i, 7).Value
            
        End If
    Next i
    
    'Formatting Percentages
    last_row = Cells(Rows.Count, 11).End(xlUp).Row
    Range("K2:K" & last_row).NumberFormat = "0.00%"
    
    'Conditional Formatting
    For i = 2 To Cells(Rows.Count, 10).End(xlUp).Row
        If Cells(i, 10) < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
        Else
            Cells(i, 10).Interior.ColorIndex = 4
        End If
    Next i
    
    'Bonus
    Range("O1").Value = "Greatest"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
   'Find max percent
   Dim m As Integer
   Dim max As Double
   max = 0
    For m = 2 To Cells(Rows.Count, 11).End(xlUp).Row
        If Cells(m, 11).Value > max Then
            max = Cells(m, 11).Value
            Range("P2") = Cells(m, 9).Value
        End If
    Next m
    Range("Q2").Value = max
    Range("Q2").NumberFormat = "0.00%"
           
   'Find min percent
   Dim n As Integer
   Dim min As Double
   min = 0
    For n = 2 To Cells(Rows.Count, 11).End(xlUp).Row
        If Cells(n, 11).Value < min Then
            min = Cells(n, 11).Value
            Range("P3") = Cells(n, 9).Value
        End If
    Next n
    Range("Q3").Value = min
    Range("Q3").NumberFormat = "0.00%"
    
    'Find max volume
    Dim a As Integer
    Dim mv As Double
    mv = 0
    For a = 2 To Cells(Rows.Count, 12).End(xlUp).Row
        If Cells(a, 12).Value > mv Then
            mv = Cells(a, 12).Value
            Range("P4") = Cells(a, 9).Value
        End If
    Next a
    Range("Q4").Value = mv
            
    
    
End Sub

