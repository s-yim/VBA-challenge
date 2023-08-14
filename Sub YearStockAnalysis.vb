Sub YearStockAnalysis()

Dim ws As Worksheet
'Looping Across Worskheet
For Each ws In Worksheets

    Dim ticker As String
    Dim price As Double
    Dim volume As Double
    Dim summary_row As Integer
    Dim change As Double
    Dim percent_change As Double
    
    volume = 0
    summary_row = 2
    start_row = 2
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
'Column Creation
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
'Retrieval of Data
        For i = 2 To last_row
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ticker = ws.Cells(i, 1).Value
                ws.Range("I" & summary_row).Value = ticker
                
                change = (ws.Cells(i, 6).Value - ws.Cells(start_row, 3))
                ws.Range("J" & summary_row).Value = change
                
                percent_change = change / ws.Cells(start_row, 3)
                ws.Range("K" & summary_row).Value = percent_change
                
                volume = volume + ws.Cells(i, 7).Value
                ws.Range("L" & summary_row).Value = volume
'Conditional Formatting
                Select Case change
                    Case Is > 0
                        ws.Range("J" & summary_row).Interior.ColorIndex = 4
                    Case Is < 0
                        ws.Range("J" & summary_row).Interior.ColorIndex = 3
                    Case Else
                        ws.Range("J" & summary_row).Interior.ColorIndex = 0
                End Select
                
                ws.Range("J" & summary_row).NumberFormat = "0.00"
                ws.Range("K" & summary_row).NumberFormat = "0.00%"
                               
                summary_row = summary_row + 1
                start_row = i + 1
                volume = 0
                                   
            Else
                volume = volume + ws.Cells(i, 7).Value
                
            End If
            
        Next i
'Calculated Values
    max_increase_row = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & last_row)), ws.Range("K2:K" & last_row), 0)
    max_decrease_row = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & last_row)), ws.Range("K2:K" & last_row), 0)
    max_volume_row = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & last_row)), ws.Range("L2:L" & last_row), 0)
    
    ws.Range("P2") = ws.Cells(max_increase_row + 1, 9)
    ws.Range("P3") = ws.Cells(max_decrease_row + 1, 9)
    ws.Range("P4") = ws.Cells(max_volume_row + 1, 9)
    
    ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & last_row)) * 100
    ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & last_row)) * 100
    ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & last_row)) * 100
            
Next ws

End Sub



