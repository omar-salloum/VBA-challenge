Sub Module2()

Dim Ticker As String, tickernumber As Integer, Opening As Double, Closing As Double, Volume As LongLong
Dim ws As Worksheet 'to work on all the worksheets

' code beg

For Each ws In Sheets

ws.Range("I1").Value = "Ticker" 'name column I ticker
ws.Range("J1").Value = "Yearly Change" 'name column J Yearly change
ws.Range("K1").Value = "Percent Change" 'name column K Percent change
ws.Range("L1").Value = "Total Stock Volume" 'name column L Total Stock Volume
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row 'Find the last active row


Volume = 0
tickernumber = 2
Opening = ws.Cells(2, 3).Value 'Initial Opening Value

    For i = 2 To Lastrow

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ws.Cells(tickernumber, 9).Value = ws.Cells(i, 1).Value 'Name each unique Ticker
            Closing = ws.Cells(i, 6).Value 'Calculates closing value
            ws.Cells(tickernumber, 10).Value = Closing - Opening 'Calculates price difference
            ws.Cells(tickernumber, 11).Value = (Closing - Opening) / Opening ' Calculates percent change
            ws.Cells(tickernumber, 11).NumberFormat = "0.00%" 'Set percent change values to percents instead of decimals
            Opening = ws.Cells(i + 1, 3).Value 'Calculates new opening
            Volume = Volume + ws.Cells(i, 7).Value
            ws.Cells(tickernumber, 12).Value = Volume
            Volume = 0
            tickernumber = tickernumber + 1 'moves next value one down
            
        Else
            Volume = Volume + ws.Cells(i, 7).Value
        End If
    Next i
    
    Lastchange = ws.Cells(Rows.Count, 10).End(xlUp).Row 'Find the last active row for Change amount and Percentage Change
    
        For r = 2 To Lastchange
        
            If ws.Cells(r, 11).Value > 0 Then 'conditional formatting
                ws.Range("K" & r).Interior.ColorIndex = 4
                ws.Range("J" & r).Interior.ColorIndex = 4
                
            ElseIf ws.Cells(r, 11).Value = 0 Then
                ws.Range("J" & r).Interior.ColorIndex = 2
                ws.Range("K" & r).Interior.ColorIndex = 2
                
            Else
                ws.Range("J" & r).Interior.ColorIndex = 3
                ws.Range("K" & r).Interior.ColorIndex = 3
                
            End If
            
        Next r
        ' Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        
        ws.Range("Q2").Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & Lastchange)) ' Greatest % inc
        ws.Range("Q2").NumberFormat = "0.00%"
    
        
        ws.Range("Q3").Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & Lastchange)) 'Greatest % dec
        ws.Range("Q3").NumberFormat = "0.00%"
        
        ws.Range("Q4").Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & Lastchange)) 'Greatest total volume
        
  
        For r = 2 To Lastchange 'Fill ticker values for Greatest % inc
            If ws.Range("K" & r) = Application.WorksheetFunction.Max(ws.Range("K2:K" & Lastchange)) Then
                ws.Range("P2").Value = ws.Range("I" & r).Value
                
            End If
            
            If ws.Range("K" & r) = Application.WorksheetFunction.Min(ws.Range("K2:K" & Lastchange)) Then
                ws.Range("P3").Value = ws.Range("I" & r).Value
                
            End If
            
            If ws.Range("L" & r) = Application.WorksheetFunction.Max(ws.Range("L2:L" & Lastchange)) Then
                ws.Range("P4").Value = ws.Range("I" & r).Value
                
            End If
            
        Next r
Next ws
    
End Sub
