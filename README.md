'Button to reset the entire workbook'
Sub Reset_Worksheets_Button():

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Range("I:P").ClearContents
        ws.Range("I:P").Interior.ColorIndex = xlNone
    Next ws
    
End Sub
'Button to complete the entire workbook'
Sub All_Stock_WS_Button():

    Application.ScreenUpdating = True
    For Each ws In Worksheets
        ws.Activate
        ws.Range("A:Q").EntireColumn.AutoFit
        ws.Range("P4, G:G, L:L").NumberFormat = "#,##0"
        Call Stock_WS
    Next ws
    
End Sub
' Process variables between sheets'
Sub Stock_WS():

    Range("I1:L1") = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    Range("O1:P1") = Array("Ticker", "Value")
    Range("N2:N4") = Application.Transpose(Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume"))
    Range("J:J").NumberFormat = "0.00"
    Range("K:K, P2:P3").NumberFormat = "0.00%"
    
    Dim ticker As String
    Dim open_price, year_changing, percent_update, volume_total As Double
    Dim increased_ticker, decreased_ticker, vol_ticker As String
    Dim best_increase, best_decrease, best_vol As Double
    Dim input_row As Long
    Dim output_row As Integer
    
    ticker = Range("A2")
    open_price = Range("C2")
    volume_total = Range("G2")
    input_row = 3
    output_row = 2
    
    While (ticker <> "")
        While (ticker = Cells(input_row, 1))
            volume_total = volume_total + Cells(input_row, 7)
            input_row = input_row + 1
        Wend
        
        year_changing = Cells(input_row - 1, 6) - open_price
        percent_update = year_changing / open_price
        Cells(output_row, 9) = ticker
        Cells(output_row, 10) = year_changing
        If (year_changing < 0) Then
            Cells(output_row, 10).Interior.Color = vbRed
        ElseIf (year_changing > 0) Then
            Cells(output_row, 10).Interior.Color = vbGreen
        End If
        
        Cells(output_row, 11) = percent_update
        
        If (percent_update < best_decrease) Then
            best_decrease = percent_update
            decreased_ticker = ticker
        ElseIf (percent_update > best_increase) Then
            best_increase = percent_update
            increased_ticker = ticker
        End If
        
        Cells(output_row, 12) = volume_total
        
        If (volume_total > best_vol) Then
            best_vol = volume_total
            vol_ticker = ticker
        End If
        ticker = Cells(input_row, 1)
        open_price = Cells(input_row, 3)
        volume_total = 0
        output_row = output_row + 1
    Wend
    
    Range("O2") = increased_ticker
    Range("P2") = best_increase
    Range("O3") = decreased_ticker
    Range("P3") = best_decrease
    Range("O4") = vol_ticker
    Range("P4") = best_vol
    
End Sub
