Sub Stockmarket()

'Establish Variables

    Dim ticker As String
    Dim row_setting As Long
    Dim Volume As Double
    Dim YearlyChange As Double
    Dim Yearlyopening As Double
    Dim YearlyClosing As Double
    Dim Lastrow As Long
    Dim ws As Worksheet
    Dim total_row As Integer
    Dim total_percent_change As Double
    

'Determine the starting values
For Each ws In ThisWorkbook.Worksheets
    row_setting = 2
    YearlyChange = 0
    Yearlyopening = 0
    YearlyClosing = 0
    Lastrow = 0
    total_row = 2
    total_percent_change = 0
    Volume = 0
    
    
'Begin the loop for the worksheets

    
    
        ws.Select
        ws.Activate
        Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
         ws.Range("H1").Value = "Ticker"
         ws.Range("I1").Value = "YearlyChange"
         ws.Range("J1").Value = "total_percent_change"
         ws.Range("K1").Value = "Total Volume"
         
         ws.Range("N2").Value = "Greatest % Increase"
         ws.Range("N3").Value = "Greatest % Decrease"
         ws.Range("N4").Value = "Greatest Total Volume"
         
         ws.Range("O1").Value = "Ticker"
         ws.Range("P1").Value = "Value"
        
        For row_setting = 2 To Lastrow
        
        ticker = ws.Cells(row_setting, 1).Value
        
        If row_setting = 2 Then
        
        Yearlyopening = Cells(row_setting, 3).Value
        End If
        
        If ticker = ws.Cells(row_setting + 1, 1).Value Then
        
                                    
        
        Volume = Volume + ws.Cells(row_setting, 7).Value

        Else
            YearlyClosing = ws.Cells(row_setting, 6).Value
            Volume = Volume + ws.Cells(row_setting, 7).Value
            YearlyChange = YearlyClosing - Yearlyopening
            
                                If YearlyChange < 0 Then
                                    ws.Range("I" & total_row).Interior.Color = RGB(255, 0, 0)
                                Else
                                    ws.Range("I" & total_row).Interior.Color = RGB(124, 252, 0)
                                End If
            

            
            If Yearlyopening <> 0 And YearlyChange <> 0 Then
                    total_percent_change = YearlyChange / Yearlyopening
            Else
                    total_percent_change = 0
            End If
            ws.Cells(total_row, 8).Value = ticker
            ws.Cells(total_row, 9).Value = YearlyChange
            ws.Cells(total_row, 10).Value = total_percent_change
            ws.Cells(total_row, 11).Value = Volume
            
            YearlyChange = 0
            total_percent_change = 0
            Volume = 0
            
            ticker = ws.Cells(row_setting + 1, 1).Value
            total_row = total_row + 1
            
            Yearlyopening = ws.Cells(row_setting + 1, 3)
            
            
            
            
            


        End If
        Next row_setting
        row_setting = 2
    Next ws
End Sub


