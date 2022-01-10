Attribute VB_Name = "Module1"

Sub VBA_StockAnalysis()

Dim ws As Worksheet

    For Each ws In Worksheets
    
        Dim ticker As String
        Dim total_volume As Double
        Dim tickercount  As Long
        Dim open_value As Double
        Dim close_value As Double
        Dim yearly_change  As Double
        Dim percent_change As Double
        Dim last_row As Long
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        tickercount = 2
        total_volume = 0
        open_value = 0
        close_value = 0
        yearly_change = 0
        percent_change = 0
        

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        For i = 2 To last_row
        
            total_volume = total_volume + ws.Cells(i, 7)

            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                open_value = ws.Cells(i, 3).Value
            End If
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                ws.Cells(tickercount, 12).Value = total_volume
                
                ws.Cells(tickercount, 9).Value = ws.Cells(i, 1).Value
                
                
                close_value = ws.Cells(i, 6).Value
                
                yearly_change = close_value - open_value
                ws.Cells(tickercount, 10).Value = yearly_change
                
                        
                If open_value = 0 And close_value <> 0 Then
                    Dim if_percentchange_0 As String
                    if_percentchange_0 = " "
                    ws.Cells(tickercount, 11).Value = percent_change
                    
                ElseIf open_value = 0 Then
                    percent_change = 0
                    ws.Cells(tickercount, 11).Value = percent_change
                    
                Else
                    percent_change = yearly_change / open_value
                    ws.Cells(tickercount, 11).Value = percent_change
                End If
                
                
                tickercount = tickercount + 1
                
                total_volume = 0
                open_value = 0
                close_value = 0
                yearly_change = 0
                percent_change = 0
        End If
    Next i
    
Next ws
    
End Sub
