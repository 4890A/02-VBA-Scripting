Attribute VB_Name = "Module11"
Sub Main():

        Dim LastRow As Long
        Dim Ticker_Name As String
        Dim Ticker_Total As Double
        Dim Year_Open As Double
        Dim Year_Close As Double
        Dim Year_Change As Double
        Dim Summary_Row As Long
        Summary_Row = 2
        Dim ws As Worksheet
        Set ws = ActiveSheet

        ' Name Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Chnange"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Find Last row for for loop
        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        ' Initialize Opening Price for first stock
        Year_Open = ws.Range("C2").Value
        
        
        ' Loop over rows, checking if the next ticker is new
        For i = 2 To LastRow
                ' Checks if the ith row is the last row for the stock
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                            
                                Ticker_Name = ws.Cells(i, 1).Value
                                ' Add the ticker to the summary
                                ws.Range("I" & Summary_Row).Value = Ticker_Name
                                
                                ' Update total and append to summary table
                                Ticker_Total = Ticker_Total + ws.Range("G" & i).Value
                                ws.Range("L" & Summary_Row).Value = Ticker_Total
                                
                                ' Set closing price and calculate change
                                Year_Close = Cells(i, 6).Value
                                Year_Change = Year_Close - Year_Open
                                Percent_Change = (Year_Change / Year_Open)
                                
                                ' Store changes in summary table, apply formatting
                                ws.Range("J" & Summary_Row).Value = Year_Change
                                If Year_Change < 0 Then
                                    ws.Range("J" & Summary_Row).Interior.ColorIndex = 3
                                Else
                                    ws.Range("J" & Summary_Row).Interior.ColorIndex = 4
                                End If
            
                                ws.Range("K" & Summary_Row).Value = Percent_Change
                                ws.Range("K" & Summary_Row).NumberFormat = "0.00%"
                                
                                
                                ' Reset Total, preparing for the next stock
                                Ticker_Total = 0
                                ' Increment the row for the next entry in the summary
                                Summary_Row = Summary_Row + 1
                                
                                Year_Open = Cells(i + 1, 3).Value
                                
                Else
                                ' Same stock, update the total
                                Ticker_Total = Ticker_Total + ws.Range("G" & i).Value
                End If
        Next i



    
End Sub
