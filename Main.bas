Attribute VB_Name = "Module11"
Sub Main():

        Dim LastRow As Long
        Dim Ticker_Name As String
        Dim Ticker_Total As Double
        Dim Year_Open As Double
        Dim Year_Close As Double
        Dim Year_Change As Double
        Dim Greatest_Inc As Double
        Dim Greatest_Dec As Double
        Dim Greatest_Vol As Double
        Dim Summary_Row As Long

        Dim ws As Worksheet
        ' Set ws = ActiveSheet
        
        For Each ws In ThisWorkbook.Worksheets
        
            ' Assign Dynamic Vars
            Greatest_Inc = -123456
            Greatest_Dec = 123456
            Greatest_Vol = 0
            Ticker_Total = 0
            Year_Close = 0
            Year_Change = 0
            
            ' (Re)set the row for the summary
            Summary_Row = 2
    
            ' Name Headers
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Chnange"
            ws.Range("L1").Value = "Total Stock Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            
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
                                    
                                    ' Avoid div by 0 which happens when stocks open at 0
                                    If Year_Open <> 0 Then
                                        Percent_change = (Year_Change / Year_Open)
                                    Else
                                        Percent_change = Null
                                    End If
                                        
                                    
                                    ' Store changes in summary table, apply formatting
                                    ws.Range("J" & Summary_Row).Value = Year_Change
                                    If Year_Change < 0 Then
                                        ws.Range("J" & Summary_Row).Interior.ColorIndex = 3
                                    Else
                                        ws.Range("J" & Summary_Row).Interior.ColorIndex = 4
                                    End If
                
                                    ws.Range("K" & Summary_Row).Value = Percent_change
                                    ws.Range("K" & Summary_Row).NumberFormat = "0.00%"
                                    
                                    If Percent_change > Greatest_Inc Then
                                        Greatest_Inc = Percent_change
                                        Inc_Ticker = Ticker_Name
                                    End If
                                    
                                    If Percent_change < Greatest_Dec Then
                                        Greatest_Dec = Percent_change
                                        Dec_Ticker = Ticker_Name
                                    End If
                                    
                                    If Ticker_Total > Greatest_Vol Then
                                        Greatest_Vol = Ticker_Total
                                        Vol_Ticker = Ticker_Name
                                    End If
                                        
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
            
                    
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Volume"
            
            ws.Range("P2").Value = Inc_Ticker
            ws.Range("P3").Value = Dec_Ticker
            ws.Range("P4").Value = Vol_Ticker
    
            ws.Range("P2").Value = Inc_Ticker
            ws.Range("P3").Value = Dec_Ticker
            ws.Range("P4").Value = Vol_Ticker
            
            ws.Range("Q2").Value = Greatest_Inc
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").Value = Greatest_Dec
            ws.Range("Q3").NumberFormat = "0.00%"
            ws.Range("Q4").Value = Greatest_Vol
            
        Next
    
    
        
End Sub
