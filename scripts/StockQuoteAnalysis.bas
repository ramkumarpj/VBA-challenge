Attribute VB_Name = "Module1"
' Create a script that loops through all the stocks for one year and outputs the following information:

' - The ticker symbol
' - Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
' - The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
' - The total stock volume of the stock.

' Assumptions

' - Stock data is sorted by ticker symbol and date
' - Each worksheet contains data for one year

Sub StockSummaryStats():

' Iterate thru each worksheet in the file
For Each ws In Worksheets
        
        ' Loop through the lines for a ticker symbol and collect following info
        ' - firstDayOpeningPrice (only once when a new ticker is found)
        ' - runningTotalVolume
        ' - lastDayClosingPrice
        ' Summarize and Print information for the current ticker when a new ticker symbol is found
        ' Continue to collect above information for the new ticker symbol
        
        ' Declare variables
        Dim ticker As String
        Dim firstDayOpeningPrice As Double
        Dim lastDayClosingPrice As Double
        
        Dim yearlyChange As Double
        Dim percentChange As Double
        Dim runningTotalVolume As Double
        
        Dim lastRow As Double
        Dim tickerRowIndex As Integer
        
        Dim greatestPercentIncreaseRowIndex As Integer
        Dim greatestPercentDecreaseRowIndex As Integer
        Dim greatestTotalVolumeRowIndex As Integer
        
        ' Initialize variables
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        tickerRowIndex = 2
        
        ' Print Headers for summary area
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Print Headers for summary aggregation area
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Iterate thru each line in current worksheet
        For i = 2 To lastRow + 1
        
            ' New ticker symbol found
            If (ws.Cells(i, 1).Value <> ticker) Then
            
                If (ticker <> "") Then
                
                    ' Summarize and Print information about the previous ticker symbol
                    
                    ' Calculate yearly and percent change
                    yearlyChange = lastDayClosingPrice - firstDayOpeningPrice
                    percentChange = yearlyChange / firstDayOpeningPrice
                    
                    ' Print information to summary area
                    ws.Range("I" & tickerRowIndex).Value = ticker
                    ws.Range("J" & tickerRowIndex).Value = yearlyChange
                    ws.Range("K" & tickerRowIndex).Value = percentChange
                    ws.Range("L" & tickerRowIndex).Value = runningTotalVolume
                    
                    ' Format yearly change cell to display as currency and do coniditional color formating
                    ws.Range("J" & tickerRowIndex).Style = "Currency"
                    If yearlyChange < 0 Then
                        ws.Range("J" & tickerRowIndex).Interior.Color = RGB(255, 0, 0)
                    Else
                         ws.Range("J" & tickerRowIndex).Interior.Color = RGB(0, 255, 0)
                    End If
                    
                    ' Format Percent change cell to display as percentage and do coniditional color formating
                    ws.Range("K" & tickerRowIndex).NumberFormat = "0.00%"
                    If percentChange < 0 Then
                        ws.Range("K" & tickerRowIndex).Interior.Color = RGB(255, 0, 0)
                    Else
                         ws.Range("K" & tickerRowIndex).Interior.Color = RGB(0, 255, 0)
                    End If
                    
                    tickerRowIndex = tickerRowIndex + 1
                    
                End If
                
                'Reset the values for new ticker
                ticker = ws.Cells(i, 1).Value
                firstDayOpeningPrice = ws.Cells(i, 3).Value
                lastDayClosingPrice = ws.Cells(i, 6).Value
                runningTotalVolume = ws.Cells(i, 7).Value
                
                Else
                    
                    ' Continue to collect information about current ticker symbol
                     lastDayClosingPrice = ws.Cells(i, 6).Value
                    runningTotalVolume = runningTotalVolume + ws.Cells(i, 7).Value
                
            End If
        Next i
        
        'Print summary aggregation
        
        'Set the cell ranges
        Set percentChangeCellRange = ws.Range("K2:K" & lastRow)
        Set volumeCellRange = ws.Range("L2:L" & lastRow)
        
        'Lookup and Print Greatest Percent Increase
        greatestPercentIncreaseRowIndex = Application.Match(Application.Max(percentChangeCellRange), percentChangeCellRange, 0)
        
        ws.Range("P2").Value = ws.Range("I" & greatestPercentIncreaseRowIndex + 1).Value
        ws.Range("K" & greatestPercentIncreaseRowIndex + 1).Copy ws.Range("Q2")
        
        'Lookup and Print Greatest Percent Decrease
        greatestPercentDecreaseRowIndex = Application.Match(Application.Min(percentChangeCellRange), percentChangeCellRange, 0)
        
        ws.Range("P3").Value = ws.Range("I" & greatestPercentDecreaseRowIndex + 1).Value
        ws.Range("K" & greatestPercentDecreaseRowIndex + 1).Copy ws.Range("Q3")
        
        'Lookup and Print Greatest Total Volume
        greatestTotalVolumeRowIndex = Application.Match(Application.Max(volumeCellRange), volumeCellRange, 0)
        
        ws.Range("P4").Value = ws.Range("I" & greatestTotalVolumeRowIndex + 1).Value
        ws.Range("L" & greatestTotalVolumeRowIndex + 1).Copy ws.Range("Q4")
Next ws

End Sub

