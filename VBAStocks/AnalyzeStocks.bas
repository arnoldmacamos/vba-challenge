Attribute VB_Name = "Module1"
Sub AnalyzeStock()

    'Declare variables
    Dim lastRow As Long 'Last Row
    Dim lastCol As Long 'Last Column
    Dim currentWS As Worksheet 'Current worksheet
    Dim ticker As String 'Current stock ticker
    Dim tickerOpeningValue As Double 'opening price at the begining of the year
    Dim tickerClosingValue As Double 'closing price at the end of the year
    Dim tickerTotalStockVolume As Long 'Total volume of stock
    Dim summaryRow As Integer 'Stock summary row
    Dim greatestIncRow As Integer 'Row with the greatest increase
    Dim greatestDecRow As Integer 'Row with the greatest decrease
    Dim greatestTotVolRow As Integer 'Row with the greatest total volume
        
    Dim isNewTicker As Boolean  'Flag to set set if current ticker has changed
        
    
    'Iterate through each worksheet of the workbook
    For Each ws In ActiveWorkbook.Worksheets
    
        Set currentWS = ws
    
        'Set last row and last column
        lastRow = currentWS.Cells(currentWS.Rows.Count, "A").End(xlUp).Row
        lastCol = currentWS.Cells(lastRow, currentWS.Columns.Count).End(xlToLeft).Column
        
        
        'Set Stock 1st Summary Header
        currentWS.Cells(1, "I").Value = "Ticker"
        currentWS.Cells(1, "J").Value = "Yearly Change"
        currentWS.Cells(1, "K").Value = "Percent Change"
        currentWS.Cells(1, "L").Value = "Total Stock Volume"
        
        'Set Stock 2nd Summary Header
        currentWS.Cells(2, "O").Value = "Greatest % Increase"
        currentWS.Cells(3, "O").Value = "Greatest % Decrease"
        currentWS.Cells(4, "O").Value = "Greatest Total Volume"
        
        currentWS.Cells(1, "P").Value = "Ticker"
        currentWS.Cells(1, "Q").Value = "Value"
            
        isNewTicker = True
        tickerTotalStockVolume = 0
        summaryRow = 2
    
        'Iterate through each row in the worksheet
        For I = 2 To lastRow
            
            If (isNewTicker) Then 'get new ticker
                ticker = currentWS.Cells(I, "A").Value
                tickerOpeningValue = currentWS.Cells(I, "C").Value
                tickerClosingValue = currentWS.Cells(I, "F").Value
                tickerTotalStock = tickerTotalStock + currentWS.Cells(I, "G").Value
                isNewTicker = False
            Else
                'if current ticker, set ticker closing value and total stock
                tickerClosingValue = currentWS.Cells(I, "F").Value
                tickerTotalStock = tickerTotalStock + currentWS.Cells(I, "G").Value
            End If
        
            'if next row is a new ticker
            If (Not (ticker = currentWS.Cells(I + 1, "A").Value)) Then
                
                'Update 1st Summary Table for the current ticker
                currentWS.Cells(summaryRow, "I").Value = ticker
                currentWS.Cells(summaryRow, "J").Value = tickerClosingValue - tickerOpeningValue
                
                If (tickerOpeningValue = 0) Then
                    currentWS.Cells(summaryRow, "K").Value = 0
                Else
                    currentWS.Cells(summaryRow, "K").Value = (tickerClosingValue - tickerOpeningValue) / tickerOpeningValue
                End If
                               
                currentWS.Cells(summaryRow, "L").Value = tickerTotalStock
                
                'Format Results
                If (currentWS.Cells(summaryRow, "J").Value >= 0) Then
                    currentWS.Cells(summaryRow, "J").Interior.ColorIndex = 4
                Else
                    currentWS.Cells(summaryRow, "J").Interior.ColorIndex = 3
                End If
                
                currentWS.Cells(summaryRow, "K").NumberFormat = "0.00%"
                
                'Initialize variables for new ticker
                isNewTicker = True
                tickerTotalStock = 0
                summaryRow = summaryRow + 1
            End If
        Next I
        
        'Set 2nd Summary table
        greatestIncRow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(currentWS.Range("K:K")), currentWS.Range("K:K"), 0)
        currentWS.Cells(2, "P").Value = currentWS.Cells(greatestIncRow, "I").Value
        currentWS.Cells(2, "Q").Value = currentWS.Cells(greatestIncRow, "K").Value
        currentWS.Cells(2, "Q").NumberFormat = "0.00%"
        
        greatestDecRow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Min(currentWS.Range("K:K")), currentWS.Range("K:K"), 0)
        currentWS.Cells(3, "P").Value = currentWS.Cells(greatestDecRow, "I").Value
        currentWS.Cells(3, "Q").Value = currentWS.Cells(greatestDecRow, "K").Value
        currentWS.Cells(3, "Q").NumberFormat = "0.00%"
            
        greatestTotVolRow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(currentWS.Range("L:L")), currentWS.Range("L:L"), 0)
        currentWS.Cells(4, "P").Value = currentWS.Cells(greatestTotVolRow, "I").Value
        currentWS.Cells(4, "Q").Value = currentWS.Cells(greatestTotVolRow, "L").Value
    
    Next
       
        
End Sub
