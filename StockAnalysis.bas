Attribute VB_Name = "Module1"
Sub StockAnalysis()

' Loop through all worksheets:
    For Each ws In Worksheets
    
' Create Variables for ticker name, opening price, closing price, counter loop, last row, stock volume and many others ...
        Dim ticker As String
        Dim opening_price, closing_price, yearly_change As Double
        Dim rowCount As Long
        Dim lastrow1, lastrow2 As Long
        Dim percentage_change, greatest_percent_increase, greatest_percent_decrease, stock_volume, max_volume As Double
            stock_volume = 0

    
' Keep track of the location for each ticker name in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

    'Define column names for the output
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
    
         ' Autofit to display column names
          ws.Columns("O").AutoFit
          ws.Columns("I:L").AutoFit
    
            ' Count the number of rows
                rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
                
            ' Get worksheet name
                WorksheetName = ws.Name
    
            ' Get the first opening price
                opening_price = ws.Range("C2").Value
    
            ' Loop through the worksheet from row 2 to last row
                For i = 2 To rowCount
        
            ' Check if we are still within the same ticker name, if it is not...
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ' Set the ticker name
                ticker = ws.Cells(i, 1).Value
           
            ' Get closing price
                closing_price = ws.Cells(i, 6).Value
                        
            ' Calculate Yearly Change
                yearly_change = closing_price - opening_price
            
            ' Calculate Percentage Change
                percentage_change = yearly_change / opening_price

            ' Add to the total volume
                stock_volume = stock_volume + ws.Cells(i, 7).Value

            ' Print the ticker name in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = ticker

            ' Print the yearly change information in the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = yearly_change
            
            ' Print the percentage change information in the Summary Table
                ws.Range("K" & Summary_Table_Row).Value = FormatPercent(percentage_change)
            
            ' Print the total stock volume to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = stock_volume

            ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
      
            ' Reset the stock_volume value to 0
                stock_volume = 0

            ' Get opening price for the next ticker
                opening_price = ws.Range("C" & (i + 1)).Value
         
            ' If the cell immediately following a row is the same name...
                 Else

                ' Add to the stock volume
                    stock_volume = stock_volume + ws.Cells(i, 7).Value

            End If

    Next i

              
        ' Get greatest total stock volume
            lastrow1 = ws.Cells(Rows.Count, 12).End(xlUp).Row
            Set DataRange = Worksheets(WorksheetName).Range("L2:L" & lastrow1)
            max_volume = Application.WorksheetFunction.Max(DataRange)
            ws.Range("Q4").Value = max_volume
            ws.Columns("Q").AutoFit
                
                For i = 2 To lastrow1
                    
                    If ws.Cells(i, 12).Value = max_volume Then
                        ws.Range("P4").Value = ws.Cells(i, 9).Value
                    End If
                
                Next i
            
        ' Get greatest % increase value
            lastrow2 = ws.Cells(Rows.Count, 11).End(xlUp).Row
            Set DataRange1 = Worksheets(WorksheetName).Range("K2:K" & lastrow2)
            greatest_percent_increase = Application.WorksheetFunction.Max(DataRange1)
            ws.Range("Q2").Value = FormatPercent(greatest_percent_increase)
            
                For i = 2 To lastrow2
                    
                    If ws.Cells(i, 11).Value = greatest_percent_increase Then
                        ws.Range("P2").Value = ws.Cells(i, 9).Value
                    End If
                
                Next i

        ' Get greatest % decrease value
            Set DataRange2 = Worksheets(WorksheetName).Range("K2:K" & lastrow2)
            greatest_percent_decrease = Application.WorksheetFunction.Min(DataRange2)
            ws.Range("Q3").Value = FormatPercent(greatest_percent_decrease)
            
                For i = 2 To lastrow2
                    
                    If ws.Cells(i, 11).Value = greatest_percent_decrease Then
                        ws.Range("P3").Value = ws.Cells(i, 9).Value
                    End If
                
                Next i
                
        ' Format cells based on negative and positive change
            ' Count the number of rows in Yearly Change column
                lastrow1 = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
                For i = 2 To lastrow1
                    
                    If ws.Cells(i, 10).Value > 0 Then
                        ws.Cells(i, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(i, 10).Interior.ColorIndex = 3
                    End If
                
                Next i
    Next ws

End Sub



