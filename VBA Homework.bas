
Sub StockAnalysis()
'Variables for the output columns
Dim ticker As String
Dim yearchange As Double
Dim pctchange As Double
Dim totalvolume As Double

'Start / end price for stock
Dim startprice As Double
Dim endprice As Double

'Set starting variables
totalvolume = 0

' Tracking of location in output - Default to second row
 Dim Summary_Table_Row As Double
 Summary_Table_Row = 2
 
 'Set Header rows for output
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"



'Iterate through worksheets
For Each ws In Worksheets

  'Find the last non-blank cell in ticker column A(1)
    lRow = ws.Cells(Rows.Count, 1).End(xlUp).row


 

'Start Price is open price on first day of year. Assuming they are in date order
 startprice = ws.Cells(2, 3).Value
 
 
'Loop through current ticker symbol
  For i = 2 To lRow

    ' Check if we are still within the same stock ticker.
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the ticker value
      ticker = ws.Cells(i, 1).Value

      ' Add to the Total Volume
      totalvolume = totalvolume + ws.Cells(i, 7).Value
      'Starting point for Close Price - Assuming the dates are sorted
       endprice = ws.Cells(i, 6).Value
       yearchange = (endprice - startprice)
       
      ' Print output rows
      Range("I" & Summary_Table_Row).Value = ticker
      Range("J" & Summary_Table_Row).Value = yearchange
      'Conditional formatting of rows
            If Range("J" & Summary_Table_Row).Value > 0 Then
             Range("J" & Summary_Table_Row).Interior.Color = vbGreen
            Else
            Range("J" & Summary_Table_Row).Interior.Color = vbRed
            End If
      
    'Error condition for divide by issues
    If startprice = 0 Then
        pctchange = 0
    Else
        pctchange = yearchange / startprice
    End If
    
    'Format percentage and totalvolume
      Range("K" & Summary_Table_Row).Value = pctchange
      Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      Range("L" & Summary_Table_Row).Value = totalvolume
      Range("L" & Summary_Table_Row).NumberFormat = "#,##"

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset Total Valume
      totalvolume = 0
        'Start Price for next iteration
    startprice = ws.Cells(i + 1, 3)
    ' If the cell immediately following a row is the same stock
    Else
      ' Update totalvolume
      totalvolume = totalvolume + ws.Cells(i, 7).Value
      End If
    Next i
      
Next ws

End Sub



