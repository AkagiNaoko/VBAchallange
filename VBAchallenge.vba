

Sub tickerscript()
 'Declear the variables
   Dim ws As Worksheet
   Dim row, column As Long
   Dim lastRow As Long
   Dim YearChange, PercentChange, TotalStock As Double
   Dim beginPrice, EndPrice As Double
   Dim NewRow, LastNewRow As Long
   Dim GreatIncreat,GreatDecrease as Double
   Dim GreatTotal as Double
   

 ' loop in each worksheet
   For Each ws In ThisWorkbook.Worksheets

   ' find the lastrow, zero value
    lastRow = ws.Cells(Rows.Count, 1).end(xlUp).Row
    NewRow = 2
    TotalStock = 0
    PercentChange=0
    beginPrice=0
    EndPrice=0
    YearChange=0
    GreatIncrease =0
    GreatDecrease =0
    Greattotal = 0
    LastNewRow =0

       'Loop to find the Year change, percent change and total stock for each ticker symbol
       For row = 2 To lastRow
          If ws.Cells(row, 1).Value <> ws.Cells(row - 1, 1).Value Then
             ws.Range("I" & NewRow).Value = ws.Cells(row, 1).Value
             NewRow = NewRow + 1
             beginPrice = ws.Cells(row,3).Value
             TotalStock = ws.Cells(row,7).Value
          Else
             TotalStock =TotalStock +ws.Cells(row,7)
             EndPrice = ws.Cells(row,6).Value
             PercentChange = (EndPrice-beginPrice)/beginPrice
             YearChange = EndPrice-beginPrice
             
             ws.Range("J" & NewRow-1).Value = YearChange
             ws.Range("K" & NewRow-1).Value = PercentChange
             ws.Range("L" & NewRow-1).Value = TotalStock

          End If
       Next row

     ' Build a new table and name the table cells
        ws.Range("I" & 1).Value = "ticker symbol"
        ws.Range("J" & 1).Value = "Year Change"
        ws.Range("K" & 1).Value = "Percent Change"
        ws.Range("L" & 1).Value = "Total Stock"

        ws.Range("N"& 2). Value = "Greatest % Increase"
        ws.Range("N"& 3). Value = "Greatest % Decrease"
        ws.Range("N"& 4). Value = "Greatest total Volume"
        ws.Range("O"& 1). Value = "Ticker"
        ws.Range("P"& 1). Value = "Value"

    'loop the new table, find last row of new table
      LastNewRow = ws.Cells(Rows.Count, 9).end(xlUp).Row

    ' loop to change the Conditional Formatting
      ws.Range("K1:K"&LastNewRow).NumberFormat = "0.00%"

      For row = 2 to LastNewRow
       if ws.Range("J"& row).Value < 0 Then
          ws.Range("J"& row).Interior.ColorIndex= 3
       Else
          ws.Range("J"& row).Interior.ColorIndex= 4
       End if 

    ' loop to find the greatest increase, decrease and great total
       if ws.Range("K"& row).Value < GreatDecrease Then  
           GreatDecrease = ws.Range("K"& row).Value
           ws.Range("O" & 3).Value = ws.Range("I" & Row).Value
           ws.Range("P" & 3).Value = ws.Range("K" & Row).Value
           ws.Range("P" & 3).NumberFormat ="0.00%"
       End if 

       if ws.Range("K"& row).Value > GreatIncrease Then
           GreatIncrease = ws.Range("K"& row).Value
           ws.Range("O" & 2).Value = ws.Range("I" & Row).Value
           ws.Range("P" & 2).Value = ws.Range("K" & Row).Value
           ws.Range("P" & 2).NumberFormat ="0.00%"
        
       End if 

       if ws.Range("L"& row).Value > GreatTotal Then
           GreatTotal = ws.Range("L"& row).Value
           ws.Range("O" & 4).Value = ws.Range("I" & Row).Value
           ws.Range("P" & 4).Value = ws.Range("L" & Row).Value
       End if
      Next Row

  Next ws   
   

End Sub




