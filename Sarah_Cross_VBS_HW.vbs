Attribute VB_Name = "Module1"
Sub Stock_Volume()

Dim ws As Worksheet
  
For Each ws In Worksheets
  
Dim Stock_Name As String

Dim Stock_Volume, Open_Price, Close_Price, Yearly_Change, Percent_Change As Double

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

' Label new columns
ws.Cells(1, 9).Value = "Ticker"

' yearly change = opening price - closing price
ws.Cells(1, 10).Value = "Yearly Change"

' % change from opening price to closing price
ws.Cells(1, 11).Value = "Percent Change"

ws.Cells(1, 12).Value = "Total Sales Volume"
  
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
Open_Price = ws.Cells(2, 3).Value

' Delete rows with 0 values, starting at the bottom
    For k = Lastrow To 2 Step -1

        If ws.Cells(k, 3).Value = 0 And ws.Cells(k, 4).Value = 0 And ws.Cells(k, 5).Value = 0 And ws.Cells(k, 6).Value = 0 And ws.Cells(k, 7).Value = 0 Then

            ws.Rows(k).Delete

        End If

    Next k
  
  For i = 2 To Lastrow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      
      Stock_Name = ws.Cells(i, 1).Value

      Close_Price = ws.Cells(i, 6).Value

      Yearly_Change = (Open_Price - Close_Price)

      Percent_Change = (Yearly_Change / Open_Price)

      Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value

      ws.Range("I" & Summary_Table_Row).Value = Stock_Name

      ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
      
        If (Yearly_Change) < 0 Then

            ' Color Index 3 is Red
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        
        ElseIf (Yearly_Change) > 0 Then
        
            ' Color Index 4 is Green
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        
        Else
        
            ' Color Index 0 is no fill
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 0
        
        End If

      ws.Range("K" & Summary_Table_Row).Value = Percent_Change
      
      ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

      ws.Range("L" & Summary_Table_Row).Value = Stock_Volume

      Summary_Table_Row = Summary_Table_Row + 1
      
      Stock_Volume = 0
      
      Yearly_Change = 0

      Open_Price = ws.Cells(i + 1, 3).Value

      Close_Price = 0

    Else

      Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value

    End If

  Next i

' Begin hard section
ws.Cells(2, 15).Value = "Greatest % Increase"

ws.Cells(3, 15).Value = "Greatest % Decrease"

ws.Cells(4, 15).Value = "Greatest Total Volume"

ws.Cells(1, 16).Value = "Ticker"

ws.Cells(1, 17).Value = "Value"

Max_Inc = WorksheetFunction.Max(ws.Range("K:K"))

Max_Dec = WorksheetFunction.Min(ws.Range("K:K"))

Max_Vol = WorksheetFunction.Max(ws.Range("L:L"))

ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(2, 17).Value = Max_Inc

ws.Cells(3, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).Value = Max_Dec

ws.Cells(4, 17).Value = Max_Vol

Lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row

  For j = 2 To Lastrow2
  
  If ws.Cells(j, 11).Value = Max_Inc Then

    Max_Inc_Tic = ws.Cells(j, 9).Value
    ws.Cells(2, 16).Value = Max_Inc_Tic

  End If

  If ws.Cells(j, 11).Value = Max_Dec Then

    Max_Dec_Tic = ws.Cells(j, 9).Value
    ws.Cells(3, 16).Value = Max_Dec_Tic

  End If

  If ws.Cells(j, 12).Value = Max_Vol Then

    Max_Vol_Tic = ws.Cells(j, 9).Value
    ws.Cells(4, 16).Value = Max_Vol_Tic

  End If

  Next j

ws.Cells.EntireColumn.AutoFit

Next ws

End Sub
