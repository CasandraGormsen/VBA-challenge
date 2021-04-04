Attribute VB_Name = "Module1"
Sub stock_market()

  ' Set an initial variable for holding the stock name
  Dim Stock_Name As String

  ' Set an initial variable for holding the total stock volume
  Dim Stock_Volume As Double
  Stock_Volume = 0
  
  'Set an initial variable for the yearly change
  Dim Yearly_Change As Double
  
  'Set an initial variable for first line tracker
  Dim First_Line As Integer
  First_Line = 1
  
  'Set an initial variable for first open
  Dim First_Open As Double
  
  ' Keep track of the location for each Stock Name in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  'Set current worksheet object variable
  Dim WS_Count As Integer
  Dim V As Integer
  
  WS_Count = ActiveWorkbook.Worksheets.Count
    'Loop through all sheets
    For V = 1 To WS_Count
  
    ' Loop through all Stock Lines
        lastRowStock = Cells(Rows.Count, "A").End(xlUp).Row - 1
        For I = 2 To lastRowStock
   

    ' Check if we are still within the same Ticker Symbol, if it is not...
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

      ' Set the Ticker name
            Ticker_Symbol = Cells(I, 1).Value

      ' Add to the Stock Volume
            Stock_Volume = Stock_Volume + Cells(I, 7).Value
      
      'Add to the Open Total
            Open_Total = Open_Total + Cells(I, 3).Value

      ' Print the Ticker in the Summary Table
            Range("J" & Summary_Table_Row).Value = Ticker_Symbol
      
      'Print the Change in opening price
            Range("K" & Summary_Table_Row).Value = Cells(I, 3).Value - First_Open
        If Range("K" & Summary_Table_Row).Value >= 0 Then
        Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
        Else
        Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
      ' Print the Stock Volume to the Summary Table
            Range("M" & Summary_Table_Row).Value = Stock_Volume
      
      'Print the Open Percentage to the Summary Table
            If First_Open > 0 And Cells(I, 3).Value > 0 Then
                Range("L" & Summary_Table_Row).Value = ((Cells(I, 3).Value - First_Open) / Cells(I, 3).Value) * 100
                Else: Range("L" & Summary_Table_Row).Value = 0
            End If

      ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Stock_Volume
            Stock_Volume = 0
      
      'Reset the First Open
            First_Open = 1
      
      'Reset First Line
            First_Line = 1

    ' If the cell immediately following a row is the same brand and the first row...
        ElseIf First_Line = 1 Then
            If Cells(I, 3).Value > 0 Then
                First_Open = Cells(I, 3).Value
                First_Line = First_Line + 1
                Stock_Volume = Stock_Volume + Cells(I, 7).Value
            Else
                Stock_Volume = Stock_Volume + Cells(I, 7).Value
            End If
    ' If the cell immediately following a row is the same stock but not the first row...
        Else
      ' Add to the Stock Volume
            Stock_Volume = Stock_Volume + Cells(I, 7).Value

        End If

        Next I
        MsgBox ActiveWorkbook.Worksheets(V).Name
    Next V

End Sub


