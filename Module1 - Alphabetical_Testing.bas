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
  Dim Current As Worksheet
  
    'Loop through all sheets
    For Each Current In Worksheets
  
    ' Loop through all Stock Lines
        For i = 2 To 70926
    'lastRowStock = i.Cells(Rows.Count, "A").End(xlUp).Row - 1

    ' Check if we are still within the same Ticker Symbol, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker name
            Ticker_Symbol = Cells(i, 1).Value

      ' Add to the Stock Volume
            Stock_Volume = Stock_Volume + Cells(i, 7).Value
      
      'Add to the Open Total
            Open_Total = Open_Total + Cells(i, 3).Value

      ' Print the Ticker in the Summary Table
            Range("J" & Summary_Table_Row).Value = Ticker_Symbol
      
      'Print the Change in opening price
            Range("K" & Summary_Table_Row).Value = Cells(i, 3).Value - First_Open
        'If Range("K" & Summary_Table_Row).Value >= 0 Then
        'Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
        'Else
        'Range("K" & Summary_Table_Row).Interior.ColorIndex = 2
      ' Print the Stock Volume to the Summary Table
            Range("M" & Summary_Table_Row).Value = Stock_Volume
      
      'Print the Open Percentage to the Summary Table
            Range("L" & Summary_Table_Row).Value = ((Cells(i, 3).Value - First_Open) / Cells(i, 3).Value) * 100

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
            First_Open = Cells(i, 3).Value
            First_Line = First_Line + 1
            Stock_Volume = Stock_Volume + Cells(i, 7).Value
    ' If the cell immediately following a row is the same stock but not the first row...
        Else
      ' Add to the Stock Volume
            Stock_Volume = Stock_Volume + Cells(i, 7).Value
      ' Add to the Line Count
            First_Line = First_Line + 1

        End If

        Next i
    Next Current

End Sub


