Attribute VB_Name = "Module2"
Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call RunCode
    Next
    Application.ScreenUpdating = True
End Sub

Sub RunCode()

  ' Set a variable for the stock ticker
  Dim stock_ticker As String

  ' Keep track of each stock ticker
  Dim Ticker_Row As Double
  Ticker_Row = 2

  ' Keep track of the total stock in the summary
  Dim summary_table_row As Integer
  summary_table_row = 2
  
  ' Keep track of the yearly change for each stock
  ' Dim yearly_change As Integer
  ' open_price = Cells(i, 3).Value
  ' close_price = Cells(i, 6).Value
  
  ' Keep track of the percent change for each stock
  Dim per_change As Integer
  
    ' Loop through all stock tickers
  For i = 2 To 797711

    ' Check if still within the same stock, if not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        ' Show each stock ticker
        stock_ticker = Cells(i, 1).Value
        
        ' ' Show change from opening price to closing price for given year
        ' yearly_change = (open_price - close_price)
        
        ' ' Show percent change from opening price to closing price for a given year
        ' per_change = ((open_price - close_price) / close_price) * 100

        ' Add the stock total volume for each stock
        stock_total = stock_total + Cells(i, 7).Value

        ' List the stock ticker in Column I
        Range("I" & summary_table_row).Value = stock_ticker

        ' ' List the yearly change in Column J
        ' Range("J" & summary_table_row).Value = yearly_change
        
        ' ' Show percent change in Column K
        ' Range("K" & summary_table_row).Value = per_change

        ' Show total volume per stock in Column L
        Range("L" & summary_table_row).Value = stock_total

        ' Add one to the ticker row
        summary_table_row = summary_table_row + 1

        ' Reset the total volume per stock
        stock_total = 0
        ' yearly_change = 0
        ' per_change = 0

        ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      stock_total = stock_total + Cells(i, 7).Value

        ' If per_change > 0 Then
        '     Cells(i, 10).Interior.ColorIndex = 33
        
        ' Else
        '     Cells(i, 10).Interior.ColorIndex = 36
        
        ' End If
  
    End If

  Next i

End Sub


