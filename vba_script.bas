Attribute VB_Name = "Module2"
Sub multi_stock_data()
    
Dim ws As Worksheet
for each ws in Worksheet

  ' Label columm headers
  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Yearly Change"
  ws.Range("K1").Value = "Percent Change"
  ws.Range("L1").Value = "Total Stock Volume"

  ' Declare variables and set default values
  Dim stock_ticker As String
  Dim last_row_a as long
  Dim last_row_k as long
  Dim stock_volume As Double
  stock_volume = 0
  Dim summary_table_row As long
  summary_table_row = 2
  Dim open_price as Double
  Dim close_price as Double
  Dim yearly_change As Double
  Dim per_change As Double
  Dim previous_amount as long
  previous_amount = 2

  ' Determine value of the last row
  last_row_a = ws.Cells(Rows.Count, 1).End(x1Up).Row

  ' Loop through the rows
  for i = 2 to last_row_a

  ' Add the stock total volume for each stock
  stock_volume = stock_volume + ws.Cells(i, 7).Value

  ' Check if still within the same stock, if not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      
      ' Show each stock ticker
        stock_ticker = ws.Cells(i, 1).Value

        ' List the stock ticker in Column I
        ws.Range("I" & summary_table_row).Value = stock_ticker
        
        ' Show total volume per stock in Column L
        ws.Range("L" & summary_table_row).Value = stock_volume
        
        ' Reset the total volume per stock
        stock_volume = 0
        
        ' Show change from opening price to closing price for given year in Column J & adjust formatting
        open_price = ws. Range("C" & previous_amount)
        close_price = ws. Range("F" & i)

        yearly_change = close_price - open_price
        ws.Range("J" & summary_table_row).Value = yearly_change
        ws.Range("J" & summary_table_row).NumberFormat = "$0.00"

        ' Keep track of the yearly change for each stock
        if open_price = 0 Then
          per_change = 0

        else
        yearly_open = ws.Range("C" & previous_amount)
        per_change = yearly_change / open_price

        end if

        ' Show percent change in Column K & adjust formatting
        ws.Range("K" & summary_table_row).Value = per_change
        ws.Range("K" & summary_table_row).NumberFormat = "0.00%"

        ' Conditional formmating = green for positive, red for negative
        if ws.Range("J" & summary_table_row).value >= 0 Then
          ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
        
        else
          ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
      
        end if

        ' Add one to the ticker row
        summary_table_row = summary_table_row + 1

    End If

    ' Go to next row
    Next i

  ' Got to next worksheet
  Next ws

End Sub


