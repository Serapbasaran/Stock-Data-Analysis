Attribute VB_Name = "Module1"
Sub Stock_Market()
' Set all variables
Dim Ticker As String
Dim Volume As LongLong
Dim Summary_table_row As Integer
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double

'Insert the the headers for each output on the summary table
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

' Keep track of location for each ticker
Summary_table_row = 2

' Determine the last row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Initialize total volume to "0" to hold total volume for each ticker
Volume = 0

'Retrieve opening price for eac ticker
open_price = Cells(2, 3).Value


' Loop through all ticker symbols
For i = 2 To lastrow

'Sums total volume for each ticker
Volume = Volume + Cells(i, 7).Value

' Check if we are still within the same ticker symbol
If Cells(i - 1, 1).Value = Cells(i, 1).Value And Cells(i + 1, 1) <> Cells(i, 1) Then

' Retrieve the close price at the current row
  close_price = Cells(i, 6).Value

' Calculate yearly change between closing price and opening price of each ticker for given year
  yearly_change = close_price - open_price
  
  
 ' Calculate the percent change for each ticker and correct the results for divison by zero error
  If open_price = 0 And close_price <> 0 Then
  
  percentchange = close_price / close_price
  
  ElseIf open_price = 0 And close_price = 0 Then
  
  percent_change = 0

  Else
  
  percent_change = yearly_change / open_price
  
  End If
  
  ' Add one to move to the next year's opening price
  open_price = Cells(i + 1, 3).Value
  
 ' Print the yearly change of each ticker on the summary table
  Range("J" & Summary_table_row).Value = yearly_change

 ' Apply conditional formating to the yearly change value
 
   If yearly_change > 0 Then
   Range("J" & Summary_table_row).Interior.ColorIndex = 4
   
   ElseIf yearly_change < 0 Then
   Range("J" & Summary_table_row).Interior.ColorIndex = 3
   
   End If
 
 ' Print the percent change of each ticker on the summary table
   Range("K" & Summary_table_row).Value = percent_change
 
 ' Format the percent change value as decimals
   Range("K" & Summary_table_row).NumberFormat = "0.00%"
 
End If

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

 Ticker = Cells(i, 1).Value

Range("I" & Summary_table_row).Value = Ticker
Range("L" & Summary_table_row).Value = Volume


Volume = 0

Summary_table_row = Summary_table_row + 1

End If
  
Next i

End Sub
