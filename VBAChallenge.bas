Attribute VB_Name = "Module1"
Sub VBAofWallStreetFinal()

'loop through all worksheets
Dim ws As Worksheet
For Each ws In Worksheets

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'declare variables
Dim stock_name As String
Dim open_price As Double
Dim closing_price As Double
Dim stock_volume As Double
stock_volume = 0
Dim Summary_Table_Row As Long
Summary_Table_Row = 2
Dim percent_change As Double
Dim LastRow As Double

LastRow = ws.Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row


'Loop through all stocks
For i = 2 To LastRow

If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
    
    open_price = ws.Cells(i, 3).Value
    
End If

'add to the stock total
stock_volume = stock_volume + ws.Cells(i, 7).Value

' Check if we are still within the same stock name
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    'If it is not, set the stock name
    stock_name = ws.Cells(i, 1).Value

    'add to the stock total
    'stock_volume = stock_volume + ws.Cells(i, 7).Value

    'print the stock name in the summary table
    ws.Range("I" & Summary_Table_Row).Value = stock_name

    'print the stock volume in the summary table
    ws.Range("L" & Summary_Table_Row).Value = stock_volume
    
    ws.Range("J" & Summary_Table_Row).Value = yearly_change

    'find the stock price on the first day of the year
    
    closing_price = ws.Cells(i, 6).Value
    
    yearly_change = closing_price - open_price
    
    
If yearly_change >= 0 Then
    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
Else
    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
End If


'calculate % change and format as %
If open_price = 0 And closing_price = 0 Then
    percent_change = 0
    ws.Range("K" & Summary_Table_Row).Value = percent_change
    ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

ElseIf open_price = 0 Then
    Dim percent_change_newstock As String
    percent_change_newstock = "New Stock"
    ws.Range("K" & Summary_Table_Row).Value = percent_change_newstock
Else
    percent_change = (closing_price - open_price) / open_price
        ws.Range("K" & Summary_Table_Row).Value = percent_change
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
    
End If

'Add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1

'yearly_change = 0
'closing_price = 0
'open_price = 0
'percent_change = 0
stock_volume = 0

End If

Next i

Next ws

End Sub









