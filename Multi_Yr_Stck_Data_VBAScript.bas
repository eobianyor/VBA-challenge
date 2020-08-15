Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data()

'Set Worksheet variable
Dim ws As Worksheet

' Loop through all of the worksheets in the active workbook.
For Each ws In Worksheets

'Set some initial variables
Dim Stock_ticker_Name As String
Dim Stock_ticker_Open As Double
Dim Stock_ticker_Close As Double

'Set an initial variable for holding the volume total per ticker symbol
Dim Stock_Vol_Total As Double
Stock_Vol_Total = 0

'---------------------------------------------------------
' CREATE SUMMARY TABLE
'---------------------------------------------------------

'Header Labels
'Ticker Symbol
ws.Cells(1, 10).Value = "TICKER"
'Ticker Year open price
ws.Cells(1, 11).Value = "Yr_Open_Price"
'Ticker Year close price
ws.Cells(1, 12).Value = "Yr_Close_Price"
'Ticker Year change
ws.Cells(1, 13).Value = "Yr_Change"
'Ticker Year change %
ws.Cells(1, 14).Value = "Yr_Change%"
'Ticker Year Volume
ws.Cells(1, 15).Value = "Yr_Vol_Total"

'Keep track of the location for each stock ticker in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'Get last row from current worksheet
Set lastRow = ws.Cells(Rows.Count, 1).End(xlUp)
lastRowNumber = lastRow.Row

'---------------------------------------------------------
' POPULATE SUMMARY TABLE
'---------------------------------------------------------

'Print 1st line ticker open price in the Summary Table
ws.Cells(2, 11).Value = ws.Cells(2, 3).Value

'Loop through all stock tickers
    For i = 2 To lastRowNumber

        'Check if we are still within the same ticker symbol, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    'Set the Stock ticker name, open and close prices
            Stock_ticker_Name = ws.Cells(i, 1).Value
            Stock_ticker_Open = ws.Cells(i + 1, 3).Value
            Stock_ticker_Close = ws.Cells(i, 6).Value

            'Add to the Stock Volume Total
            Stock_Vol_Total = Stock_Vol_Total + ws.Cells(i, 7).Value

            'Print the ticker name in the Summary Table
            ws.Cells(Summary_Table_Row, 10).Value = Stock_ticker_Name
            
            'Print the ticker open and close price in the Summary Table
            ws.Cells(Summary_Table_Row + 1, 11).Value = Stock_ticker_Open
            ws.Cells(Summary_Table_Row, 12).Value = Stock_ticker_Close
      
            'Print the Stock Volume Amount to the Summary Table
            ws.Cells(Summary_Table_Row, 15).Value = Stock_Vol_Total

            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
      
            'Reset the Stock Volume Total
            Stock_Vol_Total = 0

            ws.Cells(Summary_Table_Row, 12).Value = ws.Cells(i, 6).Value

            'If the cell immediately following a row is the same stock ticker...
        Else

            'Add to the Stock Volume Total
            Stock_Vol_Total = Stock_Vol_Total + ws.Cells(i, 7).Value

        End If

    Next i

'Set summary table row info
Set sumTableRow = ws.Cells(Rows.Count, 10).End(xlUp)
sumTableRowNumber = sumTableRow.Row

'Clean up rows below my summary table
ws.Cells(sumTableRowNumber + 1, 11).Value = " "
ws.Cells(sumTableRowNumber + 1, 12).Value = " "

'---------------------------------------------------------
' CALCULATING YEARLY CHANGE IN STOCK PRICES
'---------------------------------------------------------

'Set some initial variables
Dim Yearly_change As Double
Dim Yearly_change_percent As Double

'Loop through summary table to calculate the yearly change
    For i = 2 To sumTableRowNumber
    
        'calculate the yearly change
        If ws.Cells(i, 11).Value <> 0 Then
            Yearly_change = (ws.Cells(i, 12).Value - ws.Cells(i, 11).Value)
            Yearly_change_percent = ((ws.Cells(i, 12).Value - ws.Cells(i, 11).Value) / ws.Cells(i, 11).Value)
            ws.Cells(i, 13).Value = Yearly_change
            ws.Cells(i, 14).Value = Yearly_change_percent
        Else
            Yearly_change = 0
            Yearly_change_percent = 0
            ws.Cells(i, 13).Value = Yearly_change
            ws.Cells(i, 14).Value = Yearly_change_percent
        End If
        
    Next i
    
'---------------------------------------------------------
' CHALLENGE - FIND GREATEST INCREASE, DECREASE AND VOLUME
'---------------------------------------------------------

'Header Labels
'Ticker Symbol
ws.Cells(1, 18).Value = "Ticker"
'Ticker Value
ws.Cells(1, 19).Value = "Value"
'Greatest % increase
ws.Cells(2, 17).Value = "Greatest % increase"
'Greatest % decrease
ws.Cells(3, 17).Value = "Greatest % decrease"
'Greatest Total volume
ws.Cells(4, 17).Value = "Greatest Total Volume"

Dim PercentageRange As Range
Dim VolumeRange As Range
Set PercentageRange = ws.Range("N2:N" & sumTableRowNumber)
Set VolumeRange = ws.Range("O2:O" & sumTableRowNumber)

'Get the values for greatest increase, decrease and volume
greatestIncrease = Application.WorksheetFunction.Max(PercentageRange)
greatestDecrease = Application.WorksheetFunction.Min(PercentageRange)
greatestVolume = Application.WorksheetFunction.Max(VolumeRange)

'Populate the cells for greatest increase, decrease and volume
ws.Cells(2, 19).Value = greatestIncrease
ws.Cells(3, 19).Value = greatestDecrease
ws.Cells(4, 19).Value = greatestVolume

'Find the corresponding ticker symbol for greatest increase, decrease and volume
For i = 2 To sumTableRowNumber

    If ws.Cells(i, 14).Value = greatestIncrease Then
        ws.Cells(2, 18).Value = ws.Cells(i, 10)
    End If
    
    If ws.Cells(i, 14).Value = greatestDecrease Then
        ws.Cells(3, 18).Value = ws.Cells(i, 10)
    End If

    If ws.Cells(i, 15).Value = greatestVolume Then
        ws.Cells(4, 18).Value = ws.Cells(i, 10)
    End If

Next i

'---------------------------------------------------------
' FORMAT THE NECESSARY COLUMNS
'---------------------------------------------------------

'Color formatting the yearly change column(s)
    For i = 2 To sumTableRowNumber
        If ws.Cells(i, 13).Value >= 0 Then
            ws.Cells(i, 13).Interior.ColorIndex = 4
        ElseIf ws.Cells(i, 13).Value < 0 Then
            ws.Cells(i, 13).Interior.ColorIndex = 3
        End If
    Next i
    
'Formatting the percent column(s)
    For i = 2 To sumTableRowNumber
        ws.Cells(i, 14).Style = "Percent"
    Next i

'Formatting the currency column(s)
    For i = 2 To sumTableRowNumber
        ws.Cells(i, 11).Style = "Currency"
        ws.Cells(i, 12).Style = "Currency"
        ws.Cells(i, 13).Style = "Currency"
    Next i
    
'Formatting the percent cell(s)
    ws.Cells(2, 19).Style = "Percent"
    ws.Cells(3, 19).Style = "Percent"

'---------------------------------------------------------
' RESET VALUES
'---------------------------------------------------------

Stock_Vol_Total = 0
lastRow = 0
lastRowNumber = 0
Summary_Table_Row = 0
sumTableRowNumber = 0

'---------------------------------------------------------
' END OF SHEET QC STATEMENT
'---------------------------------------------------------

'MsgBox ("End of Sheet")

Next

'---------------------------------------------------------
' END OF LOOP END OF SCRIPT QC STATEMENT
'---------------------------------------------------------

MsgBox ("End of Script")

End Sub
