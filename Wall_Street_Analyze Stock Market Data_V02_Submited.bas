Attribute VB_Name = "Module1"
Sub Wall_Street():

For Each ws In Worksheets

'Set the Header of the Summary Table
ws.Cells(1, 10).Value = "Ticker"
ws.Cells(1, 11).Value = "Yearly_Change"
ws.Cells(1, 12).Value = "Percent_Change"
ws.Cells(1, 13).Value = "Total_Stock_Volume"
ws.Range("J1:M1").Font.Bold = True
ws.UsedRange.EntireColumn.AutoFit
ws.UsedRange.EntireRow.AutoFit


Dim Ticker_Name As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock_Volume As Long

'Set the inital
Stock_Volume = 0

'Sort the ticker symbol and add the total stock volume of the stock in the summary table
Dim Sum_Table_Row As Long
Sum_Table_Row = 2
Dim i As Long
Dim Lastrow As Long
Dim Open_Price As Double
Open_Price = ws.Cells(2, 3).Value
Dim Close_Price As Double

Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To Lastrow
   If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        'Set the ticker symbol
        Ticker_Name = ws.Cells(i, 1).Value
        ws.Range("J" & Sum_Table_Row).Value = Ticker_Name
        
        'Set Yearlt change of a given year
        Close_Price = ws.Cells(i, 6).Value
        Yearly_Change = Close_Price - Open_Price
        ws.Range("K" & Sum_Table_Row).Value = Yearly_Change
        
        'Setconditional formatting red for negative and green for postive
        If Yearly_Change > 0 Then
            ws.Range("K" & Sum_Table_Row).Font.ColorIndex = 1
            ws.Range("K" & Sum_Table_Row).Interior.ColorIndex = 10
         Else
            ws.Range("K" & Sum_Table_Row).Font.ColorIndex = 1
            ws.Range("K" & Sum_Table_Row).Interior.ColorIndex = 3
         End If
        
        'Set the percent change of a given year
        If Open_Price <> 0 Then
            Percent_Change = (Yearly_Change / Open_Price) * 100
         Else
            Percent_Change = 0
         End If
        ws.Range("L" & Sum_Table_Row).Value = Percent_Change
        ws.Range("L" & Sum_Table_Row).NumberFormat = "0.00%"
        
        ' Set the total stock volume
        Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
        ws.Range("M" & Sum_Table_Row).Value = Stock_Volume
        
        'Reset
        Sum_Table_Row = Sum_Table_Row + 1
        Open_Price = ws.Cells(i + 1, 3).Value
        Close_Price = 0
        Yearly_Change = 0
        Stock_Volume = 0
    Else
        Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
    End If

    Next i
    
  Next ws

End Sub
 

