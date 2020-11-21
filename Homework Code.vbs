Sub Analysis()
'-----Worksheet Loop Open---------------------------------
For Each ws In Worksheets
'-----Declare Variables---------------------------------
Dim Ticker As String
Dim Ticker_Row As Long
Dim Open_Value As Double
Dim Close_Value As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock_Volume As Double
'-----Declare Initial Values-----------------------------
Total_Stock_Volume = 0
Yearly_Change = 0
Ticker_Row = 2

    'Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
'-----Determine Last Row with Information-----------------
Dim Final_Row As Long
Final_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row

'-----Open Conditionals Loop-------------------------------
For i = 2 To Final_Row
'-----Conditional Set For 3 Calculations--------------------
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    'Differentiates between two different ticker names
    Ticker = ws.Cells(i, 1).Value
        ws.Range("I" & Ticker_Row).Value = Ticker
    'Establish Ticker's name then prints in summary table in order
    
'-------Yearly Change-------------------------------------
'Establish a close value from the final ticker entry
            Close_Value = ws.Range("F" & i).Value
'Establish an open value from the initial ticker entry
            Open_Value = ws.Range("C" & Ticker_Row)
'Calculate yearly change with previous values and print in column J
    Yearly_Change = Close_Value - Open_Value
    ws.Range("J" & Ticker_Row).Value = Yearly_Change

'--------Percent Change------------------------------------
'Format column K as a percentage
        ws.Range("K" & Ticker_Row).NumberFormat = "0.00%"
'Percent
    Percent_Change = (Yearly_Change / Open_Value)
    ws.Range("K" & Ticker_Row).Value = Percent_Change

'--------Total Volume---------------------------------------
    Total_Stock_Volume = Total_Stock_Volume + ws.Range("G" & i).Value
    ws.Range("L" & Ticker_Row).Value = Total_Stock_Volume
        

'Count up 1 for next summary table row
Ticker_Row = Ticker_Row + 1


Else
'Total up Yearly Volume Values for same ticker name
Total_Stock_Volume = Total_Stock_Volume + ws.Range("G" & i).Value

End If
Next i
'--------------Close Conditionals Loop-------------------------

'--------Cell Formatting Yearly Change Column (J) Loop--------
Dim Final_Row_Yearly_Change As Long
Final_Row_Yearly_Change = ws.Cells(Rows.Count, 10).End(xlUp).Row
For j = 2 To Final_Row_Yearly_Change
    'Positive Value Green Color 4
    If ws.Cells(j, 10).Value >= 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 4
    'Negative Value Red Color 3
    ElseIf ws.Cells(j, 10).Value < 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 3
    End If
Next j
'------------Closed Cell Formatting Loop-------------------
'------------Bonus-----------------------------------------
    'Headers
     ws.Cells(1, 15).Value = "Ticker"
     ws.Cells(1, 16).Value = "Value"
     ws.Cells(2, 14).Value = "Greatest % Increase"
     ws.Cells(3, 14).Value = "Greatest % Decrease"
     ws.Cells(4, 14).Value = "Greatest Total Volume"
     'New Variables
     Dim Greatest_Total_Volume As Double
     Dim Greatest_Percent_Increase As Double
     Dim Greatest_Percent_Decrease As Double
     Greatest_Percent_Increase = 0
     Greatest_Percent_Decrease = 0
     'Establish percent format for cells
     ws.Range("P2").NumberFormat = "0.00%"
     ws.Range("P3").NumberFormat = "0.00%"
     'Find bottom of percent change column
     Dim Percent_Change_Final_Row
     Percent_Change_Final_Row = ws.Cells(Rows.Count, 11).End(xlUp).Row
'-----------Open Bonus Loop----------------------------------
For i = 2 To Percent_Change_Final_Row
'Bonus Conditionals
    'sets greatest % increase to highest number
    If Greatest_Percent_Increase < ws.Range("K" & i + 1).Value Then
    Greatest_Percent_Increase = ws.Range("K" & i + 1).Value
    'Prints Ticker name of greatest increase and the value
    ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
    ws.Cells(2, 16).Value = Greatest_Percent_Increase
    
    'Sets greatest % decrease to lowest number
    ElseIf Greatest_Percent_Decrease > ws.Range("K" & i + 1).Value Then
    Greatest_Percent_Decrease = ws.Range("K" & i + 1).Value
    'Prints ticker name of greatest decrease and the value
    ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
    ws.Cells(3, 16).Value = Greatest_Percent_Decrease
    
    'Find highest total volume
    ElseIf Greatest_Total_Volume < ws.Range("L" & i + 1).Value Then
    Greatest_Total_Volume = ws.Range("L" & i + 1).Value
    'Prints ticker name of greatest volume and the value
    ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
    ws.Cells(4, 16).Value = Greatest_Total_Volume
End If
Next i
'----------Close Bonus Loop-----------------------------------
Next ws
'--------Worksheet Loop Closed---------
End Sub
