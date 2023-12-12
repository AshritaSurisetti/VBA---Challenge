Attribute VB_Name = "Module1"
'Loop through all the stocks for 1 year
'Generate the ticker symbol

Sub Stockdata()
'Loop through all worksheets
For Each ws In Worksheets
    ws.Activate

'Add headers to the columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly change"
    ws.Range("K1").Value = "Percentage change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("N2").Value = "Greatest % increase"
    ws.Range("N3").Value = "Greatest % decrease"
    ws.Range("N4").Value = "Greatest total volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Volume"
    
'Defining the variables
    Dim Ticker As String
    Dim Yearlychange As Double
    Dim Percentage As Double
    Dim Total_Stock As Double
    Dim Openpice As Double
    Dim closeprice As Double
    Dim Greatest_percent_increase As Double
    Dim Greatest_percent_decrease As Double
    Dim increase_ticker As String
    Dim Greatest_total_voulme As Double
    
    
 'Initializing variables
 
    Total_Stock = 0
    Yearlychange = 0
    openprice = ws.Cells(2, 3).Value
    Percentage = 0
    Greatest_percent_increase = 0
    Greatest_percent_decrease = 0
    
'Definig Summary table

    Dim Summarytable As Integer
    
    Summarytable = 2


'Looping through all the rows

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow

'Check if we are still in the same ticker

If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then


    Ticker = ws.Cells(i, 1).Value
    
    Total_Stock = Total_Stock + ws.Cells(i, 7).Value
    
    closeprice = ws.Cells(i, 6).Value

    Yearlychange = closeprice - openprice

    Percentage = Yearlychange / openprice
    
    
'Print the values in the appropriate cells
    
    ws.Range("I" & Summarytable).Value = Ticker
    ws.Range("L" & Summarytable).Value = Total_Stock
    ws.Range("J" & Summarytable).Value = Yearlychange
    ws.Range("k" & Summarytable).Value = Percentage
    
    
'Conditional formatting to change the column format to %
   
    ws.Range("K" & Summarytable).NumberFormat = "0.00%"
    

'conditional formatting to change the interior color of Yearlychange column

    If Yearlychange < 0 Then
    ws.Range("J" & Summarytable).Interior.ColorIndex = 3
    ElseIf Yearlychange >= 0 Then
    ws.Range("J" & Summarytable).Interior.ColorIndex = 4
    End If
    
'compare the value in column 11 to get the greatest %, total and ticker values

    If ws.Range("K" & Summarytable).Value > Greatest_percent_increase Then
    Greatest_percent_increase = ws.Range("K" & Summarytable).Value
    increase_ticker = ws.Range("I" & Summarytable).Value
    End If
    
    If ws.Range("K" & Summarytable).Value < Greatest_percent_decrease Then
    Greatest_percent_decrease = ws.Range("K" & Summarytable).Value
    decrease_ticker = ws.Range("I" & Summarytable).Value
    End If
    
    If ws.Range("L" & Summarytable).Value > Greatest_total_volume Then
    Greatest_total_volume = ws.Range("L" & Summarytable).Value
    volume_ticker = ws.Range("I" & Summarytable).Value
    End If
    
    
'To calculate the next row
Summarytable = Summarytable + 1

'Reset the values

Total_Stock = 0
openprice = ws.Cells(i + 1, 3).Value

'If the 1st condition is true, then calculate total stock
Else

Total_Stock = Total_Stock + ws.Cells(i, 7).Value

End If

Next i

ws.Cells(2, 16).Value = FormatPercent(Greatest_percent_increase)

ws.Cells(2, 15).Value = increase_ticker


ws.Cells(3, 16).Value = FormatPercent(Greatest_percent_decrease)

ws.Cells(3, 15).Value = decrease_ticker


ws.Cells(4, 16).Value = Greatest_total_volume

ws.Cells(4, 15).Value = volume_ticker


Next ws

End Sub


