Attribute VB_Name = "Module2"
Sub stockprice()

Dim i As Long
'rowcounter is a variale to keep track of the row ticks, yearly change, percentage change, and total stock volume will be placed in
Dim rowcounter As Long
rowcounter = 2

'openpricerow keeps track of the row which the first instance of each tick is in(used to find the opening price at the beggining of the year)
Dim openpricerow As Long
openpricerow = 2

Dim closepricerow As Long
Dim closingprice As Double
Dim openingprice As Double
Dim yearlychange As Double
Dim percentchange As Double
Dim stocktotal As Double
Dim greatincrease As Double
Dim greatdecrease As Double
Dim greatvolume As Double
Dim greatincrease_tic As String
Dim greatdecrease_tic As String
Dim greatvolume_tic As String

Dim ws As Worksheet


    For Each ws In Worksheets
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Setting Column Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Identifying unique tickers and adding Ticker to column k
                ws.Cells(rowcounter, 9).Value = ws.Cells(i, 1).Value
                
                'Calculates difference between opening price at the start of year and closing price at the end of year
                closepricerow = i
                openingprice = ws.Cells(openpricerow, 3).Value
                closingprice = ws.Cells(closepricerow, 6).Value
                yearlychange = closingprice - openingprice
                ws.Cells(rowcounter, 10).Value = yearlychange
                
                
                ' calculates yearly percent change of stock, if opening price is zero percent change will not be calculated
                If openingprice = 0 Or IsEmpty(openingprice) Then
                    percentchange = 0
                Else
                    percentchange = (yearlychange / openingprice)
                    ws.Cells(rowcounter, 11).Value = FormatPercent(percentchange)
                    
                End If
                
                'Conditional formating based on positive or negative yearly change
                If yearlychange > 0 Then
                    ws.Cells(rowcounter, 10).Interior.ColorIndex = 4
                ElseIf yearlychange < 0 Then
                    ws.Cells(rowcounter, 10).Interior.ColorIndex = 3
                Else
                
                End If
                
                
                'adds final value to stock total
                stocktotal = stocktotal + ws.Cells(i, 7).Value
                ws.Cells(rowcounter, 12).Value = stocktotal
                
                openpricerow = i + 1
                rowcounter = rowcounter + 1
                
                'reset stock total when moving onto new stock
                stocktotal = 0
            Else
                'adding to the total stock volume
                stocktotal = stocktotal + ws.Cells(i, 7).Value
            End If
            
            
         Next i
         
        'reset counters for nex ws
        rowcounter = 2
        openpricerow = 2
        
        'keeps track of last row for summary calculations
        lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'reset for next ws
        greatincrease = 0
        greatdecrease = 0
        greatvolume = 0
        
        'loops through to find the greates percent increase, decrease and greatest volume
        For i = 2 To lastrow2
            If ws.Cells(i, 11).Value > greatincrease Then
                greatincrease = ws.Cells(i, 11).Value
                greatincrease_tic = ws.Cells(i, 9).Value
                
            ElseIf ws.Cells(i, 11).Value < greatdecrease Then
                greatdecrease = ws.Cells(i, 11).Value
                greatdecrease_tic = ws.Cells(i, 9).Value
            End If
            
            If ws.Cells(i, 12).Value > greatvolume Then
                greatvolume = ws.Cells(i, 12).Value
                greatvolume_tic = ws.Cells(i, 9).Value
            End If
            
        Next i
        
        'Creates table with greatest % increase, decrease and greatest volume
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = greatincrease_tic
        ws.Cells(2, 17).Value = FormatPercent(greatincrease)
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = greatdecrease_tic
        ws.Cells(3, 17).Value = FormatPercent(greatdecrease)
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = greatvolume_tic
        ws.Cells(4, 17).Value = greatvolume
           
    Next



End Sub
