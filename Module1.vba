Attribute VB_Name = "Module1"
Sub Button1_Click()

For Each ws In ActiveWorkbook.Worksheets


    ' Set variables
    Dim I As Double
    Dim yearlychange As Double
    Dim PercentChange As Double
    Dim vol As Double
    Dim summary_table_row As Integer
    Dim yeardif As Double

    'Get column length
    columnlength = ws.Cells(Rows.Count, "A").End(xlUp).Row

    'set the start point
    summary_table_row = 2
    strtpoint = 2
    vol = 0

    'Set the titles for the summary table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"

    'iterate the length of the columns beginning after the title cell
        For I = strtpoint To columnlength
            
            'Add Volume up
            vol = vol + Cells(I, 7).Value
        
            'If the stockticker changes
            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
                   
                'Put the ticker symbol in its place
                ws.Range("I" & summary_table_row).Value = (ws.Cells(I, 1).Value)
                'get the yearly difference between opening and closing
                yeardif = ((ws.Cells(strtpoint, 3).Value) - (ws.Cells(I, 5).Value)) * (-1)
                'Change the color of the cell to highlight positive or negative returns
                    If yeardif >= 0 Then ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                    If yeardif < 0 Then ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                ws.Range("J" & summary_table_row).Value = yeardif

                'USE IF TO TEST IF YOU ARE DIVIDING BY 0)
    
                If ws.Cells(strtpoint, 3).Value = 0 Then PercentChange = 0
                'percent change ((newprice-oldprice)/oldprice)*100
                If ws.Cells(strtpoint, 3).Value <> 0 Then PercentChange = (((ws.Cells(I, 5).Value) - (ws.Cells(strtpoint, 3).Value)) / (ws.Cells(strtpoint, 3).Value))
                ws.Range("K" & summary_table_row).Value = PercentChange
                'Format cell as percent
                ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
                'Print the total Volume
                ws.Range("L" & summary_table_row).Value = vol
                'adjust strtpoint
                strtpoint = I + 1
                'reset vol
                vol = 0
                'adjust the output row
                summary_table_row = summary_table_row + 1
            
            End If

        Next I
        
    Next ws

End Sub


