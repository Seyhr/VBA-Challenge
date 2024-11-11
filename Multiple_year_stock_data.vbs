Sub Multiple_year_stock_data()
Dim i As Long
    Dim j As Long
    Dim ticker_name As String
    Dim open_price As Variant
    Dim close_price As Variant
    Dim price_diff As Variant
    Dim min_date As Date
    Dim max_date As Date
    Dim lastRow As Long
    Dim summary_table_row As Long
    Dim ticker_start_row As Long
    Dim total_volume As Double
    Dim ws As Worksheet
    Dim greatest_increase As Variant
    Dim greatest_increase_name As String
    Dim greatest_increase_location As Long
    Dim greatest_decrease As Variant
    Dim greatest_decrease_name As String
    Dim greatest_volume_row As Long
    Dim greatest_volume As Double
    Dim greatest_volume_location As Variant
    
          

    '-----------------------------------
    ' LOOP THROUGH ALL SHEETS
    '-------------------------------------
        
        For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        
       
    
    ' Define the last row in the data
    lastRow = Cells(Rows.Count, 2).End(xlUp).Row


    ' Set up column headers for the summary
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Quarterly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"

    ' Initialize the starting row for the summary table
    summary_table_row = 2

    ' Loop through rows to process each unique ticker
    For i = 2 To lastRow
        ' Track the start row of the current ticker
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            ticker_start_row = i
        End If

        ' Check if the current ticker changes in the next row
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker_name = ws.Cells(i, 1).Value

            ' Calculate min_date and max_date for the current ticker only
            min_date = WorksheetFunction.Min(ws.Range("B" & ticker_start_row & ":B" & i))
            max_date = WorksheetFunction.Max(ws.Range("B" & ticker_start_row & ":B" & i))

            ' Initialize total volume
            total_volume = 0

            ' Loop within the current ticker's rows to find the open and close prices and calculate total volume
            open_price = ""
            close_price = ""
            
            For j = ticker_start_row To i
                If ws.Cells(j, 2).Value = min_date Then
                    open_price = ws.Cells(j, 3).Value
                End If
                If ws.Cells(j, 2).Value = max_date Then
                    close_price = ws.Cells(j, 6).Value
                End If
                ' Calculate total volume
                total_volume = total_volume + ws.Cells(j, 7).Value
                
            Next j

        'Calculate price diff
                price_diff = close_price - open_price

            ' Write the results to the summary table
            ws.Cells(summary_table_row, 10).Value = ticker_name
            ws.Cells(summary_table_row, 11).Value = price_diff
            ws.Cells(summary_table_row, 12).Value = price_diff / open_price
            ws.Cells(summary_table_row, 12).NumberFormat = "0.00%"
            ws.Cells(summary_table_row, 13).Value = total_volume
            
            
        'Find greatest increase/decrease/volume
           greatest_increase = WorksheetFunction.Max(ws.Range("L:L"))
           ws.Cells(2, 17).Value = greatest_increase
           ws.Cells(2, 17).NumberFormat = "0.00%"
                      
           greatest_decrease = WorksheetFunction.Min(ws.Range("L:L"))
           ws.Cells(3, 17).Value = greatest_decrease
           ws.Cells(3, 17).NumberFormat = "0.00%"
           
           greatest_volume = WorksheetFunction.Max(ws.Range("M:M"))
           ws.Cells(4, 17).Value = greatest_volume
           
           
           ' Use Match to find the row number of the greatest volume in column M

            greatest_volume_location = Application.Match(greatest_volume, ws.Range("M:M"), 0)

            If Not IsError(greatest_volume_location) Then
            
            ' If the row is found, retrieve the ticker name from column K
            greatest_volume_name = ws.Cells(greatest_volume_location, 10).Value
    
            ' Output the ticker name in cell P4
            ws.Cells(4, 16).Value = greatest_volume_name
            Else
            
            ' If not found, handle the case (optional)
            MsgBox "Could not find the maximum volume in column M.", vbExclamation
            End If
           
           ' Use Match to find the row number of the greatest increase in column L

            greatest_increase_location = Application.Match(greatest_increase, ws.Range("L:L"), 0)

            If Not IsError(greatest_increase_location) Then
            ' If the row is found, retrieve the ticker name from column K
            greatest_increase_name = ws.Cells(greatest_increase_location, 10).Value
    
            ' Output the ticker name in cell P2
            ws.Cells(2, 16).Value = greatest_increase_name
            Else
            ' If not found, handle the case (optional)
            MsgBox "Could not find the greatest increase in column L.", vbExclamation
            End If
           
           
            ' Use Match to find the row number of the greatest decrease in column L

            greatest_decrease_location = Application.Match(greatest_decrease, ws.Range("L:L"), 0)

            If Not IsError(greatest_decrease_location) Then
            ' If the row is found, retrieve the ticker name from column K
            greatest_decrease_name = ws.Cells(greatest_decrease_location, 10).Value
    
            ' Output the ticker name in cell P3
            ws.Cells(3, 16).Value = greatest_decrease_name
            Else
            ' If not found, handle the case (optional)
            MsgBox "Could not find the greatest decrease in column L.", vbExclamation
            End If

                           
                    
            '---------------------------------------------
            'Conditional formatting for Quarterly Change
            '-------------------------------------------
            If ws.Cells(summary_table_row, 11).Value > 0 Then
                 
                 ws.Cells(summary_table_row, 11).Interior.ColorIndex = 4
                 
            ElseIf ws.Cells(summary_table_row, 11).Value < 0 Then
                 
                 ws.Cells(summary_table_row, 11).Interior.ColorIndex = 3
           
            End If
            
            
            
             '---------------------------------------------
            'Conditional formatting for Percent Change
            '-------------------------------------------
            If ws.Cells(summary_table_row, 12).Value > 0 Then
                 
                 ws.Cells(summary_table_row, 12).Interior.ColorIndex = 4
                 
            ElseIf ws.Cells(summary_table_row, 12).Value < 0 Then
                 
                 ws.Cells(summary_table_row, 12).Interior.ColorIndex = 3
           
            End If
            ' Move to the next row in the summary table
            summary_table_row = summary_table_row + 1
        End If
    Next i
    Next ws
End Sub

