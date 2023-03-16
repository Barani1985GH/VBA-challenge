Attribute VB_Name = "Module1"
Option Explicit

Dim ws As Worksheet              '  ws - variable to use worksheet object
Dim i As Integer, j As Integer                    ' i - For loop iteration variable
Dim last_row As Integer, report_last_row As Integer       ' last_row - total row count in a sheet,report_last_row --> last empty row number of report section
Dim ticker_sym_open As String, ticker_sym_close As String   'ticker_sym_open & ticker_sym_close --> ticker symbol on closing & opening
Dim record_year_end_row As Integer, record_year_start_row As Integer  'record_end_row &record_start_row -->start & end row numbers of a ticker
Dim opening_price As Variant, closing_price As Variant, y_change As Variant  'opening_price & closing_price --> opeing & closing price on first day and last day of the fiscal year
Dim total As Double





'************ Sub Routine to loop through all the  rows in all the sheets and find out yearly change, percent change & total volume '************

Sub list_tickers()
    record_year_start_row = 0
    record_year_start_row = 0
    

    '#####-- For each to loop throughall worksheets --#####
    For Each ws In ActiveWorkbook.Worksheets
        
        'Wrtitng column names
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change ($) "
        ws.Range("L1").Value = " Percent Change  "
        ws.Range("M1").Value = " total Volume"
        ws.Range("o2").Value = "Greatest % Inc"
        ws.Range("o3").Value = "Greatest % Dec"
        ws.Range("o4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

         
         last_row = ws.Range("B" & Rows.Count).End(xlUp).Row     'total used row count of column B


              '#####-- For loop to iterate all rows in a sheet --#####
              For i = 2 To last_row


                        If (ws.Range("B" & i).Value = 20200102) Then     ' check if it is a opening day
                                opening_price = ws.Range("C" & i).Value
                                ticker_sym_open = ws.Range("A" & i).Value
                                record_year_start_row = i
                           '     MsgBox ("OP " & opening_price & " ticker " & ticker_sym_open & " record_year_start_row " & record_year_start_row)

                        End If

                        If (ws.Range("B" & i).Value = 20201231) Then     ' check if it is a opening day
                                closing_price = ws.Range("F" & i).Value
                                ticker_sym_close = ws.Range("A" & i).Value
                                record_year_end_row = i
                              '  MsgBox ("CP - " & closing_price & " ticker " & ticker_sym_open & " record_year_end_row " & record_year_end_row)
                        End If

                        If (ticker_sym_open = ticker_sym_close) Then    'check if opening and closing dates are for same ticker

                                report_last_row = ws.Range("K" & Rows.Count).End(xlUp).Row + 1 'total used row count of column K
                                ws.Range("J" & report_last_row).Value = ticker_sym_open 'Display ticker sysmbol

                               y_change = closing_price - opening_price

                                ws.Range("K" & report_last_row).Value = y_change ' Write yearly change
                                ws.Range("K" & report_last_row).NumberFormat = "$0.00"   ' Format yearly change column

                                If ws.Range("K" & report_last_row).Value > 0 Then       'Format K column based on cell value
                                      ws.Range("K" & report_last_row).Interior.ColorIndex = 4     ' color --> Green
                                ElseIf ws.Range("K" & report_last_row).Value < 0 Then
                                      ws.Range("K" & report_last_row).Interior.ColorIndex = 3     'color --> Red
                                Else
                                     ws.Range("K" & report_last_row).Interior.ColorIndex = 6   'color --> Yellow
                                End If

                             ws.Range("L" & report_last_row).Value = y_change / opening_price ' write percent change
                             ws.Range("L" & report_last_row).NumberFormat = "0.00%"  ' Format yearly change column to with decimal



                        End If
                        ' Calculate total volume of a ticker
                        total = 0

                        If (record_year_start_row <> 0) And (record_year_end_row <> 0) Then

                            If ((record_year_start_row - record_year_end_row) < 0) Then   'Find total stock volume


                                For j = record_year_start_row To record_year_end_row
                                        total = total + ws.Range("G" & j).Value
                                Next j



                                ws.Range("M" & report_last_row).Value = total
                                record_year_start_row = 0
                                record_year_end_row = 0
                            Else
                                For j = record_year_end_row To record_year_start_row
                                     total = total + ws.Range("G" & j).Value
                                Next j
                                ws.Range("M" & report_last_row).Value = total
                                record_year_start_row = 0
                                record_year_end_row = 0
                            End If

                        End If

             Next i
             '#####-- For loop to iterate all rows in a sheet --#####

    Next ws
    '#####-- End of For each to loop--#####

    GreatestOfAll   ' calling sub routine to calulate greatest Increase, decrease& volume of ticker yearly change


End Sub

Sub GreatestOfAll()

    Dim rng As Range, rng1 As Range
    Dim row_index As Double
    Dim max_value As Double
    For Each ws In ActiveWorkbook.Worksheets

            report_last_row = ws.Range("K" & Rows.Count).End(xlUp).Row
            Set rng = ws.Range("L1:L" & report_last_row)     'Percent change yearly column
            Set rng1 = ws.Range("M1:M" & report_last_row)   'Total Volume column

            ws.Range("Q2").Value = Application.WorksheetFunction.Max(rng)    ' Greatest % Increase
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").Value = Application.WorksheetFunction.Min(rng)     'Greatest % Decrease
            ws.Range("Q3").NumberFormat = "0.00%"
            ws.Range("Q4").Value = Application.WorksheetFunction.Max(rng1)   'Greatest % volume
            max_value = ws.Range("Q2").Value

            row_index = Application.WorksheetFunction.Match(max_value, rng, 0) ' row index for % inc value
            ws.Range("P2").Value = ws.Range("J" & row_index).Value  'write ticker symbol
            row_index = 0

            row_index = Application.WorksheetFunction.Match(ws.Range("Q3").Value, rng, 0)   ' row index for  % dec value
            ws.Range("P3").Value = ws.Range("J" & row_index).Value 'write ticker value symbol
            row_index = 0

            max_value = ws.Range("Q4").Value
            row_index = Application.WorksheetFunction.Match(max_value, rng1, 0)  ' row index for Max volume value
            ws.Range("P4").Value = ws.Range("J" & row_index).Value   'write ticker value symbol
            row_index = 0


      Next ws

End Sub


