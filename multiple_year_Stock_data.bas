Attribute VB_Name = "Module1"
Option Explicit

Dim ws As Worksheet              '  ws - variable to use worksheet object

Dim i As Long, j As Long                    ' i - For loop iteration variable

Dim last_row As Long, report_last_row As Long       ' last_row - total row count in a sheet,report_last_row --> last empty row number of report section

Dim opening_date As String, closing_Date As String

Dim ticker_sym_open As String, ticker_sym_close As String   'ticker_sym_open & ticker_sym_close --> ticker symbol on opening & closing

Dim record_year_end_row As Long, record_year_start_row As Long  'record_end_row &record_start_row -->end & start row numbers of a ticker

Dim opening_price As Variant, closing_price As Variant, y_change As Variant  'opening_price & closing_price --> opening & closing price on first day and last day of the fiscal year

Dim total As Double


'************ Sub Routine to loop through all the  rows in all the sheets and find out yearly change, percent change & total volume '************

Sub list_tickers()

    record_year_start_row = 0
    record_year_start_row = 0

    '------- For each to loop through all the worksheets ------
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
        opening_date = ws.Name & "0102"                                   'stock opening & closing date of each year
        closing_Date = ws.Name & "1231"

              '------- For loop to iterate all rows in a sheet ------
              For i = 2 To last_row

                      If (ws.Range("B" & i).Value = opening_date) Then     ' check if it is an opening day
                                opening_price = ws.Range("C" & i).Value
                                ticker_sym_open = ws.Range("A" & i).Value
                                record_year_start_row = i
                     End If

                     If (ws.Range("B" & i).Value = closing_Date) Then     ' check if it is a closing day
                            closing_price = ws.Range("F" & i).Value
                            ticker_sym_close = ws.Range("A" & i).Value
                            record_year_end_row = i
                     End If

                    If (ticker_sym_open = ticker_sym_close) Then    'check if opening and closing dates are for same ticker

                            report_last_row = ws.Range("K" & Rows.Count).End(xlUp).Row + 1 'total used row count of column K
                            ws.Range("J" & report_last_row).Value = ticker_sym_open 'Display ticker sysmbol

                            ws.Range("K" & report_last_row).Value = closing_price - opening_price ' Write yearly change
                            ws.Range("K" & report_last_row).NumberFormat = "0.00" ' Format yearly change column
                            If ws.Range("K" & report_last_row).Value > 0 Then       'Format K column based on cell value
                                  ws.Range("K" & report_last_row).Interior.ColorIndex = 4     ' color --> Green
                            Else
                                 ws.Range("K" & report_last_row).Interior.ColorIndex = 3   'color --> Red
                            End If

                           'ws.Range("L" & report_last_row).Value = ((ws.Range("K" & report_last_row).Value) / opening_price) * 100 ' write percent change
                            ws.Range("L" & report_last_row).Value = (ws.Range("K" & report_last_row).Value) / opening_price
                            ws.Range("L" & report_last_row).NumberFormat = "0.00%"    ' Format yearly change column to with decimal

                    End If

                    total = 0        ' Calculate total volume of a ticker

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
             '------ For loop to iterate all rows in a sheet -------

    Next ws
    '------- End of For each loop-------

    GreatestOfAll   ' calling sub routine to calulate greatest Increase, decrease & volume of ticker yearly change


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

            row_index = Application.WorksheetFunction.Match(ws.Range("Q3").Value, rng, 0)   ' row index for  % dec value
            ws.Range("P3").Value = ws.Range("J" & row_index).Value 'write ticker value symbol
            row_index = 0

            max_value = ws.Range("Q4").Value
            row_index = Application.WorksheetFunction.Match(max_value, rng1, 0)  ' row index for Max volume value
            ws.Range("P4").Value = ws.Range("J" & row_index).Value   'write ticker value symbol

  Next ws

End Sub



