Attribute VB_Name = "Module1"
Sub homeworktest()


For Each ws In Worksheets
ws.Activate

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Precent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

Dim Volume_Amount As Double
Dim Volume_Total As Double
Dim Ticker_Name As String

Dim Close_Value As Double
Dim Open_Value As Double
Dim Change As Variant
Dim Percent_Change As Variant

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'Initialized Variables
Open_Value = Cells(2, 3).Value
Volume_Total = 0
LastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row


    For i = 2 To LastRow
             Volume_Amount = Cells(i, 7).Value
             If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
                 Volume_Total = Volume_Total + Volume_Amount
             Else
                 Ticker_Name = ws.Cells(i, 1).Value
                 Volume_Total = Volume_Total + Volume_Amount
                 
                 Range("I" & Summary_Table_Row).Value = Ticker_Name
                 Range("L" & Summary_Table_Row).Value = Volume_Total
                 
                 Volume_Total = 0
                 
             End If
             
             If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                 Close_Value = Cells(i, 6).Value
                 Change = Close_Value - Open_Value
                 
                       If Open_Value = 0 Then
                           Percent_Change = 0#
                       Else
                           Percent_Change = Change / Open_Value
                       End If
                   Range("J" & Summary_Table_Row).Value = Change
                   Range("K" & Summary_Table_Row).Value = Percent_Change
                   Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                   
                   If Change >= 0 Then
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    Else
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    End If

                   Summary_Table_Row = Summary_Table_Row + 1
                   Open_Value = Cells(i + 1, 3).Value
                   
             End If
    Next i


'hard version begins here
'result_count: pulling the "results" data

result_count = ws.Cells(Rows.Count, "I").End(xlUp).Row


ws.Cells(2, 16).Value = "Greatest % Increase"
ws.Cells(3, 16).Value = "Greatest % Decrease"
ws.Cells(4, 16).Value = "Greatest Total Volume"

ws.Cells(1, 17).Value = "Ticker"
ws.Cells(1, 18).Value = "Value"

ws.Cells(2, 18).Value = WorksheetFunction.Max(ws.Range("K:K"))
ws.Cells(3, 18).Value = WorksheetFunction.Min(ws.Range("K:K"))
ws.Cells(4, 18).Value = WorksheetFunction.Max(ws.Range("L:L"))


For a = 2 To result_count

    If ws.Cells(a, 11).Value = ws.Cells(2, 18).Value Then
        ws.Cells(2, 17).Value = ws.Cells(a, 9).Value
    ElseIf ws.Cells(a, 11).Value = ws.Cells(3, 18).Value Then
            ws.Cells(3, 17).Value = ws.Cells(a, 9).Value
    ElseIf ws.Cells(a, 12).Value = ws.Cells(4, 18).Value Then
            ws.Cells(4, 17).Value = ws.Cells(a, 9).Value
    End If
Next a

ws.Cells(2, 18).NumberFormat = "0.00%"
ws.Cells(3, 18).NumberFormat = "0.00%"
ws.Cells(4, 18).NumberFormat = "0.00"


Next ws

End Sub

