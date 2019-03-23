Attribute VB_Name = "Module1"
Sub TotalVolume()
    Dim Ticker As String
    Dim Total As Double
    Dim Total_Table As Integer
    Dim i As Long
    Dim ws As Worksheet

WS_Count = ActiveWorkbook.Worksheets.Count
    For Each ws In Worksheets
        ws.Activate
        Total_Table = 2
        Ticker = Cells(2, 1).Value
        Total = 0
        year_change = 0
        NumRows = Range("A1", Range("A1").End(xlDown)).Rows.Count
        For i = 2 To NumRows
            If Cells(i, 1) <> Cells(i + 1, 1).Value Then
                Ticker = Cells(i, 1).Value
                Total = Total + Cells(i, 7).Value
                Range("I" & Total_Table).Value = Ticker
                Range("J" & Total_Table).Value = Total
                Total_Table = Total_Table + 1
                Total = 0
            Else
                Total = Total + Cells(i, 7).Value
            
            End If
        Next i
    Next ws
End Sub

