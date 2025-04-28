Attribute VB_Name = "Module1"
Sub FilterByRed()
    Call FilterByColor(RGB(255, 0, 0))
End Sub

Sub FilterByGreen()
    Call FilterByColor(RGB(0, 176, 80))
End Sub

Sub FilterByYellow()
    Call FilterByColor(RGB(255, 255, 0))
End Sub

Sub FilterByColor(colorCode As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")

    Dim filterRange As Range
    Set filterRange = ws.Range("A3:AJ99999")

    With ws
        .AutoFilterMode = False
        filterRange.AutoFilter Field:=30, Criteria1:=colorCode, Operator:=xlFilterCellColor
    End With
End Sub

Sub ClearFilter()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
End Sub



