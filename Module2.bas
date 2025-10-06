Attribute VB_Name = "Module2"
Option Explicit

Sub CopyYellowRowsOnly()
    Dim dataWS As Worksheet, centerWS As Worksheet
    Dim lastRow As Long, pasteRow As Long
    Dim i As Long
    Dim valAD As Variant, valAA As Variant, valAB As Variant
    Dim colMap As Variant
    Dim j As Integer
    Const COL_STATUS As Long = 36

    colMap = Array(1, 6, 7, 8, 12, 26, 27, 28, 30)

    Set dataWS = ThisWorkbook.Sheets("Data")
    Set centerWS = ThisWorkbook.Sheets("Check Sheet")

    centerWS.Range("A6:Z10000").ClearContents

    lastRow = dataWS.Cells(dataWS.Rows.Count, "A").End(xlUp).Row
    pasteRow = 6

    For i = 4 To lastRow
        If IsActiveRow(dataWS, i, COL_STATUS) Then
            valAD = dataWS.Cells(i, 30).Value
            valAA = dataWS.Cells(i, 27).Value
            valAB = dataWS.Cells(i, 28).Value

            If IsNumeric(valAD) And IsNumeric(valAA) And IsNumeric(valAB) Then
                If valAD <= valAB And valAD > valAA Then
                    For j = 0 To UBound(colMap)
                        centerWS.Cells(pasteRow, j + 1).Value = dataWS.Cells(i, colMap(j)).Value
                    Next j
                    pasteRow = pasteRow + 1
                End If
            End If
        End If
    Next i
End Sub

Sub CopyRedRowsOnly()
    Dim dataWS As Worksheet, centerWS As Worksheet
    Dim lastRow As Long, pasteRow As Long
    Dim i As Long
    Dim valAD As Variant, valAA As Variant
    Dim colMap As Variant
    Dim j As Integer
    Const COL_STATUS As Long = 36

    colMap = Array(1, 6, 7, 8, 12, 26, 27, 28, 30)

    Set dataWS = ThisWorkbook.Sheets("Data")
    Set centerWS = ThisWorkbook.Sheets("Check Sheet")

    centerWS.Range("A6:Z10000").ClearContents

    lastRow = dataWS.Cells(dataWS.Rows.Count, "A").End(xlUp).Row
    pasteRow = 6

    For i = 4 To lastRow
        If IsActiveRow(dataWS, i, COL_STATUS) Then
            valAD = dataWS.Cells(i, 30).Value
            valAA = dataWS.Cells(i, 27).Value

            If IsNumeric(valAD) And IsNumeric(valAA) Then
                If valAD <= valAA Then
                    For j = 0 To UBound(colMap)
                        centerWS.Cells(pasteRow, j + 1).Value = dataWS.Cells(i, colMap(j)).Value
                    Next j
                    pasteRow = pasteRow + 1
                End If
            End If
        End If
    Next i
End Sub

Private Function IsActiveRow(ws As Worksheet, rowIndex As Long, statusCol As Long) As Boolean
    Dim v As String
    v = UCase$(Trim$(ws.Cells(rowIndex, statusCol).Value))

    If v = "" Or v = "ACTIVE" Or v = "??????" Then
        IsActiveRow = True
    Else
        If v = "INACTIVE" Or v = "???????" Or v = "?????????" Then
            IsActiveRow = False
        Else
            IsActiveRow = False
        End If
    End If
End Function

