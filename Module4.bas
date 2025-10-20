Attribute VB_Name = "Module4"
Option Explicit

Private Const TEMPLATE_SHEET As String = "PR"
Private Const SOURCE_SHEET As String = "Check Sheet"
Private Const SRC_FIRST_DATA_ROW As Long = 6

Private Const FORM_HEIGHT As Long = 26
Private Const FORM2_START_ROW As Long = 28

Private Const PR1_START As Long = 9
Private Const PR1_END As Long = 18
Private Const PR2_START As Long = 36
Private Const PR2_END As Long = 45

Private Const COL_A As String = "A"
Private Const COL_B As String = "B"
Private Const COL_I As String = "I"
Private Const COL_K As String = "K"

Public Sub Generate_PR_Files()
    Dim wsSrc As Worksheet, wsTpl As Worksheet
    Dim orderedRows As Collection
    Dim saveFolder As String
    Dim idx As Long, total As Long
    Dim wbOut As Workbook, wsOut As Worksheet
    Dim fileCount As Long
    On Error GoTo ErrHandler

    Set wsSrc = ThisWorkbook.Worksheets(SOURCE_SHEET)
    Set wsTpl = ThisWorkbook.Worksheets(TEMPLATE_SHEET)

    saveFolder = ThisWorkbook.Path & "\Print"
    If Dir(saveFolder, vbDirectory) = "" Then MkDir saveFolder

    Set orderedRows = BuildOrderedByColD(wsSrc, SRC_FIRST_DATA_ROW)
    If orderedRows Is Nothing Or orderedRows.Count = 0 Then
        MsgBox "No visible data found from row " & SRC_FIRST_DATA_ROW & " in '" & SOURCE_SHEET & "'.", vbExclamation
        Exit Sub
    End If

    total = orderedRows.Count
    idx = 1
    fileCount = 0

    Do While idx <= total
        Set wbOut = Workbooks.Add
        Set wsOut = wbOut.Worksheets(1)
        wsOut.Name = "PR"
        PrepareTwoSectionsFromTemplate wsTpl, wsOut
        idx = FillTwoSections(wsSrc, wsOut, orderedRows, idx)
        fileCount = fileCount + 1

        Static seq As Long
        Dim fn As String
        seq = seq + 1
        fn = saveFolder & "\PR_" & Format(Now, "yyyy-mm-dd_HHMMSS_") & Format(Timer * 1000 Mod 1000, "000") & "_part" & Format(fileCount, "00") & "_seq" & Format(seq, "00") & ".xlsx"
        wbOut.SaveAs Filename:=fn, FileFormat:=xlOpenXMLWorkbook
        wbOut.Close SaveChanges:=False
    Loop

    MsgBox "PR files created: " & fileCount & " file(s).", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Function BuildOrderedByColD(ws As Worksheet, startRow As Long) As Collection
    Dim lastRow As Long, r As Long
    Dim dict As Object, order As Object
    Dim out As New Collection
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set dict = CreateObject("Scripting.Dictionary")
    Set order = CreateObject("Scripting.Dictionary")
    For r = startRow To lastRow
        If Not ws.Rows(r).Hidden And Application.CountA(ws.Rows(r)) > 0 Then
            Dim key As String
            key = TrimSafe(ws.Cells(r, "D").Value)
            If Not dict.Exists(key) Then
                dict.Add key, New Collection
                order.Add order.Count + 1, key
            End If
            dict(key).Add r
        End If
    Next r
    Dim k As Variant, rowItem As Variant
    For Each k In order.Items
        For Each rowItem In dict(k)
            out.Add rowItem
        Next rowItem
    Next k
    Set BuildOrderedByColD = out
End Function

Private Sub PrepareTwoSectionsFromTemplate(wsTpl As Worksheet, wsOut As Worksheet)
    Dim rngTpl As Range
    Dim c As Long, i As Long, midRow As Long
    Set rngTpl = wsTpl.Range("A1:Q" & FORM_HEIGHT)

    wsOut.Cells.Clear
    For c = 1 To rngTpl.Columns.Count
        wsOut.Columns(c).ColumnWidth = wsTpl.Columns(c).ColumnWidth
    Next c

    rngTpl.Copy Destination:=wsOut.Range("A1")
    rngTpl.Copy Destination:=wsOut.Range("A" & FORM2_START_ROW)

    For i = 1 To FORM_HEIGHT
        wsOut.Rows(i).RowHeight = wsTpl.Rows(i).RowHeight
        wsOut.Rows(FORM2_START_ROW + (i - 1)).RowHeight = wsTpl.Rows(i).RowHeight
    Next i

    midRow = FORM2_START_ROW - 1
    wsOut.Rows(midRow).RowHeight = 6
    With wsOut.Range("A" & midRow & ":Q" & midRow)
        .Borders(xlEdgeBottom).LineStyle = xlDash
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlInsideVertical).LineStyle = xlLineStyleNone
        .Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
    End With

    With wsOut.PageSetup
        .PrintArea = "$A$1:$Q$53"
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(0)
        .BottomMargin = Application.InchesToPoints(0)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .CenterHorizontally = True
        .CenterVertically = False
        .PrintGridlines = False
    End With
End Sub

Private Function FillTwoSections(wsSrc As Worksheet, wsOut As Worksheet, _
                                 orderedRows As Collection, ByVal startIdx As Long) As Long
    Dim idx As Long: idx = startIdx
    Dim r As Long
    wsOut.Range("A" & PR1_START & ":Q" & PR1_END).ClearContents
    wsOut.Range("A" & PR2_START & ":Q" & PR2_END).ClearContents
    For r = PR1_START To PR1_END
        If idx > orderedRows.Count Then FillTwoSections = idx: Exit Function
        FillOneLine wsSrc, wsOut, orderedRows(idx), r
        idx = idx + 1
    Next r
    For r = PR2_START To PR2_END
        If idx > orderedRows.Count Then FillTwoSections = idx: Exit Function
        FillOneLine wsSrc, wsOut, orderedRows(idx), r
        idx = idx + 1
    Next r
    FillTwoSections = idx
End Function

Private Sub FillOneLine(wsSrc As Worksheet, wsOut As Worksheet, srcRow As Long, tgtRow As Long)
    Dim vA As String, vB As String, vC As String, vD As String, vE As Variant, vJ As Variant
    Dim descTxt As String
    vA = TrimSafe(wsSrc.Cells(srcRow, "A").Value)
    vB = TrimSafe(wsSrc.Cells(srcRow, "B").Value)
    vC = TrimSafe(wsSrc.Cells(srcRow, "C").Value)
    vD = TrimSafe(wsSrc.Cells(srcRow, "D").Value)
    vE = wsSrc.Cells(srcRow, "E").Value
    vJ = wsSrc.Cells(srcRow, "J").Value
    descTxt = JoinNonEmpty(Array(vB, vC), " ")
    If Len(TrimSafe(vD)) > 0 Then
        If Len(descTxt) > 0 Then
            descTxt = descTxt & " / " & TrimSafe(vD)
        Else
            descTxt = TrimSafe(vD)
        End If
    End If
    WriteToMergedTopLeft wsOut.Range(COL_A & tgtRow), vA
    WriteToMergedTopLeft wsOut.Range(COL_B & tgtRow), descTxt
    If Len(TrimSafe(vJ)) > 0 Then
        WriteToMergedTopLeft wsOut.Range(COL_I & tgtRow), vJ
    Else
        WriteToMergedTopLeft wsOut.Range(COL_I & tgtRow), ""
    End If
    If Len(TrimSafe(vE)) > 0 Then
        WriteToMergedTopLeft wsOut.Range(COL_K & tgtRow), vE
    Else
        WriteToMergedTopLeft wsOut.Range(COL_K & tgtRow), ""
    End If
End Sub

Private Function TrimSafe(v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then
        TrimSafe = ""
    Else
        TrimSafe = Trim(CStr(v))
    End If
End Function

Private Function JoinNonEmpty(arr As Variant, sep As String) As String
    Dim i As Long, t As String, s As String
    For i = LBound(arr) To UBound(arr)
        s = TrimSafe(arr(i))
        If Len(s) > 0 Then
            If Len(t) > 0 Then
                t = t & sep & s
            Else
                t = s
            End If
        End If
    Next i
    JoinNonEmpty = t
End Function

Private Sub WriteToMergedTopLeft(rng As Range, val As Variant)
    Dim tgt As Range
    If rng.MergeCells Then
        Set tgt = rng.MergeArea.Cells(1, 1)
    Else
        Set tgt = rng
    End If
    tgt.Value = val
End Sub
