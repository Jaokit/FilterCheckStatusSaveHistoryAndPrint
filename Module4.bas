Attribute VB_Name = "Module4"
Option Explicit

Private Const TEMPLATE_SHEET As String = "PR"
Private Const SOURCE_SHEET As String = "Check Sheet"
Private Const SRC_FIRST_DATA_ROW As Long = 6
Private Const FORM_HEIGHT As Long = 26
Private Const PR_LIST_START_ROW As Long = 9
Private Const PR_LIST_END_ROW As Long = 18
Private Const COL_A As String = "A"
Private Const COL_B As String = "B"
Private Const COL_K As String = "K"

Public Sub Generate_PR_Files()
    Dim wsSrc As Worksheet, wsTpl As Worksheet
    Dim orderedRows As Collection
    Dim saveFolder As String
    Dim idx As Long, total As Long, perFile As Long
    Dim wbOut As Workbook, wsOut As Worksheet
    Dim pageNo As Long
    On Error GoTo ErrHandler
    
    Set wsSrc = ThisWorkbook.Worksheets(SOURCE_SHEET)
    Set wsTpl = ThisWorkbook.Worksheets(TEMPLATE_SHEET)
    saveFolder = ThisWorkbook.Path
    perFile = 10
    
    Set orderedRows = BuildOrderedByColD(wsSrc, SRC_FIRST_DATA_ROW)
    If orderedRows Is Nothing Or orderedRows.Count = 0 Then
        MsgBox "No visible data found from row " & SRC_FIRST_DATA_ROW & " in '" & SOURCE_SHEET & "'.", vbExclamation
        Exit Sub
    End If
    
    total = orderedRows.Count
    idx = 1
    pageNo = 0
    
    Do While idx <= total
        Set wbOut = Workbooks.Add
        Set wsOut = wbOut.Worksheets(1)
        wsOut.Name = "PR"
        PrepareOneSheetFromTemplate wsTpl, wsOut
        idx = FillOnePage(wsSrc, wsOut, orderedRows, idx)
        pageNo = pageNo + 1
        
        Static seq As Long
        Dim fn As String
        seq = seq + 1
        fn = saveFolder & "\PR_" & Format(Now, "yyyy-mm-dd_HHMMSS_") & Format(Timer * 1000 Mod 1000, "000") & "_part" & Format(pageNo, "00") & "_seq" & Format(seq, "00") & ".xlsx"
        wbOut.SaveAs Filename:=fn, FileFormat:=xlOpenXMLWorkbook
        wbOut.Close SaveChanges:=False
    Loop
    
    MsgBox "PR files created: " & pageNo & " file(s).", vbInformation
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

Private Sub PrepareOneSheetFromTemplate(wsTpl As Worksheet, wsOut As Worksheet)
    Dim rngTpl As Range
    Dim c As Long, i As Long
    Set rngTpl = wsTpl.Range("A1:Q" & FORM_HEIGHT)
    wsOut.Cells.Clear
    For c = 1 To rngTpl.Columns.Count
        wsOut.Columns(c).ColumnWidth = wsTpl.Columns(c).ColumnWidth
    Next c
    rngTpl.Copy Destination:=wsOut.Range("A1")
    For i = 1 To FORM_HEIGHT
        wsOut.Rows(i).RowHeight = wsTpl.Rows(i).RowHeight
    Next i
End Sub

Private Function FillOnePage(wsSrc As Worksheet, wsOut As Worksheet, orderedRows As Collection, ByVal startIdx As Long) As Long
    Dim rPR As Long, idx As Long
    Dim srcRow As Long
    Dim vA As String, vB As String, vC As String, vD As String, vE As Variant
    Dim descTxt As String
    idx = startIdx
    wsOut.Range("A" & PR_LIST_START_ROW & ":Q" & PR_LIST_END_ROW).ClearContents
    For rPR = PR_LIST_START_ROW To PR_LIST_END_ROW
        If idx > orderedRows.Count Then Exit For
        srcRow = CLng(orderedRows(idx))
        vA = TrimSafe(wsSrc.Cells(srcRow, "A").Value)
        vB = TrimSafe(wsSrc.Cells(srcRow, "B").Value)
        vC = TrimSafe(wsSrc.Cells(srcRow, "C").Value)
        vD = TrimSafe(wsSrc.Cells(srcRow, "D").Value)
        vE = wsSrc.Cells(srcRow, "E").Value
        descTxt = JoinNonEmpty(Array(vB, vC), " ")
        If Len(TrimSafe(vD)) > 0 Then
            If Len(descTxt) > 0 Then
                descTxt = descTxt & " / " & TrimSafe(vD)
            Else
                descTxt = TrimSafe(vD)
            End If
        End If
        WriteToMergedTopLeft wsOut.Range(COL_A & rPR), vA
        WriteToMergedTopLeft wsOut.Range(COL_B & rPR), descTxt
        If Len(TrimSafe(vE)) > 0 Then
            WriteToMergedTopLeft wsOut.Range(COL_K & rPR), vE
        Else
            WriteToMergedTopLeft wsOut.Range(COL_K & rPR), ""
        End If
        idx = idx + 1
    Next rPR
    FillOnePage = idx
End Function

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
