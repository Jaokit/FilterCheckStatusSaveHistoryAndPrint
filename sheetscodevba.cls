VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim wsHistory As Worksheet
    Set wsHistory = ThisWorkbook.Sheets("Order History")
    
    If Not Intersect(Target, Range("J6:J1000,K6:K1000000")) Is Nothing Then
        Application.EnableEvents = False
        
        Dim rng As Range
        For Each rng In Target
            If rng.Column = 10 Then
                If rng.Value <> "" And IsNumeric(rng.Value) Then
                    Dim lastRow As Long
                    lastRow = wsHistory.Cells(wsHistory.Rows.Count, 1).End(xlUp).row + 1
                    wsHistory.Cells(lastRow, 1).Value = Cells(rng.row, "A").Value
                    wsHistory.Cells(lastRow, 2).Value = Cells(rng.row, "B").Value
                    wsHistory.Cells(lastRow, 3).Value = Now
                    wsHistory.Cells(lastRow, 4).Value = rng.Value
                End If
            ElseIf rng.Column = 11 Then
                If rng.Value <> "" And IsNumeric(rng.Value) Then
                    Dim prevVal As Double
                    On Error Resume Next
                    prevVal = Val(Application.Caller.Offset(0, 0).Value)
                    On Error GoTo 0
                    
                    Dim newVal As Double
                    newVal = Val(rng.Value)
                    
                    Dim totalVal As Double
                    totalVal = prevVal + newVal
                    
                    rng.Value = totalVal
                    
                    Dim lastRowRcv As Long
                    lastRowRcv = wsHistory.Cells(wsHistory.Rows.Count, 1).End(xlUp).row + 1
                    wsHistory.Cells(lastRowRcv, 1).Value = Cells(rng.row, "A").Value
                    wsHistory.Cells(lastRowRcv, 2).Value = Cells(rng.row, "B").Value
                    wsHistory.Cells(lastRowRcv, 5).Value = Now
                    wsHistory.Cells(lastRowRcv, 6).Value = newVal
                End If
            End If
        Next rng
        
        Application.EnableEvents = True
    End If
End Sub


