Attribute VB_Name = "Module3"
Sub CheckOrderStatusWithSummary()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Order History")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    Dim i As Long, j As Long
    Dim code As String
    Dim orderDateOnly As String
    Dim totalOrdered As Double
    Dim totalReceived As Double
    Dim statusMsg As String
    
    For i = 2 To lastRow
        If ws.Cells(i, 4).Value <> "" Then
            code = ws.Cells(i, 1).Value
            orderDateOnly = Format(ws.Cells(i, 3).Value, "dd-mm-yyyy")
            totalOrdered = ws.Cells(i, 4).Value
            totalReceived = 0
            
            For j = 2 To lastRow
                If ws.Cells(j, 1).Value = code And _
                   Format(ws.Cells(j, 5).Value, "dd-mm-yyyy") = orderDateOnly Then
                    totalReceived = totalReceived + Val(ws.Cells(j, 6).Value)
                End If
            Next j
            
            If totalReceived = totalOrdered Then
                statusMsg = "Completed"
            ElseIf totalReceived > totalOrdered Then
                statusMsg = "Over Received (" & (totalReceived - totalOrdered) & ")"
            ElseIf totalReceived < totalOrdered Then
                statusMsg = "Under Received (" & (totalOrdered - totalReceived) & ")"
            End If
            
            ws.Cells(i, 7).Value = statusMsg
        End If
    Next i

    MsgBox "Order status updated.", vbInformation
End Sub


