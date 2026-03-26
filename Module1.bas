Attribute VB_Name = "Module1"
Sub HighlightOldSentInvoices()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim status As String
    Dim sentDate As Date
    Dim cell As Range
    
    Set ws = ThisWorkbook.Sheets("InvoiceRegister")
    
    ' Find last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Loop through each invoice
    For i = 2 To lastRow
        status = ws.Cells(i, 4).Value        ' Column D = Status
        
        ' Only check Sent invoices
        If status = "Sent" Then
            ' Check if Sent Date exists
            If IsDate(ws.Cells(i, 7).Value) Then
                sentDate = ws.Cells(i, 7).Value
                ' More than 30 days old?
                If Date - sentDate > 30 Then
                    ' Highlight the Status cell (D) and optionally the row
                    ws.Cells(i, 4).Interior.Color = RGB(255, 200, 200)  ' Light red
                    ws.Rows(i).Interior.Color = RGB(255, 235, 235)      ' Optional: light pink for row
                Else
                    ' Clear any previous coloring if not overdue
                    ws.Cells(i, 4).Interior.ColorIndex = xlNone
                    ws.Rows(i).Interior.ColorIndex = xlNone
                End If
            End If
        Else
            ' Clear coloring for non-Sent invoices
            ws.Cells(i, 4).Interior.ColorIndex = xlNone
            ws.Rows(i).Interior.ColorIndex = xlNone
        End If
    Next i
    
    MsgBox "Sent invoices older than 30 days are highlighted.", vbInformation
End Sub

