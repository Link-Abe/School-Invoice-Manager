Attribute VB_Name = "EditInvoice"
'----------------------------------------
' Edit Invoice
'----------------------------------------
Sub EditInvoice()
    Dim wsRegister As Worksheet
    Dim invoiceRow As Long
    Dim invoiceNumber As String
    Dim invoicePath As String
    Dim wsInvoice As Workbook
    Dim found As Boolean
    
    Set wsRegister = ThisWorkbook.Sheets("InvoiceRegister")
    
    ' Select InProgress invoice
    Unload frmInvoicesSelect
    frmInvoicesSelect.InvoiceStatus = "InProgress"
    frmInvoicesSelect.LoadInvoices
    frmInvoicesSelect.Show
    
    If frmInvoicesSelect.cmbInvoices.Value = "" Then Exit Sub
    invoiceNumber = Split(frmInvoicesSelect.cmbInvoices.Value, " - ")(0)
    
    ' Find invoice
    found = False
    For invoiceRow = 2 To wsRegister.Cells(wsRegister.Rows.Count, 1).End(xlUp).Row
        If wsRegister.Cells(invoiceRow, 1).Value = invoiceNumber Then
            found = True
            Exit For
        End If
    Next invoiceRow
    
    If Not found Then
        MsgBox "Invoice not found in register.", vbExclamation
        Exit Sub
    End If
    
    ' Get invoice path
    invoicePath = wsRegister.Cells(invoiceRow, 5).Value
    If Dir(invoicePath) = "" Then
        MsgBox "Invoice Excel file not found:" & vbNewLine & invoicePath, vbCritical
        Exit Sub
    End If
    
    ' Open invoice
    Set wsInvoice = Workbooks.Open(invoicePath)
    
    ' Update InvoiceDate to today (NO protection)
    With wsInvoice.Sheets(1)
        .Range("InvoiceDate").Value = Date
    End With
    
    wsInvoice.Activate
    MsgBox "Invoice #" & invoiceNumber & " opened. Today's date added. You can edit manually.", vbInformation
End Sub

