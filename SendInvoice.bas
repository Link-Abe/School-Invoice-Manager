Attribute VB_Name = "SendInvoice"
'----------------------------------------
' Send Invoice
'----------------------------------------
Sub SendInvoice()
    Dim wsRegister As Worksheet
    Dim invoiceRow As Long
    Dim invoiceNumber As String
    Dim invoicePath As String
    Dim newExcelPath As String
    Dim schoolFolderName As String
    Dim sentFolder As String
    Dim pdfPath As String
    Dim schoolEmail As String
    Dim principalName As String
    Dim wsInvoice As Workbook
    Dim olApp As Object, olMail As Object
    Dim sentDate As Date
    Dim found As Boolean
    Dim basePath As String
    
    Set wsRegister = ThisWorkbook.Sheets("InvoiceRegister")
    basePath = GetBasePath
    
    ' Show dropdown for InProgress invoices
    Unload frmInvoicesSelect
    frmInvoicesSelect.InvoiceStatus = "InProgress"
    frmInvoicesSelect.LoadInvoices
    frmInvoicesSelect.Show
    
    If frmInvoicesSelect.cmbInvoices.Value = "" Then Exit Sub
    invoiceNumber = Split(frmInvoicesSelect.cmbInvoices.Value, " - ")(0)
    
    ' Find invoice row
    found = False
    For invoiceRow = 2 To wsRegister.Cells(wsRegister.Rows.Count, 1).End(xlUp).Row
        If wsRegister.Cells(invoiceRow, 1).Value = invoiceNumber Then
            found = True
            Exit For
        End If
    Next invoiceRow
    If Not found Then MsgBox "Invoice not found in register.", vbExclamation: Exit Sub
    
    ' School details
    schoolFolderName = Sheets("Schools").Cells(GetSchoolRow(wsRegister.Cells(invoiceRow, 2).Value), 4).Value
    principalName = Sheets("Schools").Cells(GetSchoolRow(wsRegister.Cells(invoiceRow, 2).Value), 3).Value
    schoolEmail = Sheets("Schools").Cells(GetSchoolRow(wsRegister.Cells(invoiceRow, 2).Value), 5).Value
    
    ' Paths
    invoicePath = wsRegister.Cells(invoiceRow, 5).Value
    sentFolder = basePath & "\" & schoolFolderName & "\Sent\"
    EnsureFolderExists sentFolder
    
    pdfPath = sentFolder & schoolFolderName & "-Invoice" & invoiceNumber & ".pdf"
    newExcelPath = sentFolder & schoolFolderName & "-Invoice" & invoiceNumber & ".xlsx"
    
    If Dir(invoicePath) = "" Then
        MsgBox "Invoice Excel not found:" & vbNewLine & invoicePath, vbCritical
        Exit Sub
    End If
    
    ' Open invoice
    Set wsInvoice = Workbooks.Open(invoicePath)
    
    ' Export PDF
    wsInvoice.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath
    
    ' SAVE AS XLSX (THIS FIXES THE RUNTIME ERROR)
    wsInvoice.SaveAs _
        Filename:=newExcelPath, _
        FileFormat:=xlOpenXMLWorkbook, _
        CreateBackup:=False
    
    wsInvoice.Close SaveChanges:=False
    
    ' Delete original InProgress file
    If Dir(invoicePath) <> "" Then Kill invoicePath
    
    ' Update InvoiceRegister
    sentDate = Date
    wsRegister.Cells(invoiceRow, 4).Value = "Sent"
    wsRegister.Cells(invoiceRow, 5).Value = newExcelPath
    wsRegister.Cells(invoiceRow, 6).Value = pdfPath
    wsRegister.Cells(invoiceRow, 7).Value = sentDate
    
    ' Outlook draft
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)
    
    With olMail
        .To = schoolEmail
        .Subject = GetCompanyName() & " Invoice " & invoiceNumber  ' CompanyName read from Settings sheet B2
        .Body = "Hey " & principalName & "," & vbNewLine & vbNewLine & _
                "Just sending invoice #" & invoiceNumber & " for work and equipment bought to date." & vbNewLine & vbNewLine & _
                "Thanks for your business," & vbNewLine & _
                "Kind Regards," & vbNewLine & _
                "[YOUR_NAME]" & vbNewLine & _           ' Replace with your name
                "[YOUR_EMAIL]" & vbNewLine & _          ' Replace with your email address(es)
                "[YOUR_PHONE]" & vbNewLine & _          ' Replace with your phone number
                "[YOUR_TRN]"                            ' Replace with your TRN/PPSN or remove this line
        .Attachments.Add pdfPath
        .Display
    End With
    
    MsgBox "Invoice sent successfully and saved to Sent folder: " & sentFolder, vbInformation
End Sub


