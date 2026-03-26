Attribute VB_Name = "MarkAsPaid"
'----------------------------------------
' Mark Invoice as Paid
'----------------------------------------
Sub MarkAsPaid()
    Dim wsRegister As Worksheet
    Dim invoiceRow As Long
    Dim invoiceNumber As String
    Dim invoicePath As String
    Dim newPaidExcelPath As String
    Dim tempExcelPath As String
    Dim schoolCode As String
    Dim schoolFolderName As String
    Dim paidFolder As String
    Dim sharedFolder As String
    Dim pdfPathPaid As String
    Dim pdfPathShared As String
    Dim schoolEmail As String
    Dim principalName As String
    Dim sharedLink As String
    Dim wsInvoice As Workbook
    Dim olApp As Object, olMail As Object
    Dim paidDate As String
    Dim found As Boolean
    Dim currentYear As String
    Dim pdfPathSent As String
    
    Dim invoiceTotal As Double
    
    Set wsRegister = ThisWorkbook.Sheets("InvoiceRegister")
    
    '---------------------------------------
    ' Select Sent invoice
    '---------------------------------------
    Unload frmInvoicesSelect
    frmInvoicesSelect.InvoiceStatus = "Sent"
    frmInvoicesSelect.LoadInvoices
    frmInvoicesSelect.Show
    If frmInvoicesSelect.cmbInvoices.Value = "" Then Exit Sub
    
    invoiceNumber = Split(frmInvoicesSelect.cmbInvoices.Value, " - ")(0)
    
    '---------------------------------------
    ' Find invoice row
    '---------------------------------------
    found = False
    For invoiceRow = 2 To wsRegister.Cells(wsRegister.Rows.Count, 1).End(xlUp).Row
        If wsRegister.Cells(invoiceRow, 1).Value = invoiceNumber Then
            found = True
            Exit For
        End If
    Next invoiceRow
    If Not found Then MsgBox "Invoice not found.", vbExclamation: Exit Sub
    
    '---------------------------------------
    ' Get school details
    '---------------------------------------
    schoolCode = wsRegister.Cells(invoiceRow, 2).Value
    schoolFolderName = Sheets("Schools").Cells(GetSchoolRow(schoolCode), 4).Value
    principalName = Sheets("Schools").Cells(GetSchoolRow(schoolCode), 3).Value
    schoolEmail = Sheets("Schools").Cells(GetSchoolRow(schoolCode), 5).Value
    sharedLink = Sheets("Schools").Cells(GetSchoolRow(schoolCode), 6).Value
    
    '---------------------------------------
    ' Paths
    '---------------------------------------
    invoicePath = wsRegister.Cells(invoiceRow, 5).Value
    pdfPathSent = wsRegister.Cells(invoiceRow, 6).Value
    
    Dim basePath As String
    basePath = GetBasePath
    currentYear = Year(Date)
    
    paidFolder = basePath & "\" & schoolFolderName & "\Paid\"
    sharedFolder = basePath & "\" & schoolFolderName & "\" & schoolFolderName & _
                   "-Shared\Invoices\" & currentYear & "\"
    
    EnsureFolderExists paidFolder
    EnsureFolderExists sharedFolder
    
    If Dir(invoicePath) = "" Then MsgBox "Invoice Excel not found:" & vbNewLine & invoicePath, vbCritical: Exit Sub
    
    '---------------------------------------
    ' Paid Date input
    '---------------------------------------
    paidDate = InputBox("Enter Paid Date for Invoice #" & invoiceNumber, "Paid Date", Format(Date, "DD/MM/YYYY"))
    If paidDate = "" Then Exit Sub
    paidDate = Replace(paidDate, "/", "-")
    
    '---------------------------------------
    ' New paths
    '---------------------------------------
    newPaidExcelPath = paidFolder & schoolFolderName & "-Invoice" & invoiceNumber & "-" & paidDate & ".xlsm"
    pdfPathPaid = paidFolder & schoolFolderName & "-Invoice" & invoiceNumber & "-" & paidDate & ".pdf"
    pdfPathShared = sharedFolder & schoolFolderName & "-Invoice" & invoiceNumber & "-" & paidDate & ".pdf"
    
    tempExcelPath = Environ("TEMP") & "\" & schoolFolderName & "-Invoice" & invoiceNumber & "-" & paidDate & ".xlsm"
    If Dir(tempExcelPath) <> "" Then Kill tempExcelPath
    
    '---------------------------------------
    ' Open invoice and update PaidDate
    '---------------------------------------
    Set wsInvoice = Workbooks.Open(invoicePath)
    
    With wsInvoice.Sheets(1)
        .Unprotect Password:="lock"
        .Range("PaidDate").EntireRow.Hidden = False
        .Range("PaidDate").Value = paidDate
        invoiceTotal = .Range("InvoiceTotal").Value
        .Protect Password:="lock", UserInterfaceOnly:=True
    End With
    
    '---------------------------------------
    ' Save to temp, export PDFs
    '---------------------------------------
    wsInvoice.SaveAs Filename:=tempExcelPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    wsInvoice.ExportAsFixedFormat xlTypePDF, pdfPathPaid
    wsInvoice.ExportAsFixedFormat xlTypePDF, pdfPathShared
    wsInvoice.Close SaveChanges:=False
    
    '---------------------------------------
    ' Move Excel to Paid folder
    '---------------------------------------
    If Dir(newPaidExcelPath) <> "" Then Kill newPaidExcelPath
    Name tempExcelPath As newPaidExcelPath
    
    '---------------------------------------
    ' Delete old Sent files
    '---------------------------------------
    If Dir(invoicePath) <> "" Then Kill invoicePath
    If Dir(pdfPathSent) <> "" Then Kill pdfPathSent
    
    '---------------------------------------
    ' Update InvoiceRegister
    '---------------------------------------
    wsRegister.Cells(invoiceRow, 4).Value = "Paid"
    wsRegister.Cells(invoiceRow, 5).Value = newPaidExcelPath
    wsRegister.Cells(invoiceRow, 6).Value = pdfPathPaid
    wsRegister.Cells(invoiceRow, 8).Value = paidDate
    
    '---------------------------------------
    ' Add entry to TaxTracker
    '---------------------------------------
    Dim wsTax As Worksheet
    Dim nextRow As Long
    Dim schoolName As String
    
    Set wsTax = ThisWorkbook.Sheets("TaxTracker")
    
    schoolName = Sheets("Schools").Cells(GetSchoolRow(schoolCode), 2).Value
    
    nextRow = wsTax.Cells(wsTax.Rows.Count, 1).End(xlUp).Row + 1
    
    wsTax.Cells(nextRow, 1).Value = paidDate
    wsTax.Cells(nextRow, 2).Value = invoiceNumber
    wsTax.Cells(nextRow, 3).Value = schoolCode
    wsTax.Cells(nextRow, 4).Value = schoolName
    wsTax.Cells(nextRow, 5).Value = invoiceTotal
    
    '---------------------------------------
    ' Draft Outlook email
    '---------------------------------------
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)
    
    With olMail
        .To = schoolEmail
        .Subject = GetCompanyName() & " Payment Confirmation - " & _
                   Sheets("Schools").Cells(GetSchoolRow(schoolCode), 2).Value  ' CompanyName read from Settings sheet B2
    
        .HTMLBody = _
            "<p>Hey " & principalName & ",</p>" & _
            "<p>Sending a confirmation of payment for the attached invoice <strong>#" & invoiceNumber & "</strong>.</p>" & _
            "<p>If you have any other issues or questions, please don't hesitate to contact me.</p>" & _
            "<p>Invoices, Certification of Destruction, and Network details can be found by following the link � " & _
            "<a href='" & sharedLink & "'>" & schoolFolderName & "-Shared</a>.</p>" & _
            "<br>" & _
            "<p>Kind Regards,<br>" & _
            "[YOUR_NAME]<br>" & _           ' Replace with your name
            "[YOUR_EMAIL]<br>" & _          ' Replace with your email address(es)
            "[YOUR_PHONE]<br>" & _          ' Replace with your phone number
            "[YOUR_TRN]</p>"                ' Replace with your TRN/PPSN or remove this line
    
        .Attachments.Add pdfPathPaid
        .Display
    End With
    
    MsgBox "Invoice marked as Paid, moved to Paid folder, and PDF copied to Shared folder.", vbInformation
End Sub

