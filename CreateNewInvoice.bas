Attribute VB_Name = "CreateNewInvoice"
'----------------------------------------
' Create New Invoice
'----------------------------------------
Public Sub CreateNewInvoice()
    Dim wsSchools As Worksheet, wsRegister As Worksheet
    Dim wsTemplate As Workbook, wsInvoice As Workbook
    Dim schoolCode As String, schoolFolderName As String
    Dim details As Variant
    Dim newInvoiceNumber As Long
    Dim basePath As String, schoolFolder As String, invoicePath As String
    Dim calloutFee As Double
    Dim lastRow As Long
    
    Set wsSchools = Sheets("Schools")
    Set wsRegister = Sheets("InvoiceRegister")
    basePath = GetBasePath
        
    ' Show dropdown form
    frmSchoolsSelect.Show
    If frmSchoolsSelect.cmbSchools.Value = "" Then Exit Sub
    
    ' Extract school code
    schoolCode = Split(frmSchoolsSelect.cmbSchools.Value, " - ")(0)
    
    ' Validate school
    If GetSchoolRow(schoolCode) = 0 Then
        MsgBox "School code not found!", vbExclamation
        Exit Sub
    End If
    
    ' Get school details
    details = GetSchoolDetails(schoolCode)
    If IsEmpty(details) Then
        MsgBox "School details not found.", vbCritical
        Exit Sub
    End If
    schoolFolderName = wsSchools.Cells(GetSchoolRow(schoolCode), 4).Value
    schoolFolder = basePath & "\" & schoolFolderName & "\InProgress\"
    
    ' Get Callout Fee
    calloutFee = wsSchools.Cells(GetSchoolRow(schoolCode), 9).Value
    
    ' Get next invoice number
    newInvoiceNumber = GetNextInvoiceNumber
    
    ' Open template
    Set wsTemplate = Workbooks.Open(basePath & "\InvoiceTemplate\InvoiceTemplate.xlsm")
    
    ' Save new invoice
    invoicePath = schoolFolder & schoolFolderName & "-Invoice" & newInvoiceNumber & ".xlsm"
    wsTemplate.SaveCopyAs invoicePath
    wsTemplate.Close False
    
    ' Open new invoice
    Set wsInvoice = Workbooks.Open(invoicePath)
    
    ' Fill invoice fields
    With wsInvoice.Sheets(1)
        .Range("InvoiceNumber").Value = newInvoiceNumber
        .Range("InvoiceDate").Value = Date
        .Range("SchoolName").Value = details(0)
        .Range("SchoolEmail").Value = details(1)
        .Range("SchoolPhone").Value = details(2)
        .Range("SchoolAddress").Value = details(3)
        .Range("CalloutFee").Value = calloutFee
        .Range("PaidDate").Value = ""
    End With
    
    ' Add entry to InvoiceRegister
    lastRow = wsRegister.Cells(wsRegister.Rows.Count, 1).End(xlUp).Row + 1
    
    wsRegister.Cells(lastRow, 1).Value = newInvoiceNumber
    wsRegister.Cells(lastRow, 2).Value = schoolCode
    wsRegister.Cells(lastRow, 3).Value = Date
    wsRegister.Cells(lastRow, 4).Value = "InProgress"
    wsRegister.Cells(lastRow, 5).Value = invoicePath
    wsRegister.Cells(lastRow, 6).Value = ""
    wsRegister.Cells(lastRow, 7).Value = ""
    wsRegister.Cells(lastRow, 8).Value = ""
    wsRegister.Cells(lastRow, 9).Value = calloutFee
    
    wsInvoice.Activate
    MsgBox "New invoice created: " & invoicePath, vbInformation
End Sub
