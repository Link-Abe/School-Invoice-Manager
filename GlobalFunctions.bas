Attribute VB_Name = "GlobalFunctions"
Option Explicit

'----------------------------------------
' Module: Invoice Management
'----------------------------------------

'----------------------------------------
' Global Functions
'----------------------------------------

' Return the base folder path from Settings sheet
Public Function GetBasePath() As String
    GetBasePath = Sheets("Settings").Range("B1").Value
End Function

' Return the company name from Settings sheet
Public Function GetCompanyName() As String
    GetCompanyName = Sheets("Settings").Range("B2").Value
End Function

' Return the next global invoice number (year + 001)
Public Function GetNextInvoiceNumber() As Long
    Dim ws As Worksheet
    Set ws = Sheets("Counters")
    
    Dim currentYear As Long
    currentYear = Year(Date)
    
    ' Reset counter if year changed
    If ws.Range("B2").Value <> currentYear Then
        ws.Range("B2").Value = currentYear
        ws.Range("C2").Value = 0
    End If
    
    ' Increment counter
    ws.Range("C2").Value = ws.Range("C2").Value + 1
    
    ' Combine year + counter (e.g., 2026001)
    GetNextInvoiceNumber = CLng(currentYear & Format(ws.Range("C2").Value, "000"))
End Function

' Return school details as a variant array (SchoolName, Email, Phone, Address)
Public Function GetSchoolDetails(code As String) As Variant
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Schools")

    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(r, 1).Value = code Then
            GetSchoolDetails = Array( _
                ws.Cells(r, 2).Value, _
                ws.Cells(r, 5).Value, _
                ws.Cells(r, 7).Value, _
                ws.Cells(r, 8).Value, _
                ws.Cells(r, 9).Value _
            )
            Exit Function
        End If
    Next r

    GetSchoolDetails = Empty
End Function

' Return the row number of a school in the Schools sheet
Public Function GetSchoolRow(code As String) As Long
    Dim ws As Worksheet
    Set ws = Sheets("Schools")
    
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(r, 1).Value = code Then
            GetSchoolRow = r
            Exit Function
        End If
    Next r
    
    GetSchoolRow = 0 ' not found
End Function

' Ensure folder exists, create if not
Public Sub EnsureFolderExists(ByVal folderPath As String)
    Dim fso As Object
    Dim parts As Variant
    Dim currentPath As String
    Dim i As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    parts = Split(folderPath, "\")

    ' Handle drive letter (E:\)
    currentPath = parts(0) & "\"

    For i = 1 To UBound(parts)
        currentPath = currentPath & parts(i) & "\"
        If Not fso.FolderExists(currentPath) Then
            fso.CreateFolder currentPath
        End If
    Next i
End Sub

