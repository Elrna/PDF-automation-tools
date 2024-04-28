Attribute VB_Name = "Module1"
Option Explicit

Sub ConvertDocToPDF(docPath As String, pdfPath As String)
    On Error GoTo ErrorHandler
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim normalizedDocPath As String
    Dim normalizedPdfPath As String

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False

    normalizedDocPath = NormalizePath(docPath)
    normalizedPdfPath = NormalizePath(pdfPath)

    Set wdDoc = wdApp.Documents.Open(normalizedDocPath)

    wdDoc.SaveAs2 normalizedPdfPath, FileFormat:=wdFormatPDF

    wdDoc.Close False
    Set wdDoc = Nothing

    wdApp.Quit
    Set wdApp = Nothing

    Exit Sub

ErrorHandler:
    MsgBox "Error occurred: " & Err.Description, vbCritical + vbSystemModal, "Error"

    If Not wdDoc Is Nothing Then
        wdDoc.Close False
        Set wdDoc = Nothing
    End If
    If Not wdApp Is Nothing Then
        wdApp.Quit
        Set wdApp = Nothing
    End If
End Sub

Function NormalizePath(path As String) As String
    NormalizePath = Replace(path, "/", "\")
End Function
