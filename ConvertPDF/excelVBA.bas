Attribute VB_Name = "Module1"
Option Explicit

Sub ConvertExcelToPDF(xlPath As String, pdfPath As String)
    On Error GoTo ErrorHandler
    Dim xlApp As Object
    Dim xlBook As Object
    Dim normalizedDocPath As String
    Dim normalizedPdfPath As String

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False

    normalizedDocPath = NormalizePath(xlPath)
    normalizedPdfPath = NormalizePath(pdfPath)

    Set xlBook = xlApp.Workbooks.Open(normalizedDocPath)

    xlBook.Sheets.Select
    xlBook.ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=normalizedPdfPath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False

    xlBook.Close False
    xlApp.Quit

    Set xlBook = Nothing
    Set xlApp = Nothing

    Exit Sub

ErrorHandler:
    MsgBox "Error occurred: " & Err.Description, vbCritical + vbSystemModal, "Error"

    If Not xlBook Is Nothing Then
        xlBook.Close False
        Set xlBook = Nothing
    End If
    If Not xlApp Is Nothing Then
        xlApp.Quit
        Set xlApp = Nothing
    End If
End Sub

Function NormalizePath(path As String) As String
    NormalizePath = Replace(path, "/", "\")
End Function
