Attribute VB_Name = "ExportPDF"
Option Explicit

' Exports "Declaration" and "Items" to a single PDF next to the workbook
Public Sub Export_Declaration_and_Items_PDF()
    Dim wb As Workbook
    Dim path As String, fname As String
    Dim arrSheets As Variant
    
    Set wb = ThisWorkbook
    path = wb.Path
    If Len(path) = 0 Then
        MsgBox "Please save the workbook first.", vbExclamation
        Exit Sub
    End If
    
    fname = "Customs_Declaration_" & Format(Now, "yyyymmdd_hhmm") & ".pdf"
    arrSheets = Array("Declaration", "Items")
    
    On Error GoTo ErrHandler
    wb.Sheets(arrSheets).Select
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=path & Application.PathSeparator & fname, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    
    wb.Sheets("Declaration").Select
    MsgBox "PDF exported: " & vbCrLf & path & Application.PathSeparator & fname, vbInformation
    Exit Sub
    
ErrHandler:
    MsgBox "Export failed: " & Err.Description, vbCritical
End Sub