Attribute VB_Name = "VBAexport"
Option Explicit

Function ExportModules() As Boolean
    Dim wbMacro As Workbook, varVar As Variant, bOpen As Boolean
    For Each varVar In Application.Workbooks
        If UCase(varVar.Name) = "MACROBOOK.XLSM" Then
            bOpen = True
            Set wbMacro = varVar
            Exit For
        End If
    Next
    If Not bOpen Then Set wbMacro = Workbooks.Open("\\PSACLW02\HOME\SHARED\MacroBook.xlsm")
    Application.Run "'" & wbMacro.Name & "'!ExportModules", ThisWorkbook
    If Not bOpen Then wbMacro.Close savechanges:=False
    ExportModules = True
End Function
