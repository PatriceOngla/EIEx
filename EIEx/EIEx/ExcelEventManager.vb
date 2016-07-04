Imports Microsoft.Office.Interop
Friend Module ExcelEventManager

    Public WithEvents XL As Excel.Application

    Public TargetSheet As Excel.Worksheet

    Public WithEvents TargetWindow As Excel.Window

    Friend Sub SetTargetSheet()

    End Sub

End Module
