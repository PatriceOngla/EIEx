Imports System.Diagnostics

Module Common

    Public EIExAddin As EIEx.ThisAddIn

    Friend Sub ManageError(ex As Exception, SubName As String, Optional Msg As String = Nothing)
        Debug.Print($"Erreur dans la routine {SubName}: {ex.GetType.Name} {vbCr & ex.Message}")
    End Sub

End Module
