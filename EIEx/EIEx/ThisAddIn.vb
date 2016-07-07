Imports Utils

Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Try
        Catch ex As ArgumentException
            ManageError(ex, NameOf(ThisAddIn_Startup))
        End Try
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        Try
            ExcelEventManager.CleanUp()
            EIExAddin = Nothing
        Catch ex As ArgumentException
            ManageError(ex, NameOf(ThisAddIn_Shutdown))
        End Try
    End Sub

End Class
