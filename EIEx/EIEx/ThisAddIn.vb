

Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Try
            EIExAddin = Globals.ThisAddIn
            XL = Globals.ThisAddIn.Application
        Catch ex As ArgumentException
            ManageError(ex, NameOf(ThisAddIn_Startup))
        End Try
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        Try
            XL = Nothing
            EIExAddin = Nothing
        Catch ex As ArgumentException
            ManageError(ex, NameOf(ThisAddIn_Shutdown))
        End Try
    End Sub

End Class
