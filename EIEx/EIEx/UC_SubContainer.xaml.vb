Public Class UC_SubContainer
    Private Sub Button_Click(sender As Object, e As Windows.RoutedEventArgs)
        Try
            MsgBox("ça roule")
        Catch ex As ArgumentException
            MsgBox("ça roule pas")
        End Try
    End Sub
End Class
