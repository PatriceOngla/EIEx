Imports Model

Public Class Win_Ouvrages

#Region "Méthodes"

    Public Overloads Function ShowDialog(Titre As String, Source As IEnumerable(Of Ouvrage_Base)) As Boolean?
        Try
            Me.DataContext = Source
            Me.Title = Titre
            Return Me.ShowDialog()
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Function

#End Region

#Region "Tests et debuggage"


#End Region

End Class
