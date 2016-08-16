Imports Model

Public Class Win_Ouvrages


#Region "Méthodes"
    Private Shared _DefaultInstance As Win_Ouvrages

    Public Overloads Shared Sub Show(Titre As String, Source As IEnumerable(Of Ouvrage_Base))
        Try
            If _DefaultInstance Is Nothing Then _DefaultInstance = New Win_Ouvrages
            _DefaultInstance.DataContext = Source
            _DefaultInstance.Title = Titre
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

#End Region

#Region "Tests et debuggage"


#End Region

End Class
