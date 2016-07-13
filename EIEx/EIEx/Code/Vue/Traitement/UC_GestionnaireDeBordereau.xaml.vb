Public Class UC_GestionnaireDeBordereau

#Region "Constructeurs"

    Private Sub UC_GestionnaireDeBordereau_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        Me.DataContext = New GestionnaireDeBordereau
    End Sub

#End Region

#Region "Propriétés"

#Region "GDB"
    Public ReadOnly Property GDB() As GestionnaireDeBordereau
        Get
            Return Me.DataContext
        End Get
    End Property
#End Region
#End Region

#Region "Méthodes"

    Private Sub Btn_Start_Click(sender As Object, e As Windows.RoutedEventArgs) Handles Btn_Start.Click
        Me.GDB.RécupérerLesLibellésDOuvrages()
    End Sub

#End Region

#Region "Tests et debuggage"


#End Region

End Class
