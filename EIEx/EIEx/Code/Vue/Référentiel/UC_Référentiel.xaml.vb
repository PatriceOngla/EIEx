Imports Model

Public Class UC_RéférentielView

#Region "Constructeurs"
    Private Sub UC_RéférentielView_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        Me.DataContext = Référentiel.Instance
    End Sub

#End Region

#Region "Propriétés"

#Region "Ref"
    Public ReadOnly Property Ref() As Référentiel
        Get
            Return Référentiel.Instance
        End Get
    End Property
#End Region

#End Region

End Class
