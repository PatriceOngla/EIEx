Imports Model

Public MustInherit Class AgregateRootDuRéférentiel(Of T As AgregateRootDuRéférentiel(Of T))
    Inherits AgregateRoot(Of T)

#Region "Constructeurs"

    Public Sub New(Id As Integer)
        MyBase.New(Id)
    End Sub

#End Region

#Region "Propriétés"

    Protected Ref As Référentiel = Référentiel.Instance
    Public Overrides ReadOnly Property Système As Système
        Get
            Return Ref
        End Get
    End Property

#End Region

End Class
