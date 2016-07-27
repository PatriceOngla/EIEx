Public Class FamilleDeProduit
    Inherits Entité
    Implements IAgregateRoot, IEntitéDuRéférentiel

#Region "Constructeurs"

    Public Sub New(Id As Integer)
        Me.Id = Id
    End Sub

    Protected Overrides Sub Init()
    End Sub

#End Region

#Region "Propriétés"

#Region "Système"

    Public ReadOnly Property Ref As Référentiel Implements IEntitéDuRéférentiel.Ref
        Get
            Return Référentiel.Instance
        End Get
    End Property

    Public Overrides ReadOnly Property Système As Système
        Get
            Return Ref
        End Get
    End Property

#End Region

#Region "Id"
    Public ReadOnly Property Id() As Integer? Implements IAgregateRoot.Id
#End Region

#Region "Marge (single)"
    Private _Marge As Single?
    Public Property Marge() As Single?
        Get
            Return _Marge
        End Get
        Set(ByVal value As Single?)
            If Object.Equals(value, Me._Marge) Then Exit Property
            _Marge = value
            NotifyPropertyChanged(NameOf(Marge))
        End Set
    End Property
#End Region

#End Region

#Region "Méthodes"

#Region "ToString"
    Public Overrides Function ToString() As String
        Return Me.ToStringForAgregateRoot(MyBase.ToString())
    End Function
#End Region

#End Region

#Region "Tests et debuggage"


#End Region

End Class
