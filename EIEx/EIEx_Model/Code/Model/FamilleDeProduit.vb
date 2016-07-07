Public Class FamilleDeProduit
    Inherits AgregateRoot

#Region "Constructeurs"

    Public Sub New()

    End Sub

    Public Sub New(Id As Integer)
        MyBase.New(Id)
    End Sub
    Protected Overrides Sub Init()
    End Sub

    Protected Overrides Sub SetId()
        Me._Id = Réf.GetNewId(Of FamilleDeProduit)
    End Sub

    Protected Overrides Sub SEnregistrerDansLeRéférentiel()
        Réf.EnregistrerRoot(Me)
    End Sub

#End Region

#Region "Propriétés"

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


#End Region

#Region "Tests et debuggage"


#End Region

End Class
