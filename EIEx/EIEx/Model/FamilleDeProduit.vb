Public Class FamilleDeProduit
    Inherits EIExObject

#Region "Constructeurs"


#End Region

#Region "Propriétés"

#Region "Nom (String)"
    Private _Nom As String
    Public Property Nom() As String
        Get
            Return _Nom
        End Get
        Set(ByVal value As String)
            If Object.Equals(value, Me._Nom) Then Exit Property
            _Nom = value
            NotifyPropertyChanged(NameOf(Nom))
        End Set
    End Property
#End Region

#Region "Marge (single)"
    Private _Marge As Single
    Public Property Marge() As Single
        Get
            Return _Marge
        End Get
        Set(ByVal value As Single)
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
