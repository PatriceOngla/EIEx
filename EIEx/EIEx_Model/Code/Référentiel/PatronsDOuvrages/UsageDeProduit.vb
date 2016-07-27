''' <summary>
''' Cette classe est une relation N-N entre <see cref="PatronDOuvrage"/> et <see cref="Produit"/>. Elle porte en particulier la quantité du produit dans l'ouvrage. 
''' </summary>
Public Class UsageDeProduit
    Inherits EntitéDuRéférentiel

#Region "Constructeurs"

    Friend Sub New(Parent As PatronDOuvrage)
        _Parent = Parent
    End Sub

    Protected Overrides Sub Init()
    End Sub

#End Region

#Region "Propriétés"

#Region "Parent (PatronDOuvrage)"
    Private _Parent As PatronDOuvrage
    Public ReadOnly Property Parent() As PatronDOuvrage
        Get
            Return _Parent
        End Get
    End Property
#End Region

#Region "Produit (Produit)"
    Private _Produit As Produit
    Public Property Produit() As Produit
        Get
            Return _Produit
        End Get
        Set(ByVal value As Produit)
            If Object.Equals(value, Me._Produit) Then Exit Property
            _Produit = value
            NotifyPropertyChanged(NameOf(Produit))
            Me.Parent.NotifyPropertyChanged(NameOf(PatronDOuvrage.PrixUnitaire))
            Me.Parent.NotifyPropertyChanged(NameOf(PatronDOuvrage.TempsDePauseCalculé))
        End Set
    End Property
#End Region

    '#Region "Unité"
    '    Public ReadOnly Property Unité() As Unités?
    '        Get
    '            Return Me.Produit?.Unité
    '        End Get
    '    End Property
    '#End Region

#Region "Nombre (Integer)"
    Private _Nombre As Integer
    Public Property Nombre() As Integer
        Get
            Return _Nombre
        End Get
        Set(ByVal value As Integer)
            If Object.Equals(value, Me._Nombre) Then Exit Property
            _Nombre = value
            Me.Parent.NotifyPropertyChanged(NameOf(PatronDOuvrage.PrixUnitaire))
            Me.Parent.NotifyPropertyChanged(NameOf(PatronDOuvrage.TempsDePauseCalculé))
        End Set
    End Property
#End Region

#End Region

#Region "Méthodes"


#End Region

#Region "Tests et debuggage"


#End Region

End Class
