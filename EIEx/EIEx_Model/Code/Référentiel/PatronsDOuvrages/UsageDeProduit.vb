Imports Model
''' <summary>
''' Cette classe est une relation N-N entre <see cref="Ouvrage_Base"/> et <see cref="Produit"/>. Elle porte en particulier la quantité du produit dans l'ouvrage. 
''' </summary>
Public Class UsageDeProduit
    Inherits Entité
    Implements IEntitéDuRéférentiel

#Region "Constructeurs"

    Friend Sub New(Parent As Ouvrage_Base)
        _Parent = Parent
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

#Region "Parent (Ouvrage_Base)"
    Private _Parent As Ouvrage_Base
    Public ReadOnly Property Parent() As Ouvrage_Base
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
            Me.Parent.NotifyPropertyChanged(NameOf(Ouvrage_Base.PrixUnitaire))
            Me.Parent.NotifyPropertyChanged(NameOf(Ouvrage_Base.TempsDePauseCalculé))
        End Set
    End Property
#End Region

#Region "Nombre (Integer)"
    Private _Nombre As Integer
    Public Property Nombre() As Integer
        Get
            Return _Nombre
        End Get
        Set(ByVal value As Integer)
            If Object.Equals(value, Me._Nombre) Then Exit Property
            _Nombre = value
            Me.Parent.NotifyPropertyChanged(NameOf(Ouvrage_Base.PrixUnitaire))
            Me.Parent.NotifyPropertyChanged(NameOf(Ouvrage_Base.TempsDePauseCalculé))
        End Set
    End Property
#End Region

#End Region

#Region "Méthodes"


#End Region

#Region "Tests et debuggage"


#End Region

End Class
