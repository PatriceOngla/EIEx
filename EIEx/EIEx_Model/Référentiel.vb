Imports System.Collections.ObjectModel

Public Class Référentiel
    Inherits EIExObject

#Region "Constructeurs"

    Public Sub New()
        _Produits = New ObservableCollection(Of Produit)
        _FamillesDeProduit = New ObservableCollection(Of FamilleDeProduit)
        _RéférencesDOuvrage = New ObservableCollection(Of RéférenceDOuvrage)
    End Sub
#End Region

#Region "Propriétés"

#Region "Produits (ObservableCollection(Of Produit))"
    Private _Produits As ObservableCollection(Of Produit)
    ''' <summary>Tous les <see cref="Produit"/>s du référentiel.</summary>
    Public ReadOnly Property Produits() As ObservableCollection(Of Produit)
        Get
            Return _Produits
        End Get
    End Property
#End Region

#Region "FamillesDeProduit (ObservableCollection(Of FamilleDeProduit)"
    Private _FamillesDeProduit As ObservableCollection(Of FamilleDeProduit)
    ''' <summary>Tous les <see cref="FamilleDeProduit"/>t du référentiel.</summary>
    Public ReadOnly Property FamillesDeProduit() As ObservableCollection(Of FamilleDeProduit)
        Get
            Return _FamillesDeProduit
        End Get
    End Property
#End Region

#Region "RéférencesDOuvrage (ObservableCollection(Of RéférenceDOuvrage))"
    Private _RéférencesDOuvrage As ObservableCollection(Of RéférenceDOuvrage)
    ''' <summary>Toutes les <see cref="RéférenceDOuvrage"/> du référentiel.</summary>
    Public ReadOnly Property RéférencesDOuvrage() As ObservableCollection(Of RéférenceDOuvrage)
        Get
            Return _RéférencesDOuvrage
        End Get
    End Property
#End Region

#End Region

#Region "Méthodes"


#End Region

#Region "Tests et debuggage"


#End Region

End Class
