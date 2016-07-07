﻿''' <summary>
''' Cette classe est une relation N-N entre <see cref="RéférenceDOuvrage"/> et <see cref="Produit"/>. Elle porte en particulier la quantité du produit dans l'ouvrage. 
''' </summary>
Public Class UsageDeProduit
    Inherits EIExObject

#Region "Constructeurs"

    Friend Sub New(Parent As RéférenceDOuvrage)
        _Parent = Parent
    End Sub

    Protected Overrides Sub Init()
    End Sub

#End Region

#Region "Propriétés"

#Region "Parent (RéférenceDOuvrage)"
    Private _Parent As RéférenceDOuvrage
    Public ReadOnly Property Parent() As RéférenceDOuvrage
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

        End Set
    End Property
#End Region

#End Region

#Region "Méthodes"


#End Region

#Region "Tests et debuggage"


#End Region

End Class
