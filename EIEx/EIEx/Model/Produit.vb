Imports System.Collections.ObjectModel

Public Class Produit
    Inherits EIExObject

#Region "Constructeurs"

    Public Sub New()
        _MotsClés = New ObservableCollection(Of String)
    End Sub

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

#Region "Prix (Single)"
    Private _Prix As Single
    Public Property Prix() As Single
        Get
            Return _Prix
        End Get
        Set(ByVal value As Single)
            If Object.Equals(value, Me._Prix) Then Exit Property
            _Prix = value
            NotifyPropertyChanged(NameOf(Prix))
        End Set
    End Property
#End Region

#Region "ReférenceFournisseur (String)"
    Private _ReférenceFournisseur As String
    Public Property ReférenceFournisseur() As String
        Get
            Return _ReférenceFournisseur
        End Get
        Set(ByVal value As String)
            If Object.Equals(value, Me._ReférenceFournisseur) Then Exit Property
            _ReférenceFournisseur = value
            NotifyPropertyChanged(NameOf(ReférenceFournisseur))
        End Set
    End Property
#End Region

#Region "TempsDePauseUnitaire (Integer)"
    Private _TempsDePauseUnitaire As Integer?

    ''' <summary>Le temps de pause en minutes.</summary>
    Public Property TempsDePauseUnitaire() As Integer?
        Get
            Return _TempsDePauseUnitaire
        End Get
        Set(ByVal value As Integer?)
            If Object.Equals(value, Me._TempsDePauseUnitaire) Then Exit Property
            _TempsDePauseUnitaire = value
        End Set
    End Property
#End Region

#Region "MotsClés (ObservableCollection(of String))"
    Private _MotsClés As ObservableCollection(Of String)
    Public ReadOnly Property MotsClés() As ObservableCollection(Of String)
        Get
            Return _MotsClés
        End Get
    End Property
#End Region

#End Region

#Region "Méthodes"


#End Region

#Region "Tests et debuggage"


#End Region

End Class
