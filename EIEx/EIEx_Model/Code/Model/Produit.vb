Imports System.Collections.ObjectModel

Public Class Produit
    Inherits AgregateRoot

#Region "Constructeurs"

    Public Sub New()

    End Sub

    Public Sub New(Id As Integer)
        MyBase.New(Id)
        _MotsClés = New ObservableCollection(Of String)
    End Sub

    Protected Overrides Sub Init()
        _MotsClés = New ObservableCollection(Of String)
    End Sub

    Protected Overrides Sub SetId()
        Me._Id = Réf.GetNewId(Of Produit)
    End Sub

    Protected Overrides Sub SEnregistrerDansLeRéférentiel()
        Réf.EnregistrerRoot(Me)
    End Sub

#End Region

#Region "Propriétés"

#Region "Unité"
    Private _Unité As Unités?
    Public Property Unité() As Unités?
        Get
            Return _Unité
        End Get
        Set(ByVal value As Unités?)
            If Object.Equals(value, Me._Unité) Then Exit Property
            _Unité = value
            NotifyPropertyChanged(NameOf(Unité))
        End Set
    End Property
#End Region

#Region "Prix (Single)"
    Private _Prix As Single?
    Public Property Prix() As Single?
        Get
            Return _Prix
        End Get
        Set(ByVal value As Single?)
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

#Region "Famille"
    Private _Famille As FamilleDeProduit
    Public Property Famille() As FamilleDeProduit
        Get
            Return _Famille
        End Get
        Set(ByVal value As FamilleDeProduit)
            If Object.Equals(value, Me._Famille) Then Exit Property
            _Famille = value
            NotifyPropertyChanged(NameOf(Famille))
        End Set
    End Property
#End Region

#End Region

#Region "Méthodes"


#End Region

#Region "Tests et debuggage"


#End Region

End Class

Public Enum Unités
    ''' <summary>Ensemble</summary>
    Ens

    ''' <summary>?</summary>
    PM

    ''' <summary>?</summary>
    ML

    ''' <summary>Nombre</summary>
    U
End Enum
