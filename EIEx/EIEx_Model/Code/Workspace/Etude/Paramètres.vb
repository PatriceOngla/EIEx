Imports Model

Public Class Paramètres
    Inherits Entité
    Implements IEntitéDuWorkSpace

#Region "Constructeurs"
    Protected Overrides Sub Init()
    End Sub

#End Region

#Region "Propriétés"

#Region "WS"
    Public ReadOnly Property WS As WorkSpace Implements IEntitéDuWorkSpace.WS
        Get
            Return WorkSpace.Instance
        End Get
    End Property
#End Region

#Region "Système"
    Public Overrides ReadOnly Property Système As Système
        Get
            Return Me.WS
        End Get
    End Property
#End Region

#Region "AdresseRangeLibelleOuvrage (String)"
    Private _AdresseRangeLibelleOuvrage As String
    Public Property AdresseRangeLibelleOuvrage() As String
        Get
            Return _AdresseRangeLibelleOuvrage
        End Get
        Set(ByVal value As String)
            If Object.Equals(value, Me._AdresseRangeLibelleOuvrage) Then Exit Property
            _AdresseRangeLibelleOuvrage = value
            NotifyPropertyChanged(NameOf(AdresseRangeLibelleOuvrage))
        End Set
    End Property
#End Region

#Region "AdresseRangeUnité (String)"
    Private _AdresseRangeUnité As String
    Public Property AdresseRangeUnité() As String
        Get
            Return _AdresseRangeUnité
        End Get
        Set(ByVal value As String)
            If Object.Equals(value, Me._AdresseRangeUnité) Then Exit Property
            _AdresseRangeUnité = value
            NotifyPropertyChanged(NameOf(AdresseRangeUnité))
        End Set
    End Property
#End Region

#Region "AdresseRangePrixUnitaire (String)"
    Private _AdresseRangePrixUnitaire As String
    Public Property AdresseRangePrixUnitaire() As String
        Get
            Return _AdresseRangePrixUnitaire
        End Get
        Set(ByVal value As String)
            If Object.Equals(value, Me._AdresseRangePrixUnitaire) Then Exit Property
            _AdresseRangePrixUnitaire = value
            NotifyPropertyChanged(NameOf(AdresseRangePrixUnitaire))
        End Set
    End Property
#End Region

#Region "AdresseRangeXYZ"
    Private _AdresseRangeXYZ As String
    Public Property AdresseRangeXYZ() As String
        Get
            Return _AdresseRangeXYZ
        End Get
        Set(ByVal value As String)
            If Object.Equals(value, Me._AdresseRangeXYZ) Then Exit Property
            _AdresseRangeXYZ = value
            NotifyPropertyChanged(NameOf(AdresseRangeXYZ))
        End Set
    End Property
#End Region

#End Region

#Region "Méthodes"


#End Region

#Region "Tests et debuggage"


#End Region

End Class
