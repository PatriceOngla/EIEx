Imports System.Xml.Serialization

<Serializable>
Public Class Paramètres_DAO
    Inherits SystèmesItems_DAO(Of Paramètres)

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(P As Paramètres)
        MyBase.New(P)
        Me.AdresseRangeLibelleOuvrage = P.AdresseRangeLibelleOuvrage
        Me.AdresseRangePrixUnitaire = P.AdresseRangePrixUnitaire
        Me.AdresseRangeUnité = P.AdresseRangeUnité
    End Sub

#End Region

#Region "Propriétés"

#Region "Sys"
    <XmlIgnore>
    Protected Overrides ReadOnly Property Sys As Système
        Get
            Return WorkSpace.Instance
        End Get
    End Property
#End Region

#Region "Données"

    Public Property AdresseRangeLibelleOuvrage() As String

    Public Property AdresseRangeUnité() As String

    Public Property AdresseRangePrixUnitaire() As String

#End Region

#End Region

#Region "Méthodes"

    Protected Overrides Sub UnSerialized_Ex(NouveauxParamètres As Paramètres)
        NouveauxParamètres.AdresseRangeLibelleOuvrage = AdresseRangeLibelleOuvrage
        NouveauxParamètres.AdresseRangePrixUnitaire = AdresseRangePrixUnitaire
        NouveauxParamètres.AdresseRangeUnité = AdresseRangeUnité
    End Sub

#End Region

#Region "Tests et debuggage"


#End Region

End Class
