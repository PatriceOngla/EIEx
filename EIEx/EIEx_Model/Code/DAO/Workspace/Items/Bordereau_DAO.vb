Imports System.Xml.Serialization
Imports Model

<Serializable>
Public Class Bordereau_DAO
    Inherits SystèmesItems_DAO(Of Bordereau)

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(B As Bordereau)
        MyBase.New(B)
        Me.NomFeuille = B.NomFeuille
        Me.Paramètres = New Paramètres_DAO(B.Paramètres)
    End Sub

#End Region

#Region "Propriétés"

    <XmlIgnore>
    Protected Overrides ReadOnly Property Sys As Système
        Get
            Return WorkSpace.Instance
        End Get
    End Property

    Public Property NomFeuille As String

    Public Property Paramètres() As Paramètres_DAO

#End Region

#Region "Méthodes"

    Protected Overrides Function UnSerialized_Ex() As Bordereau
        Dim r As New Bordereau()
        r.NomFeuille = Me.NomFeuille
        r.Paramètres.AdresseRangeLibelleOuvrage = Me.Paramètres.AdresseRangeLibelleOuvrage
        r.Paramètres.AdresseRangePrixUnitaire = Me.Paramètres.AdresseRangePrixUnitaire
        r.Paramètres.AdresseRangeUnité = Me.Paramètres.AdresseRangeUnité
        Return r
    End Function

#End Region

#Region "Tests et debuggage"


#End Region

End Class