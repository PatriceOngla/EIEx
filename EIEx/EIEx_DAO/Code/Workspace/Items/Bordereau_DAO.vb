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

    Protected Overrides Sub UnSerialized_Ex(NouveauBordereau As Bordereau)
        'Dim r = Me.Parent.AjouterNouveauBordereau()
        NouveauBordereau.NomFeuille = Me.NomFeuille
        Me.Paramètres.UnSerialized(NouveauBordereau.Paramètres)
    End Sub

#End Region

#Region "Tests et debuggage"


#End Region

End Class