Imports System.Xml.Serialization

<Serializable>
Public Class UsageDeProduit_DAO
    Inherits SystèmesItems_DAO(Of UsageDeProduit)

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(UP As UsageDeProduit)
        MyBase.New(UP)
        Try
            'Me.ParentId = UP.Parent?.Id
            Me.ProduitId = UP.Produit?.Id
            Me.Nombre = UP.Nombre
        Catch ex As Exception
            Dim ex2 = New Exception($"Echec de la sérialisation d'un {NameOf(UsageDeProduit)} du produit n° {Me.ProduitId}.", ex)
        End Try
    End Sub

#End Region

#Region "Propriétés"

#Region "Sys"
    Private Ref As Référentiel = Référentiel.Instance
    <XmlIgnore>
    Protected Overrides ReadOnly Property Sys As Système
        Get
            Return Ref
        End Get
    End Property
#End Region

#Region "Données"

    'Public Property ParentId As Integer?

    Public Property ProduitId() As Integer?

    <XmlAttribute>
    Public Property Nombre() As Integer

#End Region

#End Region

#Region "Méthodes"

    Protected Overrides Sub UnSerialized_Ex(NouvelUsageProduit As UsageDeProduit)
        NouvelUsageProduit.Nombre = Me.Nombre 'normalement c'est déjà fait au dessus dans AjouterProduit
    End Sub

#End Region

#Region "Tests et debuggage"


#End Region

End Class
