Imports System.Xml.Serialization

<Serializable>
Public Class UsageDeProduit_DAO
    Inherits EIEx_Object_DAO(Of UsageDeProduit)

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(UP As UsageDeProduit)
        MyBase.New(UP)
        Me.ParentId = UP.Parent.Id
        Me.ProduitId = UP.Produit.Id
        Me.Nombre = UP.Nombre
    End Sub

#End Region

#Region "Propriétés"
    <XmlAttribute>
    Public Property ParentId As Integer

    Public Property ProduitId() As Integer

    <XmlAttribute>
    Public Property Nombre() As Integer

#End Region

#Region "Méthodes"

    Protected Overrides Function UnSerialized_Ex() As UsageDeProduit
        Dim Parent = Réf.GetRéférenceDOuvrageById(Me.ParentId)
        Dim r As New UsageDeProduit(Parent)
        r.Produit = Réf.GetProduitById(Me.ProduitId)
        r.Nombre = Me.Nombre
        'Dim UsagesDeProduit = From up In Me.UsagesDeProduit Select up.UnSerialized()
        Return r
    End Function

#End Region

#Region "Tests et debuggage"


#End Region

End Class
