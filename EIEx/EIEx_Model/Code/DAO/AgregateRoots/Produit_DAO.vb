Imports System.Xml.Serialization
Imports Utils

<Serializable>
Public Class Produit_DAO
    Inherits AgregateRoot_DAO(Of Produit)

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(P As Produit)
        MyBase.New(P)
        Me.Unité = P.Unité
        Me.Prix = P.Prix
        Me.ReférenceFournisseur = P.ReférenceFournisseur
        Me.TempsDePauseUnitaire = P.TempsDePauseUnitaire
        Me.MotsClés = New List(Of String)(P.MotsClés)
        If P.Famille IsNot Nothing Then Me.FamilleId = P.Famille.Id
    End Sub

#End Region

#Region "Propriétés"

    Public Property Unité() As Unités?

    Public Property Prix() As Single?

    <XmlAttribute>
    Public Property ReférenceFournisseur() As String

    Public Property TempsDePauseUnitaire() As Integer?

    Public Property MotsClés() As List(Of String)

    Public Property FamilleId() As Integer

#End Region

#Region "Méthodes"

    Protected Overrides Function UnSerialized_Ex_Ex() As Produit
        'Dim r = Référentiel.Instance.GetNewProduit(Me.Id)
        Dim r = Réf.GetProduitById(Me.Id)
        r = If(r, New Produit(Me.Id))
        r.Unité = Unité
        r.Prix = Prix
        r.ReférenceFournisseur = ReférenceFournisseur
        r.TempsDePauseUnitaire = TempsDePauseUnitaire
        r.MotsClés.AddRange(MotsClés)
        r.Famille = Réf.GetFamilleById(Me.FamilleId)
        Return r
    End Function

#End Region

#Region "Tests et debuggage"

#End Region

End Class
