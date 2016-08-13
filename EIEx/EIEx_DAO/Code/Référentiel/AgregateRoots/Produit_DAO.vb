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
        Me.CodeLydic = P.CodeLydic
        Me.RéférenceFournisseur = P.RéférenceFournisseur
        Me.TempsDePoseUnitaire = P.TempsDePoseUnitaire
        Me.MotsClés = New List(Of String)(P.MotsClés)
        If P.Famille IsNot Nothing Then Me.FamilleId = P.Famille.Id
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

    <XmlAttribute>
    Public Property Unité() As Unités

    <XmlAttribute>
    Public Property Prix() As Single

    <XmlAttribute>
    Public Property CodeLydic() As String

    <XmlAttribute>
    Public Property RéférenceFournisseur() As String

    <XmlAttribute>
    Public Property TempsDePoseUnitaire() As Single

    Public Property MotsClés() As List(Of String)

    Public Property FamilleId() As Integer

#End Region

#End Region

#Region "Méthodes"

    Protected Overrides Sub UnSerialized_Ex_Ex(NouveauProduit As Produit)
        NouveauProduit.Unité = Unité
        NouveauProduit.Prix = Prix
        NouveauProduit.CodeLydic = CodeLydic
        NouveauProduit.RéférenceFournisseur = RéférenceFournisseur
        NouveauProduit.TempsDePoseUnitaire = TempsDePoseUnitaire
        NouveauProduit.MotsClés.AddRange(MotsClés)
        NouveauProduit.Famille = Ref.GetFamilleById(Me.FamilleId)
    End Sub

#End Region

#Region "Tests et debuggage"

#End Region

End Class
