Imports System.Xml.Serialization
Imports Utils

<Serializable>
Public Class Référentiel_DAO
    Inherits EIEx_Object_DAO(Of Référentiel)

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(R As Référentiel)
        MyBase.New(R)

        Me.DateModif = R.DateModif

        Dim Produits_DAO = From p In R.Produits Select New Produit_DAO(p)
        Me.Produits = New List(Of Produit_DAO)(Produits_DAO)

        Dim FamillesDeProduit_DAO = From f In R.FamillesDeProduit Select New FamilleDeProduit_DAO(f)
        Me.FamillesDeProduit = New List(Of FamilleDeProduit_DAO)(FamillesDeProduit_DAO)

        Dim RéférencesDOuvrage_DAO = From ro In R.RéférencesDOuvrage Select New RéférenceDOuvrage_DAO(ro)
        Me.RéférencesDOuvrage = New List(Of RéférenceDOuvrage_DAO)(RéférencesDOuvrage_DAO)

    End Sub

#End Region

#Region "Propriétés"
    <XmlAttribute>
    Public Property DateModif() As Date

    Public Property Produits() As List(Of Produit_DAO)

    Public Property FamillesDeProduit() As List(Of FamilleDeProduit_DAO)

    Public Property RéférencesDOuvrage() As List(Of RéférenceDOuvrage_DAO)

#End Region

#Region "Méthodes"

    Protected Overrides Function UnSerialized_Ex() As Référentiel
        'Dim r As New Référentiel
        Dim r = Référentiel.Instance
        r.Purger

        r.DateModif = Me.DateModif

        'Les objets s'enregistrent dans le référentiel dans leur constructeur. 
        Dim Produits = (From p In Me.Produits Select p.UnSerialized()).OfType(Of Produit)
        Produits.DoForAll(Sub(p As Produit) Réf.EnregistrerRoot(p))
        'r.Produits.AddRange(Produits)

        Dim Familles = (From f In Me.FamillesDeProduit Select f.UnSerialized()).OfType(Of FamilleDeProduit)
        Familles.DoForAll(Sub(f As FamilleDeProduit) Réf.EnregistrerRoot(f))
        'r.FamillesDeProduit.AddRange(Familles)

        Dim RéférencesDOuvrage = (From ro In Me.RéférencesDOuvrage Select ro.UnSerialized()).OfType(Of RéférenceDOuvrage)
        RéférencesDOuvrage.DoForAll(Sub(ro As RéférenceDOuvrage) Réf.EnregistrerRoot(ro))
        'r.RéférencesDOuvrage.AddRange(RéférencesDOuvrage)

        Return r

    End Function


#End Region

#Region "Tests et debuggage"


#End Region

End Class
