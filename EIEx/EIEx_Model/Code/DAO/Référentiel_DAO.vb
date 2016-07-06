
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
    Public Property DateModif() As Date

    Public Property Produits() As List(Of Produit_DAO)

    Public Property FamillesDeProduit() As List(Of FamilleDeProduit_DAO)

    Public Property RéférencesDOuvrage() As List(Of RéférenceDOuvrage_DAO)

#End Region

#Region "Méthodes"

    Public Overrides Function UnSerialized_Ex() As Référentiel
        Dim r As New Référentiel
        'r.Produit = Me.Produit.UnSerialized
        'r.Nombre = Me.Nombre
        'Dim UsagesDeProduit = From up In Me.UsagesDeProduit Select up.UnSerialized()

        r.DateModif = Me.DateModif

        Dim Produits = From p In Me.Produits Select p.UnSerialized()
        r.Produits.AddRange(Produits)

        Dim Familles = From f In Me.FamillesDeProduit Select f.UnSerialized()
        r.FamillesDeProduit.AddRange(Familles)

        Dim RéférencesDOuvrage = From ro In Me.RéférencesDOuvrage Select ro.UnSerialized()
        r.RéférencesDOuvrage.AddRange(RéférencesDOuvrage)

        Return r

    End Function


#End Region

#Region "Tests et debuggage"


#End Region

End Class
