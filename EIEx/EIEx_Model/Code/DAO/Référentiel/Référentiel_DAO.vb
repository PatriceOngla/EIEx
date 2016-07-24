Imports System.Xml.Serialization
Imports Model
Imports Utils

<Serializable>
Public Class Référentiel_DAO
    Inherits Système_DAO(Of Référentiel)

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

        Dim Ouvrages_DAO = From ro In R.Ouvrage Select New Ouvrage_DAO(ro)
        Me.Ouvrages = New List(Of Ouvrage_DAO)(Ouvrages_DAO)

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

    Public Property Produits() As List(Of Produit_DAO)

    Public Property FamillesDeProduit() As List(Of FamilleDeProduit_DAO)

    Public Property Ouvrages() As List(Of Ouvrage_DAO)

#End Region

#End Region

#Region "Méthodes"

    Protected Overrides Sub UnSerialize_Ex(NewT As Référentiel)
        Dim r = Référentiel.Instance

        r.DateModif = Me.DateModif

        Me.FamillesDeProduit.DoForAll(Sub(f)
                                          Dim NewFamille = f.UnSerialized()
                                      End Sub)

        Me.Produits.DoForAll(Sub(p)
                                 Dim NewProduit = p.UnSerialized()
                             End Sub)

        Me.Ouvrages.DoForAll(Sub(ro)
                                 Dim NewOuvrage = ro.UnSerialized()
                             End Sub)


    End Sub

#End Region

#Region "Tests et debuggage"


#End Region

End Class
