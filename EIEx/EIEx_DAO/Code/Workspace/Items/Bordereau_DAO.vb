Imports System.Xml.Serialization
Imports Model
Imports Utils

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
        Dim Ouvrages = From o In B.Ouvrages Select New Ouvrage_DAO(o)
        Me.Ouvrages = New List(Of Ouvrage_DAO)(Ouvrages)
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

    Public Property Ouvrages() As List(Of Ouvrage_DAO)


#End Region

#Region "Méthodes"

    Protected Overrides Sub UnSerialized_Ex(NouveauBordereau As Bordereau)
        NouveauBordereau.NomFeuille = Me.NomFeuille
        Me.Paramètres.UnSerialized(NouveauBordereau.Paramètres)
        Me.Ouvrages.DoForAll(Sub(o As Ouvrage_DAO)
                                 Dim NvlOuvrage = NouveauBordereau.AjouterOuvrage(o.NuméroLignePlageExcel)
                                 o.UnSerialized(NvlOuvrage)
                             End Sub)
    End Sub

#End Region

#Region "Tests et debuggage"


#End Region

End Class