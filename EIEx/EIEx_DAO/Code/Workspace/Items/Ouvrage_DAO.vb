Imports System.Xml.Serialization
Imports Utils

<Serializable>
Public Class Ouvrage_DAO
    Inherits SystèmesItems_DAO(Of Ouvrage)
    'TODO : factoriser le code entre Ouvrage_DAO et PatronDOuvrage_DAO (commun sauf complément de nom)

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(O As Ouvrage)
        MyBase.New(O)

        Me.ComplémentDeNom = O.ComplémentDeNom

        Me.Libellés = New List(Of String)(O.Libellés)

        Me.NuméroLignePlageExcel = O.NuméroLignePlageExcel

        Dim UsagesDeProduit_DAO = From up In O.UsagesDeProduit Select New UsageDeProduit_DAO(up)
        Me.UsagesDeProduit = New List(Of UsageDeProduit_DAO)(UsagesDeProduit_DAO)

        Me.MotsClés = New List(Of String)(O.MotsClés)

        If O.TempsDePoseForcé Then Me.TempsDePoseUnitaire = O.TempsDePoseUnitaire

        If O.PrixUnitaireForcé Then Me.PrixUnitaire = O.PrixUnitaire

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

    Public Property Libellés() As List(Of String)

    Public Property ComplémentDeNom() As String

    Public Property NuméroLignePlageExcel() As Integer

    Public Property UsagesDeProduit() As List(Of UsageDeProduit_DAO)

    Public Property MotsClés() As List(Of String)

    Public Property TempsDePoseUnitaire() As Integer?

    Public Property PrixUnitaire() As Single?

#End Region

#End Region

#Region "Méthodes"

    Protected Overrides Sub UnSerialized_Ex(NouvelOuvrage As Ouvrage)

        NouvelOuvrage.ComplémentDeNom = Me.ComplémentDeNom

        NouvelOuvrage.Libellés.AddRange(Me.Libellés)
        'NouvelOuvrage.NuméroLignePlageExcel = Me.NuméroLignePlageExcel

        Me.UsagesDeProduit.DoForAll(Sub(up As UsageDeProduit_DAO)
                                        Dim Produit = If(up.ProduitId Is Nothing, Nothing, Ref.GetProduitById(up.ProduitId))
                                        NouvelOuvrage.AjouterProduit(Produit, up.Nombre)
                                    End Sub)

        NouvelOuvrage.MotsClés.AddRange(Me.MotsClés)
        NouvelOuvrage.TempsDePoseUnitaire = TempsDePoseUnitaire
        NouvelOuvrage.PrixUnitaire = Me.PrixUnitaire

    End Sub

#End Region

#Region "Tests et debuggage"


#End Region

End Class
