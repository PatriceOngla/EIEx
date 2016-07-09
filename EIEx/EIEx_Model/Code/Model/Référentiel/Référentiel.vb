Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports Model

''' <summary>Singleton.</summary>
Public Class Référentiel
    Inherits Système

#Region "Constructeurs"

    Private Sub New()
        MyBase.New()
    End Sub

    Protected Overrides Sub Init()
        MyBase.Init()

        _Produits = New ObservableCollection(Of Produit)
        _Tables.Add(_Produits)

        _FamillesDeProduit = New ObservableCollection(Of FamilleDeProduit)
        _Tables.Add(_FamillesDeProduit)

        _RéférencesDOuvrage = New ObservableCollection(Of RéférenceDOuvrage)
        _Tables.Add(_RéférencesDOuvrage)
    End Sub

#End Region

#Region "Propriétés"

#Region "Instance (Référentiel)"
    Private Shared _Instance As Référentiel
    Public Shared ReadOnly Property Instance() As Référentiel
        Get
            If _Instance Is Nothing Then _Instance = New Référentiel()
            Return _Instance
        End Get
    End Property
#End Region

#Region "Produits (ObservableCollection(Of Produit))"
    Private WithEvents _Produits As ObservableCollection(Of Produit)

    ''' <summary>Tous les <see cref="Produit"/>s du référentiel.</summary>
    Public ReadOnly Property Produits() As ObservableCollection(Of Produit)
        Get
            Return _Produits
        End Get
    End Property
#End Region

#Region "FamillesDeProduit (ObservableCollection(Of FamilleDeProduit)"
    Private _FamillesDeProduit As ObservableCollection(Of FamilleDeProduit)
    ''' <summary>Tous les <see cref="FamilleDeProduit"/>t du référentiel.</summary>
    Public ReadOnly Property FamillesDeProduit() As ObservableCollection(Of FamilleDeProduit)
        Get
            Return _FamillesDeProduit
        End Get
    End Property
#End Region

#Region "RéférencesDOuvrage (ObservableCollection(Of RéférenceDOuvrage))"
    Private _RéférencesDOuvrage As ObservableCollection(Of RéférenceDOuvrage)
    ''' <summary>Toutes les <see cref="RéférenceDOuvrage"/> du référentiel.</summary>
    Public ReadOnly Property RéférencesDOuvrage() As ObservableCollection(Of RéférenceDOuvrage)
        Get
            Return _RéférencesDOuvrage
        End Get
    End Property
#End Region

#End Region

#Region "Méthodes"

#Region "Persistance"

    Public Overrides Sub Charger(Chemin As String)
        Dim WS_DAO = Utils.DéSérialisation(Of Référentiel_DAO)(Chemin)
        WS_DAO.UnSerialize(Me)
    End Sub

#End Region

#Region "Plomberie"

    Protected Overrides Function GetTable(Of Tr As AgregateRoot_Base)() As IList(Of Tr)
        Dim r As IList(Of Tr)
        Dim leType = GetType(Tr)
        Select Case leType
            Case GetType(Produit)
                r = Me.Produits
            Case GetType(RéférenceDOuvrage)
                r = Me.RéférencesDOuvrage
            Case GetType(FamilleDeProduit)
                r = Me.FamillesDeProduit
            Case Else
                Throw New InvalidOperationException($"Le référentiel ne gère pas le type ""{leType.Name}"".")
        End Select
        Return r
    End Function

#End Region

#Region "Factory"

#Region "Produit"

    Public Function GetNewProduit(newId As Integer) As Produit
        Dim r = New Produit(newId)
        Me.Produits.Add(r)
        Return r
    End Function

    Public Function GetNewProduit() As Produit
        Dim newId = GetNewId(Of Produit)()
        Dim r = GetNewProduit(newId)
        Return r
    End Function

#End Region

#Region "RéférenceDOuvrage"

    Public Function GetNewRéférenceDOuvrage(newId As Integer) As RéférenceDOuvrage
        Dim r = New RéférenceDOuvrage(newId)
        Me.RéférencesDOuvrage.Add(r)
        Return r
    End Function

    Public Function GetNewRéférenceDOuvrage() As RéférenceDOuvrage
        Dim newId = GetNewId(Of RéférenceDOuvrage)()
        Dim r = GetNewRéférenceDOuvrage(newId)
        Return r
    End Function

#End Region

#Region "FamilleDeProduit"

    Public Function GetNewFamilleDeProduit(newId As Integer) As FamilleDeProduit
        Dim r = New FamilleDeProduit(newId)
        Me.FamillesDeProduit.Add(r)
        Return r
    End Function

    Public Function GetNewFamilleDeProduit() As FamilleDeProduit
        Dim newId = Me.GetNewId(Of FamilleDeProduit)
        Dim r = GetNewFamilleDeProduit(newId)
        Return r
    End Function

#End Region

#End Region

#Region "Accès aux objets"

    Public Function GetProduitById(id As Integer) As Produit
        Dim r = GetObjectById(Of Produit)(id)
        Return r
    End Function

    Public Function GetFamilleById(id As Integer) As FamilleDeProduit
        Dim r = GetObjectById(Of FamilleDeProduit)(id)
        Return r
    End Function

    Public Function GetRéférenceDOuvrageById(id As Integer) As RéférenceDOuvrage
        Dim r = GetObjectById(Of RéférenceDOuvrage)(id)
        Return r
    End Function

#End Region

#End Region

#Region "Tests et debuggage"

    Private Sub CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs) Handles _Produits.CollectionChanged

    End Sub

#End Region

End Class
