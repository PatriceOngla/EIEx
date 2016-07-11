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

        Me.Nom = "Référentiel"

        _Produits = New ObservableCollection(Of Produit)
        _Tables.Add(_Produits)

        _FamillesDeProduit = New ObservableCollection(Of FamilleDeProduit)
        _Tables.Add(_FamillesDeProduit)

        _PatronsDOuvrage = New ObservableCollection(Of PatronDOuvrage)
        _Tables.Add(_PatronsDOuvrage)
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

#Region "PatronsDOuvrage (ObservableCollection(Of PatronDOuvrage))"
    Private _PatronsDOuvrage As ObservableCollection(Of PatronDOuvrage)
    ''' <summary>Toutes les <see cref="PatronDOuvrage"/> du référentiel.</summary>
    Public ReadOnly Property PatronsDOuvrage() As ObservableCollection(Of PatronDOuvrage)
        Get
            Return _PatronsDOuvrage
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
            Case GetType(PatronDOuvrage)
                r = Me.PatronsDOuvrage
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

    Public Function CheckUnicityRefProduit(RefProduit As String) As Boolean
        Try
            If LaRéfProduitExisteDéjà(RefProduit) Then
                Throw New InvalidOperationException($"Un produit portant la référence ""{RefProduit}"" existe déjà. {vbCr}Modification ignorée.")
            End If
            Return True
        Catch ex As Exception
            Me.RaiseExceptionRaisedEvent(ex, True)
            Return False
        End Try
    End Function

    Public Function LaRéfProduitExisteDéjà(RéfProduit As String) As Boolean
        Dim r = (From p In Me.Produits Where p.RéférenceProduit.Equals(RéfProduit)).Any()
        Return r
    End Function

#End Region

#Region "PatronDOuvrage"

    Public Function GetNewPatronDOuvrage(newId As Integer) As PatronDOuvrage
        Dim r = New PatronDOuvrage(newId)
        Me.PatronsDOuvrage.Add(r)
        Return r
    End Function

    Public Function GetNewPatronDOuvrage() As PatronDOuvrage
        Dim newId = GetNewId(Of PatronDOuvrage)()
        Dim r = GetNewPatronDOuvrage(newId)
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

    Public Function GetProduitById(id As Integer, Optional FailIfNotFound As Boolean = False) As Produit
        Dim r = GetObjectById(Of Produit)(id, FailIfNotFound)
        Return r
    End Function

    Public Function GetFamilleById(id As Integer, Optional FailIfNotFound As Boolean = False) As FamilleDeProduit
        Dim r = GetObjectById(Of FamilleDeProduit)(id, FailIfNotFound)
        Return r
    End Function

    Public Function GetPatronDOuvrageById(id As Integer, Optional FailIfNotFound As Boolean = False) As PatronDOuvrage
        Dim r = GetObjectById(Of PatronDOuvrage)(id, FailIfNotFound)
        Return r
    End Function

#End Region

#End Region

#Region "Tests et debuggage"

    Private Sub CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs) Handles _Produits.CollectionChanged

    End Sub

#End Region

End Class
