Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports Model

''' <summary>Singleton.</summary>
Public Class Référentiel
    Inherits EIExObject

#Region "Constructeurs"

    Private Sub New()
        MyBase.New()
    End Sub

    Protected Overrides Sub Init()
        _Produits = New ObservableCollection(Of Produit)
        _FamillesDeProduit = New ObservableCollection(Of FamilleDeProduit)
        _RéférencesDOuvrage = New ObservableCollection(Of RéférenceDOuvrage)
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

#Region "DateModif"
    Private _DateModif As Date
    Public Property DateModif() As Date
        Get
            Return _DateModif
        End Get
        Set(ByVal value As Date)
            If Object.Equals(value, Me._DateModif) Then Exit Property
            _DateModif = value
            NotifyPropertyChanged(NameOf(DateModif))
        End Set
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

#Region "Sérialisation"

    ''' <summary>Peuple le référentiel à partir du fichier de persistance <see cref="EIExData.CheminRéférentiel"/>.</summary>
    Public Shared Sub Charger(Chemin As String)
        Dim r_DAO = Utils.DéSérialisation(Of Référentiel_DAO)(Chemin)
        _Instance = r_DAO.UnSerialized()
    End Sub

    Public Shared Sub Enregistrer(Chemin As String)
        Instance.DateModif = Now()
        Dim DAO = New Référentiel_DAO(Instance)
        Utils.Sérialiser(DAO, Chemin)
    End Sub

#End Region

#Region "Factory"

#Region "Factory générique"
    Friend Sub EnregistrerRoot(Of T As AgregateRoot)(root As T)
        Dim Table = GetTable(Of T)()
        If Not Table.Contains(root) Then Table.Add(root)
    End Sub


    'Public Function GetNewRoot(Of T As {AgregateRoot, New})(newId As Integer) As T
    '    Dim r = New T()
    '    EnregistrerRoot(r)
    '    Return r
    'End Function

    Public Function GetNewRoot(Of T As {AgregateRoot, New})() As T
        Dim r = New T()
        EnregistrerRoot(r)
        Return r
    End Function

#End Region


#Region "Produit"

    'Public Function GetNewProduit(newId As Integer) As Produit
    '    Dim r = New Produit(newId)
    '    Me.Produits.Add(r)
    '    Return r
    'End Function

    'Public Function GetNewProduit() As Produit
    '    Dim newId = (From p In Me.Produits Select p.Id).Max
    '    newId = If(newId Is Nothing, 0, newId + 1)
    '    Dim r = GetNewProduit(newId)
    '    Return r
    'End Function

#End Region

#Region "RéférenceDOuvrage"

    'Public Function GetNewRéférenceDOuvrage(newId As Integer) As RéférenceDOuvrage
    '    Dim r = New RéférenceDOuvrage(newId)
    '    Me.RéférencesDOuvrage.Add(r)
    '    Return r
    'End Function

    'Public Function GetNewRéférenceDOuvrage() As RéférenceDOuvrage
    '    Dim newId = (From p In Me.RéférencesDOuvrage Select p.Id).Max
    '    newId = If(newId Is Nothing, 0, newId + 1)
    '    Dim r = GetNewRéférenceDOuvrage(newId)
    '    Return r
    'End Function

#End Region

#Region "FamilleDeProduit"

    'Public Function GetNewFamilleDeProduit(newId As Integer) As FamilleDeProduit
    '    Dim r = New FamilleDeProduit(newId)
    '    Me.FamillesDeProduit.Add(r)
    '    Return r
    'End Function

    'Public Function GetNewFamilleDeProduit() As FamilleDeProduit
    '    Dim newId = Me.GetNewId(Of FamilleDeProduit)
    '    newId = If(newId Is Nothing, 0, newId + 1)
    '    Dim r = GetNewFamilleDeProduit(newId)
    '    Return r
    'End Function

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

    Private Function GetObjectById(Of T As AgregateRoot)(id As Integer) As T
        Dim Table As ObservableCollection(Of T) = GetTable(Of T)()
        Dim r = (Table.Where(Function(o) o.Id = id)).FirstOrDefault
        Return r
    End Function

#End Region

#Region "Divers"

    Private Function GetTable(Of T As EIExObject)() As ObservableCollection(Of T)
        Dim r As IEnumerable(Of T)
        Dim leType = GetType(T)
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

    Friend Function GetNewId(Of T As AgregateRoot)() As Integer?
        Dim Table As ObservableCollection(Of T) = GetTable(Of T)()
        Dim newId = (From p In Table Select p.Id).Max
        newId = If(newId, 0)
        newId += 1
        Return newId
    End Function

#Region "Purger"

    Public Sub Purger()
        Me.Produits.Clear()
        Me.FamillesDeProduit.Clear()
        Me.RéférencesDOuvrage.Clear()
    End Sub

#End Region

#Region "EstVide"
    Public Function EstVide() As Boolean
        Dim r = Me.Produits.Count = 0
        r = r And Me.FamillesDeProduit.Count = 0
        r = r And Me.RéférencesDOuvrage.Count = 0
        Return r
    End Function
#End Region

#End Region

#End Region

#Region "Tests et debuggage"

    Private Sub CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs) Handles _Produits.CollectionChanged

    End Sub

#End Region

End Class
