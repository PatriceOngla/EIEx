Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports Model

Public Class Etude
    Inherits Entité
    Implements IAgregateRoot, IEntitéDuWorkSpace

#Region "Constructeurs"

    Public Sub New(Id As Integer)
        Me.Id = Id
    End Sub

    Protected Overrides Sub Init()
        Me.Nom = "Nouvelle étude"
        _ClasseursExcel = New ObservableCollection(Of ClasseurExcel)
    End Sub

#End Region

#Region "Propriétés"

#Region "WS"
    Public ReadOnly Property WS As WorkSpace Implements IEntitéDuWorkSpace.WS
        Get
            Return WorkSpace.Instance
        End Get
    End Property
#End Region

#Region "Système"
    Public Overrides ReadOnly Property Système As Système
        Get
            Return Me.WS
        End Get
    End Property
#End Region

#Region "Id"
    Public ReadOnly Property Id() As Integer? Implements IAgregateRoot.Id
#End Region

#Region "Client"
    Private _Client As String
    Public Property Client() As String
        Get
            Return _Client
        End Get
        Set(ByVal value As String)
            If Object.Equals(value, Me._Client) Then Exit Property
            _Client = value
            NotifyPropertyChanged(NameOf(Client))
        End Set
    End Property
#End Region

#Region "EstOuverte"

    Private _EstOuverte As Boolean = False

    ''' <summary>
    ''' Indique qu'il s'agit de l'étude courante (cf. <see cref="WorkSpace.EtudeCourante"/>) . Il ne peut y en avoir qu'une. 
    ''' </summary>
    ''' <returns></returns>
    Public Property EstOuverte() As Boolean
        Get
            Return _EstOuverte
        End Get
        Set(ByVal value As Boolean)
            If Object.Equals(value, Me._EstOuverte) Then Exit Property
            _EstOuverte = value
            If value Then WS.EtudeCourante = Me
            NotifyPropertyChanged(NameOf(EstOuverte))
        End Set
    End Property

#End Region

#Region "ClasseursExcel"
    Private WithEvents _ClasseursExcel As ObservableCollection(Of ClasseurExcel)
    Public ReadOnly Property ClasseursExcel() As ObservableCollection(Of ClasseurExcel)
        Get
            Return _ClasseursExcel
        End Get
    End Property

#Region "NbClasseursExcel"

    Public ReadOnly Property NbClasseursExcel() As Integer
        Get
            Return Me.ClasseursExcel.Count()
        End Get
    End Property

    Private Sub _ClasseursExcel_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs) Handles _ClasseursExcel.CollectionChanged
        Me.NotifyPropertyChanged(NameOf(NbClasseursExcel))
    End Sub

#End Region

#End Region

#Region "Ouvrages"
    Public ReadOnly Property Ouvrages As IEnumerable(Of Ouvrage)
        Get
            Dim r = From c In Me.ClasseursExcel From b In c.Bordereaux From o In b.Ouvrages Select o
            Return r
        End Get
    End Property

#Region "NbOuvrages"

    Public ReadOnly Property NbOuvrages() As Integer
        Get
            Return Me.Ouvrages.Count()
        End Get
    End Property

#End Region

#End Region

#End Region

#Region "Méthodes"

    Public Function AjouterNouveauClasseur() As ClasseurExcel
        Return AjouterNouveauClasseur(Nothing)
    End Function

    Public Function AjouterNouveauClasseur(Chemin As String) As ClasseurExcel
        Dim newC = New ClasseurExcel(Chemin)
        Me.ClasseursExcel.Add(newC)
        Return newC
    End Function

    Public Overrides Function ToString() As String
        Return Me.ToStringForAgregateRoot(MyBase.ToString())
    End Function

#End Region

#Region "Tests et debuggage"


#End Region

End Class
