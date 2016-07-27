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

#Region "EstOuverte"

    Private _EstOuverte As Boolean = False

    Public Property EstOuverte() As Boolean
        Get
            Return _EstOuverte
        End Get
        Set(ByVal value As Boolean)
            If Object.Equals(value, Me._EstOuverte) Then Exit Property
            _EstOuverte = value
            NotifyPropertyChanged(NameOf(EstOuverte))
            ManageEstOuverteChanged()
        End Set
    End Property

    Private Sub ManageEstOuverteChanged()
        If Me.EstOuverte Then
            Dim WS = WorkSpace.Instance
            Dim EC = WS.EtudeCourante
            If EC IsNot Nothing AndAlso EC IsNot Me Then
                EC.EstOuverte = False
            End If
            If EC IsNot Me Then WS.EtudeCourante = Me
        End If
    End Sub

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
