Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel

Public Class ClasseurExcel
    Inherits EntitéDuWorkSpace

#Region "Constructeurs"

    Friend Sub New(FullName As String)
        If Not String.IsNullOrEmpty(FullName) Then
            Me.Nom = IO.Path.GetFileName(FullName)
            Me.CheminFichier = FullName
        Else
            Me.Nom = "Nouveau classeur"
        End If
    End Sub

    Protected Overrides Sub Init()
        Me.Nom = "Nouvelle étude"
        _Bordereaux = New ObservableCollection(Of Bordereau)
    End Sub

#End Region

#Region "Propriétés"

#Region "Système"
    Public Overrides ReadOnly Property Système As Système
        Get
            Return WorkSpace.Instance
        End Get
    End Property
#End Region

#Region "CheminFichier"
    Private _CheminFichier As String
    Public Property CheminFichier() As String
        Get
            Return _CheminFichier
        End Get
        Set(ByVal value As String)
            If Object.Equals(value, Me._CheminFichier) Then Exit Property
            _CheminFichier = value
            NotifyPropertyChanged(NameOf(CheminFichier))
        End Set
    End Property
#End Region

#Region "Bordereaux"
    Private WithEvents _Bordereaux As ObservableCollection(Of Bordereau)
    Public ReadOnly Property Bordereaux() As ObservableCollection(Of Bordereau)
        Get
            Return _Bordereaux
        End Get
    End Property

#Region "NbBordereaux"

    Public ReadOnly Property NbBordereaux() As Integer
        Get
            Return Me.Bordereaux.Count()
        End Get
    End Property

    Private Sub _Bordereaux_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs) Handles _Bordereaux.CollectionChanged
        Me.NotifyPropertyChanged(NameOf(NbBordereaux))
    End Sub

#End Region

#End Region

#End Region

#Region "Méthodes"

    Public Function AjouterNouveauBordereau() As Bordereau
        Dim newB = New Bordereau()
        newB.Parent = Me
        Me.Bordereaux.Add(newB)
        Return newB
    End Function

#End Region

#Region "Tests et debuggage"


#End Region

End Class
