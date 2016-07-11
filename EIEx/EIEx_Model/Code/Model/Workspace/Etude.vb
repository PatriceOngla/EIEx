Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel

Public Class Etude
    Inherits AgregateRoot(Of Etude)

#Region "Constructeurs"
    Public Sub New(Id As Integer)
        MyBase.New(Id)
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
        Dim newB = New Bordereau
        Me.Bordereaux.Add(newB)
        Return newB
    End Function

#End Region

#Region "Tests et debuggage"


#End Region

End Class
