Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports Utils

Public Class ClasseurExcel
    Inherits Entité
    Implements IEntitéDuWorkSpace

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

#End Region

#End Region

#Region "MêmeStructurePourTousLesBordereaux"
    Private _MêmeStructurePourTousLesBordereaux As Boolean
    Public Property MêmeStructurePourTousLesBordereaux() As Boolean
        Get
            Return _MêmeStructurePourTousLesBordereaux
        End Get
        Set(ByVal value As Boolean)
            If Object.Equals(value, Me._MêmeStructurePourTousLesBordereaux) Then Exit Property
            _MêmeStructurePourTousLesBordereaux = value
            NotifyPropertyChanged(NameOf(MêmeStructurePourTousLesBordereaux))
        End Set
    End Property
#End Region

#End Region

#Region "Méthodes"

    Public Function AjouterNouveauBordereau() As Bordereau
        Dim newB = New Bordereau()
        newB.Parent = Me
        Me.Bordereaux.Add(newB)
        Return newB
    End Function

    Private Sub _Bordereaux_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs) Handles _Bordereaux.CollectionChanged
        Me.NotifyPropertyChanged(NameOf(NbBordereaux))

        If e.NewItems IsNot Nothing Then
            For Each b As Bordereau In e.NewItems
                SABonnerAuxModificationsDuBordereau(b, True)
            Next
        End If

        If e.OldItems IsNot Nothing Then
            For Each b As Bordereau In e.OldItems
                SABonnerAuxModificationsDuBordereau(b, True)
            Next
        End If

    End Sub

#Region "SABonnerAuModificationsDuBordereau"

    Private Sub SABonnerAuxModificationsDuBordereau(b As Bordereau, OuiNon As Boolean)
        If OuiNon Then
            AddHandler b.PropertyChanged, BordereauPropertyChangedHandler
        Else
            RemoveHandler b.PropertyChanged, BordereauPropertyChangedHandler
        End If
    End Sub

    Private BordereauPropertyChangedHandler As PropertyChangedEventHandler = Sub(sender As Bordereau, e As PropertyChangedEventArgs) BordereauPropertyChanged(sender, e)

    Private Sub BordereauPropertyChanged(sender As Bordereau, e As PropertyChangedEventArgs)
        If Me.MêmeStructurePourTousLesBordereaux Then
            If e.PropertyName = NameOf(Bordereau.Paramètres) Then
                RecopierLesParamètresDuBordereau(sender)
            End If
        End If
    End Sub

    Private RecopieDesParamètresEnCours As Boolean

    Private Sub RecopierLesParamètresDuBordereau(Model As Bordereau)
        If RecopieDesParamètresEnCours Then Exit Sub
        Try

            RecopieDesParamètresEnCours = True
            Me.Bordereaux.DoForAll(Sub(b As Bordereau)
                                       If b IsNot Model Then
                                           With Model.Paramètres
                                               b.Paramètres.AdresseRangeLibelleOuvrage = .AdresseRangeLibelleOuvrage
                                               b.Paramètres.AdresseRangePrixUnitaire = .AdresseRangePrixUnitaire
                                               b.Paramètres.AdresseRangeUnité = .AdresseRangeUnité
                                           End With
                                       End If
                                   End Sub)
        Catch ex As Exception
            Throw ex
        Finally
            RecopieDesParamètresEnCours = False
        End Try
    End Sub

#End Region

#End Region

#Region "Tests et debuggage"


#End Region

End Class
