Imports System.Collections.ObjectModel
Imports System.Collections.Specialized

''' <summary>
''' On distingue <see cref="RéférenceDOuvrage"/> et Ouvrage (pas encore implémenté). Les ouvrages sont les entrées des bordereau et sont associé à des <see cref="RéférenceDOuvrage"/> afin de calculer leur prix sur la base du <see cref="RéférenceDOuvrage.PrixUnitaire"/>. 
''' </summary>
Public Class RéférenceDOuvrage
    Inherits EIExObject

#Region "Constructeurs"
    Public Sub New()
        _Libellés = New ObservableCollection(Of String)
        _UsagesDeProduit = New ObservableCollection(Of UsageDeProduit)
        _MotsClés = New ObservableCollection(Of String)
    End Sub

    Public Sub New(Id As Integer)
        MyBase.New(Id)
    End Sub

#End Region

#Region "Propriétés"

#Region "LibelléPrincipal (String)"
    Private _LibelléPrincipal As String
    Public Property LibelléPrincipal() As String
        Get
            Return _LibelléPrincipal
        End Get
        Set(ByVal value As String)
            If Object.Equals(value, Me._LibelléPrincipal) Then Exit Property
            _LibelléPrincipal = value
            NotifyPropertyChanged(NameOf(LibelléPrincipal))
            If Not (Me.Libellés.Contains(value)) Then Me.Libellés.Add(value)
        End Set
    End Property

#End Region

#Region "Libellés"
    Private WithEvents _Libellés As ObservableCollection(Of String)
    Public ReadOnly Property Libellés() As ObservableCollection(Of String)
        Get
            Return _Libellés
        End Get
    End Property

    Private Sub _Libellés_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs) Handles _Libellés.CollectionChanged
        ForcerLaCohérenceEntreLibelléPrincipalEtLaCollection()
    End Sub

    Private Sub ForcerLaCohérenceEntreLibelléPrincipalEtLaCollection()
        If Not (Me.Libellés.Contains(Me.LibelléPrincipal)) Then Me.Libellés.Add(Me.LibelléPrincipal)
    End Sub
#End Region

#Region "UsagesDeProduit "
    Private _UsagesDeProduit As ObservableCollection(Of UsageDeProduit)
    Public ReadOnly Property UsagesDeProduit() As ObservableCollection(Of UsageDeProduit)
        Get
            Return _UsagesDeProduit
        End Get
    End Property
#End Region

#Region "MotsClés (ObservableCollection(of String))"
    Private _MotsClés As ObservableCollection(Of String)
    Public ReadOnly Property MotsClés() As ObservableCollection(Of String)
        Get
            Return _MotsClés
        End Get
    End Property

#End Region

#Region "TempsDePauseUnitaire (Integer)"
    Private _TempsDePauseUnitaire As Integer?

    ''' <summary>Le temps de pause en minutes.</summary>
    Public Property TempsDePauseUnitaire() As Integer?
        Get
            If _TempsDePauseUnitaire Is Nothing Then
                Return TempsDePauseCalculé
            Else
                Return _TempsDePauseUnitaire
            End If
        End Get
        Set(ByVal value As Integer?)
            If Object.Equals(value, Me._TempsDePauseUnitaire) Then Exit Property
            _TempsDePauseUnitaire = value
            NotifyPropertyChanged(NameOf(TempsDePauseUnitaire))
        End Set
    End Property


    Public ReadOnly Property TempsDePauseCalculé As Single
        Get
            Dim r = (From up In UsagesDeProduit Select up.Nombre * up.Produit.TempsDePauseUnitaire).Sum()
            Return r
        End Get
    End Property

#End Region

#Region "PrixUnitaire (Single)"
    Private _PrixUnitaire As Single?

    ''' <summary>Le prix unitaire. Forcé en attendant de </summary>
    Public Property PrixUnitaire() As Single?
        Get
            If _PrixUnitaire Is Nothing Then
                Return PrixUnitaireCalculé
            Else
                Return _PrixUnitaire
            End If
        End Get
        Set(ByVal value As Single?)
            If Object.Equals(value, Me._PrixUnitaire) Then Exit Property
            _PrixUnitaire = value
            NotifyPropertyChanged(NameOf(PrixUnitaire))
        End Set
    End Property

    Public ReadOnly Property PrixUnitaireCalculé As Single
        Get
            Dim r = (From up In UsagesDeProduit Select up.Nombre * up.Produit.Prix).Sum()
            Return r
        End Get
    End Property

#End Region

#End Region

#Region "Méthodes"


#End Region

#Region "Tests et debuggage"


#End Region

End Class
