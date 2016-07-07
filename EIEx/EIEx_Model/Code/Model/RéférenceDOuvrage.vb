Imports System.Collections.ObjectModel
Imports System.Collections.Specialized

''' <summary>
''' On distingue <see cref="RéférenceDOuvrage"/> et Ouvrage (pas encore implémenté). Les ouvrages sont les entrées des bordereau et sont associé à des <see cref="RéférenceDOuvrage"/> afin de calculer leur prix sur la base du <see cref="RéférenceDOuvrage.PrixUnitaire"/>. 
''' </summary>
Public Class RéférenceDOuvrage
    Inherits AgregateRoot

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(Id As Integer)
        MyBase.New(Id)
    End Sub

    Protected Overrides Sub Init()
        _Libellés = New ObservableCollection(Of String)
        _UsagesDeProduit = New ObservableCollection(Of UsageDeProduit)
        _MotsClés = New ObservableCollection(Of String)
    End Sub

    Protected Overrides Sub SetId()
        Me._Id = Réf.GetNewId(Of RéférenceDOuvrage)
    End Sub

    Protected Overrides Sub SEnregistrerDansLeRéférentiel()
        Réf.EnregistrerRoot(Me)
    End Sub

#End Region

#Region "Propriétés"

#Region "Nom (String)"
    Private _Nom As String
    Public Overrides Property Nom() As String
        Get
            Return _Nom
        End Get
        Set(ByVal value As String)
            If Object.Equals(value, Me._Nom) Then Exit Property
            _Nom = value
            NotifyPropertyChanged(NameOf(Nom))
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
        If Not (Me.Libellés.Contains(Me.Nom)) Then Me.Libellés.Add(Me.Nom)
    End Sub
#End Region

#Region "UsagesDeProduit "
    Private WithEvents _UsagesDeProduit As ObservableCollection(Of UsageDeProduit)
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

#Region "AjouterProduit"

    Public Sub AjouterProduit(IdProduit As Integer, Nombre As Integer)
        Dim p = Réf.GetProduitById(IdProduit)
        Dim up = New UsageDeProduit(Me) With {.Produit = p, .Nombre = Nombre}
        Me.UsagesDeProduit.Add(up)
    End Sub

#Region "VerifierLesElémentsAjoutés"
    Private VérificationDesUsageDeProduitAjoutésEnCours As Boolean
    Private Sub _UsagesDeProduit_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs) Handles _UsagesDeProduit.CollectionChanged
        If Not VérificationDesUsageDeProduitAjoutésEnCours Then
            VerifierLesElémentsAjoutés(e.NewItems.OfType(Of UsageDeProduit))
        End If
    End Sub

    ''' <summary>S'assure que tous les <paramref name="UsagesDeProduitAjoutés"/> ont bien pour parent la <see cref="RéférenceDOuvrage"/> courante.</summary>
    ''' <param name="UsagesDeProduitAjoutés"></param>
    Private Sub VerifierLesElémentsAjoutés(UsagesDeProduitAjoutés As IEnumerable(Of UsageDeProduit))
        Try
            VérificationDesUsageDeProduitAjoutésEnCours = True

            Dim ItemsEnErreur = New List(Of UsageDeProduit)()
            For Each up In UsagesDeProduitAjoutés
                If up.Parent IsNot Me Then
                    ItemsEnErreur.Add(up)
                End If
            Next
            Dim NbErr = ItemsEnErreur.Count
            If NbErr > 0 Then
                For Each up In ItemsEnErreur
                    If Me.UsagesDeProduit.Contains(up) Then
                        Me.UsagesDeProduit.Remove(up)
                    End If
                Next
                Dim Pluriel = NbErr > 1
                Dim Msg = $"{NbErr} des '{NameOf(UsageDeProduit)}' ajouté{If(Pluriel, "s", "")} {If(Pluriel, "sont", "est")} déjà associé{If(Pluriel, "s", "")} à une autre '{NameOf(RéférenceDOuvrage)}'. Ce{If(Pluriel, "s", "t")} élément{If(Pluriel, "s", "")} {If(Pluriel, "n'ont", "n'a")} pas été ajouté{If(Pluriel, "s", "")}."
                Throw New InvalidOperationException(Msg)
            End If
        Catch ex As Exception
            Throw ex
        Finally
            VérificationDesUsageDeProduitAjoutésEnCours = False
        End Try

    End Sub

#End Region

#End Region

#End Region

#Region "Tests et debuggage"


#End Region

End Class
