Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports Model
Imports Utils

''' <summary>
''' On distingue <see cref="PatronDOuvrage"/> et <see cref="Ouvrage"/>. Les <see cref="Ouvrage"/>s sont les entrées des bordereau et sont associés à des <see cref="PatronDOuvrage"/> (pas acquis, à valider !) afin de calculer leur prix sur la base du <see cref="PatronDOuvrage.PrixUnitaire"/>. 
''' </summary>
Public MustInherit Class Ouvrage_Base
    Inherits Entité
    Implements ICloneable

#Region "Constructeurs"

    Protected Overrides Sub Init()
        _Libellés = New ObservableCollection(Of String)
        _UsagesDeProduit = New ObservableCollection(Of UsageDeProduit)
        _MotsClés = New List(Of String)
    End Sub

#End Region

#Region "Propriétés"

#Region "Gestion du nommage"

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

#Region "ComplémentDeNom"
    Private _ComplémentDeNom As String
    Public Property ComplémentDeNom() As String
        Get
            Return _ComplémentDeNom
        End Get
        Set(ByVal value As String)
            If Object.Equals(value, Me._ComplémentDeNom) Then Exit Property
            _ComplémentDeNom = value
            NotifyPropertyChanged(NameOf(ComplémentDeNom))
        End Set
    End Property
#End Region

#Region "NomComplet"
    ''' <summary>Le nom saisi + le complément de nom s'il y a en a un. </summary>
    Public ReadOnly Property NomComplet() As String
        Get
            Return Me.Nom & If(String.IsNullOrEmpty(ComplémentDeNom), "", " - " & ComplémentDeNom)
        End Get
    End Property
#End Region

#Region "Libellés"

    Private WithEvents _Libellés As ObservableCollection(Of String)
    Public ReadOnly Property Libellés() As ObservableCollection(Of String)
        Get
            Return _Libellés
        End Get
    End Property

    Public ReadOnly Property NbLibellés() As Integer
        Get
            Return Me.Libellés?.Count
        End Get
    End Property

    Private Sub _Libellés_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs) Handles _Libellés.CollectionChanged
        ForcerLaCohérenceEntreLibelléPrincipalEtLaCollection()
        Me.NotifyPropertyChanged(NameOf(NbLibellés))

    End Sub

    Private Sub ForcerLaCohérenceEntreLibelléPrincipalEtLaCollection()
        If Not (String.IsNullOrEmpty(Me.Nom) OrElse Me.Libellés.Contains(Me.Nom)) Then Me.Libellés.Add(Me.Nom)
    End Sub

#End Region

#End Region

#Region "Bordereau"
    Private _Bordereau As Bordereau
    Public ReadOnly Property Bordereau() As Bordereau
        Get
            Return _Bordereau
        End Get
    End Property
#End Region

#Region "UsagesDeProduit "

    Private WithEvents _UsagesDeProduit As ObservableCollection(Of UsageDeProduit)
    Public ReadOnly Property UsagesDeProduit() As ObservableCollection(Of UsageDeProduit)
        Get
            Return _UsagesDeProduit
        End Get
    End Property


    Public ReadOnly Property NbProduits() As Integer
        Get
            Return Me.UsagesDeProduit.Count()
        End Get
    End Property

#End Region

#Region "MotsClés (List(of String))"
    Protected _MotsClés As List(Of String)
    Public Property MotsClés() As List(Of String)
        Get
            Return _MotsClés
        End Get
        Set(value As List(Of String))
            _MotsClés = value
            NotifyPropertyChanged(NameOf(MotsClés))
        End Set
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
            NotifyPropertyChanged(NameOf(TempsDePauseForcé))
        End Set
    End Property


    Public ReadOnly Property TempsDePauseCalculé As Single
        Get
            Dim r = (From up In UsagesDeProduit Select up.Nombre * up.Produit?.TempsDePauseUnitaire).Sum()
            Return r
        End Get
    End Property

    Public ReadOnly Property TempsDePauseForcé() As Boolean
        Get
            Return _TempsDePauseUnitaire IsNot Nothing
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
            NotifyPropertyChanged(NameOf(PrixUnitaireForcé))
        End Set
    End Property

    Public ReadOnly Property PrixUnitaireCalculé As Single
        Get
            Dim r = (From up In UsagesDeProduit Select up.Nombre * up.Produit?.Prix).Sum()
            Return r
        End Get
    End Property

    Public ReadOnly Property PrixUnitaireForcé() As Boolean
        Get
            Return _PrixUnitaire IsNot Nothing
        End Get
    End Property

#End Region

#End Region

#Region "Méthodes"

#Region "AjouterProduit"

    Public Function AjouterProduit(P As Produit, Nombre As Short) As UsageDeProduit
        Dim up = New UsageDeProduit(Me) With {.Produit = P, .Nombre = Nombre}
        Me.UsagesDeProduit.Add(up)
        Return up
    End Function

#Region "VerifierLesElémentsAjoutés"
    Private VérificationDesUsageDeProduitAjoutésEnCours As Boolean
    Private Sub _UsagesDeProduit_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs) Handles _UsagesDeProduit.CollectionChanged
        If Not VérificationDesUsageDeProduitAjoutésEnCours Then
            If e.NewItems IsNot Nothing Then
                VerifierLesElémentsAjoutés(e.NewItems.OfType(Of UsageDeProduit))
            End If
        End If
        Me.NotifyPropertyChanged(NameOf(NbProduits))
        NotifyPropertyChanged(NameOf(TempsDePauseUnitaire))
        NotifyPropertyChanged(NameOf(PrixUnitaire))

    End Sub

    ''' <summary>S'assure que tous les <paramref name="UsagesDeProduitAjoutés"/> ont bien pour parent le <see cref="Ouvrage_Base"/> courant.</summary>
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
                Dim Msg = $"{NbErr} des '{NameOf(UsageDeProduit)}' ajouté{If(Pluriel, "s", "")} {If(Pluriel, "sont", "est")} déjà associé{If(Pluriel, "s", "")} à une autre '{NameOf(Ouvrage_Base)}'. Ce{If(Pluriel, "s", "t")} élément{If(Pluriel, "s", "")} {If(Pluriel, "n'ont", "n'a")} pas été ajouté{If(Pluriel, "s", "")}."
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

#Region "UtiliseProduit"
    Public Function UtiliseProduit(p As Produit) As Boolean
        Dim r = (From up In Me.UsagesDeProduit Where up.Produit Is p).Any()
        Return r
    End Function
#End Region

#Region "Gestion du templating"

    Public Function Clone() As Object Implements ICloneable.Clone
        Return Me.Copie
    End Function

    Public Function Copie() As Ouvrage_Base
        'TODO: implémenter
        Throw New NotImplementedException()
        'Dim r = Ref.GetNewOuvrage()
        'r.Nom = Me.NomComplet
        'r.ComplémentDeNom = "?"
        'r.Libellés.AddRange(Me.Libellés)
        'r.MotsClés.AddRange(Me.MotsClés)
        'If Me.PrixUnitaireForcé Then r.PrixUnitaire = Me.PrixUnitaire
        'If Me.TempsDePauseForcé Then r.TempsDePauseUnitaire = Me.TempsDePauseUnitaire

        'Me.UsagesDeProduit.DoForAll(Sub(up As UsageDeProduit)
        '                                r.AjouterProduit(up.Produit, up.Nombre)
        '                            End Sub)
        'Return r
    End Function

#End Region

#End Region

#Region "Tests et debuggage"


#End Region

End Class
