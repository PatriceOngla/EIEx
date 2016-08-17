Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Data
Imports System.Windows.Input
Imports Model
Imports Utils

Public Class Win_SélecteurDEtude

#Region "Constructeurs"

    Private Sub UC_SélecteurDEtude_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        SetCollectionViewSourceEtudes()
        Me.SLtr_RésultatRecherche.ItemsSource = CollectionViewSourceEtudes.View
    End Sub

#End Region

#Region "Propriétés"

#Region "WS"
    Public ReadOnly Property WS() As WorkSpace
        Get
            Return WorkSpace.Instance
        End Get
    End Property
#End Region

#Region "Critères"

#Region "CritèreNom (String)"

#Region "Déclaration et registration de CritèreNomProperty"

    Private Shared FlagsMDCritèreNom As FrameworkPropertyMetadataOptions = 0
    Private Shared MDCritèreNom As New FrameworkPropertyMetadata(Nothing, FlagsMDCritèreNom, New PropertyChangedCallback(AddressOf OnCritèreNomInvalidated))
    Public Shared CritèreNomProperty As DependencyProperty = DependencyProperty.Register("CritèreNom", GetType(String), GetType(Win_SélecteurDEtude), MDCritèreNom)

#End Region

#Region "Wrapper CLR de CritèreNomProperty"
    Public Property CritèreNom() As String
        Get
            Return GetValue(CritèreNomProperty)
        End Get
        Set(ByVal value As String)
            SetValue(CritèreNomProperty, value)
        End Set
    End Property
#End Region

#Region "Gestion évennementielle de la mise à jour de CritèreNomProperty"

#Region "Evènnement CritèreNomChangedEvent et son Wrapper CLR (Non testé !!!)"

    Public Shared ReadOnly CritèreNomChangedEvent As RoutedEvent =
                  EventManager.RegisterRoutedEvent("CritèreNomChangedEvent", RoutingStrategy.Bubble,
                                                                                      GetType(RoutedPropertyChangedEventHandler(Of String)), GetType(Win_SélecteurDEtude))

    Custom Event CritèreNomChanged As RoutedEventHandler
        AddHandler(ByVal value As RoutedEventHandler)
            Me.AddHandler(CritèreNomChangedEvent, value)
        End AddHandler

        RemoveHandler(ByVal value As RoutedEventHandler)
            Me.RemoveHandler(CritèreNomChangedEvent, value)
        End RemoveHandler

        RaiseEvent(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs)
            sender.RaiseEvent(e)
        End RaiseEvent
    End Event


#End Region

    Private Shared Sub OnCritèreNomInvalidated(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)

        Dim Sender As Win_SélecteurDEtude = d
        Dim OldValue As String = e.OldValue
        Dim NewValue As String = e.NewValue

        Sender.OnCritèreNomChanged(OldValue, NewValue)

    End Sub

    Private Sub OnCritèreNomChanged(ByVal OldValue As String, ByVal NewValue As String)

        If Object.Equals(NewValue, OldValue) Then Exit Sub

        Dim args As New RoutedPropertyChangedEventArgs(Of String)(OldValue, NewValue)
        args.RoutedEvent = Win_SélecteurDEtude.CritèreNomChangedEvent

        'Insérer ici le code spécifique à la gestion du changement de la propriété "CritèreNom"
        If Not Me.RechercheSurDemande Then FiltrerLesEtudes()

        'Signalement de l'évennement au framework
        If args IsNot Nothing Then Me.RaiseEvent(args)

    End Sub

#End Region

#End Region

#Region "CritèreClient (String)"

#Region "Déclaration et registration de CritèreClientProperty"

    Private Shared FlagsMDCritèreClient As FrameworkPropertyMetadataOptions = 0
    Private Shared MDCritèreClient As New FrameworkPropertyMetadata(Nothing, FlagsMDCritèreClient, New PropertyChangedCallback(AddressOf OnCritèreClientInvalidated))
    Public Shared CritèreClientProperty As DependencyProperty = DependencyProperty.Register("CritèreClient", GetType(String), GetType(Win_SélecteurDEtude), MDCritèreClient)

#End Region

#Region "Wrapper CLR de CritèreClientProperty"
    Public Property CritèreClient() As String
        Get
            Return GetValue(CritèreClientProperty)
        End Get
        Set(ByVal value As String)
            SetValue(CritèreClientProperty, value)
        End Set
    End Property
#End Region

#Region "Gestion évennementielle de la mise à jour de CritèreClientProperty"

#Region "Evènnement CritèreClientChangedEvent et son Wrapper CLR (Non testé !!!)"

    Public Shared ReadOnly CritèreClientChangedEvent As RoutedEvent =
                  EventManager.RegisterRoutedEvent("CritèreClientChangedEvent", RoutingStrategy.Bubble,
                                                                                      GetType(RoutedPropertyChangedEventHandler(Of String)), GetType(Win_SélecteurDEtude))

    Custom Event CritèreClientChanged As RoutedEventHandler
        AddHandler(ByVal value As RoutedEventHandler)
            Me.AddHandler(CritèreClientChangedEvent, value)
        End AddHandler

        RemoveHandler(ByVal value As RoutedEventHandler)
            Me.RemoveHandler(CritèreClientChangedEvent, value)
        End RemoveHandler

        RaiseEvent(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs)
            sender.RaiseEvent(e)
        End RaiseEvent
    End Event


#End Region

    Private Shared Sub OnCritèreClientInvalidated(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)

        Dim Sender As Win_SélecteurDEtude = d
        Dim OldValue As String = e.OldValue
        Dim NewValue As String = e.NewValue

        Sender.OnCritèreClientChanged(OldValue, NewValue)

    End Sub

    Private Sub OnCritèreClientChanged(ByVal OldValue As String, ByVal NewValue As String)

        If Object.Equals(NewValue, OldValue) Then Exit Sub

        Dim args As New RoutedPropertyChangedEventArgs(Of String)(OldValue, NewValue)
        args.RoutedEvent = CritèreClientChangedEvent

        'Insérer ici le code spécifique à la gestion du changement de la propriété "CritèreClient"
        If Not Me.RechercheSurDemande Then FiltrerLesEtudes()

        'Signalement de l'évennement au framework
        If args IsNot Nothing Then Me.RaiseEvent(args)

    End Sub

#End Region

#End Region

#End Region

#Region "Options de recherche"

#Region "RechercheSurDemande (Boolean)"

    Public Shared ReadOnly RechercheSurDemandeProperty As DependencyProperty =
            DependencyProperty.Register("RechercheSurDemande", GetType(Boolean), GetType(Win_SélecteurDEtude), New UIPropertyMetadata(False))

    Public Property RechercheSurDemande As Boolean
        Get
            Return DirectCast(GetValue(RechercheSurDemandeProperty), Boolean)
        End Get

        Set(ByVal value As Boolean)
            SetValue(RechercheSurDemandeProperty, value)
        End Set
    End Property

#End Region

#Region "Source Etudes"

#Region "ViewSourceEtudes"
    Public ReadOnly Property ViewSourceEtudes() As ICollectionView
        Get
            Return CollectionViewSource.GetDefaultView(Me.SLtr_RésultatRecherche.ItemsSource)
        End Get
    End Property
#End Region

#Region "SourceEtudes"

    Public ReadOnly Property SourceEtudes As IEnumerable(Of Etude)
        Get
            Return WS.Etudes
        End Get
    End Property

    Private WithEvents _CollectionViewSourceEtudes As CollectionViewSource
    Public ReadOnly Property CollectionViewSourceEtudes() As CollectionViewSource
        Get
            Return _CollectionViewSourceEtudes
        End Get
    End Property

    Private Sub SetCollectionViewSourceEtudes()
        _CollectionViewSourceEtudes = New CollectionViewSource()
        _CollectionViewSourceEtudes.Source = Me.SourceEtudes
        Me.SLtr_RésultatRecherche.ItemsSource = CollectionViewSourceEtudes.View
    End Sub

#End Region

#End Region

#Region "DistanceTolérée (short)"

    Public Shared ReadOnly DistanceToléréeProperty As DependencyProperty =
            DependencyProperty.Register("DistanceTolérée", GetType(Short), GetType(Win_SélecteurDEtude),
                                        New UIPropertyMetadata(CShort(0), New PropertyChangedCallback(
                                                               Sub(Sender As Win_SélecteurDEtude,
                                                                   e As DependencyPropertyChangedEventArgs)
                                                                   Sender.FiltrerLesEtudes()

                                                                   'If e.NewValue > 0 Then
                                                                   '    Sender.RechercheSurDemande = True
                                                                   'End If
                                                               End Sub)
                                                               ))

    Public Property DistanceTolérée As Short
        Get
            Return DirectCast(GetValue(DistanceToléréeProperty), Short)
        End Get

        Set(ByVal value As Short)
            SetValue(DistanceToléréeProperty, value)
        End Set
    End Property

#End Region

#End Region

#Region "EtudeSélectionnée (Etude)"

#Region "Déclaration et registration de EtudeSélectionnéeProperty"

    Private Shared FlagsMDEtudeSélectionnée As FrameworkPropertyMetadataOptions = FrameworkPropertyMetadataOptions.BindsTwoWayByDefault
    Private Shared MDEtudeSélectionnée As New FrameworkPropertyMetadata(Nothing, FlagsMDEtudeSélectionnée, New PropertyChangedCallback(AddressOf OnEtudeSélectionnéeInvalidated))
    Public Shared EtudeSélectionnéeProperty As DependencyProperty = DependencyProperty.Register("EtudeSélectionnée", GetType(Etude), GetType(Win_SélecteurDEtude), MDEtudeSélectionnée)

#End Region

#Region "Wrapper CLR de EtudeSélectionnéeProperty"
    Public Property EtudeSélectionnée() As Etude
        Get
            Return GetValue(EtudeSélectionnéeProperty)
        End Get
        Set(ByVal value As Etude)
            SetValue(EtudeSélectionnéeProperty, value)
        End Set
    End Property
#End Region

#Region "Gestion évennementielle de la mise à jour de EtudeSélectionnéeProperty"

#Region "Evènnement EtudeSélectionnéeChangedEvent et son Wrapper CLR (Non testé !!!)"

    Public Shared ReadOnly EtudeSélectionnéeChangedEvent As RoutedEvent =
                  EventManager.RegisterRoutedEvent("EtudeSélectionnéeChangedEvent", RoutingStrategy.Bubble,
                                                                                      GetType(RoutedPropertyChangedEventHandler(Of Etude)), GetType(Win_SélecteurDEtude))

    Custom Event EtudeSélectionnéeChanged As RoutedEventHandler
        AddHandler(ByVal value As RoutedEventHandler)
            Me.AddHandler(EtudeSélectionnéeChangedEvent, value)
        End AddHandler

        RemoveHandler(ByVal value As RoutedEventHandler)
            Me.RemoveHandler(EtudeSélectionnéeChangedEvent, value)
        End RemoveHandler

        RaiseEvent(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs)
            sender.RaiseEvent(e)
        End RaiseEvent
    End Event


#End Region

    Private Shared Sub OnEtudeSélectionnéeInvalidated(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)

        Dim Sender As Win_SélecteurDEtude = d
        Dim OldValue As Etude = e.OldValue
        Dim NewValue As Etude = e.NewValue

        Sender.OnEtudeSélectionnéeChanged(OldValue, NewValue)

    End Sub

    Private Sub OnEtudeSélectionnéeChanged(ByVal OldValue As Etude, ByVal NewValue As Etude)

        If Object.Equals(NewValue, OldValue) Then Exit Sub

        Dim args As New RoutedPropertyChangedEventArgs(Of Etude)(OldValue, NewValue)
        args.RoutedEvent = Win_SélecteurDEtude.EtudeSélectionnéeChangedEvent
        'Insérer ici le code spécifique à la gestion du changement de la propriété "EtudeSélectionnée"

        'Signalement de l'évennement au framework
        If args IsNot Nothing Then Me.RaiseEvent(args)

    End Sub

#End Region

#End Region

#Region "InfosResultat (String)"

#Region "Déclaration et registration de InfosResultatProperty"

    Private Shared MDInfosResultat As New FrameworkPropertyMetadata(Nothing)
    Public Shared InfosResultatPropertyKey As DependencyPropertyKey = DependencyProperty.RegisterReadOnly("InfosResultat", GetType(String), GetType(Win_SélecteurDEtude), MDInfosResultat)
    Public Shared InfosResultatProperty As DependencyProperty = InfosResultatPropertyKey.DependencyProperty

#End Region

#Region "Wrapper CLR de InfosResultatProperty"

    Public ReadOnly Property InfosResultat() As String
        Get
            Return GetValue(InfosResultatProperty)
        End Get
    End Property

#End Region

#End Region

#Region "NbRésultats"
    Public ReadOnly Property NbRésultats() As Integer
        Get
            Return Me.ViewSourceEtudes.OfType(Of Etude).Count()
        End Get
    End Property
#End Region

#Region "Résultat"
    Private _Résultat As Etude
    Public ReadOnly Property Résultat() As Etude
        Get
            Return _Résultat
        End Get
    End Property
#End Region

#End Region

#Region "Méthodes"

#Region "Tri, filtre, regroupement"

#Region "Filtre"

    Private Sub FiltrerLesEtudes()
        'Dim r = From o In Me.SourceEtudes Where MatcheCritères(o)
        'Me.SetValue(InfosResultatPropertyKey, r.Count & " Etude(s)")
        Me.ViewSourceEtudes.Filter = Function(o As Etude) MatcheCritères(o)
        Me.SetValue(InfosResultatPropertyKey, $"{Me.NbRésultats} Etude(s) trouvé(s)")
    End Sub

    Private Sub _SourceProduits_Filter(sender As Object, e As FilterEventArgs) Handles _CollectionViewSourceEtudes.Filter
        Dim O As Etude = e.Item
        If O IsNot Nothing Then
            e.Accepted = MatcheCritères(e.Item)
        End If
    End Sub

#Region "MatcheCritères"

    Private Function MatcheCritères(E As Etude) As Boolean

        Dim MatchNom, MatchClient As Boolean

        MatchNom = MatcheCritèresNoms(E, Me.CritèreNom, Me.DistanceTolérée)
        If MatchNom Then
            MatchClient = MatcheCritèresClients(E, Me.CritèreClient, Me.DistanceTolérée)
        End If

        Dim r = MatchNom AndAlso MatchClient
        Return r

    End Function

    Private Shared Function MatcheCritèresNoms(E As Etude, CritèreNom As String, DistanceTolérée As Short) As Boolean
        Dim r = MatcheCritères(E.Nom, CritèreNom, DistanceTolérée)
        Return r
    End Function

    Private Shared Function MatcheCritèresClients(E As Etude, CritèreClient As String, DistanceTolérée As Short) As Boolean
        Dim r = MatcheCritères(E.Client, CritèreClient, DistanceTolérée)
        Return r
    End Function

    ''' <param name="DistanceTolérée">Distance de Levenstein tolérée.</param>
    ''' <returns></returns>
    Private Shared Function MatcheCritères(TargetPropValue As String, Critère As String, DistanceTolérée As Short) As Boolean
        Dim AucunCritère = String.IsNullOrEmpty(Critère)
        Dim r As Boolean
        If AucunCritère Then
            r = True
        Else
            If String.IsNullOrEmpty(TargetPropValue) Then
                r = False
            Else
                Dim Mots = TargetPropValue.Split({" "c, "'"c})
                Dim TabMotsCherchés = Critère.Split({" "c, "'"c}, StringSplitOptions.RemoveEmptyEntries)
                r = Mots.ContainsList_String(TabMotsCherchés, True, True, DistanceTolérée)
            End If
        End If
        Return r
    End Function

#End Region

#End Region

#Region "Tri"

    'Private Sub TrierLesProduitsParFournisseurEtParRéférence()
    '    With Me.ViewSourceEtudes
    '        If?.CanSort Then
    '            .SortDescriptions.Clear()
    '            .SortDescriptions.Add(New SortDescription(NameOf(Produit.CodeLydic), ListSortDirection.Ascending))
    '            .SortDescriptions.Add(New SortDescription(NameOf(Produit.RéférenceFournisseur), ListSortDirection.Ascending))
    '        End If
    '    End With
    'End Sub

#End Region

#Region "Regroupements"

    Private Sub GrouperLesProduitsParFournisseur()
        With Me.ViewSourceEtudes
            If?.CanGroup Then
                '.GroupDescriptions.Clear()
                .GroupDescriptions.Add(New PropertyGroupDescription(NameOf(Produit.CodeLydic)))
            End If
        End With
    End Sub

    Private Sub GrouperLesProduitsParUnité()
        With Me.ViewSourceEtudes
            If?.CanGroup Then
                .GroupDescriptions.Add(New PropertyGroupDescription("Unité"))
                '.GroupDescriptions.Add(New PropertyGroupDescription("Complete"))
            End If
        End With
    End Sub

#End Region

#End Region

#Region "Gestionnaires d'évennements"

#Region "Validation du choix de produit"

    Private Sub SLtr_RésultatRecherche_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles SLtr_RésultatRecherche.MouseDoubleClick
        Me.ValiderLeChoix()
    End Sub

    Private Sub SLtr_RésultatRecherche_KeyDown(sender As Object, e As KeyEventArgs) Handles SLtr_RésultatRecherche.KeyDown
        If e.Key = Key.Return Then
            e.Handled = True
            Me.ValiderLeChoix()
        End If
    End Sub

#End Region

    Private Sub UC_SélecteurDEtude_KeyDown(sender As Object, e As KeyEventArgs) Handles TBx_CritèreNom.KeyDown, TBx_CritèreClient.KeyDown
        If e.Key = Key.Return Then
            Me.FiltrerLesEtudes()
            Me.SLtr_RésultatRecherche.Focus()
        End If
    End Sub

    Private Sub Btn_Chercher_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Chercher.Click
        Me.FiltrerLesEtudes()
    End Sub

#End Region

#Region "Cherche"

    Public Shared Function Cherche() As Etude
        Dim w = New Win_SélecteurDEtude
        With w
            .Reset()
            .ShowDialog()
            Return .Résultat
        End With
    End Function

    Private Sub Reset()
        Me._Résultat = Nothing
        Dim bckup = Me.RechercheSurDemande
        Me.RechercheSurDemande = True
        Me.TBx_CritèreNom.Clear()
        Me.TBx_CritèreClient.Clear()
        Me.TBx_CritèreNom.Focus()
        Me.RechercheSurDemande = bckup
        Me.FiltrerLesEtudes()
    End Sub

#End Region

#Region "ValiderLeChoix"
    Private Sub ValiderLeChoix()
        Me._Résultat = Me.EtudeSélectionnée
        Me.DialogResult = True
        If (Me.Résultat IsNot Nothing) Then
            RaiseEvent EtudeTrouvé(Me.Résultat)
            Me.Close()
        End If
    End Sub
#End Region

#End Region

#Region "Events"

#Region "EtudeTrouvé"

    Public Event EtudeTrouvé(O As Etude)

#End Region

#End Region

#Region "Tests et debuggage"


#End Region

End Class

