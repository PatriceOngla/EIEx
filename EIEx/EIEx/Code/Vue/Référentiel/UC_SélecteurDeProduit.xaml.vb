Imports System.ComponentModel
Imports System.Threading
Imports System.Windows
Imports System.Windows.Data
Imports System.Windows.Input
Imports Model
Imports Utils

Public Class UC_SélecteurDeProduit

#Region "Constructeurs"

    Private Sub New()

        ' Cet appel est requis par le concepteur.
        InitializeComponent()

        ' Ajoutez une initialisation quelconque après l'appel InitializeComponent().

    End Sub

    Private Sub UC_SélecteurDeProduit_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        'If Me.DataContext Is Nothing Then
        'Me.Lbx_RésultatsRecherche.ItemsSource = Ref.Produits
        'End If
        InitSource()
        Me.SLtr_RésultatRecherche.ItemsSource = SourceProduits.View
    End Sub

    Private Sub InitSource()
        _SourceProduits = New CollectionViewSource()
        _SourceProduits.Source = Ref.Produits

    End Sub

#End Region

#Region "Propriétés"

#Region "Instance"
    Private Shared _Instance As UC_SélecteurDeProduit
    Public Shared ReadOnly Property Instance As UC_SélecteurDeProduit
        Get
            Return _Instance
        End Get
    End Property

#End Region

#Region "FenêtreParente"
    Private _FenêtreParente As Window
    Public ReadOnly Property FenêtreParente() As Window
        Get
            Return _FenêtreParente
        End Get
    End Property
#End Region

#Region "Ref"
    Public ReadOnly Property Ref() As Référentiel
        Get
            Return Référentiel.Instance
        End Get
    End Property
#End Region

#Region "RechercheSurDemande (Boolean)"

    Public Shared ReadOnly RechercheSurDemandeProperty As DependencyProperty =
            DependencyProperty.Register("RechercheSurDemande", GetType(Boolean), GetType(UC_SélecteurDeProduit), New UIPropertyMetadata(False))

    Public Property RechercheSurDemande As Boolean
        Get
            Return DirectCast(GetValue(RechercheSurDemandeProperty), Boolean)
        End Get

        Set(ByVal value As Boolean)
            SetValue(RechercheSurDemandeProperty, value)
        End Set
    End Property

#End Region

#Region "Critères"

#Region "CritèreMotsClés (String)"

#Region "Déclaration et registration de CritèreMotsClésProperty"

    Private Shared FlagsMDCritèreMotsClés As FrameworkPropertyMetadataOptions = 0
    Private Shared MDCritèreMotsClés As New FrameworkPropertyMetadata(Nothing, FlagsMDCritèreMotsClés, New PropertyChangedCallback(AddressOf OnCritèreMotsClésInvalidated))
    Public Shared CritèreMotsClésProperty As DependencyProperty = DependencyProperty.Register("CritèreMotsClés", GetType(String), GetType(UC_SélecteurDeProduit), MDCritèreMotsClés)

#End Region

#Region "Wrapper CLR de CritèreMotsClésProperty"
    Public Property CritèreMotsClés() As String
        Get
            Return GetValue(CritèreMotsClésProperty)
        End Get
        Set(ByVal value As String)
            SetValue(CritèreMotsClésProperty, value)
        End Set
    End Property
#End Region

#Region "Gestion évennementielle de la mise à jour de CritèreMotsClésProperty"

#Region "Evènnement CritèreMotsClésChangedEvent et son Wrapper CLR (Non testé !!!)"

    Public Shared ReadOnly CritèreMotsClésChangedEvent As RoutedEvent =
                  EventManager.RegisterRoutedEvent("CritèreMotsClésChangedEvent", RoutingStrategy.Bubble,
                                                                                      GetType(RoutedPropertyChangedEventHandler(Of String)), GetType(UC_SélecteurDeProduit))

    Custom Event CritèreMotsClésChanged As RoutedEventHandler
        AddHandler(ByVal value As RoutedEventHandler)
            Me.AddHandler(CritèreMotsClésChangedEvent, value)
        End AddHandler

        RemoveHandler(ByVal value As RoutedEventHandler)
            Me.RemoveHandler(CritèreMotsClésChangedEvent, value)
        End RemoveHandler

        RaiseEvent(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs)
            sender.RaiseEvent(e)
        End RaiseEvent
    End Event


#End Region

    Private Shared Sub OnCritèreMotsClésInvalidated(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)

        Dim Sender As UC_SélecteurDeProduit = d
        Dim OldValue As String = e.OldValue
        Dim NewValue As String = e.NewValue

        Sender.OnCritèreMotsClésChanged(OldValue, NewValue)

    End Sub

    Private Sub OnCritèreMotsClésChanged(ByVal OldValue As String, ByVal NewValue As String)

        If Object.Equals(NewValue, OldValue) Then Exit Sub

        Dim args As New RoutedPropertyChangedEventArgs(Of String)(OldValue, NewValue)
        args.RoutedEvent = UC_SélecteurDeProduit.CritèreMotsClésChangedEvent

        'Insérer ici le code spécifique à la gestion du changement de la propriété "CritèreMotsClés"
        If Not Me.RechercheSurDemande Then FiltrerLesProduits()

        'Signalement de l'évennement au framework
        If args IsNot Nothing Then Me.RaiseEvent(args)

    End Sub

#End Region

#End Region

#Region "CritèreCodeLydic (String)"

#Region "Déclaration et registration de CritèreCodeLydicProperty"

    Private Shared FlagsMDCritèreCodeLydic As FrameworkPropertyMetadataOptions = 0
    Private Shared MDCritèreCodeLydic As New FrameworkPropertyMetadata(Nothing, FlagsMDCritèreCodeLydic, New PropertyChangedCallback(AddressOf OnCritèreCodeLydicInvalidated))
    Public Shared CritèreCodeLydicProperty As DependencyProperty = DependencyProperty.Register("CritèreCodeLydic", GetType(String), GetType(UC_SélecteurDeProduit), MDCritèreCodeLydic)

#End Region

#Region "Wrapper CLR de CritèreCodeLydicProperty"
    Public Property CritèreCodeLydic() As String
        Get
            Return GetValue(CritèreCodeLydicProperty)
        End Get
        Set(ByVal value As String)
            SetValue(CritèreCodeLydicProperty, value)
        End Set
    End Property
#End Region

#Region "Gestion évennementielle de la mise à jour de CritèreCodeLydicProperty"

#Region "Evènnement CritèreCodeLydicChangedEvent et son Wrapper CLR (Non testé !!!)"

    Public Shared ReadOnly CritèreCodeLydicChangedEvent As RoutedEvent =
                  EventManager.RegisterRoutedEvent("CritèreCodeLydicChangedEvent", RoutingStrategy.Bubble,
                                                                                      GetType(RoutedPropertyChangedEventHandler(Of String)), GetType(UC_SélecteurDeProduit))

    Custom Event CritèreCodeLydicChanged As RoutedEventHandler
        AddHandler(ByVal value As RoutedEventHandler)
            Me.AddHandler(CritèreCodeLydicChangedEvent, value)
        End AddHandler

        RemoveHandler(ByVal value As RoutedEventHandler)
            Me.RemoveHandler(CritèreCodeLydicChangedEvent, value)
        End RemoveHandler

        RaiseEvent(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs)
            sender.RaiseEvent(e)
        End RaiseEvent
    End Event


#End Region

    Private Shared Sub OnCritèreCodeLydicInvalidated(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)

        Dim Sender As UC_SélecteurDeProduit = d
        Dim OldValue As String = e.OldValue
        Dim NewValue As String = e.NewValue

        Sender.OnCritèreCodeLydicChanged(OldValue, NewValue)

    End Sub

    Private Sub OnCritèreCodeLydicChanged(ByVal OldValue As String, ByVal NewValue As String)

        If Object.Equals(NewValue, OldValue) Then Exit Sub

        Dim args As New RoutedPropertyChangedEventArgs(Of String)(OldValue, NewValue)
        args.RoutedEvent = UC_SélecteurDeProduit.CritèreCodeLydicChangedEvent

        'Insérer ici le code spécifique à la gestion du changement de la propriété "CritèreCodeLydic"
        If Not Me.RechercheSurDemande Then FiltrerLesProduits()

        'Signalement de l'évennement au framework
        If args IsNot Nothing Then Me.RaiseEvent(args)

    End Sub

#End Region

#End Region

#Region "CritèreRefFournisseur (String)"

#Region "Déclaration et registration de CritèreRefFournisseurProperty"

    Private Shared FlagsMDCritèreRefFournisseur As FrameworkPropertyMetadataOptions = 0
    Private Shared MDCritèreRefFournisseur As New FrameworkPropertyMetadata(Nothing, FlagsMDCritèreRefFournisseur, New PropertyChangedCallback(AddressOf OnCritèreRefFournisseurInvalidated))
    Public Shared CritèreRefFournisseurProperty As DependencyProperty = DependencyProperty.Register("CritèreRefFournisseur", GetType(String), GetType(UC_SélecteurDeProduit), MDCritèreRefFournisseur)

#End Region

#Region "Wrapper CLR de CritèreRefFournisseurProperty"
    Public Property CritèreRefFournisseur() As String
        Get
            Return GetValue(CritèreRefFournisseurProperty)
        End Get
        Set(ByVal value As String)
            SetValue(CritèreRefFournisseurProperty, value)
        End Set
    End Property
#End Region

#Region "Gestion évennementielle de la mise à jour de CritèreRefFournisseurProperty"

#Region "Evènnement CritèreRefFournisseurChangedEvent et son Wrapper CLR (Non testé !!!)"

    Public Shared ReadOnly CritèreRefFournisseurChangedEvent As RoutedEvent =
                  EventManager.RegisterRoutedEvent("CritèreRefFournisseurChangedEvent", RoutingStrategy.Bubble,
                                                                                      GetType(RoutedPropertyChangedEventHandler(Of String)), GetType(UC_SélecteurDeProduit))

    Custom Event CritèreRefFournisseurChanged As RoutedEventHandler
        AddHandler(ByVal value As RoutedEventHandler)
            Me.AddHandler(CritèreRefFournisseurChangedEvent, value)
        End AddHandler

        RemoveHandler(ByVal value As RoutedEventHandler)
            Me.RemoveHandler(CritèreRefFournisseurChangedEvent, value)
        End RemoveHandler

        RaiseEvent(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs)
            sender.RaiseEvent(e)
        End RaiseEvent
    End Event


#End Region

    Private Shared Sub OnCritèreRefFournisseurInvalidated(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)

        Dim Sender As UC_SélecteurDeProduit = d
        Dim OldValue As String = e.OldValue
        Dim NewValue As String = e.NewValue

        Sender.OnCritèreRefFournisseurChanged(OldValue, NewValue)

    End Sub

    Private Sub OnCritèreRefFournisseurChanged(ByVal OldValue As String, ByVal NewValue As String)

        If Object.Equals(NewValue, OldValue) Then Exit Sub

        Dim args As New RoutedPropertyChangedEventArgs(Of String)(OldValue, NewValue)
        args.RoutedEvent = CritèreRefFournisseurChangedEvent

        'Insérer ici le code spécifique à la gestion du changement de la propriété "CritèreRefFournisseur"
        If Not Me.RechercheSurDemande Then FiltrerLesProduits()

        'Signalement de l'évennement au framework
        If args IsNot Nothing Then Me.RaiseEvent(args)

    End Sub

#End Region

#End Region

#End Region

#Region "Source produits"

#Region "VueSourceProduits"
    Public ReadOnly Property VueSourceProduits() As ICollectionView
        Get
            Return CollectionViewSource.GetDefaultView(Me.SLtr_RésultatRecherche.ItemsSource)
        End Get
    End Property
#End Region

#Region "SourceProduits"

    Private WithEvents _SourceProduits As CollectionViewSource
    Public ReadOnly Property SourceProduits() As CollectionViewSource
        Get
            Return _SourceProduits
        End Get
    End Property
#End Region

#End Region

#Region "ProduitSélectionné (Produit)"

#Region "Déclaration et registration de ProduitSélectionnéProperty"

    Private Shared FlagsMDProduitSélectionné As FrameworkPropertyMetadataOptions = FrameworkPropertyMetadataOptions.BindsTwoWayByDefault
    Private Shared MDProduitSélectionné As New FrameworkPropertyMetadata(Nothing, FlagsMDProduitSélectionné, New PropertyChangedCallback(AddressOf OnProduitSélectionnéInvalidated))
    Public Shared ProduitSélectionnéProperty As DependencyProperty = DependencyProperty.Register("ProduitSélectionné", GetType(Produit), GetType(UC_SélecteurDeProduit), MDProduitSélectionné)

#End Region

#Region "Wrapper CLR de ProduitSélectionnéProperty"
    Public Property ProduitSélectionné() As Produit
        Get
            Return GetValue(ProduitSélectionnéProperty)
        End Get
        Set(ByVal value As Produit)
            SetValue(ProduitSélectionnéProperty, value)
        End Set
    End Property
#End Region

#Region "Gestion évennementielle de la mise à jour de ProduitSélectionnéProperty"

#Region "Evènnement ProduitSélectionnéChangedEvent et son Wrapper CLR (Non testé !!!)"

    Public Shared ReadOnly ProduitSélectionnéChangedEvent As RoutedEvent =
                  EventManager.RegisterRoutedEvent("ProduitSélectionnéChangedEvent", RoutingStrategy.Bubble,
                                                                                      GetType(RoutedPropertyChangedEventHandler(Of Produit)), GetType(UC_SélecteurDeProduit))

    Custom Event ProduitSélectionnéChanged As RoutedEventHandler
        AddHandler(ByVal value As RoutedEventHandler)
            Me.AddHandler(ProduitSélectionnéChangedEvent, value)
        End AddHandler

        RemoveHandler(ByVal value As RoutedEventHandler)
            Me.RemoveHandler(ProduitSélectionnéChangedEvent, value)
        End RemoveHandler

        RaiseEvent(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs)
            sender.RaiseEvent(e)
        End RaiseEvent
    End Event


#End Region

    Private Shared Sub OnProduitSélectionnéInvalidated(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)

        Dim Sender As UC_SélecteurDeProduit = d
        Dim OldValue As Produit = e.OldValue
        Dim NewValue As Produit = e.NewValue

        Sender.OnProduitSélectionnéChanged(OldValue, NewValue)

    End Sub

    Private Sub OnProduitSélectionnéChanged(ByVal OldValue As Produit, ByVal NewValue As Produit)

        If Object.Equals(NewValue, OldValue) Then Exit Sub

        Dim args As New RoutedPropertyChangedEventArgs(Of Produit)(OldValue, NewValue)
        args.RoutedEvent = UC_SélecteurDeProduit.ProduitSélectionnéChangedEvent

        'Insérer ici le code spécifique à la gestion du changement de la propriété "ProduitSélectionné"

        'Signalement de l'évennement au framework
        If args IsNot Nothing Then Me.RaiseEvent(args)

    End Sub

#End Region

#End Region

#Region "InfosResultat (String)"

#Region "Déclaration et registration de InfosResultatProperty"

    Private Shared MDInfosResultat As New FrameworkPropertyMetadata(Nothing)
    Public Shared InfosResultatPropertyKey As DependencyPropertyKey = DependencyProperty.RegisterReadOnly("InfosResultat", GetType(String), GetType(UC_SélecteurDeProduit), MDInfosResultat)
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
            Return Me.VueSourceProduits.OfType(Of Produit).Count()
        End Get
    End Property
#End Region

#Region "EntêteRésultats"
    Public ReadOnly Property EntêteRésultats() As String
        Get
            Return Produit.ProductsListHeader
        End Get
    End Property
#End Region

#End Region

#Region "Méthodes"

#Region "Tri, filtre, regroupement"

#Region "Filtre"

    Private Sub FiltrerLesProduits()
        'Dim r = From p In Ref.Produits Where MatcheCritères(p)
        'Me.SetValue(InfosResultatPropertyKey, r.Count & " produit(s)")
        'Exit Sub
        Me.VueSourceProduits.Filter = Function(p As Produit) MatcheCritères(p, Me.CritèreMotsClés, Me.CritèreCodeLydic, Me.CritèreRefFournisseur)
        Me.SetValue(InfosResultatPropertyKey, $"{Me.NbRésultats} produit(s) trouvé(s)")
    End Sub

    Private Sub _SourceProduits_Filter(sender As Object, e As FilterEventArgs) Handles _SourceProduits.Filter
        Dim p As Produit = e.Item
        If p IsNot Nothing Then
            e.Accepted = MatcheCritères(e.Item, Me.CritèreMotsClés, Me.CritèreCodeLydic, Me.CritèreRefFournisseur)
        End If
    End Sub

#Region "MatcheCritères"

    Friend Shared Function MatcheCritères(P As Produit, CritèreMotsClés As String, CritèreCodeLydic As String, CritèreRefFournisseur As String) As Boolean

        'Return True

        Dim MatchMotsClés, MatchCodeLydic, MatchCodeFournisseur As Boolean

        MatchMotsClés = MatcheCritèresMotsClés(P, CritèreMotsClés)
        If MatchMotsClés Then
            MatchCodeLydic = MatcheCritèresCodeLydic(P, CritèreCodeLydic)
            If MatchCodeLydic Then
                MatchCodeFournisseur = MatcheCritèresNumFournisseur(P, CritèreRefFournisseur)
            End If
        End If

        Dim r = MatchMotsClés AndAlso MatchCodeLydic AndAlso MatchCodeFournisseur
        Return r

    End Function

    Private Shared Function MatcheCritèresCodeLydic(P As Produit, C As String) As Boolean
        Dim r = String.IsNullOrEmpty(C)
        r = r OrElse P.CodeLydic?.StartsWith(C, StringComparison.CurrentCultureIgnoreCase)
        Return r
    End Function

    Private Shared Function MatcheCritèresNumFournisseur(P As Produit, C As String) As Boolean
        Dim r = String.IsNullOrEmpty(C)
        r = r OrElse P.RéférenceFournisseur.Contains(C)
        Return r
    End Function

    Private Shared Function MatcheCritèresMotsClés(P As Produit, C As String) As Boolean
        Dim r = String.IsNullOrEmpty(C)
        If Not r Then
            Dim TabMotsClés = C.Split({" "c, "'"c}, StringSplitOptions.RemoveEmptyEntries)
            r = P.Mots.ContainsList_String(TabMotsClés, True, True)
        End If
        Return r
    End Function

#End Region

#End Region

#Region "Tri"

    Private Sub TrierLesProduitsParFournisseurEtParRéférence()
        With Me.VueSourceProduits
            If?.CanSort Then
                .SortDescriptions.Clear()
                .SortDescriptions.Add(New SortDescription(NameOf(Produit.CodeLydic), ListSortDirection.Ascending))
                .SortDescriptions.Add(New SortDescription(NameOf(Produit.RéférenceFournisseur), ListSortDirection.Ascending))
            End If
        End With
    End Sub

#End Region

#Region "Regroupements"

    Private Sub GrouperLesProduitsParFournisseur()
        With Me.VueSourceProduits
            If?.CanGroup Then
                '.GroupDescriptions.Clear()
                .GroupDescriptions.Add(New PropertyGroupDescription(NameOf(Produit.CodeLydic)))
            End If
        End With
    End Sub

    Private Sub GrouperLesProduitsParUnité()
        With Me.VueSourceProduits
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

    Private Sub UC_SélecteurDeProduit_KeyDown(sender As Object, e As KeyEventArgs) Handles TBx_CritèreCodeLydic.KeyDown, TBx_CritèreMotsClés.KeyDown, TBx_CritèreRefFournisseur.KeyDown
        If e.Key = Key.Return Then
            Me.FiltrerLesProduits()
            Me.SLtr_RésultatRecherche.Focus()
        End If
    End Sub

    Private Sub Btn_Chercher_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Chercher.Click
        Me.FiltrerLesProduits()
    End Sub

#End Region


#Region "Show"

    Private Shared Sub Show_Core()
        'Pas d'erreur d'exécution mais ne montre rien ! Je pense que c'est du au STA qui isole complétement le thread. Après exécution, l'objet reste de toute façon innaccessible directement depuis le reste du code qui tourne dans le thread principal. Donc galère. 
        Dim ShowTask As New Thread(Sub()
                                       _Instance = New UC_SélecteurDeProduit()
                                       _Instance.Margin = New Thickness(0)

                                       Dim w = New Windows.Window With {.Title = "Recherche de produit"}
                                       _Instance._FenêtreParente = w

                                       Dim aw = XL.ActiveWindow
                                       If aw IsNot Nothing Then
                                           Dim hwndHelper = New Interop.WindowInteropHelper(w)
                                           hwndHelper.Owner = New IntPtr(CLng(Globals.ThisAddIn.Application.ActiveWindow?.Hwnd))
                                       End If

                                       With w
                                           w.Content = Instance
                                           Instance._FenêtreParente = w
                                           .Padding = New Thickness(5)
                                           Instance.LoadResourceDict()
                                       End With

                                       AddHandler Instance._FenêtreParente.Closed,
                                                    Sub(sender2, e2)
                                                        w.Dispatcher.InvokeShutdown()
                                                        Threading.Dispatcher.Run()
                                                    End Sub
                                       w.Show()
                                   End Sub)
        ShowTask.SetApartmentState(ApartmentState.STA)
        ShowTask.Start()

    End Sub
    Public Sub Show()
        'Instance.Dispatcher.Invoke(Sub()
        '                               Me._FenêtreParente = w
        '                           End Sub)
        'Dispatcher.Invoke(Sub()
        '                      Me._FenêtreParente = w
        '                  End Sub)

    End Sub
    Public Shared Sub Show2()

        Produit.SetMotsPourTousLesProduits()

        Instance?.Reset()

        Show_Core()
        'Me._FenêtreParente = New Windows.Window With {.Title = "Recherche de produit"}

    End Sub

    'Public Sub Show()

    '    Produit.SetMotsPourTousLesProduits()

    '    Me._FenêtreParente = New Windows.Window With {.Title = "Recherche de produit"}

    '    Dim aw = XL.ActiveWindow
    '    Dim OpenModal = aw IsNot Nothing
    '    If OpenModal Then
    '        Dim hwndHelper = New Interop.WindowInteropHelper(_FenêtreParente)
    '        hwndHelper.Owner = New IntPtr(CLng(Globals.ThisAddIn.Application.ActiveWindow?.Hwnd))
    '        'hwndHelper.Owner = New IntPtr(CLng(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle))
    '    End If

    '    Me.Reset()

    '    With _FenêtreParente
    '        LoadResourceDict()
    '        .Padding = New Thickness(5)
    '        Me.Margin = New Thickness(0)
    '        .Content = Me
    '        'If OpenModal Then
    '        '    .ShowDialog()
    '        'Else
    '        '    .Show()
    '        '    .AddHandler(Windows.Window.LostFocusEvent, New RoutedEventHandler(Sub(w2 As Window, e As RoutedEventArgs)
    '        '                                                                          w2.Topmost = True
    '        '                                                                          'xxx
    '        '                                                                      End Sub))
    '        '    .Topmost = True
    '        'End If
    '    End With
    'End Sub

    'Private Sub Show(w As Window)

    '    Dim ShowTask As New Thread(Sub()
    '                                   w.Show()
    '                                   AddHandler w.Closed,
    '                                                Sub(sender2, e2)
    '                                                    w.Dispatcher.InvokeShutdown()
    '                                                    Threading.Dispatcher.Run()
    '                                                End Sub
    '                               End Sub)
    '    ShowTask.SetApartmentState(ApartmentState.STA)
    '    ShowTask.Start()
    'End Sub

    ''' <summary>Charge dynamiquement le dictionnaire de ressource pour la fenêtre créée. </summary>
    Private Sub LoadResourceDict()
        Application.ResourceAssembly = My.Application.GetType.Assembly 'la prop. est nulle sinon et ça plante après avec un msg qui demande explicitement de la définir.
        'prefix to the relative Uri for resource (xaml file)
        Dim _prefix = $"/{Globals.ThisAddIn.GetType.Namespace};component/"
        Dim URIDico = New Uri(_prefix & "Code/Dico.xaml", UriKind.Relative)
        Dim Dico = New ResourceDictionary With {.Source = URIDico}
        Me.FenêtreParente.Resources.MergedDictionaries.Add(New ResourceDictionary With {.Source = URIDico})
    End Sub

    Private Sub Reset()
        Dim bckup = Me.RechercheSurDemande
        Me.RechercheSurDemande = True
        Me.TBx_CritèreMotsClés.Clear()
        Me.TBx_CritèreCodeLydic.Clear()
        Me.TBx_CritèreRefFournisseur.Clear()
        Me.TBx_CritèreMotsClés.Focus()
        Me.RechercheSurDemande = bckup
        Me.FiltrerLesProduits()
    End Sub

#End Region

#Region "ValiderLeChoix"
    Private Sub ValiderLeChoix()
        If (Me.ProduitSélectionné IsNot Nothing) Then
            RaiseEvent ProduitTrouvé(Me.ProduitSélectionné)
            Me.FenêtreParente?.Close()
        End If
    End Sub
#End Region

#End Region

#Region "Events"

#Region "ProduitTrouvé"

    Public Event ProduitTrouvé(P As Produit)

#End Region

#End Region

#Region "Tests et debuggage"


#End Region

End Class

