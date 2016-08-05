Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Data
Imports System.Windows.Input
Imports Model
Imports Utils

Public Class UC_SélecteurDOuvrage

#Region "Constructeurs"

    Private Sub UC_SélecteurDOuvrage_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        SetCollectionViewSourceOuvrages()
        Produit.SetMotsPourTousLesProduits()
        Me.SLtr_RésultatRecherche.ItemsSource = CollectionViewSourceOuvrages.View
    End Sub

#End Region

#Region "Propriétés"

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

#Region "Critères"

#Region "CritèreMotsClés (String)"

#Region "Déclaration et registration de CritèreMotsClésProperty"

    Private Shared FlagsMDCritèreMotsClés As FrameworkPropertyMetadataOptions = 0
    Private Shared MDCritèreMotsClés As New FrameworkPropertyMetadata(Nothing, FlagsMDCritèreMotsClés, New PropertyChangedCallback(AddressOf OnCritèreMotsClésInvalidated))
    Public Shared CritèreMotsClésProperty As DependencyProperty = DependencyProperty.Register("CritèreMotsClés", GetType(String), GetType(UC_SélecteurDOuvrage), MDCritèreMotsClés)

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
                                                                                      GetType(RoutedPropertyChangedEventHandler(Of String)), GetType(UC_SélecteurDOuvrage))

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

        Dim Sender As UC_SélecteurDOuvrage = d
        Dim OldValue As String = e.OldValue
        Dim NewValue As String = e.NewValue

        Sender.OnCritèreMotsClésChanged(OldValue, NewValue)

    End Sub

    Private Sub OnCritèreMotsClésChanged(ByVal OldValue As String, ByVal NewValue As String)

        If Object.Equals(NewValue, OldValue) Then Exit Sub

        Dim args As New RoutedPropertyChangedEventArgs(Of String)(OldValue, NewValue)
        args.RoutedEvent = UC_SélecteurDOuvrage.CritèreMotsClésChangedEvent

        'Insérer ici le code spécifique à la gestion du changement de la propriété "CritèreMotsClés"
        If Not Me.RechercheSurDemande Then FiltrerLesOuvrages()

        'Signalement de l'évennement au framework
        If args IsNot Nothing Then Me.RaiseEvent(args)

    End Sub

#End Region

#End Region

#Region "Critères produit"

#Region "CritèreMotsClésProduits (String)"

    Public Shared ReadOnly CritèreMotsClésProduitsProperty As DependencyProperty =
            DependencyProperty.Register("CritèreMotsClésProduits", GetType(String), GetType(UC_SélecteurDOuvrage), New UIPropertyMetadata(Nothing, New PropertyChangedCallback(Sub(Sender As UC_SélecteurDOuvrage, e As DependencyPropertyChangedEventArgs)
                                                                                                                                                                                   Sender.FiltrerLesOuvrages()
                                                                                                                                                                               End Sub)))

    Public Property CritèreMotsClésProduits As String
        Get
            Return DirectCast(GetValue(CritèreMotsClésProduitsProperty), String)
        End Get

        Set(ByVal value As String)
            SetValue(CritèreMotsClésProduitsProperty, value)
        End Set
    End Property

#End Region

#Region "CritèreCodeLydic (String)"

#Region "Déclaration et registration de CritèreCodeLydicProperty"

    Private Shared FlagsMDCritèreCodeLydic As FrameworkPropertyMetadataOptions = 0
    Private Shared MDCritèreCodeLydic As New FrameworkPropertyMetadata(Nothing, FlagsMDCritèreCodeLydic, New PropertyChangedCallback(AddressOf OnCritèreCodeLydicInvalidated))
    Public Shared CritèreCodeLydicProperty As DependencyProperty = DependencyProperty.Register("CritèreCodeLydic", GetType(String), GetType(UC_SélecteurDOuvrage), MDCritèreCodeLydic)

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
                                                                                      GetType(RoutedPropertyChangedEventHandler(Of String)), GetType(UC_SélecteurDOuvrage))

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

        Dim Sender As UC_SélecteurDOuvrage = d
        Dim OldValue As String = e.OldValue
        Dim NewValue As String = e.NewValue

        Sender.OnCritèreCodeLydicChanged(OldValue, NewValue)

    End Sub

    Private Sub OnCritèreCodeLydicChanged(ByVal OldValue As String, ByVal NewValue As String)

        If Object.Equals(NewValue, OldValue) Then Exit Sub

        Dim args As New RoutedPropertyChangedEventArgs(Of String)(OldValue, NewValue)
        args.RoutedEvent = UC_SélecteurDOuvrage.CritèreCodeLydicChangedEvent

        'Insérer ici le code spécifique à la gestion du changement de la propriété "CritèreCodeLydic"
        If Not Me.RechercheSurDemande Then FiltrerLesOuvrages()

        'Signalement de l'évennement au framework
        If args IsNot Nothing Then Me.RaiseEvent(args)

    End Sub

#End Region

#End Region

#Region "CritèreRefFournisseur (String)"

#Region "Déclaration et registration de CritèreRefFournisseurProperty"

    Private Shared FlagsMDCritèreRefFournisseur As FrameworkPropertyMetadataOptions = 0
    Private Shared MDCritèreRefFournisseur As New FrameworkPropertyMetadata(Nothing, FlagsMDCritèreRefFournisseur, New PropertyChangedCallback(AddressOf OnCritèreRefFournisseurInvalidated))
    Public Shared CritèreRefFournisseurProperty As DependencyProperty = DependencyProperty.Register("CritèreRefFournisseur", GetType(String), GetType(UC_SélecteurDOuvrage), MDCritèreRefFournisseur)

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
                                                                                      GetType(RoutedPropertyChangedEventHandler(Of String)), GetType(UC_SélecteurDOuvrage))

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

        Dim Sender As UC_SélecteurDOuvrage = d
        Dim OldValue As String = e.OldValue
        Dim NewValue As String = e.NewValue

        Sender.OnCritèreRefFournisseurChanged(OldValue, NewValue)

    End Sub

    Private Sub OnCritèreRefFournisseurChanged(ByVal OldValue As String, ByVal NewValue As String)

        If Object.Equals(NewValue, OldValue) Then Exit Sub

        Dim args As New RoutedPropertyChangedEventArgs(Of String)(OldValue, NewValue)
        args.RoutedEvent = CritèreRefFournisseurChangedEvent

        'Insérer ici le code spécifique à la gestion du changement de la propriété "CritèreRefFournisseur"
        If Not Me.RechercheSurDemande Then FiltrerLesOuvrages()

        'Signalement de l'évennement au framework
        If args IsNot Nothing Then Me.RaiseEvent(args)

    End Sub

#End Region

#End Region

#End Region

#End Region

#Region "Options de recherche"

#Region "RechercheSurDemande (Boolean)"

    Public Shared ReadOnly RechercheSurDemandeProperty As DependencyProperty =
            DependencyProperty.Register("RechercheSurDemande", GetType(Boolean), GetType(UC_SélecteurDOuvrage), New UIPropertyMetadata(False))

    Public Property RechercheSurDemande As Boolean
        Get
            Return DirectCast(GetValue(RechercheSurDemandeProperty), Boolean)
        End Get

        Set(ByVal value As Boolean)
            SetValue(RechercheSurDemandeProperty, value)
        End Set
    End Property

#End Region

#Region "Source ouvrages"

#Region "ViewSourceOuvrages"
    Public ReadOnly Property ViewSourceOuvrages() As ICollectionView
        Get
            Return CollectionViewSource.GetDefaultView(Me.SLtr_RésultatRecherche.ItemsSource)
        End Get
    End Property
#End Region

#Region "SourceOuvrages"

    Public ReadOnly Property SourceOuvrages As IEnumerable(Of Ouvrage_Base)
        Get
            If Me.LaSourceEstLeRéférentiel Then
                Return Ref.PatronsDOuvrage
            Else
                Return WorkSpace.Instance.EtudeCourante.Ouvrages
            End If
        End Get
    End Property

    Private WithEvents _CollectionViewSourceOuvrages As CollectionViewSource
    Public ReadOnly Property CollectionViewSourceOuvrages() As CollectionViewSource
        Get
            Return _CollectionViewSourceOuvrages
        End Get
    End Property

    Private Sub SetCollectionViewSourceOuvrages()
        _CollectionViewSourceOuvrages = New CollectionViewSource()
        _CollectionViewSourceOuvrages.Source = Me.SourceOuvrages
    End Sub

#End Region

#Region "LaSourceEstLeRéférentiel (Boolean)"

    Public Shared ReadOnly LaSourceEstLeRéférentielProperty As DependencyProperty =
            DependencyProperty.Register("LaSourceEstLeRéférentiel", GetType(Boolean), GetType(UC_SélecteurDOuvrage),
                                        New UIPropertyMetadata(True, New PropertyChangedCallback(
                                                               Sub(Sender As UC_SélecteurDOuvrage, e As DependencyPropertyChangedEventArgs)
                                                                   Ouvrage_Base.SetMotsPourTousLesOuvrages(Sender.LaSourceEstLeRéférentiel, Not Sender.LaSourceEstLeRéférentiel)

                                                                   Sender.SetCollectionViewSourceOuvrages()
                                                               End Sub))
                                                               )

    Public Property LaSourceEstLeRéférentiel As Boolean
        Get
            Return DirectCast(GetValue(LaSourceEstLeRéférentielProperty), Boolean)
        End Get

        Set(ByVal value As Boolean)
            SetValue(LaSourceEstLeRéférentielProperty, value)
        End Set
    End Property

#End Region

#End Region

#Region "DistanceTolérée (short)"

    Public Shared ReadOnly DistanceToléréeProperty As DependencyProperty =
            DependencyProperty.Register("DistanceTolérée", GetType(Short), GetType(UC_SélecteurDOuvrage),
                                        New UIPropertyMetadata(CShort(0), New PropertyChangedCallback(
                                                               Sub(Sender As UC_SélecteurDOuvrage,
                                                                   e As DependencyPropertyChangedEventArgs)
                                                                   If e.NewValue > 0 Then
                                                                       Sender.RechercheSurDemande = True
                                                                   End If
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

#Region "OuvrageSélectionné (Ouvrage_Base)"

#Region "Déclaration et registration de OuvrageSélectionnéProperty"

    Private Shared FlagsMDOuvrageSélectionné As FrameworkPropertyMetadataOptions = FrameworkPropertyMetadataOptions.BindsTwoWayByDefault
    Private Shared MDOuvrageSélectionné As New FrameworkPropertyMetadata(Nothing, FlagsMDOuvrageSélectionné, New PropertyChangedCallback(AddressOf OnOuvrageSélectionnéInvalidated))
    Public Shared OuvrageSélectionnéProperty As DependencyProperty = DependencyProperty.Register("OuvrageSélectionné", GetType(Ouvrage_Base), GetType(UC_SélecteurDOuvrage), MDOuvrageSélectionné)

#End Region

#Region "Wrapper CLR de OuvrageSélectionnéProperty"
    Public Property OuvrageSélectionné() As Ouvrage_Base
        Get
            Return GetValue(OuvrageSélectionnéProperty)
        End Get
        Set(ByVal value As Ouvrage_Base)
            SetValue(OuvrageSélectionnéProperty, value)
        End Set
    End Property
#End Region

#Region "Gestion évennementielle de la mise à jour de OuvrageSélectionnéProperty"

#Region "Evènnement OuvrageSélectionnéChangedEvent et son Wrapper CLR (Non testé !!!)"

    Public Shared ReadOnly OuvrageSélectionnéChangedEvent As RoutedEvent =
                  EventManager.RegisterRoutedEvent("OuvrageSélectionnéChangedEvent", RoutingStrategy.Bubble,
                                                                                      GetType(RoutedPropertyChangedEventHandler(Of Ouvrage)), GetType(UC_SélecteurDOuvrage))

    Custom Event OuvrageSélectionnéChanged As RoutedEventHandler
        AddHandler(ByVal value As RoutedEventHandler)
            Me.AddHandler(OuvrageSélectionnéChangedEvent, value)
        End AddHandler

        RemoveHandler(ByVal value As RoutedEventHandler)
            Me.RemoveHandler(OuvrageSélectionnéChangedEvent, value)
        End RemoveHandler

        RaiseEvent(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs)
            sender.RaiseEvent(e)
        End RaiseEvent
    End Event


#End Region

    Private Shared Sub OnOuvrageSélectionnéInvalidated(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)

        Dim Sender As UC_SélecteurDOuvrage = d
        Dim OldValue As Ouvrage_Base = e.OldValue
        Dim NewValue As Ouvrage_Base = e.NewValue

        Sender.OnOuvrageSélectionnéChanged(OldValue, NewValue)

    End Sub

    Private Sub OnOuvrageSélectionnéChanged(ByVal OldValue As Ouvrage_Base, ByVal NewValue As Ouvrage_Base)

        If Object.Equals(NewValue, OldValue) Then Exit Sub

        Dim args As New RoutedPropertyChangedEventArgs(Of Ouvrage_Base)(OldValue, NewValue)
        args.RoutedEvent = UC_SélecteurDOuvrage.OuvrageSélectionnéChangedEvent

        'Insérer ici le code spécifique à la gestion du changement de la propriété "OuvrageSélectionné"

        'Signalement de l'évennement au framework
        If args IsNot Nothing Then Me.RaiseEvent(args)

    End Sub

#End Region

#End Region

#Region "InfosResultat (String)"

#Region "Déclaration et registration de InfosResultatProperty"

    Private Shared MDInfosResultat As New FrameworkPropertyMetadata(Nothing)
    Public Shared InfosResultatPropertyKey As DependencyPropertyKey = DependencyProperty.RegisterReadOnly("InfosResultat", GetType(String), GetType(UC_SélecteurDOuvrage), MDInfosResultat)
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
            Return Me.ViewSourceOuvrages.OfType(Of Ouvrage_Base).Count()
        End Get
    End Property
#End Region

#Region "EntêteRésultats"
    Public ReadOnly Property EntêteRésultats() As String
        Get
            Return Ouvrage_Base.OuvragesListHeader
        End Get
    End Property
#End Region

#End Region

#Region "Méthodes"

#Region "Tri, filtre, regroupement"

#Region "Filtre"

    Private Sub FiltrerLesOuvrages()
        'Dim r = From o In Me.SourceOuvrages Where MatcheCritères(o)
        'Me.SetValue(InfosResultatPropertyKey, r.Count & " ouvrage(s)")
        Me.ViewSourceOuvrages.Filter = Function(o As Ouvrage_Base) MatcheCritères(o)
        Me.SetValue(InfosResultatPropertyKey, $"{Me.NbRésultats} ouvrage(s) trouvé(s)")
    End Sub

    Private Sub _SourceProduits_Filter(sender As Object, e As FilterEventArgs) Handles _CollectionViewSourceOuvrages.Filter
        Dim O As Ouvrage_Base = e.Item
        If O IsNot Nothing Then
            e.Accepted = MatcheCritères(e.Item)
        End If
    End Sub

#Region "MatcheCritères"

    Private Function MatcheCritères(O As Ouvrage_Base) As Boolean

        'Return True

        Dim MatchMotsClés, MatchProduits As Boolean

        MatchMotsClés = MatcheCritèresMotsClés(O, Me.CritèreMotsClés, Me.DistanceTolérée)

        If MatchMotsClés Then
            MatchProduits = MatcheCritèresProduits(O, Me.CritèreMotsClésProduits, Me.CritèreCodeLydic, Me.CritèreRefFournisseur)
        End If

        Dim r = MatchMotsClés AndAlso MatchProduits
        Return r

    End Function

    ''' <param name="DistanceTolérée">Distance de Levenstein tolérée.</param>
    ''' <returns></returns>
    Private Shared Function MatcheCritèresMotsClés(O As Ouvrage_Base, C As String, DistanceTolérée As Short) As Boolean
        Dim r = String.IsNullOrEmpty(C)
        If Not r Then
            Dim TabMotsClés = C.Split({" "c, "'"c}, StringSplitOptions.RemoveEmptyEntries)
            r = O.Mots.ContainsList_String(TabMotsClés, True, True, DistanceTolérée)
        End If
        Return r
    End Function

    Private Function MatcheCritèresProduits(o As Ouvrage_Base, CritèreMotsClés As String, critèreCodeLydic As String, critèreRefFournisseur As String) As Boolean
        Dim r = (From up In o.UsagesDeProduit Where UC_SélecteurDeProduit.MatcheCritères(up.Produit, CritèreMotsClés, critèreCodeLydic, critèreRefFournisseur)).Any()
        Return r
    End Function

#End Region

#End Region

#Region "Tri"

    'Private Sub TrierLesProduitsParFournisseurEtParRéférence()
    '    With Me.ViewSourceOuvrages
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
        With Me.ViewSourceOuvrages
            If?.CanGroup Then
                '.GroupDescriptions.Clear()
                .GroupDescriptions.Add(New PropertyGroupDescription(NameOf(Produit.CodeLydic)))
            End If
        End With
    End Sub

    Private Sub GrouperLesProduitsParUnité()
        With Me.ViewSourceOuvrages
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

    Private Sub UC_SélecteurDOuvrage_KeyDown(sender As Object, e As KeyEventArgs) Handles TBx_CritèreCodeLydic.KeyDown, TBx_CritèreMotsClés.KeyDown, TBx_CritèreRefFournisseur.KeyDown
        If e.Key = Key.Return Then
            Me.FiltrerLesOuvrages()
            Me.SLtr_RésultatRecherche.Focus()
        End If
    End Sub

    Private Sub Btn_Chercher_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Chercher.Click
        Me.FiltrerLesOuvrages()
    End Sub

#End Region

#Region "Show"

    Public Sub Show()

        Ouvrage_Base.SetMotsPourTousLesOuvrages(Me.LaSourceEstLeRéférentiel, Not Me.LaSourceEstLeRéférentiel)

        Me._FenêtreParente = New Windows.Window With {.Title = "Recherche d'ouvrage"}

        Dim aw = XL.ActiveWindow
        Dim OpenModal = aw IsNot Nothing
        If OpenModal Then
            Dim hwndHelper = New Interop.WindowInteropHelper(_FenêtreParente)
            hwndHelper.Owner = New IntPtr(CLng(Globals.ThisAddIn.Application.ActiveWindow?.Hwnd))
            'hwndHelper.Owner = New IntPtr(CLng(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle))
        End If

        Me.Reset()

        With _FenêtreParente
            LoadResourceDict()
            .Padding = New Thickness(5)
            Me.Margin = New Thickness(0)
            .Content = Me
            If OpenModal Then
                .ShowDialog()
            Else
                .Show()
                .AddHandler(Windows.Window.LostFocusEvent, New RoutedEventHandler(Sub(w2 As Window, e As RoutedEventArgs)
                                                                                      w2.Topmost = True
                                                                                      'xxx
                                                                                  End Sub))
                .Topmost = True
            End If
        End With

    End Sub

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
        Me.FiltrerLesOuvrages()
    End Sub

#End Region

#Region "ValiderLeChoix"
    Private Sub ValiderLeChoix()
        If (Me.OuvrageSélectionné IsNot Nothing) Then
            RaiseEvent OuvrageTrouvé(Me.OuvrageSélectionné)
            Me.FenêtreParente?.Close()
        End If
    End Sub
#End Region

#End Region

#Region "Events"

#Region "OuvrageTrouvé"

    Public Event OuvrageTrouvé(O As Ouvrage_Base)

#End Region

#End Region

#Region "Tests et debuggage"


#End Region

End Class

