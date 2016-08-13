Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports Model
Imports Utils

Public Class UC_OuvragesView

#Region "Champs privés"

    Friend WithEvents UCSO As UC_SélecteurDOuvrage

#End Region

#Region "Constructeurs"

    Private Sub UC_OuvragesView_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized

        'UCSO = ThisAddIn.UC_Container.EIEx_Manager_UI.UCSO
        UCSO = UC_EIEx_Manager_UI.Instance.UCSO

        With UC_CmdesCRUD_Ouvrages

            '.MsgAlerteCohérenceSuppression = "Attention, ce produit est associé à au moins un patron d'ouvrage. En cas de suppression, ce(s) patron(s) predra(ont) leur référence à ce produit."

            .NomEntité = "patron d'ouvrage"

            .AssociatedSelector = Me.DG_Master

        End With

        With UC_CmdCRUD_UsagesProduit

            .NomEntité = "usage de produit"

            .AssociatedSelector = Me.DG_Produits

        End With

        With UC_CmdCRUD_Libellés

            .NomEntité = "libellé"

            .AssociatedSelector = Me.LBx_Libellés

        End With

    End Sub

#End Region

#Region "Propriétés"

#Region "Ref"
    Public ReadOnly Property Ref() As Référentiel
        Get
            Return Référentiel.Instance
        End Get
    End Property
#End Region

#Region "OuvrageCourant (Ouvrage)"

    Public Shared ReadOnly OuvrageCourantProperty As DependencyProperty =
            DependencyProperty.Register(NameOf(OuvrageCourant), GetType(Ouvrage_Base), GetType(UC_OuvragesView), New UIPropertyMetadata(Nothing, New PropertyChangedCallback(
                                            Sub(ucov As UC_OuvragesView, e As DependencyPropertyChangedEventArgs)
                                                Try
                                                    Dim o = TryCast(ucov.OuvrageCourant, Ouvrage)
                                                    If o IsNot Nothing Then
                                                        Dim r As Excel.Range
                                                        r = o.GetCelluleExcelAssociée
                                                        If r IsNot Nothing Then
                                                            SélectionnerPlageExcel(r)
                                                        End If
                                                    End If
                                                Catch ex As Exception
                                                    ManageErreur(ex, "Echec de la sélection de la cellule Excel associée.")
                                                End Try
                                            End Sub)))


    Public Property OuvrageCourant As Ouvrage_Base
        Get
            Return DirectCast(GetValue(OuvrageCourantProperty), Ouvrage_Base)
        End Get

        Set(ByVal value As Ouvrage_Base)
            SetValue(OuvrageCourantProperty, value)
        End Set
    End Property

#End Region

#Region "SélecteurDeProduit"
    Private WithEvents _SélecteurDeProduit As UC_SélecteurDeProduit
    Public ReadOnly Property SélecteurDeProduit() As UC_SélecteurDeProduit
        Get
            If _SélecteurDeProduit Is Nothing Then
                _SélecteurDeProduit = New UC_SélecteurDeProduit
            End If
            Return _SélecteurDeProduit
        End Get
    End Property
#End Region

    '#Region "EditeUnOuvrage"
    '    Public ReadOnly Property EditeUnOuvrage() As Boolean
    '        Get
    '            Return TypeOf Me.OuvrageCourant Is Ouvrage
    '        End Get
    '    End Property
    '#End Region

#Region "FenêtreParente"
    Private _FenêtreParente As Window
    Public ReadOnly Property FenêtreParente() As Window
        Get
            Return _FenêtreParente
        End Get
    End Property
#End Region

    'Attention ! Ce snippet est une ébauche !!! 
#Region "CanModify (Boolean)"

#Region "Déclaration et registration de CanModifyProperty"

    Private Shared MDCanModify As New FrameworkPropertyMetadata(True)
    Public Shared CanModifyPropertyKey As DependencyPropertyKey = DependencyProperty.RegisterReadOnly(NameOf(CanModify), GetType(Boolean), GetType(UC_OuvragesView), MDCanModify)
    Public Shared CanModifyProperty As DependencyProperty = CanModifyPropertyKey.DependencyProperty

#End Region

#Region "Wrapper CLR de CanModifyProperty"

    Public ReadOnly Property CanModify() As Boolean
        Get
            Return GetValue(CanModifyProperty)
        End Get
    End Property

#End Region

#Region "Gestion évennementielle de la mise à jour de CanModifyProperty"

#End Region

#End Region

#End Region

#Region "Méthodes"

#Region "Gestionnaires d'évennements"

#Region "CRUD"

#Region "CRUD Ouvrage"
    Private Sub UC_CmdesCRUD_DemandeAjout() Handles UC_CmdesCRUD_Ouvrages.DemandeAjout
        Try
            Dim OuvrageProduit = Ref.GetNewOuvrage()
            Me.DG_Master.SelectedItem = OuvrageProduit
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

    Private Sub UC_CmdesCRUD_DemandeSuppression() Handles UC_CmdesCRUD_Ouvrages.DemandeSuppression
        Try
            Dim Ouvrage As Ouvrage_Base = Me.DG_Master.SelectedItem
            Ref.PatronsDOuvrage.Remove(Ouvrage)
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

#End Region

#Region "CRUD Usage de produits"
    Private Sub UC_CmdCRUD_UsagesProduit_DemandeAjout() Handles UC_CmdCRUD_UsagesProduit.DemandeAjout
        Try
            If Me.OuvrageCourant IsNot Nothing Then
                Dim UsageProduit = Me.OuvrageCourant.AjouterProduit(Nothing, 1)
                Me.DG_Produits.SelectedItem = UsageProduit
            Else
                AlertePasDOuvrageSélectionné()
            End If
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

    Private Sub UC_CmdCRUD_UsagesProduit_DemandeSuppression() Handles UC_CmdCRUD_UsagesProduit.DemandeSuppression
        Try
            If Me.OuvrageCourant IsNot Nothing Then
                Dim UsageProduit As UsageDeProduit = Me.DG_Produits.SelectedItem
                Me.OuvrageCourant.UsagesDeProduit.Remove(UsageProduit)
            Else
                AlertePasDOuvrageSélectionné()
            End If
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

    Private Sub Btn_RechercherProduit_Click(sender As Object, e As RoutedEventArgs)
        'Dim BoutonRecherche As Button = TryCast(e.OriginalSource, Button)
        'If BoutonRecherche?.Tag = "" Then
        '    MsgBox("ok")
        'End If
        Me.SélecteurDeProduit.Show()
    End Sub

    Private Sub _SélecteurDeProduit_ProduitTrouvé(P As Produit) Handles _SélecteurDeProduit.ProduitTrouvé
        Dim up As UsageDeProduit = Me.DG_Produits.SelectedItem
        If up IsNot Nothing Then up.Produit = P
    End Sub

#End Region

#Region "CRUD libellés"

    Private Sub UC_CmdCRUD_Libellés_DemandeAjout() Handles UC_CmdCRUD_Libellés.DemandeAjout
        Try
            If Me.OuvrageCourant IsNot Nothing Then
                Dim NouveauLibellé = InputBox("Nouveau libellé : ", ThisAddIn.Nom)
                If Not String.IsNullOrEmpty(NouveauLibellé) Then
                    Me.OuvrageCourant.Libellés.Add(NouveauLibellé)
                End If
            Else
                AlertePasDOuvrageSélectionné()
            End If
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

    Private Sub UC_CmdCRUD_Libellés_DemandeSuppression() Handles UC_CmdCRUD_Libellés.DemandeSuppression
        Try
            If Me.OuvrageCourant IsNot Nothing Then
                Dim LibelléSélectionné As String = Me.LBx_Libellés.SelectedItem
                If LibelléSélectionné IsNot Nothing Then
                    Me.OuvrageCourant.Libellés.Remove(LibelléSélectionné)
                Else
                    Message("Aucun libellé sélectionné.")
                End If
            Else
                AlertePasDOuvrageSélectionné()
            End If
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

#End Region

    Private Sub AlertePasDOuvrageSélectionné()
        Message("Aucun patron d'ouvrage sélectionné.")
    End Sub

#End Region

#Region "Divers"

    Private Sub Btn_ResetTempsDePose_Click(sender As Object, e As RoutedEventArgs) Handles Btn_ResetTempsDePose.Click
        If Me.OuvrageCourant IsNot Nothing Then
            Me.OuvrageCourant.TempsDePoseUnitaire = Nothing
        End If
    End Sub

    Private Sub Btn_ResetPrixUnitaire_Click(sender As Object, e As RoutedEventArgs) Handles Btn_ResetPrixUnitaire.Click
        If Me.OuvrageCourant IsNot Nothing Then
            Me.OuvrageCourant.PrixUnitaire = Nothing
        End If
    End Sub

#Region "Recherche"

    Private Sub UC_OuvragesView_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        Try
            If e.Key = Key.F AndAlso e.KeyboardDevice.Modifiers = ModifierKeys.Control Then
                Dim result = RechercheOuvrage()
                If result IsNot Nothing Then Me.OuvrageCourant = UCSO.Résultat
            End If
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

    Private Function RechercheOuvrage() As Ouvrage_Base
        Dim result = UCSO.Show()
        Return result
    End Function

    'Private Sub UCSO_OuvrageTrouvé(O As Ouvrage_Base) Handles UCSO.OuvrageTrouvé
    '    Try
    '        Me.OuvrageCourant = O
    '    Catch ex As Exception
    '        ManageErreur(ex)
    '    End Try
    'End Sub

#End Region

    Private Sub Btn_AppliquerModèle_Click(sender As Object, e As RoutedEventArgs) Handles Btn_AppliquerModèle.Click
        Dim Modèle = RechercheOuvrage()
        If Modèle IsNot Nothing Then
            Me.OuvrageCourant.Copier(Modèle)
        End If
    End Sub

#End Region

#End Region

#Region "Show"

    Public Sub Show(Titre As String, Source As IEnumerable(Of Ouvrage_Base))

        Me.SetValue(CanModifyPropertyKey, False)

        Me.DataContext = Source

        SetParentWindow(Titre)

    End Sub

    Private Sub SetParentWindow(Titre As String)

        Me._FenêtreParente = New Windows.Window With {.Title = Titre}

        Dim hwndHelper = New Interop.WindowInteropHelper(_FenêtreParente)
        hwndHelper.Owner = New IntPtr(CLng(Globals.ThisAddIn.Application.ActiveWindow?.Hwnd))
        'hwndHelper.Owner = New IntPtr(CLng(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle))

        With _FenêtreParente
            LoadResourceDict()
            .Padding = New Thickness(5)
            Me.Margin = New Thickness(0)
            .Content = Me
            .ShowDialog()
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

#End Region


#End Region

End Class
