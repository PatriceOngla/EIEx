Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports Excel = Microsoft.Office.Interop.Excel
Imports Model
Imports Utils

Public Class UC_OuvragesView

#Region "Constructeurs"

    Private Sub UC_OuvragesView_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized

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

    '#Region "SélecteurDeProduit"
    '    Private WithEvents _SélecteurDeProduit As New Win_SélecteurDeProduit

    '    Public ReadOnly Property SélecteurDeProduit() As Win_SélecteurDeProduit
    '        Get
    '            Return _SélecteurDeProduit
    '        End Get
    '    End Property
    '#End Region

    '#Region "EditeUnOuvrage"
    '    Public ReadOnly Property EditeUnOuvrage() As Boolean
    '        Get
    '            Return TypeOf Me.OuvrageCourant Is Ouvrage
    '        End Get
    '    End Property
    '#End Region

    '#Region "FenêtreParente"
    '    Private _FenêtreParente As Window
    '    Public ReadOnly Property FenêtreParente() As Window
    '        Get
    '            Return _FenêtreParente
    '        End Get
    '    End Property
    '#End Region

#Region "CanModify (Boolean)"

    Public Shared ReadOnly CanModifyProperty As DependencyProperty =
            DependencyProperty.Register("CanModify", GetType(Boolean), GetType(UC_OuvragesView), New UIPropertyMetadata(True))

    Public Property CanModify As Boolean
        Get
            Return DirectCast(GetValue(CanModifyProperty), Boolean)
        End Get

        Set(ByVal value As Boolean)
            SetValue(CanModifyProperty, value)
        End Set
    End Property

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
        Dim ProduitTrouvé = Win_SélecteurDeProduit.Cherche()
        If ProduitTrouvé IsNot Nothing Then
            Dim up As UsageDeProduit = Me.DG_Produits.SelectedItem
            If up IsNot Nothing Then up.Produit = ProduitTrouvé
        End If
    End Sub

#End Region

#Region "CRUD libellés"

    Private Sub UC_CmdCRUD_Libellés_DemandeAjout() Handles UC_CmdCRUD_Libellés.DemandeAjout
        Try
            If Me.OuvrageCourant IsNot Nothing Then
                Dim NouveauLibellé = InputBox("Nouveau libellé : ", Application.Nom)
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

    Private Sub Btn_Load_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Load.Click
        Win_Main.RechargerRéférentiel()
    End Sub

    Private Sub Btn_Save_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Save.Click
        Win_Main.EnregistrerRéférentiel()
    End Sub

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
                If result IsNot Nothing Then Me.OuvrageCourant = result
            End If
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

    Private Function RechercheOuvrage() As Ouvrage_Base
        Dim result = Win_SélecteurDOuvrage.Cherche()
        Return result
    End Function

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

#End Region

#End Region

End Class
