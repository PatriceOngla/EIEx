Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports Model
Imports Utils

Public Class UC_OuvragesView

#Region "Champs privés"

    Private WithEvents UCSO As New UC_SélecteurDOuvrage()

#End Region

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
            DependencyProperty.Register(NameOf(OuvrageCourant), GetType(PatronDOuvrage), GetType(UC_OuvragesView), New UIPropertyMetadata(Nothing))

    Public Property OuvrageCourant As PatronDOuvrage
        Get
            Return DirectCast(GetValue(OuvrageCourantProperty), PatronDOuvrage)
        End Get

        Set(ByVal value As PatronDOuvrage)
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
            Dim Ouvrage As PatronDOuvrage = Me.DG_Master.SelectedItem
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

    Private Sub Btn_ResetTempsDePause_Click(sender As Object, e As RoutedEventArgs) Handles Btn_ResetTempsDePause.Click
        If Me.OuvrageCourant IsNot Nothing Then
            Me.OuvrageCourant.TempsDePauseUnitaire = Nothing
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
                UCSO.Show()
            End If
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

    Private Sub UCSO_OuvrageTrouvé(O As Ouvrage_Base) Handles UCSO.OuvrageTrouvé
        Try
            Me.OuvrageCourant = O
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

#End Region

#End Region

#End Region

#End Region

End Class
