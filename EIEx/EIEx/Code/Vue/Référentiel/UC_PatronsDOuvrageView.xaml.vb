Imports System.Windows
Imports System.Windows.Controls
Imports Model
Imports Utils

Public Class UC_PatronsDOuvrageView

#Region "Constructeurs"

    Private Sub UC_PatronsDOuvrageView_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized

        With UC_CmdesCRUD_Ouvrages

            '.MsgAlerteCohérenceSuppression = "Attention, ce produit est associé à au moins un patron d'ouvrage. En cas de suppression, ce(s) patron(s) predra(ont) leur référence à ce produit."

            .NomEntité = "patron d'ouvrage"

            .AssociatedSelector = Me.DG_Master

            '.SuppressionAConfirmer = Function(p As Produit)
            '                             Dim r = (From ro In Ref.PatronsDOuvrage
            '                                      Where ro.UtiliseProduit(p)).Any
            '                             Return r
            '                         End Function
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

#Region "PatronDOuvrageCourant (PatronDOuvrage)"

    Public Shared ReadOnly PatronDOuvrageCourantProperty As DependencyProperty =
            DependencyProperty.Register(NameOf(PatronDOuvrageCourant), GetType(PatronDOuvrage), GetType(UC_PatronsDOuvrageView), New UIPropertyMetadata(Nothing))

    Public Property PatronDOuvrageCourant As PatronDOuvrage
        Get
            Return DirectCast(GetValue(PatronDOuvrageCourantProperty), PatronDOuvrage)
        End Get

        Set(ByVal value As PatronDOuvrage)
            SetValue(PatronDOuvrageCourantProperty, value)
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

#Region "Gestionnaires d'évennements"

#Region "CRUD"

#Region "CRUD Ouvrage"
    Private Sub UC_CmdesCRUD_DemandeAjout() Handles UC_CmdesCRUD_Ouvrages.DemandeAjout
        Try
            Dim OuvrageProduit = Ref.GetNewPatronDOuvrage()
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
            If Me.PatronDOuvrageCourant IsNot Nothing Then
                Dim UsageProduit = Me.PatronDOuvrageCourant.AjouterProduit(Nothing, 1)
                Me.DG_Produits.SelectedItem = UsageProduit
            Else
                AlertePasDePatronDOuvrageSélectionné()
            End If
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

    Private Sub UC_CmdCRUD_UsagesProduit_DemandeSuppression() Handles UC_CmdCRUD_UsagesProduit.DemandeSuppression
        Try
            If Me.PatronDOuvrageCourant IsNot Nothing Then
                Dim UsageProduit As UsageDeProduit = Me.DG_Produits.SelectedItem
                Me.PatronDOuvrageCourant.UsagesDeProduit.Remove(UsageProduit)
            Else
                AlertePasDePatronDOuvrageSélectionné()
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
            If Me.PatronDOuvrageCourant IsNot Nothing Then
                Dim NouveauLibellé = InputBox("Nouveau libellé : ", ThisAddIn.Nom)
                If Not String.IsNullOrEmpty(NouveauLibellé) Then
                    Me.PatronDOuvrageCourant.Libellés.Add(NouveauLibellé)
                End If
            Else
                AlertePasDePatronDOuvrageSélectionné()
            End If
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

    Private Sub UC_CmdCRUD_Libellés_DemandeSuppression() Handles UC_CmdCRUD_Libellés.DemandeSuppression
        Try
            If Me.PatronDOuvrageCourant IsNot Nothing Then
                Dim LibelléSélectionné As String = Me.LBx_Libellés.SelectedItem
                If LibelléSélectionné IsNot Nothing Then
                    Me.PatronDOuvrageCourant.Libellés.Remove(LibelléSélectionné)
                Else
                    Message("Aucun libellé sélectionné.")
                End If
            Else
                AlertePasDePatronDOuvrageSélectionné()
            End If
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

#End Region

    Private Sub AlertePasDePatronDOuvrageSélectionné()
        Message("Aucun patron d'ouvrage sélectionné.")
    End Sub

#End Region

#Region "Divers"

    Private Sub Btn_ResetTempsDePause_Click(sender As Object, e As RoutedEventArgs) Handles Btn_ResetTempsDePause.Click
        Me.PatronDOuvrageCourant.TempsDePauseUnitaire = Nothing
    End Sub

    Private Sub Btn_ResetPrixUnitaire_Click(sender As Object, e As RoutedEventArgs) Handles Btn_ResetPrixUnitaire.Click
        Me.PatronDOuvrageCourant.PrixUnitaire = Nothing
    End Sub



#End Region

#End Region

End Class
