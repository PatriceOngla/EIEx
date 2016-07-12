Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports Model

Public Class UC_EtudesView

#Region "Constructeurs"

    Private Sub UC_EtudesView_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        ConnecterLesGestionnairesCRUD()
        AddExcelEventsHandlers()
    End Sub

    Private Sub ConnecterLesGestionnairesCRUD()

        With UC_CmdesCRUD_Etudes

            '.MsgAlerteCohérenceSuppression = "Attention, ce produit est associé à au moins un patron d'ouvrage. En cas de suppression, ce(s) patron(s) predra(ont) leur référence à ce produit."

            .NomEntité = "étude"

            .AssociatedSelector = Me.DG_Master

            '.SuppressionAConfirmer = Function(p As Produit)
            '                             Dim r = (From ro In Ref.PatronsDOuvrage
            '                                      Where ro.UtiliseProduit(p)).Any
            '                             Return r
            '                         End Function
        End With

        With UC_CRUD_Bordereaux

            .NomEntité = "bordereau"

            .AssociatedSelector = Me.DG_Bordereaux

        End With

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

#Region "EtudeCourante"
    Public Property EtudeCourante() As Etude
        Get
            Return WS.EtudeCourante
        End Get
        Set(ByVal value As Etude)
            WS.EtudeCourante = value
        End Set
    End Property
#End Region

#Region "BordereauCourant"
    Public Property BordereauCourant() As Bordereau
        Get
            Return WS.BordereauCourant
        End Get
        Set(ByVal value As Bordereau)
            WS.BordereauCourant = value
        End Set
    End Property
#End Region

#End Region

#Region "Gestionnaires d'évennements"

#Region "CRUD Etudes"

    Private Sub UC_CmdesCRUD_Etudes_DemandeAjout() Handles UC_CmdesCRUD_Etudes.DemandeAjout
        Try
            Dim E = WS.GetNewEtude()
            Me.DG_Master.SelectedItem = E
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

    Private Sub UC_CmdesCRUD_Etudes_DemandeSuppression() Handles UC_CmdesCRUD_Etudes.DemandeSuppression
        Try
            Dim e As Etude = Me.DG_Master.SelectedItem
            WS.Etudes.Remove(e)
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

#End Region

#Region "CRUD Bordereaux"

    Private Sub UC_CRUD_Bordereaux_DemandeAjout() Handles UC_CRUD_Bordereaux.DemandeAjout
        Try
            If Me.EtudeCourante IsNot Nothing Then
                Dim B = Me.EtudeCourante.AjouterNouveauBordereau()
                Me.DG_Bordereaux.SelectedItem = B
            Else
                AlertePasDeDEtudeSélectionnée()
            End If
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

    Private Sub UC_CRUD_Bordereaux_DemandeSuppression() Handles UC_CRUD_Bordereaux.DemandeSuppression
        Try
            If Me.EtudeCourante IsNot Nothing Then
                Dim B As Bordereau = Me.DG_Bordereaux.SelectedItem
                Me.EtudeCourante.Bordereaux.Remove(B)
            Else
                AlertePasDeDEtudeSélectionnée()
            End If
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

    Private Sub AlertePasDeDEtudeSélectionnée()
        Message("Aucune étude sélectionnée.")
    End Sub

#End Region

#Region "Excel events"

    Private Sub AddExcelEventsHandlers()
        'AddHandler ExcelEventManager.TargetSelectedRangeChanged, AddressOf XLSelectionChangeHandling
    End Sub

    'Private Sub XLSelectionChangeHandling(NewSelection As Excel.Range)
    'End Sub

#End Region

#End Region

End Class
