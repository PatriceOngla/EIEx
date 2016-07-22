Imports System.Diagnostics
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

        With UC_CRUD_Classeurs

            .NomEntité = "classeur Excel"

            .AssociatedSelector = Me.DG_ClasseursExcel

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

#Region "XL"
    Public ReadOnly Property XL() As Excel.Application
        Get
            Return ExcelCommander.XL
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

#Region "ClasseurExcelCourant"
    Public Property ClasseurExcelCourant() As ClasseurExcel
        Get
            Return WS.ClasseurExcelCourant
        End Get
        Set(ByVal value As ClasseurExcel)
            WS.ClasseurExcelCourant = value
        End Set
    End Property

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

#Region "Gestion des classeurs associés à l'étude"

#Region "InitialiserLesClasseursExcelDeLEtudeCopurante"

    Friend Sub InitialiserLesClasseursExcelDeLEtudeCopurante()
        Try
            Dim NewC As ClasseurExcel

            If Me.EtudeCourante Is Nothing Then Throw New Exception("Pas d'étude courante.")
            With Me.EtudeCourante
                For Each wb As Excel.Workbook In XL.Workbooks
                    If Not ContientLeClasseur(wb.FullName) Then
                        NewC = .AjouterNouveauClasseur()
                        NewC.CheminFichier = wb.FullName
                        AjouterLesFeuilles(NewC)
                        NewC.Nom = wb.Name
                    End If
                Next
                If Me.ClasseurExcelCourant Is Nothing Then
                    Me.ClasseurExcelCourant = Me.EtudeCourante.ClasseursExcel.FirstOrDefault
                End If
            End With
        Catch ex As Exception
            ManageErreur(ex, , True, False)
        End Try
    End Sub

    Private Sub AjouterLesFeuilles(C As ClasseurExcel)
        Dim WShts As Microsoft.Office.Interop.Excel.Sheets = C.ClasseurRéel?.Worksheets
        'Dim WShts2 As Microsoft.Office.Tools.Excel.Worksheet.Worksheets = C.ClasseurRéel?.Worksheets
        'Dim WShts As Excel.Worksheets = C.ClasseurRéel?.Worksheets
        Dim B As Bordereau
        Debug.Print(WShts.Count())
        If WShts IsNot Nothing Then
            For Each Wsht In WShts
                B = C.AjouterNouveauBordereau()
                B.Nom = Wsht.Name
                B.NomFeuille = B.Nom
            Next
        End If

    End Sub

    Private Function ContientLeClasseur(Chemin As String) As Boolean
        Dim r As Boolean
        If Me.EtudeCourante Is Nothing Then
            r = False
        Else
            r = (From c In Me.EtudeCourante.ClasseursExcel Where Object.Equals(c.CheminFichier, Chemin)).Any()
        End If
        Return r
    End Function

#End Region

#Region "ChargerLesClasseursExcelDeLEtudeCopurante"

    Friend Sub ChargerLesClasseursExcelDeLEtudeCopurante()
        Try
            With Me.EtudeCourante
                For Each c In .ClasseursExcel
                    Try
                        If IO.File.Exists(c.CheminFichier) Then
                            XL.Workbooks.Open(c.CheminFichier)
                        Else
                            Message($"Le classeur ""{c.CheminFichier}"" est introuvable.)", MsgBoxStyle.Exclamation)
                        End If
                    Catch ex As Exception
                        ManageErreur(ex, $"Echec de la tentative d'ouverture du classeur ""{c.CheminFichier}""", True, False)
                    End Try
                Next
            End With
        Catch ex As Exception
            ManageErreur(ex, , True, False)
        End Try
    End Sub

#End Region

#End Region

#End Region

#Region "CRUD ClasseursExcel"

    Private Sub UC_CRUD_ClasseursExcel_DemandeAjout() Handles UC_CRUD_Classeurs.DemandeAjout
        Try
            If Me.EtudeCourante IsNot Nothing Then
                Dim C = Me.EtudeCourante.AjouterNouveauClasseur()
                Me.DG_ClasseursExcel.SelectedItem = C
            Else
                AlertePasDeDEtudeSélectionnée()
            End If
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

    Private Sub UC_CRUD_ClasseursExcel_DemandeSuppression() Handles UC_CRUD_Classeurs.DemandeSuppression
        Try
            If Me.EtudeCourante IsNot Nothing Then
                Dim C As ClasseurExcel = Me.DG_ClasseursExcel.SelectedItem
                Me.EtudeCourante.ClasseursExcel.Remove(C)
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

#Region "CRUD Bordereaux"

    Private Sub UC_CRUD_Bordereaux_DemandeAjout() Handles UC_CRUD_Bordereaux.DemandeAjout
        Try
            If Me.ClasseurExcelCourant IsNot Nothing Then
                Dim B = Me.ClasseurExcelCourant.AjouterNouveauBordereau()
                Me.DG_Bordereaux.SelectedItem = B
            Else
                AlertePasDeDeClasseurSélectionné()
            End If
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

    Private Sub UC_CRUD_Bordereaux_DemandeSuppression() Handles UC_CRUD_Bordereaux.DemandeSuppression
        Try
            If Me.ClasseurExcelCourant IsNot Nothing Then
                Dim B As Bordereau = Me.DG_Bordereaux.SelectedItem
                Me.ClasseurExcelCourant.Bordereaux.Remove(B)
            Else
                AlertePasDeDeClasseurSélectionné()
            End If
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

    Private Sub AlertePasDeDeClasseurSélectionné()
        Message("Aucun classeur sélectionné.")
    End Sub

#End Region

#End Region

#Region "Excel events"

    Private Sub AddExcelEventsHandlers()
        'AddHandler ExcelEventManager.TargetSelectedRangeChanged, AddressOf XLSelectionChangeHandling
    End Sub

    Private Sub UC_EtudesView_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        If e.Key = Key.S AndAlso e.KeyboardDevice.Modifiers = ModifierKeys.Control Then
            Try
                EIExData.EnregistrerLeWorkspace()
            Catch ex As Exception
                ManageErreur(ex, "Echec de l'enregistrement de l'espace de travail.")
            End Try
        End If
    End Sub

    'Private Sub XLSelectionChangeHandling(NewSelection As Excel.Range)
    'End Sub

#End Region

End Class
