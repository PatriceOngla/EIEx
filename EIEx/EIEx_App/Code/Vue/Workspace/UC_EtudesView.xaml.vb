﻿Imports System.Diagnostics
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports EIEx_DAO
Imports Excel = Microsoft.Office.Interop.Excel
Imports Model
Imports Utils
Imports System.ComponentModel

Public Class UC_EtudesView
    Implements INotifyPropertyChanged

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
            '                             Dim r = (From ro In Ref.Ouvrages
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
            NotifyPropertyChanged(NameOf(EtudeCourante))
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

#End Region

#Region "BordereauCourant"
    Public Property BordereauCourant() As Bordereau
        Get
            Return WS.BordereauCourant
        End Get
        Set(ByVal value As Bordereau)
            WS.BordereauCourant = value
            ActiverLaWorksheetDuBordereauCourant()
        End Set
    End Property

    Private Sub ActiverLaWorksheetDuBordereauCourant()
        Try
            If Me.BordereauCourant IsNot Nothing Then
                With Me.BordereauCourant
                    If .Parent.ClasseurRéel IsNot Nothing Then
                        .Parent.ClasseurRéel.Activate()
                        .Worksheet.Activate()
                    End If
                End With
            End If
        Catch ex As Exception
            Debug.Print("ActiverLaWorksheetDuBordereauCourant a échoué" & vbCr & ex.ToString)
        End Try
    End Sub
#End Region

#End Region

#Region "Méthodes"

#Region "Gestionnaires d'évennements"

#Region "Global"

    Private Sub Btn_Save_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Save.Click
        Win_Main.SaveWorkspace()
    End Sub

    Private Sub Btn_Reload_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Reload.Click
        Win_Main.RechargerWorkspace()
    End Sub

    Private Sub Btn_OuvrirUneAutreEtude_Click(sender As Object, e As RoutedEventArgs) Handles Btn_OuvrirUneAutreEtude.Click
        Try
            Dim resultat = Win_SélecteurDEtude.Cherche
            If resultat IsNot Nothing Then
                Me.EtudeCourante = resultat
            End If
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

    Private Sub Btn_Analyser_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Analyser.Click
        Try
            MsgBox("Btn_Analyser_Click")
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

#End Region

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

    Private Function CheckEtudeCourante() As Boolean
        If WS.EtudeCourante Is Nothing Then
            Message("Aucune étude n'est sélectionnée.")
            Return False
        Else
            Return True
        End If
    End Function

#Region "InitialiserLesClasseursExcelDeLEtudeCopurante"

    Friend Sub InitialiserLesClasseursExcelDeLEtudeCourante()
        Try

            If Not CheckEtudeCourante() Then Exit Sub

            Dim NewC As ClasseurExcel

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

            If Not CheckEtudeCourante() Then Exit Sub

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

    Private Sub Btn_OuvrirTousLesClasseurs_Click(sender As Object, e As RoutedEventArgs) Handles Btn_OuvrirTousLesClasseurs.Click
        Me.ChargerLesClasseursExcelDeLEtudeCopurante()
    End Sub

    Private Sub Btn_InitialiserLesClasseurs_Click(sender As Object, e As RoutedEventArgs) Handles Btn_InitialiserLesClasseurs.Click
        Me.InitialiserLesClasseursExcelDeLEtudeCourante()
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

#Region "GotoOuvrages"

    Private Sub Btn_GotoOuvrages_Click(sender As Object, e As RoutedEventArgs) Handles Btn_GotoOuvrages.Click
        Try
            With Me.EtudeCourante
                If .Ouvrages.Count = 0 Then
                    Message("Aucun ouvrage pour cette étude.", vbInformation)
                Else
                    Dim wo = New Win_Ouvrages
                    wo.ShowDialog($"Ouvrages du bordereau ""{Me.BordereauCourant.Nom}"" de l'étude ""{Me.EtudeCourante.Nom}""", Me.BordereauCourant.Ouvrages)
                End If
            End With
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

#End Region

#Region "Excel events"

    Private Sub AddExcelEventsHandlers()
        'AddHandler ExcelEventManager.TargetSelectedRangeChanged, AddressOf XLSelectionChangeHandling
    End Sub

    Private Sub UC_EtudesView_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        If e.Key = Key.S AndAlso e.KeyboardDevice.Modifiers = ModifierKeys.Control Then
            Try
                PersistancyManager.EnregistrerLeWorkspace()
                Win_Main.AfficherMessage($"Espace de travail {Application.Nom} enregistré à {Now().ToLongTimeString()}.")
            Catch ex As Exception
                ManageErreur(ex, "Echec de l'enregistrement de l'espace de travail.")
            End Try
        End If
    End Sub

    Private Sub Btn_UpdateExcel_Click(sender As Object, e As RoutedEventArgs) Handles Btn_UpdateExcel.Click
        Try
            WS.EtudeCourante.Ouvrages.DoForAll(Sub(o)
                                                   o.MAJExcel()
                                               End Sub)
            Message("Mise à jour effectuée.")
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

    'Private Sub XLSelectionChangeHandling(NewSelection As Excel.Range)
    'End Sub

#End Region

#End Region

#Region "INotifyPropertyChanged"

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Friend Sub NotifyPropertyChanged(PropertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(PropertyName))
    End Sub


#End Region

#End Region

End Class
