Imports Microsoft.Office.Tools
Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop
Imports Model

Public Class EIExRibbon

#Region "Champs privés"
    Public WithEvents EIExTaskPane As CustomTaskPane
#End Region

#Region "Constructeurs"


#End Region

#Region "Propriétés"

#Region "XL"
    Public Shared WithEvents _XL As Excel.Application
    Public Shared ReadOnly Property XL() As Excel.Application
        Get
            If _XL Is Nothing Then _XL = EIExAddin.Application
            Return _XL
        End Get
    End Property
#End Region

#Region "UC_EIEx_Manager_UI (shared)"
    Private _UC_EIEx_Manager_UI As UC_EIEx_Manager_UI
    Public ReadOnly Property UC_EIEx_Manager_UI() As UC_EIEx_Manager_UI
        Get
            Return _UC_EIEx_Manager_UI
        End Get
    End Property
#End Region

#Region "EtudeCourante"
    Public ReadOnly Property EtudeCourante() As Etude
        Get
            Return UC_EIEx_Manager_UI?.EtudeCourante
        End Get
    End Property
#End Region

#End Region

#Region "Méthodes"

    'Private Sub EIExRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
    '    HideOrShowAndAttachPanel(True)
    'End Sub
#Region "Gestion du panel"

    Private Shared Sub Application_SheetActivate(Sh As Object) Handles _XL.SheetActivate
        'TBt_ShowPanel.Checked = EIExTaskPane.Visible
    End Sub

    Private Sub TBt_ShowPanel_Click(sender As Object, e As RibbonControlEventArgs) Handles TBt_ShowPanel.Click
        Try
            Dim FenêtreAppelante = ExcelEventManager.TargetWindow
            Dim OuvrirPannel = TBt_ShowPanel.Checked
            If EIExTaskPane IsNot Nothing AndAlso FenêtreAppelante IsNot EIExTaskPane.Window Then
                EIExTaskPane.Dispose()
                EIExTaskPane = Nothing
            End If

            If EIExTaskPane Is Nothing OrElse TBt_ShowPanel.Checked <> EIExTaskPane.Visible Then
                HideOrShowAndAttachPanel(OuvrirPannel)
            End If
        Catch ex As System.Runtime.InteropServices.COMException
            EIExTaskPane = Nothing
            ManageErreur(ex, "Une erreur est survenue. Merci d'essayer à nouveau.")
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

    Private Sub HideOrShowAndAttachPanel(show As Boolean)
        If ExcelEventManager.TargetWindow Is Nothing Then
            MsgBox("Aucune fenêtre active.")
        Else
            If EIExTaskPane Is Nothing Then
                Dim c = New UC_Container()
                EIExTaskPane = EIExAddin.CustomTaskPanes.Add(c, "EIEx", ExcelEventManager.TargetWindow)
                _UC_EIEx_Manager_UI = c.EIEx_Manager_UI
            End If
            EIExTaskPane.Visible = show
            EIExTaskPane.Width = 800
            If Not show Then
                EIExTaskPane.Dispose()
                EIExTaskPane = Nothing
            End If
        End If

    End Sub

    Private Sub EIExTaskPane_VisibleChanged(sender As Object, e As EventArgs) Handles EIExTaskPane.VisibleChanged
        TBt_ShowPanel.Checked = EIExTaskPane.Visible
    End Sub

    Private Sub EIExRibbon_Close(sender As Object, e As EventArgs) Handles Me.Close
        Try
            EIExTaskPane?.Dispose()
            EIExTaskPane = Nothing
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

#End Region

#Region "Gestion des enregistrements"
    Private Sub Btn_EnregistrerRéférentiel_Click(sender As Object, e As RibbonControlEventArgs) Handles Btn_EnregistrerRéférentiel.Click
        Try
            Model.EIExData.EnregistrerLeRéférentiel()
            Message("Enregistrement effectué.")
        Catch ex As Exception
            ManageErreur(ex, "Echec de l'enregistrement du référentiel.", True, False)
        End Try
    End Sub

    Private Sub Btn_RechargerRéférentiel_Click(sender As Object, e As RibbonControlEventArgs) Handles Btn_RechargerRéférentiel.Click
        Try
            Model.EIExData.ChargerLeRéférentiel()
            Message("Rechargement effectué.")
        Catch ex As Exception
            ManageErreur(ex, "Echec du chargement du référentiel.", True, False)
        End Try
    End Sub

    Private Sub Btn_SaveWorkspace_Click(sender As Object, e As RibbonControlEventArgs) Handles Btn_SaveWorkspace.Click
        Try
            Model.EIExData.EnregistrerLeWorkspace()
            Message("Enregistrement effectué.")
        Catch ex As Exception
            ManageErreur(ex, "Echec de l'enregistrement de l'espace de travail.", True, False)
        End Try
    End Sub

    Private Sub Btn_RechargerWorkspace_Click(sender As Object, e As RibbonControlEventArgs) Handles Btn_RechargerWorkspace.Click
        Try
            Model.EIExData.ChargerLeWorkspace()
            Message("Rechargement effectué.")
        Catch ex As Exception
            ManageErreur(ex, "Echec du chargement de l'espace de travail.", True, False)
        End Try
    End Sub

#End Region

    Private Sub Btn_ChargerDepuisExcel_Click(sender As Object, e As RibbonControlEventArgs) Handles Btn_ImporterProduitsDepuisExcel.Click
        ImporterProduitsDepuisExcel()
    End Sub

    Private Sub Btn_InitialiserLesClasseursExcelDeLEtudeCourante_Click(sender As Object, e As RibbonControlEventArgs) Handles Btn_InitialiserLesClasseursExcelDeLEtudeCourante.Click
        Me.UC_EIEx_Manager_UI.UC_Etude.InitialiserLesClasseursExcelDeLEtudeCopurante()
    End Sub

    Private Sub Btn_ChargerLesClasseursExcelDeLEtudeCourante_Click(sender As Object, e As RibbonControlEventArgs) Handles Btn_ChargerLesClasseursExcelDeLEtudeCourante.Click
        Me.UC_EIEx_Manager_UI.UC_Etude.ChargerLesClasseursExcelDeLEtudeCopurante()
    End Sub


#End Region

#Region "Tests et debuggage"

#End Region

End Class
