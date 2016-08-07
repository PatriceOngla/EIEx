Partial Class EIExRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Requis pour la prise en charge du Concepteur de composition de classes Windows.Forms
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'Cet appel est requis par le Concepteur de composants.
        InitializeComponent()

    End Sub

    'Component remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requise par le Concepteur de composants
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur de composants
    'Elle peut être modifiée à l'aide du Concepteur de composants.
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Grp_Unic = Me.Factory.CreateRibbonGroup
        Me.TBt_ShowPanel = Me.Factory.CreateRibbonToggleButton
        Me.Grp_Workspace = Me.Factory.CreateRibbonGroup
        Me.Btn_SaveWorkspace = Me.Factory.CreateRibbonButton
        Me.Btn_RechargerWorkspace = Me.Factory.CreateRibbonButton
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.Box_Classeurs = Me.Factory.CreateRibbonBox
        Me.Label1 = Me.Factory.CreateRibbonLabel
        Me.Btn_InitialiserLesClasseursExcelDeLEtudeCourante = Me.Factory.CreateRibbonButton
        Me.Btn_ChargerLesClasseursExcelDeLEtudeCourante = Me.Factory.CreateRibbonButton
        Me.Grp_Référentiel = Me.Factory.CreateRibbonGroup
        Me.Btn_EnregistrerRéférentiel = Me.Factory.CreateRibbonButton
        Me.Btn_RechargerRéférentiel = Me.Factory.CreateRibbonButton
        Me.Btn_ImporterProduitsDepuisExcel = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Grp_Unic.SuspendLayout()
        Me.Grp_Workspace.SuspendLayout()
        Me.Box_Classeurs.SuspendLayout()
        Me.Grp_Référentiel.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Grp_Unic)
        Me.Tab1.Groups.Add(Me.Grp_Workspace)
        Me.Tab1.Groups.Add(Me.Grp_Référentiel)
        Me.Tab1.Label = "EIEx"
        Me.Tab1.Name = "Tab1"
        '
        'Grp_Unic
        '
        Me.Grp_Unic.Items.Add(Me.TBt_ShowPanel)
        Me.Grp_Unic.Label = "Affichage"
        Me.Grp_Unic.Name = "Grp_Unic"
        '
        'TBt_ShowPanel
        '
        Me.TBt_ShowPanel.Description = "Afficher le panel de contrôle."
        Me.TBt_ShowPanel.Label = "Afficher le panel"
        Me.TBt_ShowPanel.Name = "TBt_ShowPanel"
        '
        'Grp_Workspace
        '
        Me.Grp_Workspace.Items.Add(Me.Btn_SaveWorkspace)
        Me.Grp_Workspace.Items.Add(Me.Btn_RechargerWorkspace)
        Me.Grp_Workspace.Items.Add(Me.Separator1)
        Me.Grp_Workspace.Items.Add(Me.Box_Classeurs)
        Me.Grp_Workspace.Label = "Espace de travail"
        Me.Grp_Workspace.Name = "Grp_Workspace"
        '
        'Btn_SaveWorkspace
        '
        Me.Btn_SaveWorkspace.Label = "Enregistrer l'espace de travail"
        Me.Btn_SaveWorkspace.Name = "Btn_SaveWorkspace"
        '
        'Btn_RechargerWorkspace
        '
        Me.Btn_RechargerWorkspace.Label = "Recharger l'espace de travail"
        Me.Btn_RechargerWorkspace.Name = "Btn_RechargerWorkspace"
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        '
        'Box_Classeurs
        '
        Me.Box_Classeurs.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical
        Me.Box_Classeurs.Items.Add(Me.Label1)
        Me.Box_Classeurs.Items.Add(Me.Btn_InitialiserLesClasseursExcelDeLEtudeCourante)
        Me.Box_Classeurs.Items.Add(Me.Btn_ChargerLesClasseursExcelDeLEtudeCourante)
        Me.Box_Classeurs.Name = "Box_Classeurs"
        '
        'Label1
        '
        Me.Label1.Label = "Classeurs de l'étude courante :"
        Me.Label1.Name = "Label1"
        '
        'Btn_InitialiserLesClasseursExcelDeLEtudeCourante
        '
        Me.Btn_InitialiserLesClasseursExcelDeLEtudeCourante.Label = "Initialiser tous"
        Me.Btn_InitialiserLesClasseursExcelDeLEtudeCourante.Name = "Btn_InitialiserLesClasseursExcelDeLEtudeCourante"
        '
        'Btn_ChargerLesClasseursExcelDeLEtudeCourante
        '
        Me.Btn_ChargerLesClasseursExcelDeLEtudeCourante.Label = "Ouvrir tous"
        Me.Btn_ChargerLesClasseursExcelDeLEtudeCourante.Name = "Btn_ChargerLesClasseursExcelDeLEtudeCourante"
        '
        'Grp_Référentiel
        '
        Me.Grp_Référentiel.Items.Add(Me.Btn_EnregistrerRéférentiel)
        Me.Grp_Référentiel.Items.Add(Me.Btn_RechargerRéférentiel)
        Me.Grp_Référentiel.Items.Add(Me.Btn_ImporterProduitsDepuisExcel)
        Me.Grp_Référentiel.Label = "Référentiel"
        Me.Grp_Référentiel.Name = "Grp_Référentiel"
        '
        'Btn_EnregistrerRéférentiel
        '
        Me.Btn_EnregistrerRéférentiel.Label = "Enregistrer le référentiel"
        Me.Btn_EnregistrerRéférentiel.Name = "Btn_EnregistrerRéférentiel"
        '
        'Btn_RechargerRéférentiel
        '
        Me.Btn_RechargerRéférentiel.Label = "Recharger le référentiel"
        Me.Btn_RechargerRéférentiel.Name = "Btn_RechargerRéférentiel"
        '
        'Btn_ImporterProduitsDepuisExcel
        '
        Me.Btn_ImporterProduitsDepuisExcel.Description = "Charger les produits depuis la feuille excel courante."
        Me.Btn_ImporterProduitsDepuisExcel.Label = "Importer des produits"
        Me.Btn_ImporterProduitsDepuisExcel.Name = "Btn_ImporterProduitsDepuisExcel"
        '
        'EIExRibbon
        '
        Me.Name = "EIExRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Grp_Unic.ResumeLayout(False)
        Me.Grp_Unic.PerformLayout()
        Me.Grp_Workspace.ResumeLayout(False)
        Me.Grp_Workspace.PerformLayout()
        Me.Box_Classeurs.ResumeLayout(False)
        Me.Box_Classeurs.PerformLayout()
        Me.Grp_Référentiel.ResumeLayout(False)
        Me.Grp_Référentiel.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Grp_Unic As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents TBt_ShowPanel As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents Btn_EnregistrerRéférentiel As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Btn_SaveWorkspace As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Grp_Référentiel As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Btn_RechargerRéférentiel As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Btn_RechargerWorkspace As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Grp_Workspace As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Btn_ImporterProduitsDepuisExcel As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Btn_InitialiserLesClasseursExcelDeLEtudeCourante As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Btn_ChargerLesClasseursExcelDeLEtudeCourante As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Box_Classeurs As Microsoft.Office.Tools.Ribbon.RibbonBox
    Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Label1 As Microsoft.Office.Tools.Ribbon.RibbonLabel
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property EIExRibbon() As EIExRibbon
        Get
            Return Me.GetRibbon(Of EIExRibbon)()
        End Get
    End Property
End Class
