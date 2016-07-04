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
        Me.Tab1.SuspendLayout()
        Me.Grp_Unic.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Grp_Unic)
        Me.Tab1.Label = "EIEx"
        Me.Tab1.Name = "Tab1"
        '
        'Grp_Unic
        '
        Me.Grp_Unic.Items.Add(Me.TBt_ShowPanel)
        Me.Grp_Unic.Name = "Grp_Unic"
        '
        'TBt_ShowPanel
        '
        Me.TBt_ShowPanel.Description = "Afficher le panel de contrôle."
        Me.TBt_ShowPanel.Label = "Afficher le panel"
        Me.TBt_ShowPanel.Name = "TBt_ShowPanel"
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
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Grp_Unic As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents TBt_ShowPanel As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property EIExRibbon() As EIExRibbon
        Get
            Return Me.GetRibbon(Of EIExRibbon)()
        End Get
    End Property
End Class
