﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class UC_Container
    Inherits System.Windows.Forms.UserControl

    'UserControl remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.WPFHost = New System.Windows.Forms.Integration.ElementHost()
        Me.UC_SubContainer1 = New EIEx.UC_SubContainer()
        Me.SuspendLayout()
        '
        'WPFHost
        '
        Me.WPFHost.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.WPFHost.AutoSize = True
        Me.WPFHost.BackColorTransparent = True
        Me.WPFHost.Location = New System.Drawing.Point(3, 0)
        Me.WPFHost.Name = "WPFHost"
        Me.WPFHost.Size = New System.Drawing.Size(294, 127)
        Me.WPFHost.TabIndex = 0
        Me.WPFHost.Text = "ElementHost1"
        Me.WPFHost.Child = Me.UC_SubContainer1
        '
        'UC_Container
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.WPFHost)
        Me.Name = "UC_Container"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents WPFHost As Windows.Forms.Integration.ElementHost
    Friend UC_SubContainer1 As UC_SubContainer
End Class
