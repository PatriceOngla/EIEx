Imports Microsoft.Office.Tools
Imports Microsoft.Office.Tools.Ribbon
Imports Utils

Public Class EIExRibbon

    Private Sub EIExRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
    End Sub

    Public WithEvents EIExTaskPane As CustomTaskPane

    Private Sub TBt_ShowPanel_Click(sender As Object, e As RibbonControlEventArgs) Handles TBt_ShowPanel.Click
        Try
            HideOrShowAndAttachPanel(TBt_ShowPanel.Checked)
        Catch ex As ArgumentException
            ManageError(ex, NameOf(TBt_ShowPanel_Click))
        End Try
    End Sub

    Private Sub HideOrShowAndAttachPanel(show As Boolean)
        If ExcelEventManager.TargetWindow Is Nothing Then
            MsgBox("Aucune fenêtre active.")
        Else
            If EIExTaskPane Is Nothing Then
                Dim c = New UC_Container()
                EIExTaskPane = EIExAddin.CustomTaskPanes.Add(c, "EIEx", ExcelEventManager.TargetWindow)
            End If
            EIExTaskPane.Visible = show
            If Not show Then
                EIExTaskPane.Dispose()
                EIExTaskPane = Nothing
            End If
        End If

    End Sub

    Private Sub EIExRibbon_Close(sender As Object, e As EventArgs) Handles Me.Close
        Try
            EIExTaskPane?.Dispose()
            EIExTaskPane = Nothing
        Catch ex As Exception
            ManageError(ex, NameOf(EIExRibbon_Close))
        End Try
    End Sub



End Class
