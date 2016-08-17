Imports System.Windows
Imports Excel = Microsoft.Office.Interop.Excel

Public Class UC_ExcelRangeSelector

#Region "Constructeurs"

    Private Sub UC_ExcelRangeSelector_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        AddExcelEventsHandlers()
    End Sub
    Private Sub AddExcelEventsHandlers()
        AddHandler ExcelEventManager.TargetSelectedRangeChanged, AddressOf XLSelectionChangeHandling
    End Sub

#End Region

#Region "Propriétés"

#Region "XLRange (String)"

    Public Shared ReadOnly XLRangeProperty As DependencyProperty =
            DependencyProperty.Register("XLRange", GetType(String), GetType(UC_ExcelRangeSelector), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

    Public Property XLRange As String
        Get
            Return DirectCast(GetValue(XLRangeProperty), String)
        End Get

        Set(ByVal value As String)
            SetValue(XLRangeProperty, value)
        End Set
    End Property

#End Region

#Region "EstActif (Boolean)"

    Public Shared ReadOnly EstActifProperty As DependencyProperty =
            DependencyProperty.Register("EstActif", GetType(Boolean), GetType(UC_ExcelRangeSelector),
                                        New UIPropertyMetadata(False, New PropertyChangedCallback(
                                                               Sub(sender As UC_ExcelRangeSelector, e As DependencyPropertyChangedEventArgs)
                                                                   ExcelCommander.FocusExcel()
                                                               End Sub))
                                                               )

    Public Property EstActif As Boolean
        Get
            Return DirectCast(GetValue(EstActifProperty), Boolean)
        End Get

        Set(ByVal value As Boolean)
            SetValue(EstActifProperty, value)
        End Set
    End Property

#End Region

#End Region

#Region "Gestionnaire d'évennement"
    Private Sub XLSelectionChangeHandling(NewSelection As Excel.Range)
        If Me.EstActif Then
            Me.XLRange = NewSelection?.Address
            Me.EstActif = False
            ReFocusApplication()
        End If
    End Sub

    ''' <summary>Parce que le focus a été transmis à Excel. </summary>
    Private Sub ReFocusApplication()
        Try
            Dim w = Win_Main.Instance
            w.Activate()
            Me.Focus()
        Catch ex As Exception
            ManageErreur(ex, , False)
        End Try

    End Sub
#End Region
End Class
