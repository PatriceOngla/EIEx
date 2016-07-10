Imports Microsoft.Office.Interop

Friend Class ExcelEventManager

#Region "Properties"

#Region "XL"
    Public Shared WithEvents _XL As Excel.Application
    Public Shared ReadOnly Property XL() As Excel.Application
        Get
            If _XL Is Nothing Then _XL = EIExAddin.Application
            Return _XL
        End Get
    End Property
#End Region

#Region "TargetWindow"
    Private Shared _TargetWindow As Excel.Window
    Public Shared Property TargetWindow() As Excel.Window
        Get
            If _TargetWindow Is Nothing Then _TargetWindow = XL.ActiveWindow
            Return _TargetWindow
        End Get
        Set(ByVal value As Excel.Window)
            _TargetWindow = value
        End Set
    End Property
#End Region

#Region "TargetSheet"
    Private Shared WithEvents _TargetSheet As Excel.Worksheet
    Public Shared Property TargetSheet() As Excel.Worksheet
        Get
            If _TargetSheet Is Nothing Then
                _TargetSheet = XL.ActiveSheet
            End If
            Return _TargetSheet
        End Get
        Set(ByVal value As Excel.Worksheet)
            _TargetSheet = value
        End Set
    End Property

#End Region

#Region "UCSC"

    Private Shared _UCSC As UC_SubContainer
    Public Shared Property UCSC() As UC_SubContainer
        Get
            Return _UCSC
        End Get
        Set(ByVal value As UC_SubContainer)
            _UCSC = value
        End Set
    End Property

#End Region

#End Region

#Region "Methods"

    Private Shared Sub _XL_SheetChange(Sh As Object, Target As Excel.Range) Handles _XL.SheetSelectionChange
        Try
            UCSC.SelectedRange = Target.Address
        Catch ex As ArgumentException

        End Try

    End Sub

    Public Shared Sub CleanUp()
        _XL = Nothing
    End Sub

#End Region

End Class
