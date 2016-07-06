Imports System.Windows
Imports System.Windows.Input

Public Class UC_SubContainer

#Region "Constructeurs"

    Public Sub New()
        ' Cet appel est requis par le concepteur.
        InitializeComponent()
        ' Ajoutez une initialisation quelconque après l'appel InitializeComponent().
        ExcelEventManager.UCSC = Me
    End Sub

#End Region

#Region "Properties"

#Region "SelectedRange"

    Public Property SelectedRange As String ' Excel.Range
        Get
            Return GetValue(SelectedRangeProperty)
        End Get

        Set(ByVal value As String)
            SetValue(SelectedRangeProperty, value)
        End Set
    End Property

    Public Shared ReadOnly SelectedRangeProperty As DependencyProperty =
                           DependencyProperty.Register("SelectedRange",
                           GetType(String), GetType(UC_SubContainer),
                           New PropertyMetadata(Nothing))

#End Region

#End Region

#Region "Methods"

#Region "UI event handlers"

    Private Sub Button_Click(sender As Object, e As Windows.RoutedEventArgs)
        Try
            MsgBox("ça roule")
        Catch ex As ArgumentException
            MsgBox("ça roule pas")
        End Try
    End Sub

#End Region

#End Region

#Region "Tests and debug"
    Private Sub UC_SubContainer_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseRightButtonUp

    End Sub

#End Region

End Class
