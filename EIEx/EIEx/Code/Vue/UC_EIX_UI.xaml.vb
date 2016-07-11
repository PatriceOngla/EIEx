Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports Model

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

#Region "Ref"
    Public ReadOnly Property Ref() As Référentiel
        Get
            Return Référentiel.Instance
        End Get
    End Property
#End Region

#Region "WS"
    Public ReadOnly Property WS() As WorkSpace
        Get
            Return WorkSpace.Instance
        End Get
    End Property
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

    Private Sub UC_SubContainer_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseDoubleClick
        Try
            Dim FESource = TryCast(e.OriginalSource, FrameworkElement)
            If FESource IsNot Nothing Then
                Dim DonnéeNavigationCible = TryCast(FESource.DataContext, Entité)
                If DonnéeNavigationCible Is Nothing Then
                    Dim CBxParent = GetPrentCombobox(FESource)
                    If CBxParent IsNot Nothing Then
                        DonnéeNavigationCible = TryCast(CBxParent.SelectedItem, Entité)
                    End If
                End If
                If DonnéeNavigationCible IsNot Nothing Then
                    NaviguerVers(DonnéeNavigationCible)
                End If
            End If
        Catch ex As Exception
            ManageErreur(ex, NameOf(UC_SubContainer_MouseDoubleClick))
        End Try
    End Sub

    Private Function GetPrentCombobox(fe As FrameworkElement) As ComboBox
        Dim Parent = VisualTreeHelper.GetParent(fe)
        Do While Parent IsNot Nothing AndAlso TypeOf Parent IsNot ComboBox
            Parent = VisualTreeHelper.GetParent(Parent)
        Loop
        Return Parent
    End Function

    Private Sub NaviguerVers(donnéeNavigationCible As Entité)
        If TypeOf donnéeNavigationCible Is UsageDeProduit Then
            donnéeNavigationCible = TryCast(donnéeNavigationCible, UsageDeProduit).Produit
        End If
        MsgBox($"En route pour {donnéeNavigationCible.ToString()}.")
    End Sub

#End Region

#End Region

#Region "Tests and debug"

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

    Private Sub UC_SubContainer_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseRightButtonUp

    End Sub

    'Private Sub UC_SubContainer_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
    '    Try

    '        Dim WS = WorkSpace.Instance
    '        Me.UC_Etude.DataContext = WS.EtudeCourante
    '    Catch ex As Exception
    '        ManageErreur(ex, "Echec de la récupération de l'étude courante.")
    '    End Try
    'End Sub

#End Region

End Class
