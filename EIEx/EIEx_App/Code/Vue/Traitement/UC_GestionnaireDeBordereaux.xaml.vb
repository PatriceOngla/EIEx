Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Windows
Imports System.Windows.Controls
Imports Microsoft.Office.Interop.Excel
Imports Utils

Public Class UC_GestionnaireDeBordereau

#Region "Constructeurs"

    Private Sub UC_GestionnaireDeBordereau_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        Me._GdB = New GestionnaireDeBordereaux
        Me.DataContext = GdB
    End Sub

#End Region

#Region "Propriétés"

#Region "GdB"
    Private WithEvents _GdB As GestionnaireDeBordereaux
    Public ReadOnly Property GdB() As GestionnaireDeBordereaux
        Get
            Return Me._GdB
        End Get
    End Property
#End Region

    '#Region "SynchronizeWithExcelSelections_To (Boolean)"

    '    Public Shared ReadOnly SynchronizeWithExcelSelections_ToProperty As DependencyProperty =
    '            DependencyProperty.Register(NameOf(SynchronizeWithExcelSelections_To), GetType(Boolean), GetType(UC_GestionnaireDeBordereau),
    '                                        New UIPropertyMetadata(True))

    '    Public Property SynchronizeWithExcelSelections_To As Boolean
    '        Get
    '            Return DirectCast(GetValue(SynchronizeWithExcelSelections_ToProperty), Boolean)
    '        End Get

    '        Set(ByVal value As Boolean)
    '            SetValue(SynchronizeWithExcelSelections_ToProperty, value)
    '        End Set
    '    End Property

    '#End Region

    '#Region "SynchronizeWithExcelSelections_From (Boolean)"

    '    Public Shared ReadOnly SynchronizeWithExcelSelections_FromProperty As DependencyProperty =
    '            DependencyProperty.Register(NameOf(SynchronizeWithExcelSelections_From), GetType(Boolean), GetType(UC_GestionnaireDeBordereau),
    '                                        New UIPropertyMetadata(False,
    '                                                               New PropertyChangedCallback(Sub(sender As UC_GestionnaireDeBordereau,
    '                                                                                               args As DependencyPropertyChangedEventArgs)

    '                                                                                               If sender.SynchronizeWithExcelSelections_From Then
    '                                                                                                   sender.AddExcelSelectionChangedEventHandler()
    '                                                                                               Else
    '                                                                                                   sender.RemoveExcelSelectionChangedEventHandler()
    '                                                                                               End If
    '                                                                                           End Sub)))

    '    Public Property SynchronizeWithExcelSelections_From As Boolean
    '        Get
    '            Return DirectCast(GetValue(SynchronizeWithExcelSelections_FromProperty), Boolean)
    '        End Get

    '        Set(ByVal value As Boolean)
    '            SetValue(SynchronizeWithExcelSelections_FromProperty, value)
    '        End Set
    '    End Property

    '#End Region

#End Region

#Region "Méthodes"

    Private Sub Btn_Start_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Start.Click
        Try
            RécupérerLesOuvrages()
        Catch ex As Exception
            ManageErreur(ex)
            'Finally
            '    Me.IsRunnig = False
        End Try
    End Sub

#Region "GérerLaQualificationDuDoublon"

#Region "Gérer choix type doublon"

    Private Sub ChoixVraiDoublon(sender As Object, e As RoutedEventArgs)
        'Me.GérerLaQualificationDuDoublon(e, True)
        Me.GdB.GérerLaQualificationDesDoublonsSélectionnés(True)
    End Sub

    Private Sub ChoixFauxDoublon(sender As Object, e As RoutedEventArgs)
        'Me.GérerLaQualificationDuDoublon(e, False)
        Me.GdB.GérerLaQualificationDesDoublonsSélectionnés(False)
    End Sub

    'Private Sub GérerLaQualificationDuDoublon(e As RoutedEventArgs, VraiDoublon As Boolean)
    '    Dim L = GetLibelléCible(e)
    '    Dim LS = GetLibelléSuivant(L)
    '    If L Is Nothing Then
    '        Throw New InvalidOperationException("Le libellé d'ouvrage est nul. Opération incorrecte.")
    '    Else
    '        Me.GDB.GérerLaQualificationDuDoublon(L, VraiDoublon)
    '    End If
    '    If LS IsNot Nothing Then GDB.LibelléEnDoublonCourant = LS
    'End Sub

    'Private Function GetLibelléCible(e As RoutedEventArgs) As LibelléDouvrage
    '    Dim fe = TryCast(e.OriginalSource, FrameworkElement)
    '    Dim L = TryCast(fe.DataContext, LibelléDouvrage)
    '    Return L
    'End Function

    'Private Function GetLibelléSuivant(L As LibelléDouvrage) As LibelléDouvrage
    '    Try
    '        Dim r As LibelléDouvrage = Nothing
    '        If Me.GDB.LibellésEnDoublonEncoreATraiter.Count > 1 Then
    '            r = Me.GDB.LibellésEnDoublonEncoreATraiter.GetNextOrPrevious(L, True)
    '        End If
    '        Return r
    '    Catch ex As Exception
    '        ManageErreur(ex)
    '        Return Nothing
    '    End Try
    'End Function

    ''' <summary>
    ''' Pour éviter l'interception des clics trop rapprochés en tant que double clics et l'appel à la navigation (TraiterDemandeDeNavigation dans <see cref="Win_Main"/>).
    ''' </summary>
    Private Sub DG_LibellésOuvrages_MouseDoubleClick(sender As Object, e As Input.MouseButtonEventArgs)
        e.Handled = True
    End Sub

#End Region

    Private Sub Btn_PurgerLeTransit_Click(sender As Object, e As RoutedEventArgs) Handles Btn_PurgerLeTransit.Click
        Try
            Me.GdB.PurgerLeTransit()
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

#End Region

#Region "Récupération des ouvrages"

    Private Sub RécupérerLesOuvrages()
        Me.GdB.RécupérerLesLibellésDOuvrages()
        Message("Opération terminée.")
        XL.StatusBar = ""
    End Sub

#End Region

#Region "Interfaçage Excel"

    Private Sub Btn_Go_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Go.Click
        Try
            Dim Confirmation = Message($"Tu vas créer {Me.GdB.NbLibellésRetenus} ouvrages mon coco, c'est pas de la commande de pédé ça. T'est sûr ?", MsgBoxStyle.YesNo)
            If Confirmation = MsgBoxResult.Yes Then
                Dim Confirmation2 = Message($"Faudra pas venir pleurer après, ok ?", MsgBoxStyle.OkCancel)
                If Confirmation2 Then
                    Dim NbOK, NbKO As Integer
                    Me.GdB.CréerLesOuvrages(NbOK, NbKO)
                    Message($"Opération effectué : {NbOK} créations, {NbKO} échecs.")
                End If
            End If
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

#Region "Sélection des ranges"

    Private Sub Btn_SélectionnerRangePrécédent(sender As Object, e As RoutedEventArgs)
        Dim L = TryCast((CType(sender, System.Windows.Controls.Button)).DataContext, LibelléDouvrage)
        If L IsNot Nothing Then
            L.SelectPreviousRange()
            If Me.GdB.SynchronizeWithExcelSelections_To Then
                Me.GdB.SélectionnerLeRangeAssociéCourant(L)
            End If
        End If
    End Sub

    Private Sub Btn_SélectionnerRangeSuivant(sender As Object, e As RoutedEventArgs)
        Dim L = TryCast((CType(sender, System.Windows.Controls.Button)).DataContext, LibelléDouvrage)
        If L IsNot Nothing Then
            L.SelectNextRange()
            If Me.GdB.SynchronizeWithExcelSelections_To Then
                Me.GdB.SélectionnerLeRangeAssociéCourant(L)
            End If
        End If
    End Sub

#Region "Bring into view des sélections dans les DG"

    'Private Sub _GdB_PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Handles _GdB.PropertyChanged
    '    Select Case e.PropertyName
    '        Case NameOf(GestionnaireDeBordereaux.LibelléEnDoublonCourant)
    '            RendreVisibleLeSelectedItem(DG_LibellésOuvrages)
    '    End Select
    'End Sub

    Private Sub DG_LibellésOuvrages_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles DG_LibellésOuvrages.SelectionChanged
        If GdB.SynchronizeWithExcelSelections_To AndAlso e.AddedItems.Count = 1 Then Me.GdB.SélectionnerLeRangeAssociéCourant(e.AddedItems(0))
        RendreVisibleLeSelectedItem(DG_LibellésOuvrages)
    End Sub

    Private Sub DG_OuvragesAQualifier_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles DG_OuvragesAQualifier.SelectionChanged
        RendreVisibleLeSelectedItem(DG_OuvragesAQualifier)
    End Sub

    Private Sub DG_Ouvragesidentifiés_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles DG_Ouvragesidentifiés.SelectionChanged
        RendreVisibleLeSelectedItem(DG_Ouvragesidentifiés)
    End Sub

    Private Sub RendreVisibleLeSelectedItem(DG As DataGrid)
        With DG
            If .SelectedItem IsNot Nothing Then
                .ScrollIntoView(.SelectedItem)
                .Focus()
            End If
        End With
    End Sub

#End Region

#End Region

#End Region

#End Region

#Region "Tests et debuggage"


#End Region

End Class
