Imports System.Diagnostics
Imports System.Windows
Imports Microsoft.Office.Interop.Excel

Public Class UC_GestionnaireDeBordereau

#Region "Constructeurs"

    Private Sub UC_GestionnaireDeBordereau_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        Me.DataContext = New GestionnaireDeBordereaux
        Me.AddHandler(Controls.RadioButton.CheckedEvent, New RoutedEventHandler(Sub(sender2 As Object, e2 As RoutedEventArgs)
                                                                                    GérerLaQualificationDuDoublon(sender2, e2)
                                                                                End Sub))
    End Sub

#End Region

#Region "Propriétés"

#Region "GDB"
    Public ReadOnly Property GDB() As GestionnaireDeBordereaux
        Get
            Return Me.DataContext
        End Get
    End Property
#End Region

#Region "IsRunnig"
    Private _IsRunnig As Boolean
    Public Property IsRunnig() As Boolean
        Get
            Return _IsRunnig
        End Get
        Private Set(ByVal value As Boolean)
            If Object.Equals(value, Me._IsRunnig) Then Exit Property
            _IsRunnig = value
        End Set
    End Property
#End Region

#Region "TrackExcelEvent (Boolean)"

    Public Shared ReadOnly TrackExcelEventProperty As DependencyProperty =
            DependencyProperty.Register(NameOf(TrackExcelEvent), GetType(Boolean), GetType(UC_GestionnaireDeBordereau),
                                        New UIPropertyMetadata(False,
                                                               New PropertyChangedCallback(Sub(sender As UC_GestionnaireDeBordereau,
                                                                                               args As DependencyPropertyChangedEventArgs)

                                                                                               If sender.TrackExcelEvent Then
                                                                                                   sender.AddExcelEventsHandlers()
                                                                                               Else
                                                                                                   sender.RemoveExcelEventsHandlers()
                                                                                               End If
                                                                                           End Sub)))

    Public Property TrackExcelEvent As Boolean
        Get
            Return DirectCast(GetValue(TrackExcelEventProperty), Boolean)
        End Get

        Set(ByVal value As Boolean)
            SetValue(TrackExcelEventProperty, value)
        End Set
    End Property

#End Region

#End Region

#Region "Méthodes"

    Private Sub Btn_Start_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Start.Click
        Try
            RécupérerLesOuvrages()
        Catch ex As Exception
            ManageErreur(ex)
        Finally
            Me.IsRunnig = False
        End Try
    End Sub

#Region "GérerLaQualificationDuDoublon"

    Private Sub GérerLaQualificationDuDoublon(sender As Object, e As RoutedEventArgs)
        If EstUnChoixDeDoublon(sender, e) Then
            Dim rb = TryCast(e.OriginalSource, Controls.RadioButton)
            Dim Choix As Boolean? = Object.Equals(rb?.Tag, "true")
            Dim fe = TryCast(e.OriginalSource, FrameworkElement)

            Dim L = TryCast(fe.DataContext, LibelléDouvrage)
            If (Choix Is Nothing OrElse L Is Nothing) Then
                Throw New Exception("Opération inattendue.")
            Else
                Me.GDB.GérerLaQualificationDuDoublon(Choix, L)
            End If
        End If
    End Sub

    Private Function EstUnChoixDeDoublon(sender As Object, e As RoutedEventArgs) As Boolean

        Dim rb = TryCast(e.OriginalSource, Controls.RadioButton)
        Dim LeSenderEstUnRadio = rb IsNot Nothing
        Dim fe = TryCast(e.OriginalSource, FrameworkElement)
        Dim LeContexteEstUnLibellé = (TypeOf fe?.DataContext Is LibelléDouvrage)

        Return LeSenderEstUnRadio AndAlso LeContexteEstUnLibellé

    End Function

    Private Sub Btn_PurgerLeTransit_Click(sender As Object, e As RoutedEventArgs) Handles Btn_PurgerLeTransit.Click
        Try
            Me.GDB.PurgerLeTransit()
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

#End Region

#Region "Récupération des ouvrages"

    Private Sub RécupérerLesOuvrages()
        Me.IsRunnig = True
        Me.RemoveExcelEventsHandlers()
        Me.GDB.RécupérerLesLibellésDOuvrages()
        Message("Opération terminée.")
        XL.StatusBar = ""
        If Me.TrackExcelEvent Then Me.AddExcelEventsHandlers()
    End Sub

#End Region

#Region "Interfaçage Excel"

    Private Sub AddExcelEventsHandlers()
        If Not Me.IsRunnig Then
            AddHandler ExcelEventManager.TargetSelectedRangeChanged, AddressOf XLSelectionChangeHandling
        End If
    End Sub

    Private Sub RemoveExcelEventsHandlers()
        RemoveHandler ExcelEventManager.TargetSelectedRangeChanged, AddressOf XLSelectionChangeHandling
    End Sub

    Private Sub XLSelectionChangeHandling(newSelectedRange As Range)
        Try
            If Not Me.IsRunnig Then
                GDB.Sélectionner(newSelectedRange)
            End If
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

    Private Sub Btn_Go_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Go.Click
        Try
            Dim Confirmation = Message($"Tu vas créer {Me.GDB.NbLibellésRetenus} ouvrages mon coco, c'est pas de la commande de pédé ça. T'est sûr ?", MsgBoxStyle.YesNo)
            If Confirmation = MsgBoxResult.Yes Then
                Dim Confirmation2 = Message($"Faudra pas venir pleurer après, ok ?", MsgBoxStyle.OkCancel)
                If Confirmation2 Then
                    Dim NbOK, NbKO As Integer
                    Me.GDB.CréerLesOuvrages(NbOK, NbKO)
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
            GestionnaireDeBordereaux.SélectionnerLeRangesAssociéCourant(XL, L)
        End If

    End Sub

    Private Sub Btn_SélectionnerRangeSuivant(sender As Object, e As RoutedEventArgs)
        Dim L = TryCast((CType(sender, System.Windows.Controls.Button)).DataContext, LibelléDouvrage)
        If L IsNot Nothing Then
            L.SelectNextRange()
            GestionnaireDeBordereaux.SélectionnerLeRangesAssociéCourant(XL, L)
        End If
    End Sub

#End Region

#End Region

#End Region

#Region "Tests et debuggage"


#End Region

End Class
