Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives

Public Class UC_CommandesCRUD

#Region "Propriétés"

#Region "AssociatedSelector (Selector)"

    Public Shared ReadOnly AssociatedItemsControlProperty As DependencyProperty =
            DependencyProperty.Register(NameOf(AssociatedSelector), GetType(Selector), GetType(UC_CommandesCRUD),
                                        New UIPropertyMetadata(Nothing, New PropertyChangedCallback(
                                        Sub(ucCrud As UC_CommandesCRUD, e As DependencyPropertyChangedEventArgs)
                                            ucCrud._AssociatedSelector = e.NewValue
                                        End Sub)))

    Public Property AssociatedSelector As Selector
        Get
            Return DirectCast(GetValue(AssociatedItemsControlProperty), Selector)
        End Get

        Set(ByVal value As Selector)
            SetValue(AssociatedItemsControlProperty, value)
        End Set
    End Property

    Private WithEvents _AssociatedSelector As Selector

#End Region

#Region "NomEntité"
    Private _NomEntité As String
    Public Property NomEntité() As String
        Get
            If String.IsNullOrEmpty(_NomEntité) Then
                Return "?"
            Else
                Return _NomEntité
            End If
        End Get
        Set(value As String)
            _NomEntité = value
            SetTooltips()
        End Set
    End Property
#End Region

#Region "ContrôleDeCohérence"
    Private _SuppressionAConfirmer As Predicate(Of Model.Entité)
    Public Property SuppressionAConfirmer() As Predicate(Of Model.Entité)
        Get
            Return _SuppressionAConfirmer
        End Get
        Set(ByVal value As Predicate(Of Model.Entité))
            If Object.Equals(value, Me._SuppressionAConfirmer) Then Exit Property
            _SuppressionAConfirmer = value
        End Set
    End Property
#End Region

#Region "MsgAlerteCohérenceSuppression"
    Private _MsgAlerteCohérenceSuppression As String
    Public Property MsgAlerteCohérenceSuppression() As String
        Get
            Return _MsgAlerteCohérenceSuppression
        End Get
        Set(ByVal value As String)
            If Object.Equals(value, Me._MsgAlerteCohérenceSuppression) Then Exit Property
            _MsgAlerteCohérenceSuppression = value
        End Set
    End Property
#End Region

#Region "ItemSélectionné"
    Public ReadOnly Property ItemSélectionné() As Object
        Get
            Return Me.AssociatedSelector?.SelectedItem
        End Get
    End Property
#End Region

#Region "Tooltips"

#Region "TooltipAjout (String)"

#Region "Déclaration et registration de TooltipAjoutProperty"

    Private Shared MDTooltipAjout As New FrameworkPropertyMetadata(Nothing)
    Public Shared TooltipAjoutPropertyKey As DependencyPropertyKey = DependencyProperty.RegisterReadOnly("TooltipAjout", GetType(String), GetType(UC_CommandesCRUD), MDTooltipAjout)
    Public Shared TooltipAjoutProperty As DependencyProperty = TooltipAjoutPropertyKey.DependencyProperty

#End Region

    Public ReadOnly Property TooltipAjout() As String
        Get
            Return GetValue(TooltipAjoutProperty)
        End Get
    End Property

#Region "Calcul de la valeur"
    Private Sub SetTooltipAjout()
        Dim v = $"Ajouter {If(String.IsNullOrEmpty(Me.NomEntité), "", "un(e) " & Me.NomEntité)}."
        Me.SetValue(TooltipAjoutPropertyKey, v)
    End Sub
#End Region

#End Region

#Region "TooltipSuppression (String)"

#Region "Déclaration et registration de TooltipSuppressionProperty"

    Private Shared MDTooltipSuppression As New FrameworkPropertyMetadata(Nothing)
    Public Shared TooltipSuppressionPropertyKey As DependencyPropertyKey = DependencyProperty.RegisterReadOnly("TooltipSuppression", GetType(String), GetType(UC_CommandesCRUD), MDTooltipSuppression)
    Public Shared TooltipSuppressionProperty As DependencyProperty = TooltipSuppressionPropertyKey.DependencyProperty

#End Region

#Region "Wrapper CLR de TooltipSuppressionProperty"

    Public ReadOnly Property TooltipSuppression() As String
        Get
            Return GetValue(TooltipSuppressionProperty)
        End Get
    End Property

#End Region

#Region "Calcul de la valeur"
    Private Sub SetTooltipSuppressionx()
        Dim v As String
        If Me.ItemSélectionné Is Nothing Then
            v = $"Suppression impossible. Aucun(e) {If(String.IsNullOrEmpty(Me.NomEntité), "élément", Me.NomEntité)} n'est sélectionné(e)."
        Else
            v = $"Supprimer {If(String.IsNullOrEmpty(Me.NomEntité), "l'élément", "le(a) " & Me.NomEntité)} sélectionné(e)."
        End If
        Me.SetValue(TooltipSuppressionPropertyKey, v)
    End Sub
#End Region

#End Region

    Private Sub SetTooltips()
        Me.SetTooltipAjout()
        Me.SetTooltipSuppressionx()
    End Sub

#End Region

#End Region

#Region "déclarations d'évennements"

    Public Event DemandeAjout()

    Public Event DemandeSuppression()

#End Region

#Region "Gestionnaires d'évennements"

    Private Function ContextOK(ElementSélectionnéRequis As Boolean) As Boolean
        Dim ElementSélectionnéOK = (Not ElementSélectionnéRequis) OrElse Me.ItemSélectionné IsNot Nothing
        Return Me.AssociatedSelector IsNot Nothing AndAlso ElementSélectionnéOK
    End Function

    Private Sub AlertContexteKO()
        Dim Msg = If(Me.AssociatedSelector Is Nothing,
            $"Aucun contrôle Selector n'est associé à cette barre de commande.",
            $"Aucun(e) {NomEntité} sélectionné(e)")
        MsgBox(Msg, MsgBoxStyle.Exclamation, ThisAddIn.Nom)
    End Sub

    Private Sub Btn_Ajouter_Click(sender As Object, e As Windows.RoutedEventArgs) Handles Btn_Ajouter.Click
        Try
            If ContextOK(False) Then
                RaiseEvent DemandeAjout()
            Else
                AlertContexteKO()
            End If
        Catch ex As Exception
            ManageErreur(ex, NameOf(Btn_Ajouter_Click))
        End Try
    End Sub

    Private Sub Btn_Supprimer_Click(sender As Object, e As Windows.RoutedEventArgs) Handles Btn_Supprimer.Click
        Try
            If ContextOK(True) Then
                LeverEventDeSSuppression()
            Else
                AlertContexteKO()
            End If
        Catch ex As Exception
            ManageErreur(ex, NameOf(Btn_Supprimer_Click))
        End Try
    End Sub

    Private Sub LeverEventDeSSuppression()

        Dim MsgConfirmation = $"Supprimer le(a) ""{NomEntité}"" courant(e) ?"
        If Me.SuppressionAConfirmer IsNot Nothing AndAlso Me.SuppressionAConfirmer(AssociatedSelector.SelectedItem) Then
            MsgConfirmation &= vbCr & "Attention : " & MsgAlerteCohérenceSuppression
        End If
        If MsgBox(MsgConfirmation, vbOKCancel) = MsgBoxResult.Ok Then
            RaiseEvent DemandeSuppression()
        End If

    End Sub

    Private Sub _AssociatedSelector_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles _AssociatedSelector.SelectionChanged
        SetTooltips()
    End Sub

#End Region

End Class
