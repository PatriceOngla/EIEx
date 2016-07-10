Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives

Public Class UC_CommandesCRUD

#Region "Propriétés"

#Region "AssociatedSelector (Selector)"

    Public Shared ReadOnly AssociatedItemsControlProperty As DependencyProperty =
            DependencyProperty.Register(NameOf(AssociatedSelector), GetType(Selector), GetType(UC_CommandesCRUD), New UIPropertyMetadata(Nothing))

    Public Property AssociatedSelector As Selector
        Get
            Return DirectCast(GetValue(AssociatedItemsControlProperty), Selector)
        End Get

        Set(ByVal value As Selector)
            SetValue(AssociatedItemsControlProperty, value)
        End Set
    End Property

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
#End Region

#Region "déclarations d'évennements"

    Public Event DemandeAjout()

    Public Event DemandeSuppression()

#End Region

#Region "Gestionnaires d'évennements"

    Private Function ContextOK(ElementSélectionnéRequis As Boolean) As Boolean
        Dim ElementSélectionnéOK = (Not ElementSélectionnéRequis) OrElse Me.AssociatedSelector.SelectedItem IsNot Nothing
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
#End Region

End Class
