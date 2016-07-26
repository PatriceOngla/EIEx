Imports System.Windows.Input
Imports EIEx_DAO
Imports Model

Public Class UC_RéférentielView

#Region "Constructeurs"
    Private Sub UC_RéférentielView_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        Me.DataContext = Référentiel.Instance
    End Sub

#End Region

#Region "Propriétés"

#Region "Ref"
    Public ReadOnly Property Ref() As Référentiel
        Get
            Return Référentiel.Instance
        End Get
    End Property
#End Region

#End Region

#Region "Méthodes"

    Private Sub UC_EtudesView_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        If e.Key = Key.S AndAlso e.KeyboardDevice.Modifiers = ModifierKeys.Control Then
            Try
                PersistancyManager.EnregistrerLeRéférentiel()
                XL.StatusBar = $"Référentiel {ThisAddIn.Nom} enregistré à {Now().ToLongTimeString()}."
            Catch ex As Exception
                ManageErreur(ex, "Echec de l'enregistrement du référentiel.")
            End Try
        End If
    End Sub

#End Region

End Class
