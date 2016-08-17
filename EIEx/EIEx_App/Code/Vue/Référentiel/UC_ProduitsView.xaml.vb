Imports System.Windows
Imports System.Windows.Input
Imports Model
Imports Utils

Public Class UC_ProduitsView

#Region "Constructeurs"

    Private Sub UC_ProduitsView_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized

        With UC_CmdesCRUD_Produits

            .MsgAlerteCohérenceSuppression = "Attention, ce produit est associé à au moins un patron d'ouvrage. En cas de suppression, ce(s) patron(s) predra(ont) leur référence à ce produit."

            .NomEntité = "produit"

            .AssociatedSelector = Me.DG_Master

            .SuppressionAConfirmer = Function(p As Produit)
                                         Dim r = (From o In Ref.PatronsDOuvrage
                                                  Where o.UtiliseProduit(p)).Any
                                         Return r
                                     End Function
        End With
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

#Region "ProduitCourant (Produit)"

    Public Shared ReadOnly ProduitCourantProperty As DependencyProperty =
            DependencyProperty.Register("ProduitCourant", GetType(Produit), GetType(UC_ProduitsView), New UIPropertyMetadata(Nothing))

    Public Property ProduitCourant As Produit
        Get
            Return DirectCast(GetValue(ProduitCourantProperty), Produit)
        End Get

        Set(ByVal value As Produit)
            SetValue(ProduitCourantProperty, value)
        End Set
    End Property

#End Region

#End Region

#Region "Méthodes"

#Region "Gestionnaires d'évennements"

#Region "CRUD"

    Private Sub UC_CmdesCRUD_DemandeAjout() Handles UC_CmdesCRUD_Produits.DemandeAjout
        Try
            Dim newProduit = Ref.GetNewProduit()
            Me.DG_Master.SelectedItem = newProduit
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

    Private Sub UC_CmdesCRUD_DemandeSuppression() Handles UC_CmdesCRUD_Produits.DemandeSuppression
        Try
            Dim Produit = Me.DG_Master.SelectedItem
            SupprimerLeProduitsDesOuvragesAssociés(Produit)
            Ref.Produits.Remove(Produit)
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

    Private Sub SupprimerLeProduitsDesOuvragesAssociés(Pdt As Produit)
        Dim UsageAssociés = From po In Ref.PatronsDOuvrage From up In po.UsagesDeProduit Where up.Produit Is Pdt Select up
        Dim UPASupprimer = New List(Of UsageDeProduit)(UsageAssociés)
        UPASupprimer.DoForAll(Sub(up) up.Parent.UsagesDeProduit.Remove(up))
    End Sub

#End Region

#Region "Recherche"

    Private Sub UC_ProduitsView_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        Try
            If e.Key = Key.F AndAlso e.KeyboardDevice.Modifiers = ModifierKeys.Control Then
                Dim result = RechercheProduit()
                If result IsNot Nothing Then Me.ProduitCourant = result
            End If
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

    Private Function RechercheProduit() As Produit
        Dim result = Win_SélecteurDeProduit.Cherche()
        Return result
    End Function

#End Region

    Private Sub Btn_ImporterProduits_Click(sender As Object, e As RoutedEventArgs) Handles Btn_ImporterProduits.Click
        ExcelCommander.ImporterProduitsDepuisExcel()
    End Sub

    Private Sub Btn_Load_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Load.Click
        Win_Main.RechargerRéférentiel()
    End Sub

    Private Sub Btn_Save_Click(sender As Object, e As RoutedEventArgs) Handles Btn_Save.Click
        Win_Main.EnregistrerRéférentiel()
    End Sub

#End Region

#End Region

End Class
