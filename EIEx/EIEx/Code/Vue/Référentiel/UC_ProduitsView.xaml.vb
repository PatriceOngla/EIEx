Imports Model
Imports Utils

Public Class UC_ProduitsView

#Region "Constructeurs"

    Private Sub UC_FamillesProduitView_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        With UC_CmdesCRUD_Produits

            .MsgAlerteCohérenceSuppression = "Attention, ce produit est associé à au moins un patron d'ouvrage. En cas de suppression, ce(s) patron(s) predra(ont) leur référence à ce produit."

            .NomEntité = "produit"

            .AssociatedSelector = Me.DG_Master

            .SuppressionAConfirmer = Function(p As Produit)
                                         Dim r = (From ro In Ref.PatronsDOuvrage
                                                  Where ro.UtiliseProduit(p)).Any
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
            SupprimerLeProduitsDesPatronsDOuvragesAssociés(Produit)
            Ref.Produits.Remove(Produit)
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

    Private Sub SupprimerLeProduitsDesPatronsDOuvragesAssociés(famille As FamilleDeProduit)
        Dim ProduitsAssociés = From p In Ref.Produits Where p.Famille Is famille
        ProduitsAssociés.DoForAll(Sub(p) p.Famille = Nothing)
    End Sub

#End Region

#End Region

#End Region

End Class
