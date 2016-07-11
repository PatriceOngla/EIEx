Imports Model
Imports Utils

Public Class UC_FamillesProduitView

#Region "Constructeurs"

    Private Sub UC_FamillesProduitView_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized

        Me.UC_CmdesCRUD.MsgAlerteCohérenceSuppression = "Attention, cette famille est associé à au moins un produit. En cas de suppression, ce(s) produit(s) ne sera (seront) plus associé(s) à une famille."

        Me.UC_CmdesCRUD.NomEntité = "famille de produit"

        Me.UC_CmdesCRUD.AssociatedSelector = Me.DG_Master

        Me.UC_CmdesCRUD.SuppressionAConfirmer = Function(F As FamilleDeProduit)
                                                    Dim r = (From p In Ref.Produits Where p.Famille Is F).Any
                                                    Return r
                                                End Function
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

#Region "Gestionnaires d'évennements"

    Private Sub UC_CmdesCRUD_DemandeAjout() Handles UC_CmdesCRUD.DemandeAjout
        Try
            Dim NewFamille = Ref.GetNewFamilleDeProduit
            Me.DG_Master.SelectedItem = NewFamille
        Catch ex As Exception
            ManageErreur(ex)
        End Try
    End Sub

    Private Sub UC_CmdesCRUD_DemandeSuppression() Handles UC_CmdesCRUD.DemandeSuppression
        Try
            Dim Famille = Me.DG_Master.SelectedItem
            ResseterLesFamillesDesProduitsAssociés(Famille)
            Ref.FamillesDeProduit.Remove(Famille)
        Catch ex As Exception
            ManageErreur(ex, "Echec de la suppression du produit.")
        End Try
    End Sub

    Private Sub ResseterLesFamillesDesProduitsAssociés(famille As Object)
        Dim ProduitsAssociés = From p In Ref.Produits Where p.Famille Is famille
        ProduitsAssociés.DoForAll(Sub(p) p.Famille = Nothing)
    End Sub


#End Region

End Class
