Imports Model

Public Class XXX


    Public Shared Function ToutesLesFamillesDeProduit() As IEnumerable(Of FamilleDeProduit)
        Return Référentiel.Instance.FamillesDeProduit
    End Function

    Public Shared Function TousLesProduits() As IEnumerable(Of Produit)
        Return Référentiel.Instance.Produits
    End Function

End Class
