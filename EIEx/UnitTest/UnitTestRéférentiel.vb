Imports System.Text
Imports EIEx_DAO
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Model
Imports Utils

<TestClass()> Public Class UnitTestRéférentiel

#Region "Ref (Référentiel)"
    Public ReadOnly Property Ref() As Référentiel
        Get
            Return Référentiel.Instance
        End Get
    End Property
#End Region

    <TestMethod()> Public Sub Référentiel_TesterSérialisation()

        Assert.IsTrue(Ref IsNot Nothing)

        Dim NbObjets = 10

        PeuplerRéférentiel(NbObjets)

        PersistancyManager.EnregistrerLeRéférentiel()
        CopierLeFichier(PersistancyManager.CheminRéférentiel)

        Assert.IsTrue(Ref.Produits.Count = NbObjets)
        Assert.IsTrue(Ref.Ouvrage.Count = NbObjets)
        Assert.IsTrue(Ref.FamillesDeProduit.Count = NbObjets)

        Assert.IsTrue(IO.File.Exists(PersistancyManager.CheminRéférentiel))

        Assert.IsTrue(Ref IsNot Nothing)

        Ref.Purger()

        Assert.IsTrue(Ref.EstVide())

        PersistancyManager.ChargerLeRéférentiel()

        Assert.IsTrue(Ref IsNot Nothing)
        Assert.IsTrue(Ref.Produits.Count = NbObjets)
        Assert.IsTrue(Ref.Ouvrage.Count = NbObjets)
        Assert.IsTrue(Ref.FamillesDeProduit.Count = NbObjets)

        PersistancyManager.EnregistrerLeRéférentiel()

    End Sub

    Private Sub PeuplerRéférentiel(NbObjets As Integer)
        For i = 1 To 10
            NewFamille(i)
        Next
        For i = 1 To 10
            NewProduit(i)
        Next
        For i = 1 To 10
            NewOuvrage(i)
        Next

    End Sub

    Private Function NewProduit(i As Integer) As Produit
        Dim r = Ref.GetNewProduit()
        Dim f = Ref.GetFamilleById(i)
        r.Nom = "Produit " & i : r.Unité = Unités.U : r.Prix = 100 + i : r.RéférenceFournisseur = "100" & i : r.CodeLydic = "CONS" & i : r.Famille = f : r.TempsDePauseUnitaire = i
        r.MotsClés.AddRange({"keyWord " & i, "keyWord " & i + 1})
        Return r
    End Function

    Private Function NewFamille(i As Integer) As FamilleDeProduit
        Dim r = Ref.GetNewFamilleDeProduit()
        r.Nom = "Famille " & i : r.Marge = i
        Return r
    End Function

    Private Function NewOuvrage(i As Integer) As Ouvrage
        Dim r = Ref.GetNewOuvrage
        r.Nom = "Ouvrage " & i : r.TempsDePauseUnitaire = i : r.PrixUnitaire = i : r.Libellés.Add("Libellé supplémentaire " & i)
        AjouterProduitsALouvrage(r, i)
        Return r
    End Function

    Private Sub AjouterProduitsALouvrage(po As Ouvrage, i As Integer)

        For j = 1 To i
            Dim p = Ref.GetProduitById(j)
            po.AjouterProduit(p, j)
        Next

    End Sub
End Class