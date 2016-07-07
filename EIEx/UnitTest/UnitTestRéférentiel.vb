Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Model

<TestClass()> Public Class UnitTestRéférentiel

#Region "Ref (Référentiel)"
    Public ReadOnly Property Ref() As Référentiel
        Get
            Return Référentiel.Instance
        End Get
    End Property
#End Region

    <TestMethod()> Public Sub TesterSérialisation()

        Assert.IsTrue(Ref IsNot Nothing)

        Dim NbObjets = 10

        PeuplerRéférentiel(NbObjets)

        EIExData.EnregistrerRéférentiel()
        Assert.IsTrue(Ref.Produits.Count = NbObjets)
        Assert.IsTrue(Ref.RéférencesDOuvrage.Count = NbObjets)
        Assert.IsTrue(Ref.FamillesDeProduit.Count = NbObjets)

        Assert.IsTrue(IO.File.Exists(EIExData.CheminRéférentiel))

        Assert.IsTrue(Ref IsNot Nothing)

        Ref.Purger()

        Assert.IsTrue(Ref.EstVide())

        EIExData.ChargerLeRéférentiel()

        Assert.IsTrue(Ref IsNot Nothing)
        Assert.IsTrue(Ref.Produits.Count = NbObjets)
        Assert.IsTrue(Ref.RéférencesDOuvrage.Count = NbObjets)
        Assert.IsTrue(Ref.FamillesDeProduit.Count = NbObjets)

    End Sub

    Private Sub PeuplerRéférentiel(NbObjets As Integer)
        For i = 1 To 10
            NewFamille(i)
        Next
        For i = 1 To 10
            NewProduit(i)
        Next
        For i = 1 To 10
            NewRéférenceDOuvrage(i)
        Next

    End Sub

    Private Function NewProduit(i As Integer) As Produit
        Dim r = New Produit()
        Dim f = Ref.GetFamilleById(i)
        r.Nom = "Produit " & i : r.Unité = Unités.U : r.Prix = 100 + i : r.ReférenceFournisseur = "Ref_" & i : r.Famille = f : r.TempsDePauseUnitaire = i
        Return r
    End Function

    Private Function NewFamille(i As Integer) As FamilleDeProduit
        'Dim r = Ref.GetNewFamilleDeProduit
        Dim r = New FamilleDeProduit
        r.Nom = "Famille " & i : r.Marge = i
        Return r
    End Function

    Private Function NewRéférenceDOuvrage(i As Integer) As RéférenceDOuvrage
        Dim r = New RéférenceDOuvrage
        r.Nom = "Ouvrage " & i : r.TempsDePauseUnitaire = i : r.PrixUnitaire = i : r.Libellés.Add("Libellé supplémentaire " & i)
        r.AjouterProduit(i, i)
        Return r
    End Function

End Class