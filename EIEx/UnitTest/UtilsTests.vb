Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Utils

<TestClass()> Public Class UtilsTests

    <TestMethod()> Public Sub TestCollectionMethods()

        Dim L = {1, 2, 3, 4, 5, 6, 7, 8, 9}

        Dim It = 5

        Assert.IsTrue(L.GetNextOrPrevious(It, True) = 6)
        Assert.IsTrue(L.GetNextOrPrevious(It, False) = 4)

        CheckException(Sub() L.GetNextOrPrevious(12, True), "l'item n'est pas dans la liste")
        CheckException(Sub() L.GetNextOrPrevious(9, True), "l'item n'a pas de successeur")
        CheckException(Sub() L.GetNextOrPrevious(1, False), "l'item n'a pas de prédécesseur")

    End Sub

    Private Sub CheckException(A As Action, CaseDescription As String)
        Try
            A()
            Assert.Fail($"Pas d'exception déclenchée alors que {CaseDescription}.")
        Catch ex As Exception
            Assert.IsInstanceOfType(ex, GetType(InvalidOperationException), $"{CaseDescription}. {NameOf(InvalidOperationException)} attendue.")
        End Try

    End Sub

End Class