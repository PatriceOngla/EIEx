Imports System.IO
Imports System.Text
Imports System.Xml.Serialization
Imports EIEx_DAO
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Model
Imports Utils

<TestClass()> Public Class UnitTestWorkspace

#Region "WS (Workspace)"
    Public ReadOnly Property WS() As WorkSpace
        Get
            Return WorkSpace.Instance
        End Get
    End Property
#End Region

    <TestMethod()> Public Sub Workspace_TesterSérialisation()

        Assert.IsTrue(WS IsNot Nothing)

        Dim NbObjets = 10

        PeuplerWorkspace(NbObjets)

        PersistancyManager.EnregistrerLeWorkspace()
        CopierLeFichier(PersistancyManager.CheminWorkspace)

        Assert.IsTrue(WS.Etudes.Count = NbObjets)

        Assert.IsTrue(IO.File.Exists(PersistancyManager.CheminWorkspace))

        Assert.IsTrue(WS IsNot Nothing)

        WS.Purger()

        Assert.IsTrue(WS.EstVide())

        PersistancyManager.ChargerLeWorkspace()

        Assert.IsTrue(WS IsNot Nothing)
        Assert.IsTrue(WS.Etudes.Count = NbObjets)

        PersistancyManager.EnregistrerLeWorkspace()

    End Sub

    Private Sub PeuplerWorkspace(NbObjets As Integer)
        For i = 1 To 10
            NewEtudes(i)
        Next
    End Sub

    Private Function NewEtudes(i As Integer) As Etude
        Dim r = WS.GetNewEtude()
        r.Nom = $"Etude {i}"
        r.Client = $"Client {i}"
        Dim Clsrs As New List(Of ClasseurExcel)
        Dim Clsr As ClasseurExcel
        For j = 1 To i
            Clsr = r.AjouterNouveauClasseur("c:\dossier " & i & j)
            FillClasseur(Clsr, i)
        Next
        r.ClasseursExcel.AddRange(Clsrs)
        Return r
    End Function

    Private Sub FillClasseur(XC As ClasseurExcel, i As Integer)
        XC.Nom = "Classeur " & i
        For j = 1 To i
            Dim Bd = XC.AjouterNouveauBordereau()
            FillBordereau(Bd, i, j)
        Next
    End Sub

    Private Sub FillBordereau(B As Bordereau, i As Integer, j As Integer)
        B.NomFeuille = "Feuille " & i
        B.Nom = "Bordereau " & i & " - " & j
        B.Paramètres.AdresseRangeLibelleOuvrage = $"A{i}L{i}"
        B.Paramètres.AdresseRangeUnité = $"A{i}U{i}"
        B.Paramètres.AdresseRangePrixUnitaire = $"A{i}P{i}"
        B.Paramètres.AdresseRangeXYZ = $"A{i}P{i}"
        Dim NO As Ouvrage
        For k = 1 To j
            NO = B.AjouterOuvrage(i * 10 + k)
            NO.Nom = $"Ouvrage {i}-{j}-{k}"
        Next
    End Sub

#Region "Test & debug"
    <TestMethod()>
    Public Sub Test()

        Dim O2 As New Class2() With {.Nom = "Titi"}
        Dim O1 As New Class1() With {.Nom = "Toto", .Objet = O2, .INom = "IToto"}

        Dim Chemin = "C:\Temp\FTest.xml"

        Dim xsz = New XmlSerializer(GetType(Class1))
        Using sw As New StreamWriter(Chemin)
            xsz.Serialize(sw, O1)
        End Using

    End Sub

    <TestMethod()>
    Public Sub TestLevenshtein()

        Debug.Print(DistanceDeLevenshtein("toto", "totoxx"))

    End Sub

#End Region

End Class