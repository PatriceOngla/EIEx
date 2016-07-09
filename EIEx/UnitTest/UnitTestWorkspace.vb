Imports System.IO
Imports System.Text
Imports System.Xml.Serialization
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

        EIExData.EnregistrerLeWorkspace()
        CopierLeFichier(EIExData.CheminWorkspace)

        Assert.IsTrue(WS.Etudes.Count = NbObjets)

        Assert.IsTrue(IO.File.Exists(EIExData.CheminWorkspace))

        Assert.IsTrue(WS IsNot Nothing)

        WS.Purger()

        Assert.IsTrue(WS.EstVide())

        EIExData.ChargerLeWorkspace()

        Assert.IsTrue(WS IsNot Nothing)
        Assert.IsTrue(WS.Etudes.Count = NbObjets)

        EIExData.EnregistrerLeWorkspace()

    End Sub

    Private Sub PeuplerWorkspace(NbObjets As Integer)
        For i = 1 To 10
            NewEtudes(i)
        Next
    End Sub

    Private Function NewEtudes(i As Integer) As Etude
        Dim r = WS.GetNewEtude()
        r.Nom = $"Etude {i}"
        Dim Bdx As New List(Of Bordereau)
        For i = 1 To i
            Bdx.Add(NewBordereau(i))
        Next
        r.Bordereaux.AddRange(Bdx)
        Return r
    End Function

    Private Function NewBordereau(i As Integer) As Bordereau
        Dim r = New Bordereau
        r.Nom = "Bordereau " & i : r.CheminFichier = "c:\dossier " & i
        r.Paramètres.AdresseRangeLibelleOuvrage = $"A{i}L{i}"
        r.Paramètres.AdresseRangeUnité = $"A{i}U{i}"
        r.Paramètres.AdresseRangePrixUnitaire = $"A{i}P{i}"
        Return r
    End Function

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

#End Region

End Class