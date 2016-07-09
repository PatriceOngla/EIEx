Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Xml.Serialization

Public Module Utils

#Region "Sérialisation XML"

    Private Sub CheckPath(Chemin As String)
        Dim FolderPath = IO.Path.GetDirectoryName(Chemin)
        If Not IO.Directory.Exists(FolderPath) Then IO.Directory.CreateDirectory(FolderPath)
    End Sub

    'Public Sub Sérialiser(Of T)(Objet As T, Chemin As String)
    '    Try
    '        CheckPath(Chemin)
    '        Dim xsz = New XmlSerializer(GetType(T))
    '        Using sw As New StreamWriter(Chemin)
    '            xsz.Serialize(sw, Objet)
    '        End Using
    '        Debug.Print("Fin")
    '    Catch ex As Exception
    '        ManageError(ex, NameOf(Sérialiser))
    '    End Try
    'End Sub
    Public Sub Sérialiser(Objet As Object, Chemin As String)
        Try
            CheckPath(Chemin)
            Dim xsz = New XmlSerializer(Objet.GetType())
            Using sw As New StreamWriter(Chemin)
                xsz.Serialize(sw, Objet)
            End Using
            'Debug.Print("Fin")
        Catch ex As Exception
            ManageError(ex, NameOf(Sérialiser))
        End Try
    End Sub

    Public Function DéSérialisation(Of T)(Chemin As String) As T
        Try
            Dim r As T
            Dim xsz = New XmlSerializer(GetType(T))
            Using sw As New StreamReader(Chemin)
                r = xsz.Deserialize(sw)
            End Using
            Return r
        Catch ex As Exception
            ManageError(ex, NameOf(DéSérialisation))
        End Try
    End Function

#End Region

#Region "Divers"

#Region "Gestion des collection"

    <Extension>
    Public Sub AddRange(Of T)(List As IList(Of T), AddedRange As IEnumerable(Of T))
        For Each itemToAdd In AddedRange
            List.Add(itemToAdd)
        Next
    End Sub

    <Extension>
    Public Sub Clear(L As IList)
        For i = 1 To L.Count()
            L.Remove(0)
        Next
    End Sub

    <Extension>
    Public Sub DoForAll(Of T)(L As IEnumerable(Of T), Action As Action(Of T))
        For Each item In L
            Action(item)
        Next
    End Sub

    '<Extension>
    'Public Sub DoForAll(L As IEnumerable, Action As Action(Of Object))
    '    For Each item In L
    '        Action(item)
    '    Next
    'End Sub

    <Extension>
    Public Function TrueForAll(Of T)(L As IEnumerable(Of T), Test As Predicate(Of T)) As Boolean
        For Each item In L
            If Not Test(item) Then Return False
        Next
        Return True
    End Function

#End Region

    Public Sub ManageError(ex As Exception, SubName As String, Optional Msg As String = Nothing)
        Debug.Print($"Erreur dans la routine {SubName}: {ex.GetType.Name} {vbCr & ex.Message}")
    End Sub

#End Region

#Region "Test & debug"

    'Public Sub Test()
    '    Dim xsz = New XmlSerializer(GetType(T))
    '    Using sw As New StreamWriter(Chemin)
    '        xsz.Serialize(sw, Objet)
    '    End Using

    'End Sub

#End Region

End Module
