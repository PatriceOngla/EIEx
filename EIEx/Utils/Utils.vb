﻿Imports System.IO
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

#Region "ContainsList_String"
    Private scb As BlurStringComparer = BlurStringComparer.DefaultInstance
    ''' <summary>
    ''' Attention, si <paramref name="DistanceTolérée"/> est supérieur à 0 alors <paramref name="StartWith"/> est ignoré. 
    ''' </summary>
    ''' <param name="DistanceTolérée">Distance de Levenshtein</param>
    <Extension>
    Public Function ContainsList_String(List As IEnumerable(Of String), OtherList As IEnumerable(Of String), Optional ignoreCase As Boolean = True, Optional StartWith As Boolean = False, Optional DistanceTolérée As Short = 0) As Boolean
        Dim sc As StringComparer
        If DistanceTolérée > 0 Then
            scb.IgnoreCase = ignoreCase
            scb.Blur = DistanceTolérée
            sc = scb
        Else
            sc = If(ignoreCase, StringComparer.CurrentCultureIgnoreCase, StringComparer.CurrentCulture)
        End If
        Dim scn = If(ignoreCase, StringComparison.CurrentCultureIgnoreCase, StringComparison.CurrentCulture)

        Return ContainsList_String(List, OtherList, StartWith, sc, scn, DistanceTolérée)
    End Function

    Private Function ContainsList_String(List As IEnumerable(Of String), OtherList As IEnumerable(Of String), StartWith As Boolean, sc As StringComparer, scn As StringComparison, DistanceTolérée As Short) As Boolean
        'Dim r As Boolean
        For Each item In OtherList

            If StartWith AndAlso DistanceTolérée = 0 Then
                Dim r = (From s In List Where s.StartsWith(item, scn)).Any
                If Not r Then Return False
            Else
                If Not List.Contains(item, sc) Then Return False
            End If
        Next
        Return True

        'If DistanceTolérée = 0 Then
        '    r = ContainsList_String_Strict(List, OtherList, StartWith, sc, scn)
        'Else
        '    r = ContainsList_String_Strict(List, OtherList, StartWith, sc, scn)
        '    'r = ContainsList_String_Blur(List, OtherList, sc, scn, DistanceTolérée)
        'End If
        'Return r
    End Function
    'Private Function ContainsList_String_Strict(List As IEnumerable(Of String), OtherList As IEnumerable(Of String), StartWith As Boolean, sc As StringComparer, scn As StringComparison) As Boolean

    '    For Each item In OtherList
    '        If StartWith Then
    '            Dim r = (From s In List Where s.StartsWith(item, scn)).Any
    '            If Not r Then Return False
    '        Else
    '            If Not List.Contains(item, sc) Then Return False
    '        End If
    '    Next
    '    Return True
    'End Function
    'Private Function ContainsList_String_Blur(List As IEnumerable(Of String), OtherList As IEnumerable(Of String), sc As StringComparer, scn As StringComparison, DistanceTolérée As Short) As Boolean

    '    For Each item In OtherList
    '        If StartWith Then
    '            Dim r = (From s In List Where s.StartsWith(item, scn)).Any
    '            If Not r Then Return False
    '        Else
    '            If Not List.Contains(item, sc) Then Return False
    '        End If
    '    Next
    '    Return True
    'End Function
    ''' <summary>
    ''' Based on Levenshtein distance. 
    ''' Be carrefull : Blur is only used for equality test. It is ignored for sorting.
    ''' </summary>
    Private Class BlurStringComparer
        Inherits StringComparer

#Region "Constructeurs"
        Public Sub New(ByVal Blur As Short, ignoreCase As Boolean)
            Me.Blur = Blur
            Me.IgnoreCase = ignoreCase
        End Sub
#End Region

#Region "DefaultInstance"
        Private Shared _DefaultInstance As BlurStringComparer = New BlurStringComparer(1, False)
        Public Shared ReadOnly Property DefaultInstance As BlurStringComparer
            Get
                Return _DefaultInstance
            End Get
        End Property
#End Region

#Region "Propriétés"
        Public Property Blur As Short
        Public Property IgnoreCase As Boolean
#End Region

#Region "Méthodes"

        Public Overrides Function Compare(x As String, y As String) As Integer
            Return String.Compare(x, y, Me.IgnoreCase)
        End Function

        Public Overrides Function Equals(x As String, y As String) As Boolean
            If String.Equals(x, y, If(Me.IgnoreCase, StringComparison.CurrentCultureIgnoreCase, StringComparison.CurrentCulture)) Then
                Return True
            Else
                If Me.IgnoreCase Then
                    x = UCase(x)
                    y = UCase(y)
                End If
                Dim d = DistanceDeLevenshtein(x, y)
                Return d <= Blur
            End If
        End Function

        Public Overrides Function GetHashCode(obj As String) As Integer
            Return obj.GetHashCode
        End Function

#End Region

    End Class

#End Region

    <Extension>
    Public Function ContainsList(Of T)(List As IEnumerable(Of T), OtherList As IEnumerable(Of T)) As Boolean
        For Each item In OtherList
            If Not List.Contains(item) Then Return False
        Next
        Return True
    End Function

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

    '''<summary>
    '''Returns the item following or preceding <paramref name="Item"/> according to the value of <paramref name="[Next]"/>.
    '''</summary>
    ''' <param name="[Next]">If true, returns the item following <paramref name="Item"/>, otherwise returns the item preceding <paramref name="Item"/>.</param>
    ''' <returns></returns>
    <Extension> Public Function GetNextOrPrevious(Of T)(L As IEnumerable(Of T), Item As T, [Next] As Boolean) As T
        Dim r = GetNextForIEnumerable(L, Item, [Next])
        Return r
    End Function

    'Private Function GetNextForIList(Of T)(L As IList(Of T), Item As T, [Next] As Boolean) As T
    '    Dim i = L.IndexOf(Item)
    '    If (i = 0 And [Next]) OrElse (i = L.Count - 1 And Not [Next]) Then
    '        Throw New IndexOutOfRangeException()
    '    Else
    '        Dim ir = If([Next], i + 1, i - 1)
    '        Return L(ir)
    '    End If
    'End Function

    Private Function GetNextForIEnumerable(Of T)(L As IEnumerable(Of T), Item As T, [Next] As Boolean) As T
        Dim L_loc = If([Next], L, L.Reverse)
        Dim L_skipped = L_loc.SkipWhile(Function(it As T) Not (Object.Equals(it, Item)))
        Dim r = L_skipped.Skip(1).First()
        Return r
    End Function

#End Region

#Region "Text processing"
    Public Function DistanceDeLevenshtein(ByVal s As String, ByVal t As String) As Integer
        'Dim n = s.Length
        'Dim m = t.Length
        'Dim d(n + 1, m + 1) As Integer

        '' Step 1
        'If (n = 0) Then
        '    Return m
        'End If

        'If (m = 0) Then
        '    Return n
        'End If

        '' Step 2
        'Dim i As Integer = 0
        'Do While (i <= n)
        '    i = (i + 1)
        'Loop

        'd(i, 0) = i
        'Dim j As Integer = 0
        'Do While (j <= m)
        '    j = (j + 1)
        'Loop

        'd(0, j) = j
        '' Step 3

        'Mouline(s, t, n, m, d)

        '' Step 7
        'Return d(n, m)

        Dim Matrix(s.Length, t.Length) As Integer
        Dim Key As Integer
        For Key = 0 To s.Length
            Matrix(Key, 0) = Key
        Next
        For Key = 0 To t.Length
            Matrix(0, Key) = Key
        Next
        For Key1 As Integer = 1 To t.Length
            For Key2 As Integer = 1 To s.Length
                If s(Key2 - 1) = t(Key1 - 1) Then
                    Matrix(Key2, Key1) = Matrix(Key2 - 1, Key1 - 1)
                Else
                    Matrix(Key2, Key1) = Math.Min(Matrix(Key2 - 1, Key1) + 1, Math.Min(Matrix(Key2, Key1 - 1) + 1, Matrix(Key2 - 1, Key1 - 1) + 1))
                End If
            Next
        Next
        'Return Matrix(s.Length - 1, t.Length - 1)
        Return Matrix(s.Length, t.Length)

    End Function

    Private Sub Mouline(ByVal s As String, ByVal t As String, n As Integer, m As Integer, d(,) As Integer)
        Dim i As Integer = 1
        Do While (i <= n)
            'Step 4
            Dim j As Integer = 1
            Do While (j <= m)
                ' Step 5
                Dim cost = If(t((j - 1)) = s((i - 1)), 0, 1)

                ' Step 6
                d(i, j) = Math.Min(Math.Min((d((i - 1), j) + 1), (d(i, (j - 1)) + 1)), (d((i - 1), (j - 1)) + cost))
                j = (j + 1)
            Loop

            i = (i + 1)
        Loop
    End Sub

#End Region
    Public Sub ManageError(ex As Exception, SubName As String, Optional Msg As String = Nothing, Optional DlgBox As Boolean = False, Optional Titre As String = Nothing)
        Dim msg2 = $"Erreur dans la routine ""{SubName}"".{vbCr}{If(String.IsNullOrEmpty(Msg), "", Msg & vbCr)}{ex.GetType.Name} : { ex.Message}"
        Debug.Print(Msg)
        If DlgBox Then MsgBox(msg2, vbExclamation, Titre)
    End Sub

#End Region

#Region "Test & debug"

#End Region

End Module
