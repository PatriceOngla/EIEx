Imports System.Runtime.CompilerServices
Module Utils

    <Extension>
    Public Sub AddRange(Of T)(List As IList(Of T), AddedRange As IEnumerable(Of T))
        For Each itemToAdd In AddedRange
            List.Add(itemToAdd)
        Next
    End Sub

End Module
