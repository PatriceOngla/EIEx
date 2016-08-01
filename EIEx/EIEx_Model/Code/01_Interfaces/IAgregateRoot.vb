Imports System.Runtime.CompilerServices

Public Interface IAgregateRoot

    ReadOnly Property Id() As Integer?

End Interface

Friend Module AgregateRootHelper

    <Extension>
    Public Function ToStringForAgregateRoot(ar As IAgregateRoot, ToStringBase As String) As String
        Try
            Dim r = $"{ToStringBase} n° {If(ar.Id, "???")}"
            Return r
        Catch ex As Exception
            Utils.ManageError(ex, NameOf(ToString))
            Return ToStringBase
        End Try
    End Function

End Module