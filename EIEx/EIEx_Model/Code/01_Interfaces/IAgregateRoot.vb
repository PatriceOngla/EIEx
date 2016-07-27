'Public MustInherit Class AgregateRoot_Base
'    Inherits Entité

'    Public Sub New(Id As Integer)
'        Me.Id = Id
'    End Sub

'    Public ReadOnly Property Id() As Integer?

'    Public Overrides Function ToString() As String
'        Try
'            Dim r = $"{MyBase.ToString} n° {If(Id, "???")}"
'            Return r
'        Catch ex As Exception
'            Utils.ManageError(ex, NameOf(ToString))
'            Return MyBase.ToString()
'        End Try
'    End Function

'End Class


'''' <typeparam name="TypeAgregateRoot">Le type AgregateRoot lui-même. Permet les appels de méthodes génériques sur <see cref="Système"/></typeparam>
'Public MustInherit Class AgregateRoot(Of TypeAgregateRoot As AgregateRoot_Base)
'    Inherits AgregateRoot_Base
'    Public Sub New(Id As Integer)
'        MyBase.New(Id)
'    End Sub

'End Class

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