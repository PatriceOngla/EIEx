Public MustInherit Class AgregateRoot_Base
    Inherits Entité

#Region "Id (Integer)"
    Protected _Id As Integer?
    Public ReadOnly Property Id() As Integer?
        Get
            Return _Id
        End Get
    End Property
#End Region

#Region "TosTring"
    Public Overrides Function ToString() As String
        Try
            Dim r = $"{MyBase.ToString} n° {If(Id, "???")}"
            Return r
        Catch ex As Exception
            Utils.ManageError(ex, NameOf(ToString))
            Return MyBase.ToString()
        End Try
    End Function
#End Region

End Class


''' <typeparam name="TypeAgregateRoot">Le type AgregateRoot lui-même. Permet les appels de méthodes génériques sur <see cref="Système"/></typeparam>
Public MustInherit Class AgregateRoot(Of TypeAgregateRoot As AgregateRoot_Base)
    Inherits AgregateRoot_Base

#Region "Constructeurs"

    Public Sub New(Id As Integer)
        Me._Id = Id
    End Sub

#End Region

End Class
