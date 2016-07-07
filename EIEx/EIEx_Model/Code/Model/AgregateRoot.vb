Imports System.Collections.ObjectModel

Public MustInherit Class AgregateRoot
    Inherits EIExObject

#Region "Constructeurs"

    Public Sub New()
        Me.SetId()
        Me.SEnregistrerDansLeRéférentiel()
    End Sub

    Public Sub New(Id As Integer)
        Me._Id = Id
        Me.SEnregistrerDansLeRéférentiel()
    End Sub

    Protected MustOverride Sub SetId()

    Protected MustOverride Sub SEnregistrerDansLeRéférentiel()

#End Region

#Region "Propriétés"

#Region "Id (Integer)"
    Protected _Id As Integer?
    Public ReadOnly Property Id() As Integer?
        Get
            Return _Id
        End Get
    End Property
#End Region

#End Region

#Region "Méthodes"

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


#End Region

#Region "Tests et debuggage"


#End Region

End Class
