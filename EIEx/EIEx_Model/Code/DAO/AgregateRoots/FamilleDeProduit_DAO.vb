Imports System.Xml.Serialization

<Serializable>
Public Class FamilleDeProduit_DAO
    Inherits AgregateRoot_DAO(Of FamilleDeProduit)

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(F As FamilleDeProduit)
        MyBase.New(F)
        Me.Marge = F.Marge
    End Sub

#End Region

#Region "Propriétés"

    Public Property Marge() As Single?

#End Region

#Region "Méthodes"

    Protected Overrides Function UnSerialized_Ex_Ex() As FamilleDeProduit
        Dim r = Réf.GetFamilleById(Me.Id)
        r = If(r, New FamilleDeProduit(Me.Id))
        r.Marge = Marge
        Return r
    End Function


#End Region

#Region "Tests et debuggage"


#End Region

End Class
