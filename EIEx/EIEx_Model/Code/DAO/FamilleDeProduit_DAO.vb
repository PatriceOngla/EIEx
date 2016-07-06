Imports EIEx_Model

Public Class FamilleDeProduit_DAO
    Inherits EIEx_Object_DAO(Of FamilleDeProduit)

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(F As FamilleDeProduit)
        MyBase.New(F)
        Me.Marge = F.Marge
    End Sub

#End Region

#Region "Propriétés"

    Public Property Marge() As Single

#End Region

#Region "Méthodes"

    Public Overrides Function UnSerialized_Ex() As FamilleDeProduit
        Dim r As New FamilleDeProduit
        r.Marge = Marge
        Return r
    End Function


#End Region

#Region "Tests et debuggage"


#End Region

End Class
