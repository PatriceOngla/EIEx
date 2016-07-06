Public Class UsageDeProduit_DAO
    Inherits EIEx_Object_DAO(Of UsageDeProduit)

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(UP As UsageDeProduit)
        MyBase.New(UP)
        Me.Produit = New Produit_DAO(UP.Produit)
        Me.Nombre = UP.Nombre
    End Sub

#End Region

#Region "Propriétés"

    Public Property Produit() As Produit_DAO

    Public Property Nombre() As Integer

#End Region

#Region "Méthodes"

    Public Overrides Function UnSerialized_Ex() As UsageDeProduit
        Dim r As New UsageDeProduit
        r.Produit = Me.Produit.UnSerialized
        r.Nombre = Me.Nombre
        'Dim UsagesDeProduit = From up In Me.UsagesDeProduit Select up.UnSerialized()
        Return r
    End Function

#End Region

#Region "Tests et debuggage"


#End Region

End Class
