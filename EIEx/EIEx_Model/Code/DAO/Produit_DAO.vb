Imports EIEx_Model

Public Class Produit_DAO
    Inherits EIEx_Object_DAO(Of Produit)

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(P As Produit)
        MyBase.New(P)
        Me.Unité = P.Unité
        Me.Prix = P.Prix
        Me.ReférenceFournisseur = P.ReférenceFournisseur
        Me.TempsDePauseUnitaire = P.TempsDePauseUnitaire
        Me.MotsClés = New List(Of String)(P.MotsClés)
        Me.Famille = New FamilleDeProduit_DAO(P.Famille)
    End Sub

#End Region

#Region "Propriétés"

    Public Property Unité() As Unités

    Public Property Prix() As Single

    Public Property ReférenceFournisseur() As String

    Public Property TempsDePauseUnitaire() As Integer?

    Public Property MotsClés() As List(Of String)

    Public Property Famille() As FamilleDeProduit_DAO

#End Region

#Region "Méthodes"

    Public Overrides Function UnSerialized_Ex() As Produit
        Dim r As New Produit
        r.Unité = Unité
        r.Prix = Prix
        r.ReférenceFournisseur = ReférenceFournisseur
        r.TempsDePauseUnitaire = TempsDePauseUnitaire
        r.MotsClés.AddRange(MotsClés)
        Return r
    End Function

#End Region

#Region "Tests et debuggage"

#End Region

End Class
