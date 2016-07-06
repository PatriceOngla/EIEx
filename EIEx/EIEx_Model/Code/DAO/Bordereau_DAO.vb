Imports System.Windows
Imports EIEx_Model

Public Class Bordereau_DAO
    Inherits EIEx_Object_DAO(Of Bordereau)

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(B As Bordereau)
        MyBase.New(B)
        Me.Paramètres = New Paramètres_DAO(B.Paramètres)
        Me.CheminFichier = B.CheminFichier
    End Sub

#End Region

#Region "Propriétés"

    Public Property CheminFichier() As String

    Public Property Paramètres() As Paramètres_DAO

#End Region

#Region "Méthodes"

    Public Overrides Function UnSerialized_Ex() As Bordereau
        Dim r As New Bordereau
        r.CheminFichier = Me.CheminFichier
        r.Paramètres.AdresseRangeLibelleOuvrage = Me.Paramètres.AdresseRangeLibelleOuvrage
        r.Paramètres.AdresseRangePrixUnitaire = Me.Paramètres.AdresseRangePrixUnitaire
        r.Paramètres.AdresseRangeUnité = Me.Paramètres.AdresseRangeUnité
        Return r
    End Function

#End Region

#Region "Tests et debuggage"


#End Region

End Class
