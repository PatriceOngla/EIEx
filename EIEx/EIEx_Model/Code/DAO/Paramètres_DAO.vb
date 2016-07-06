Imports EIEx_Model

Public Class Paramètres_DAO
    Inherits EIEx_Object_DAO(Of Paramètres)

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(P As Paramètres)
        MyBase.New(P)
        Me.AdresseRangeLibelleOuvrage = P.AdresseRangeLibelleOuvrage
        Me.AdresseRangePrixUnitaire = P.AdresseRangePrixUnitaire
        Me.AdresseRangeUnité = P.AdresseRangeUnité
    End Sub

#End Region

#Region "Propriétés"

    Public Property AdresseRangeLibelleOuvrage() As String

    Public Property AdresseRangeUnité() As String

    Public Property AdresseRangePrixUnitaire() As String
#End Region

#Region "Méthodes"
    Public Overrides Function UnSerialized_Ex() As Paramètres
        Dim r As New Paramètres(Me.Id)
        r.AdresseRangeLibelleOuvrage = AdresseRangeLibelleOuvrage
        r.AdresseRangePrixUnitaire = AdresseRangePrixUnitaire
        r.AdresseRangeUnité = AdresseRangeUnité
        Return r
    End Function

#End Region

#Region "Tests et debuggage"


#End Region

End Class
