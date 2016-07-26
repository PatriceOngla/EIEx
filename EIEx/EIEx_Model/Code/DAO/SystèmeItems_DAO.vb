Imports System.Runtime.Serialization

Public MustInherit Class SystèmesItems_DAO(Of T As {Entité})
    Inherits EIEx_Object_DAO(Of T)

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(Modèle As T)
        MyBase.New(Modèle)
        Me.HoroDateDeCréation = Modèle.HoroDateDeCréation
    End Sub

#End Region

#Region "Propriétés"

    Public ReadOnly Property HoroDateDeCréation() As Date

#End Region

#Region "Méthodes"

    Public Function UnSerialized() As T
        Dim r = UnSerialized_Ex()
        r.ForcerHoroDateDeCréation(Me.HoroDateDeCréation)
        r.Nom = Me.Nom
        r.Commentaires = Me.Commentaires
        Return r
    End Function

    Protected MustOverride Function UnSerialized_Ex() As T

#End Region

End Class
