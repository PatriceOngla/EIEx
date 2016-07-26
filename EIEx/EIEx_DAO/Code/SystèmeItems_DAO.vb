Imports System.Runtime.Serialization

''' <summary>Classe système. Ces classes forment un tout fonctionnel cohérent et sont objet de sérialisation.</summary>
''' <typeparam name="T">Une <see cref="Entité"/>.</typeparam>
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

    Public Sub UnSerialized(nouvelleEntité As T)
        UnSerialized_Ex(nouvelleEntité)
        nouvelleEntité.ForcerHoroDateDeCréation(Me.HoroDateDeCréation)
        nouvelleEntité.Nom = Me.Nom
        nouvelleEntité.Commentaires = Me.Commentaires
    End Sub

    Protected MustOverride Sub UnSerialized_Ex(nouvelleEntité As T)

#End Region

End Class
