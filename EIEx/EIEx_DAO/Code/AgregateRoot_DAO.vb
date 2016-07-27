
Imports System.Xml.Serialization
Imports Model

<Serializable>
Public MustInherit Class AgregateRoot_DAO(Of T As {Model.Entité, IAgregateRoot})
    Inherits SystèmesItems_DAO(Of T)

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(Model As T)
        MyBase.New(Model)
        Me.Id = Model.Id
    End Sub

#End Region

#Region "Propriétés"

    <XmlAttribute>
    Public Property Id As Integer

#End Region

#Region "Méthodes"

    Protected Overrides Sub UnSerialized_Ex(nouvelleEntité As T)
        UnSerialized_Ex_Ex(nouvelleEntité)
        If nouvelleEntité.Id Is Nothing Then Throw New Exception($"L'objet désérialisé ""{GetType(T).Name}"" n'a pas d'Id.")
    End Sub

    Protected MustOverride Sub UnSerialized_Ex_Ex(nouvelleEntité As T)

#End Region

End Class
