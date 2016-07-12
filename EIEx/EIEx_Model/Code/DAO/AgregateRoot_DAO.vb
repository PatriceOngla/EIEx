﻿
Imports System.Xml.Serialization
Imports Model

<Serializable>
Public MustInherit Class AgregateRoot_DAO(Of T As {AgregateRoot_Base})
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

    Protected Overrides Function UnSerialized_Ex() As T
        'Protected Overrides Function UnSerialized_Ex() As T
        Dim r = UnSerialized_Ex_Ex()
        If r.Id Is Nothing Then Throw New Exception($"L'objet désérialisé ""{GetType(T).Name}"" n'a pas d'Id.")
        Return r
    End Function

    Protected MustOverride Function UnSerialized_Ex_Ex() As T

#End Region

#Region "Tests et debuggage"


#End Region

End Class