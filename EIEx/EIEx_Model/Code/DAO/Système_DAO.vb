Imports System.Runtime.Serialization
Imports System.Xml.Serialization
Imports Model

Public Interface ISystèmeDAO
    <XmlIgnore>
    Property DateModif As Date
    Sub UnSerialize(NewT As Système)
End Interface

Public MustInherit Class Système_DAO(Of T As Système)
    Inherits EIEx_Object_DAO(Of T)
    Implements ISystèmeDAO

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(Modèle As T)
        MyBase.New(Modèle)
        Me.DateModif = Modèle.DateModif
    End Sub

#End Region

#Region "Propriétés"

    Public Property DateModif() As Date Implements ISystèmeDAO.DateModif

#End Region

#Region "Méthodes"

    Public Sub UnSerialize(NewT As T)
        NewT.Purger()
        NewT.DateModif = Me.DateModif
        NewT.Nom = Me.Nom
        NewT.Commentaires = Me.Commentaires
        UnSerialize_Ex(NewT)
    End Sub

    Protected MustOverride Sub UnSerialize_Ex(NewT As T)

    Public Sub UnSerialize2(NewT As Système) Implements ISystèmeDAO.UnSerialize
        UnSerialize(NewT)
    End Sub

#End Region

End Class
