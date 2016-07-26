
Imports System.Xml.Serialization

<Serializable>
Public MustInherit Class EIEx_Object_DAO(Of T As {EIExObject})

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(Model As T)
        'Me.Modèle = Model
        Me.Nom = Model.Nom
        Me.Commentaires = Model.Commentaires
    End Sub

#End Region

#Region "Propriétés"

#Region "Modèle"
    '<XmlIgnore>
    'Protected ReadOnly Property Modèle() As T
#End Region

#Region "Sys (Système)"
    <XmlIgnore>
    Protected MustOverride ReadOnly Property Sys() As Système
#End Region

    <XmlAttribute>
    Public Property Nom As String

    Public Property Commentaires As String

#End Region

#Region "Tests et debuggage"


#End Region

End Class
