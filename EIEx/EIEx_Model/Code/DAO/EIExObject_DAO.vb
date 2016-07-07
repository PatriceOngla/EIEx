
Imports System.Xml.Serialization

<Serializable>
Public MustInherit Class EIEx_Object_DAO(Of T As {EIExObject})

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(Model As T)
        Me.Nom = Model.Nom
    End Sub

#End Region

#Region "Propriétés"

#Region "Réf (Référentiel)"
    <XmlIgnore>
    Public ReadOnly Property Réf() As Référentiel
        Get
            Return Référentiel.Instance
        End Get
    End Property
#End Region

    '<XmlAttribute>
    'Public Property Id As Integer

    <XmlAttribute>
    Public Property Nom As String

#End Region

#Region "Méthodes"

    Public Function UnSerialized() As T
        Dim r = UnSerialized_Ex()
        r.Nom = Me.Nom
        Return r
    End Function

    Protected MustOverride Function UnSerialized_Ex() As T

#End Region

#Region "Tests et debuggage"


#End Region

End Class
