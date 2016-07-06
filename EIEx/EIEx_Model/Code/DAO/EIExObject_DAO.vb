
<Serializable>
Public MustInherit Class EIEx_Object_DAO(Of T As {EIExObject, New})

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(Model As T)
        Me.Id = Model.Id
        Me.Nom = Model.Nom
    End Sub

#End Region

#Region "Propriétés"

    Public Property Id As Integer

    Public Property Nom As String

#End Region

#Region "Méthodes"

    Public Function UnSerialized() As T
        Dim r = UnSerialized_Ex()
        If r.Id Is Nothing Then Throw New Exception($"L'objet désérialisé ""{GetType(T).Name}"" n'a pas d'Id.")
        r.Nom = Me.Nom
        Return r
    End Function

    Public MustOverride Function UnSerialized_Ex() As T

#End Region

#Region "Tests et debuggage"


#End Region

End Class
