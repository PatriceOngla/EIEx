Public MustInherit Class EntitéDuWorkSpace
    Inherits Entité

#Region "Système"
    Public Overrides ReadOnly Property Système As Système
        Get
            Return WorkSpace.Instance
        End Get
    End Property

#End Region

End Class
