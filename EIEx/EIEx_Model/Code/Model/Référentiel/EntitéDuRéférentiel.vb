Public MustInherit Class EntitéDuRéférentiel
    Inherits Entité

#Region "Propriétés"

    Protected Ref As Référentiel = Référentiel.Instance
    Public Overrides ReadOnly Property Système As Système
        Get
            Return Ref
        End Get
    End Property

#End Region

End Class
