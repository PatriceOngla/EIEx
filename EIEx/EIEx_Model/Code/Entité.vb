Public MustInherit Class Entité
    Inherits EIExObject

    Public Sub New()
        Me._HoroDateDeCréation = Now
    End Sub

#Region "Système"
    Public MustOverride ReadOnly Property Système() As Système '(Of AgregateRoot)
#End Region

#Region "HoroDateDeCréation"
    Protected _HoroDateDeCréation As Date
    Public ReadOnly Property HoroDateDeCréation() As Date
        Get
            Return _HoroDateDeCréation
        End Get
    End Property

    ''' <summary>Nécessaire uniquement en désérialisation.</summary>
    <ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
    Public Sub ForcerHoroDateDeCréation(value As Date)
        _HoroDateDeCréation = value
    End Sub

#End Region

End Class
