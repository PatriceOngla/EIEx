Public MustInherit Class Entité
    Inherits EIExObject

    Public Sub New()
        Me._HoroDateDeCréation = Now
    End Sub

#Region "Système"
    Public MustOverride ReadOnly Property Système() As Système '(Of AgregateRoot)
#End Region

#Region "HoroDateDeCréation"
    Friend _HoroDateDeCréation As Date
    Public ReadOnly Property HoroDateDeCréation() As Date
        Get
            Return _HoroDateDeCréation
        End Get
    End Property
#End Region

End Class
