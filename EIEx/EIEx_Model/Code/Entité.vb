Public MustInherit Class Entité
    Inherits EIExObject

#Region "Constructeurs"

    Public Sub New()
        Me._HoroDateDeCréation = Now
    End Sub

#End Region

#Region "Propriétés"

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

#End Region

#Region "Méthodes"

#Region "FormateForColumn"

    Protected Shared Function FormateForColumn(s As String, width As Short, Optional AddSep As Boolean = False) As String
        Const Margin = "  "
        Dim r = Margin & Left(s, width).PadRight(width) & Margin & (If(AddSep, "|", ""))
        Return r
    End Function

#End Region

#End Region

End Class
