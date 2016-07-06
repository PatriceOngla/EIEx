
Public Class RéférenceDOuvrage_DAO
    Inherits EIEx_Object_DAO(Of RéférenceDOuvrage)

#Region "Constructeurs"
    Public Sub New()
    End Sub

    Public Sub New(R As RéférenceDOuvrage)
        MyBase.New(R)

        Me.LibelléPrincipal = R.LibelléPrincipal

        Me.Libellés = New List(Of String)(R.Libellés)

        Dim UsagesDeProduit_DAO = From up In R.UsagesDeProduit Select New UsageDeProduit_DAO(up)
        Me.UsagesDeProduit = New List(Of UsageDeProduit_DAO)(UsagesDeProduit_DAO)

        Me.MotsClés = New List(Of String)(R.MotsClés)

        Me.TempsDePauseUnitaire = R.TempsDePauseUnitaire

        Me.PrixUnitaire = R.PrixUnitaire

    End Sub

#End Region

#Region "Propriétés"

    Public Property LibelléPrincipal() As String

    Public Property Libellés() As List(Of String)

    Public Property UsagesDeProduit() As List(Of UsageDeProduit_DAO)

    Public Property MotsClés() As List(Of String)

    Public Property TempsDePauseUnitaire() As Integer?

    Public Property PrixUnitaire() As Single?

#End Region

#Region "Méthodes"

    Public Overrides Function UnSerialized_Ex() As RéférenceDOuvrage
        Dim r As New RéférenceDOuvrage
        r.LibelléPrincipal = Me.LibelléPrincipal
        r.Libellés.AddRange(Me.Libellés)
        Dim UsagesDeProduit = From up In Me.UsagesDeProduit Select up.UnSerialized()
        r.UsagesDeProduit.AddRange(UsagesDeProduit)
        r.MotsClés.AddRange(Me.MotsClés)
        r.TempsDePauseUnitaire = TempsDePauseUnitaire
        r.PrixUnitaire = Me.PrixUnitaire
        Return r
    End Function

#End Region

#Region "Tests et debuggage"


#End Region

End Class
