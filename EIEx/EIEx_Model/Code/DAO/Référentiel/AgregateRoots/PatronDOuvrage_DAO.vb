Imports System.Xml.Serialization
Imports Utils

<Serializable>
Public Class PatronDOuvrage_DAO
    Inherits AgregateRoot_DAO(Of PatronDOuvrage)

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(R As PatronDOuvrage)
        MyBase.New(R)

        Me.ComplémentDeNom = R.ComplémentDeNom

        Me.Libellés = New List(Of String)(R.Libellés)

        Dim UsagesDeProduit_DAO = From up In R.UsagesDeProduit Select New UsageDeProduit_DAO(up)
        Me.UsagesDeProduit = New List(Of UsageDeProduit_DAO)(UsagesDeProduit_DAO)

        Me.MotsClés = New List(Of String)(R.MotsClés)

        If R.TempsDePauseForcé Then Me.TempsDePauseUnitaire = R.TempsDePauseUnitaire

        If R.PrixUnitaireForcé Then Me.PrixUnitaire = R.PrixUnitaire

    End Sub

#End Region

#Region "Propriétés"

#Region "Sys"
    Private Ref As Référentiel = Référentiel.Instance
    <XmlIgnore>
    Protected Overrides ReadOnly Property Sys As Système
        Get
            Return Ref
        End Get
    End Property
#End Region

#Region "Données"

    Public Property Libellés() As List(Of String)

    Public Property ComplémentDeNom() As String

    Public Property UsagesDeProduit() As List(Of UsageDeProduit_DAO)

    Public Property MotsClés() As List(Of String)

    Public Property TempsDePauseUnitaire() As Integer?

    Public Property PrixUnitaire() As Single?

#End Region

#End Region

#Region "Méthodes"

    Protected Overrides Function UnSerialized_Ex_Ex() As PatronDOuvrage
        Dim r = Ref.GetNewPatronDOuvrage(Me.Id)
        r.ComplémentDeNom = Me.ComplémentDeNom
        r = If(r, New PatronDOuvrage(Me.Id))
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
