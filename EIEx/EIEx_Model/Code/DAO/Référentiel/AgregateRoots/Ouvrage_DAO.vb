Imports System.Xml.Serialization
Imports Utils

<Serializable>
Public Class Ouvrage_DAO
    Inherits AgregateRoot_DAO(Of Ouvrage)

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(O As Ouvrage)
        MyBase.New(O)

        Me.ComplémentDeNom = O.ComplémentDeNom

        Me.Libellés = New List(Of String)(O.Libellés)

        Me.EstModèle = O.EstModèle

        Dim UsagesDeProduit_DAO = From up In O.UsagesDeProduit Select New UsageDeProduit_DAO(up)
        Me.UsagesDeProduit = New List(Of UsageDeProduit_DAO)(UsagesDeProduit_DAO)

        Me.MotsClés = New List(Of String)(O.MotsClés)

        If O.TempsDePauseForcé Then Me.TempsDePauseUnitaire = O.TempsDePauseUnitaire

        If O.PrixUnitaireForcé Then Me.PrixUnitaire = O.PrixUnitaire

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

    Public Property EstModèle As Boolean

    Public Property UsagesDeProduit() As List(Of UsageDeProduit_DAO)

    Public Property MotsClés() As List(Of String)

    Public Property TempsDePauseUnitaire() As Integer?

    Public Property PrixUnitaire() As Single?

#End Region

#End Region

#Region "Méthodes"

    Protected Overrides Function UnSerialized_Ex_Ex() As Ouvrage
        Dim r = Ref.GetNewOuvrage(Me.Id)
        r.ComplémentDeNom = Me.ComplémentDeNom
        r.EstModèle = Me.EstModèle
        r = If(r, New Ouvrage(Me.Id))
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
