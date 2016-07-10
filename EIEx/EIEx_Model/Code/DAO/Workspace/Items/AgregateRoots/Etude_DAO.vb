Imports System.Xml.Serialization
Imports Model
Imports Utils

<Serializable>
Public Class Etude_DAO
    Inherits AgregateRoot_DAO(Of Etude)

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(E As Etude)
        MyBase.New(E)

        Me.EstOuverte = E.EstOuverte

        Dim SBordereaux = From b In E.Bordereaux Select New Bordereau_DAO(b)
        Me.Bordereaux = New List(Of Bordereau_DAO)(SBordereaux)

    End Sub

#End Region

#Region "Propriétés"

#Region "Sys"
    Private WS As WorkSpace = WorkSpace.Instance
    <XmlIgnore>
    Protected Overrides ReadOnly Property Sys As Système
        Get
            Return WS
        End Get
    End Property
#End Region

#Region "Données"

    <XmlAttribute>
    Public Property EstOuverte() As Boolean

    Public Property Bordereaux() As List(Of Bordereau_DAO)

#End Region

#End Region

#Region "Méthodes"

    Protected Overrides Function UnSerialized_Ex_Ex() As Etude
        Dim r = WS.GetNewEtude(Me.Id)

        r.EstOuverte = Me.EstOuverte

        Dim Bdx = (From b In Me.Bordereaux Select b.UnSerialized())
        r.Bordereaux.AddRange(Bdx)
        Return r
    End Function

#End Region

#Region "Tests et debuggage"


#End Region

End Class