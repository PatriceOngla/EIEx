Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.Xml.Serialization
Imports Model

''' <summary>Singleton.</summary>
Public Class WorkSpace
    Inherits Système

#Region "Constructeurs"

    Private Sub New()
        MyBase.New()
    End Sub

    Protected Overrides Sub Init()
        MyBase.Init()
        _Etudes = New ObservableCollection(Of Etude)()
        _Tables.Add(_Etudes)
    End Sub

#End Region

#Region "Propriétés"

#Region "Instance (WorkSpace)"
    Private Shared _Instance As WorkSpace
    Public Shared ReadOnly Property Instance() As WorkSpace
        Get
            If _Instance Is Nothing Then _Instance = New WorkSpace()
            Return _Instance
        End Get
    End Property
#End Region

#Region "Etudes"
    Private WithEvents _Etudes As ObservableCollection(Of Etude)

    ''' <summary>Toutes les <see cref="Etudes"/>s du <see cref="WorkSpace"/>.</summary>
    Public ReadOnly Property Etudes() As ObservableCollection(Of Etude)
        Get
            Return _Etudes
        End Get
    End Property
#End Region

#End Region

#Region "Méthodes"

#Region "Persistance"

    Public Overrides Sub Charger(Chemin As String)
        Dim WS_DAO = Utils.DéSérialisation(Of Workspace_DAO)(Chemin)
        WS_DAO.UnSerialize(Me)
    End Sub

#End Region

#Region "Plomberie"

    Protected Overrides Function GetTable(Of Tr As AgregateRoot_Base)() As IList(Of Tr)
        Return Me._Etudes
    End Function

#End Region

#Region "Factory"

#Region "Etude"

    Public Function GetNewEtude(newId As Integer) As Etude
        Dim r = New Etude(newId)
        Me.Etudes.Add(r)
        Return r
    End Function

    Public Function GetNewEtude() As Etude
        Dim newId = GetNewId(Of Etude)()
        Dim r = GetNewEtude(newId)
        Return r
    End Function

#End Region

#End Region

#Region "Accès aux données"
    Public Function GetEtudeById(id As Integer) As Etude
        Dim r = GetObjectById(Of Etude)(id)
        Return r
    End Function

#End Region

#End Region

#Region "Tests et debuggage"

    Private Sub _Etudes_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs) Handles _Etudes.CollectionChanged

    End Sub

#End Region

End Class
