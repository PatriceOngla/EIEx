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
        Me.Nom = "Espace de travail"
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

#Region "EtudeCourante"
    Private _EtudeCourante As Etude
    Public Property EtudeCourante As Etude
        Get
            If _EtudeCourante Is Nothing Then
                _EtudeCourante = (From e In Etudes Where e.EstOuverte).FirstOrDefault()
                If _EtudeCourante Is Nothing Then
                    _EtudeCourante = Me.Etudes.FirstOrDefault()
                    If _EtudeCourante IsNot Nothing Then _EtudeCourante.EstOuverte = True
                End If
            End If
            Return _EtudeCourante
        End Get
        Set(ByVal value As Etude)
            If value IsNot Nothing AndAlso Object.Equals(value, Me.EtudeCourante) Then Exit Property
            Try
                If Not (value Is Nothing OrElse Etudes.Contains(value)) Then Throw New InvalidOperationException("L'étude n'appartient pas à l'espace de travail.")
                _EtudeCourante = value
                NotifyPropertyChanged(NameOf(EtudeCourante))
            Catch ex As Exception
                Me.RaiseExceptionRaisedEvent(ex, True)
            End Try
        End Set
    End Property


#End Region

#Region "ClasseurExcelCourant"
    Private _ClasseurExcelCourant As ClasseurExcel
    Public Property ClasseurExcelCourant As ClasseurExcel
        Get
            Return _ClasseurExcelCourant
        End Get
        Set(ByVal value As ClasseurExcel)
            If Object.Equals(value, Me.ClasseurExcelCourant) Then Exit Property
            Try
                If value IsNot Nothing AndAlso Not EtudeCourante?.ClasseursExcel.Contains(value) Then Throw New InvalidOperationException("Le classeur Excel n'appartient pas à l'étude courante.")
                _ClasseurExcelCourant = value
                NotifyPropertyChanged(NameOf(ClasseurExcelCourant))
            Catch ex As Exception
                Me.RaiseExceptionRaisedEvent(ex, True)
            End Try
        End Set
    End Property


#End Region

#Region "BordereauCourant"
    Private _BordereauCourant As Bordereau
    Public Property BordereauCourant As Bordereau
        Get
            Return _BordereauCourant
        End Get
        Set(ByVal value As Bordereau)
            If Object.Equals(value, Me.BordereauCourant) Then Exit Property
            Try
                If value IsNot Nothing AndAlso Not ClasseurExcelCourant?.Bordereaux.Contains(value) Then Throw New InvalidOperationException("Le bordereau n'appartient pas au classeur courant.")
                _BordereauCourant = value
                NotifyPropertyChanged(NameOf(BordereauCourant))
            Catch ex As Exception
                Me.RaiseExceptionRaisedEvent(ex, True)
            End Try
        End Set
    End Property


#End Region

#End Region

#Region "Méthodes"

    '#Region "Persistance"

    '    Public Overrides Sub Charger(Chemin As String)
    '        If IO.File.Exists(Chemin) Then
    '            Dim WS_DAO As Workspace_DAO
    '            WS_DAO = Utils.DéSérialisation(Of Workspace_DAO)(Chemin)
    '            WS_DAO.UnSerialize(Me)
    '        End If
    '    End Sub

    '#End Region

#Region "Plomberie"

    Protected Overrides Function GetTable(Of Tr As IAgregateRoot)() As IList(Of Tr)
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
    Public Function GetEtudeById(id As Integer, Optional FailIfNotFound As Boolean = False) As Etude
        Dim r = GetObjectById(Of Etude)(id, FailIfNotFound)
        Return r
    End Function

#End Region

#End Region

#Region "Tests et debuggage"

    Private Sub _Etudes_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs) Handles _Etudes.CollectionChanged

    End Sub

#End Region

End Class
