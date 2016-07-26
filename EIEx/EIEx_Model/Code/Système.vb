Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.Runtime.Serialization
Imports Utils

'Public MustInherit Class Système(Of T  As {AgregateRoot(Of T2)}, T2 As Entité)
Public MustInherit Class Système '(Of Ts As Système)
    Inherits EIExObject

#Region "Constructeurs"
    Protected Overrides Sub Init()
        _Tables = New List(Of IList)()
        AddHandler ExceptionRaised, Sub(e As Exception, S As Système, Attendue As Boolean) Système_ExceptionRaised(e, S, Attendue)
    End Sub

#End Region

#Region "Propriétés"

#Region "DateModif"
    Private _DateModif As Date
    Public Property DateModif() As Date
        Get
            Return _DateModif
        End Get
        Set(ByVal value As Date)
            If Object.Equals(value, Me._DateModif) Then Exit Property
            _DateModif = value
            NotifyPropertyChanged(NameOf(DateModif))
        End Set
    End Property
#End Region

#Region "Tables"

    Protected _Tables As List(Of IList)
    Protected ReadOnly Property Tables As IEnumerable(Of IList)
        '--> Impossible de mieux typer à cause de l'éternelle absence de covariance (des erreurs de types à l'exécution pour éviter... des erreurs de type à l'exécution... qui n'arriveraient quasiment jamais). 
        Get
            Return _Tables
        End Get
    End Property

#End Region

    'TODO : renommer en EntityCollections
    'Protected MustOverride ReadOnly Property Tables As IEnumerable(Of IList) '(Of AgregateRoot_Base)) --> Impossible de mieux typer à cause de l'éternelle absence de covariance (des erreurs de types à l'exécution pour éviter... des erreurs de type à l'exécution... qui n'arriveraient quasiment jamais). 

    'Public MustOverride ReadOnly Property System_DAO As ISystèmeDAO
#End Region

#Region "Méthodes"

#Region "Sérialisation"

    '''' <summary>Peuple le référentiel à partir du fichier de persistance <see cref="EIExData.CheminRéférentiel"/>.</summary>
    'Public MustOverride Sub Charger(Chemin As String)

    'Public Sub Enregistrer(Chemin As String)
    '    Me.DateModif = Now()

    'End Sub

#End Region

#Region "Factory"

#Region "Factory générique"
    Friend Sub EnregistrerRoot(Of Tr As AgregateRoot_Base)(root As Object) 'pas réussi à mieux typer. 
        Dim Table = GetTable(Of Tr)()
        CheckUnicity(Table, root.id)
        If Not Table.Contains(root) Then Table.Add(root)
    End Sub

    Private Sub CheckUnicity(Of Tr As AgregateRoot_Base)(Table As IEnumerable(Of Tr), Id As Integer)
        Dim Doublon = (Table.Where(Function(item) item.Id = Id)).Any()
        If Doublon Then Throw New Exception($"Un élément ""{GetType(Tr).Name}"" portant l'identifiant{Id} existe déjà.")
    End Sub

    'Public Function GetNewRoot(Of T As {AgregateRoot, New})(newId As Integer) As T
    '    Dim r = New T()
    '    EnregistrerRoot(r)
    '    Return r
    'End Function

    Public Function GetNewRoot(Of Tr As {AgregateRoot_Base, New})() As Tr
        Dim r = New Tr()
        EnregistrerRoot(Of Tr)(r)
        Return r
    End Function

#End Region

#End Region

#Region "Accès aux objets"

    Protected Function GetObjectById(Of Tr As AgregateRoot_Base)(id As Integer, FailIfNotFound As Boolean) As Tr
        Dim Table As ObservableCollection(Of Tr) = GetTable(Of Tr)()
        Dim r = (Table.Where(Function(o) o.Id = id)).FirstOrDefault
        If r Is Nothing AndAlso FailIfNotFound Then
            Dim Msg = $"L'objet ""{GetType(Tr).Name}"" d'id {id} n'existe pas."
            Throw New InvalidOperationException(Msg)
        End If
        Return r
    End Function

#End Region

#Region "Divers"

    Protected MustOverride Function GetTable(Of Tr As AgregateRoot_Base)() As IList(Of Tr)

#Region "GetNewId"
    Friend Function GetNewId(Of Tr As AgregateRoot_Base)() As Integer?
        Dim Table As ObservableCollection(Of Tr) = GetTable(Of Tr)()
        Dim newId = (From p In Table Select p.Id).Max
        newId = If(newId, 0)
        newId += 1
        Return newId
    End Function
#End Region

#Region "Purger"

    Public Sub Purger()
        Me.Tables.DoForAll(Sub(Tb As IList) Tb.Clear())
    End Sub

#End Region

#Region "EstVide"
    Public Function EstVide() As Boolean
        Dim r = Me.Tables.TrueForAll(Function(Tb As IList) Tb.Count() = 0)
        Return r
    End Function
#End Region

#Region "Système_ExceptionRaised"
    Private Sub Système_ExceptionRaised(e As Exception, S As Système, Attendue As Boolean)

    End Sub
#End Region

#End Region

#End Region

#Region "Evénnements"

    ''' <summary>Contournement de l'absence de centralisation des traitements d'exception en VSTO.</summary>
    ''' <param name="Attendue">Indique s'il s'agit d'une exception métier ou technique.</param>
    Public Shared Event ExceptionRaised(e As Exception, S As Système, Attendue As Boolean)

    Friend Sub RaiseExceptionRaisedEvent(e As Exception, Attendue As Boolean)
        RaiseEvent ExceptionRaised(e, Me, Attendue)
    End Sub

#End Region

#Region "Tests et debuggage"

#End Region

End Class
