
Imports System.ComponentModel
Imports Microsoft.Office.Interop.Excel
Imports Model

Public Class LibelléDouvrage
    Implements INotifyPropertyChanged

#Region "Constructeurs"

    Public Sub New(b As Bordereau, ByVal r As Excel.Range)
        Me.Bordereau = b
        Me._Ranges = New List(Of Excel.Range)({r})
    End Sub

#End Region

#Region "Propriétés"

#Region "Bordereau"
    Public ReadOnly Property Bordereau() As Bordereau
#End Region

#Region "Libellés"

#Region "LibelléSource"
    Public ReadOnly Property LibelléSource() As String
        Get
            Return Me.PremierRange?.Value
        End Get
    End Property
#End Region

#Region "Libellé"
    Private _Libellé As String

    Public Property Libellé() As String
        Get
            If String.IsNullOrEmpty(_Libellé) Then
                Return Me.LibelléSource
            Else
                Return _Libellé
            End If
        End Get
        Set(ByVal value As String)
            If Object.Equals(value, Me._Libellé) Then Exit Property
            _Libellé = value
            NotifyPropertyChanged(NameOf(Libellé))
        End Set
    End Property
#End Region

#Region "ComplémentDeNom"
    Private _ComplémentDeNom As String
    Public Property ComplémentDeNom() As String
        Get
            Return _ComplémentDeNom
        End Get
        Set(ByVal value As String)
            If Object.Equals(value, Me._ComplémentDeNom) Then Exit Property
            _ComplémentDeNom = value
            NotifyPropertyChanged(NameOf(ComplémentDeNom))
        End Set
    End Property
#End Region

#End Region

#Region "Ranges infos"

#Region "Feuille"
    Public ReadOnly Property Feuille() As Excel.Worksheet
        Get
            Return Me.PremierRange.Parent
        End Get
    End Property
#End Region

#Region "Ranges"

#Region "Ranges"
    Private _Ranges As List(Of Excel.Range)
    Public ReadOnly Property Ranges() As IEnumerable(Of Excel.Range)
        Get
            Return _Ranges.AsReadOnly
        End Get
    End Property
#End Region

#Region "PremierRange"
    Public ReadOnly Property PremierRange() As Excel.Range
        Get
            Return Me.Ranges.FirstOrDefault()
        End Get
    End Property
#End Region

#End Region

#Region "SourceRangeInfo"
    Public ReadOnly Property SourceRangeInfo() As String
        Get
            Dim r = $"{Me.Bordereau.Nom} - {Me.PremierRange.Address}"
            Return r
        End Get
    End Property
#End Region

#Region "SourceFileInfo"
    Public ReadOnly Property SourceFileInfo() As String
        Get
            Dim r = Me.Bordereau.Parent.Nom
            Return r
        End Get
    End Property
#End Region

#Region "SourceFilePathInfo"
    Public ReadOnly Property SourceFilePathInfo() As String
        Get
            Dim r = Me.Bordereau.Parent.CheminFichier
            Return r
        End Get
    End Property
#End Region

#End Region

#Region "NbOccurrences"
    Public ReadOnly Property NbOccurrences() As Integer
        Get
            Return Me.Ranges.Count
        End Get
    End Property
#End Region

#Region "Message"
    Private _Message As String
    Public Property Message() As String
        Get
            Return _Message
        End Get
        Set(ByVal value As String)
            If Object.Equals(value, Me._Message) Then Exit Property
            _Message = value
            NotifyPropertyChanged(NameOf(Message))
        End Set
    End Property
#End Region

#End Region

#Region "Méthodes"

    Public Sub AjouterRange(r As Excel.Range)
        Me._Ranges.Add(r)
    End Sub

    Friend Function EstAssociéA(range As Range) As Boolean
        Dim r = Ranges.Contains(range)
        Return r
    End Function

    Private Sub NotifyPropertyChanged(v As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(v))
    End Sub

    Public Overrides Function ToString() As String
        Try
            Dim r = $"{Me.Libellé} - Nb : {Me.NbOccurrences}"
            Return r
        Catch ex As Exception
            Utils.ManageError(ex, NameOf(ToString))
            Return MyBase.ToString()
        End Try
    End Function
#End Region

#Region "Events"

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

#End Region

#Region "Tests et debuggage"


#End Region

End Class
