
Imports System.ComponentModel
Imports Model

Public Class LibelléDouvrage
    Implements INotifyPropertyChanged

#Region "Constructeurs"

    Public Sub New(b As Bordereau, ByVal r As Excel.Range, nbOccurrences As Integer)
        Me.Bordereau = b
        Me.Range = r
        Me.NbOccurrences = nbOccurrences
    End Sub

#End Region

#Region "Propriétés"

#Region "Bordereau"
    Public ReadOnly Property Bordereau() As Bordereau
#End Region

#Region "LibelléSource"
    Public ReadOnly Property LibelléSource() As String
        Get
            Return Me.Range.Value
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

#Region "Range"
    Public ReadOnly Property Range() As Excel.Range
#End Region

#Region "NbOccurrences"
    Public ReadOnly Property NbOccurrences() As Integer
#End Region

#Region "SourceRangeInfo"
    Public ReadOnly Property SourceRangeInfo() As String
        Get
            Dim r = $"{Me.Bordereau.Nom} - {Me.Range.Address}"
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

#Region "Méthodes"

    Private Sub NotifyPropertyChanged(v As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(v))
    End Sub

#End Region

#Region "Events"

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

#End Region

#Region "Tests et debuggage"


#End Region

End Class
