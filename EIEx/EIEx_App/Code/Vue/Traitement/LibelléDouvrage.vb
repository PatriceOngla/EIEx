
Imports System.ComponentModel
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports Model

Public Class LibelléDouvrage
    Implements INotifyPropertyChanged

#Region "Constructeurs"

    'Public Sub New(Parent As GestionnaireDeBordereaux, b As Bordereau, ByVal premierRange As Excel.Range, statut As StatutDeLibellé)
    Public Sub New(b As Bordereau, ByVal premierRange As Excel.Range, statut As StatutDeLibellé)
        'Me.Parent = Parent
        Me.Statut = statut
        Me.Bordereau = b
        Me._Ranges = New List(Of Excel.Range)({premierRange})
        Me.Feuille = premierRange.Worksheet
        Me.LignePremierRange = premierRange.Row
        Me.LibelléSource = premierRange.Value
    End Sub

#End Region

#Region "Propriétés"

    '#Region "Parent"
    '    Public ReadOnly Property Parent() As GestionnaireDeBordereaux
    '#End Region

#Region "Bordereau"
    Public ReadOnly Property Bordereau() As Bordereau
#End Region

#Region "Libellés"

#Region "LibelléSource"
    Public ReadOnly Property LibelléSource() As String
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

#Region "LignePremierRange"
    Public ReadOnly Property LignePremierRange() As Integer
#End Region

#Region "Feuille"
    Public ReadOnly Property Feuille() As Excel.Worksheet
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

#Region "Range sélectionné"

#Region "SelectedRangeIndex"
    Private _SelectedRangeIndex As Integer
    Public Property SelectedRangeIndex() As Integer
        Get
            Return _SelectedRangeIndex
        End Get
        Private Set(ByVal value As Integer)
            If Object.Equals(value, Me._SelectedRangeIndex) Then Exit Property
            If value > _SelectedRangeIndex Then
                Me.IncrémenteSelectedRange(True)
            ElseIf value < _SelectedRangeIndex Then
                Me.IncrémenteSelectedRange(False)
            End If
        End Set
    End Property

    Private Sub IncrémenteSelectedRange(Avancer As Boolean)
        If Avancer Then
            _SelectedRangeIndex = ((Me.SelectedRangeIndex + 1) Mod Me.NbOccurrences)
        Else
            If SelectedRangeIndex = 0 Then
                _SelectedRangeIndex = Me.NbOccurrences - 1
            Else
                _SelectedRangeIndex = ((Me.SelectedRangeIndex - 1) Mod Me.NbOccurrences)
            End If
        End If
        NotifyPropertyChanged(NameOf(SelectedRangeIndex))
        NotifyPropertyChanged(NameOf(SelectedRange))
        Me.NotifyPropertyChanged(NameOf(SelectedRangeIndex_Base1))
        'Diagnostics.Debug.Print($"_SelectedRangeIndex : {_SelectedRangeIndex}")
    End Sub

#End Region

#Region "SelectedRangeIndex_Base1"
    Private _SelectedRangeIndex_Base1 As Integer
    Public ReadOnly Property SelectedRangeIndex_Base1() As Integer
        Get
            Return Me.SelectedRangeIndex + 1
        End Get
    End Property
#End Region
#Region "SelectedRange"

    Public ReadOnly Property SelectedRange() As Excel.Range
        Get
            Return Me.Ranges(Me.SelectedRangeIndex)
        End Get
    End Property

#End Region

#End Region

#End Region

#Region "SourceRangeInfo"
    Public ReadOnly Property SourceRangeInfo() As String
        Get
            Dim r = $"{Me.Bordereau.Nom} - {Me.SelectedRange.Address}"
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

#Region "EstSélectionnéPourQualification"
    Private _EstSélectionnéPourQualification As Boolean
    ''' <summary>
    ''' Indique que le <see cref="LibelléDouvrage"/> est sélectionné pour être qualifié en tant que vrai ou faux doublon (seulement dans la première liste pour traitement multiple). 
    ''' </summary>
    ''' <returns></returns>
    Public Property EstSélectionnéPourQualification() As Boolean
        Get
            If Me.Statut <> StatutDeLibellé.AQualifier Then
                Return False
            Else
                Return _EstSélectionnéPourQualification
            End If
        End Get
        Set(ByVal value As Boolean)
            ''System.Diagnostics.Debug.Print($"{Me.Libellé} : {value}")
            'SetEstSélectionnéPourQualification(value, True)
            If Object.Equals(value, Me._EstSélectionnéPourQualification) Then Exit Property
            If (value AndAlso Me.Statut <> StatutDeLibellé.AQualifier) Then Throw New InvalidOperationException($"Un {NameOf(LibelléDouvrage)} ne peut être sélectionné que s'il a le statut ""{StatutDeLibellé.AQualifier}"".")
            _EstSélectionnéPourQualification = value
            NotifyPropertyChanged(NameOf(EstSélectionnéPourQualification))
        End Set
    End Property

    'Friend Sub SetEstSélectionnéPourQualification(value As Boolean, NotifierParent As Boolean)
    '    If Object.Equals(value, Me._EstSélectionnéPourQualification) Then Exit Sub
    '    If (value AndAlso Me.Statut <> StatutDeLibellé.AQualifier) Then Throw New InvalidOperationException($"Un {NameOf(LibelléDouvrage)} ne peut être sélectionné que s'il a le statut ""{StatutDeLibellé.AQualifier}"".")
    '    _EstSélectionnéPourQualification = value
    '    NotifyPropertyChanged(NameOf(EstSélectionnéPourQualification))
    '    'If NotifierParent Then Me.Parent.GérerSélectionChanges(Me)
    'End Sub

#End Region

#Region "Statut"

    Private _Statut As StatutDeLibellé = StatutDeLibellé.AQualifier
    Public Property Statut() As StatutDeLibellé
        Get
            Return _Statut
        End Get
        Set(ByVal value As StatutDeLibellé)
            If Object.Equals(value, Me._Statut) Then Exit Property
            'If value <> StatutDeLibellé.AQualifier Then Me.SetEstSélectionnéPourQualification(False, False)
            If value <> StatutDeLibellé.AQualifier Then Me.EstSélectionnéPourQualification = False
            _Statut = value
            NotifyPropertyChanged(NameOf(Statut))
        End Set
    End Property

    Public Enum StatutDeLibellé
        ''' <summary>Le libellé est dans la première liste et doit être qualifié en tant que vrai ou faux doublon.</summary>
        AQualifier

        ''' <summary>Le libellé est un faux doublon, il est donc dans la deuxième liste et doit être complété (complément de nom).</summary>
        ACompléter

        ''' <summary>Le libellé est traité et prêt pour génération d'un <see cref="Ouvrage"/>.</summary>
        Traité
    End Enum
#End Region
#End Region

#Region "Méthodes"

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

#Region "Gestion des ranges"

    Public Sub AjouterRange(r As Excel.Range)
        Me._Ranges.Add(r)
    End Sub

    Public Sub SelectNextRange()
        Me.IncrémenteSelectedRange(True)
    End Sub

    Public Sub SelectPreviousRange()
        Me.IncrémenteSelectedRange(False)
    End Sub

#End Region

#End Region

#Region "Events"

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

#End Region

#Region "Tests et debuggage"


#End Region

End Class
