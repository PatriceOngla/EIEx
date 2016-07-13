Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports Microsoft.Office.Interop.Excel
Imports Model
Imports Utils

Public Class GestionnaireDeBordereau
    Implements INotifyPropertyChanged

#Region "Constructeurs"

    Public Sub New()
        _LibellésEnDoublonEncoreATraiter = New ObservableCollection(Of LibelléDouvrage)
        _LibellésRetenus = New ObservableCollection(Of LibelléDouvrage)
        _TousLesLibellés = New List(Of Excel.Range)
    End Sub

#End Region

#Region "Propriétés"

#Region "WS"
    Public ReadOnly Property WS() As WorkSpace
        Get
            Return WorkSpace.Instance
        End Get
    End Property
#End Region

#Region "XL"
    Public ReadOnly Property XL() As Excel.Application
        Get
            Return ExcelCommander.XL
        End Get
    End Property

#End Region

#Region "TousLesLibellés"
    Private _TousLesLibellés As List(Of Excel.Range)
    Public ReadOnly Property TousLesLibellés() As IEnumerable(Of Excel.Range)
        Get
            Return _TousLesLibellés?.AsReadOnly
        End Get
    End Property
#End Region

#Region "Libellés en doublon encore à traiter"

#Region "LibellésEnDoublonEncoreATraiter"
    Private WithEvents _LibellésEnDoublonEncoreATraiter As ObservableCollection(Of LibelléDouvrage)
    Public ReadOnly Property LibellésEnDoublonEncoreATraiter() As ObservableCollection(Of LibelléDouvrage)
        Get
            Return _LibellésEnDoublonEncoreATraiter
        End Get
    End Property
#End Region

#Region "LibelléEnDoublonCourant"
    Private _LibelléEnDoublonCourant As LibelléDouvrage
    Public Property LibelléEnDoublonCourant() As LibelléDouvrage
        Get
            Return _LibelléEnDoublonCourant
        End Get
        Set(ByVal value As LibelléDouvrage)
            If Object.Equals(value, Me._LibelléEnDoublonCourant) Then Exit Property
            _LibelléEnDoublonCourant = value
            NotifyPropertyChanged(NameOf(LibelléEnDoublonCourant))
        End Set
    End Property
#End Region

#Region "NbLibellésEnDoublonEncoreATraiter"
    Public ReadOnly Property NbLibellésEnDoublonEncoreATraiter() As Integer
        Get
            Return LibellésEnDoublonEncoreATraiter.Count()
        End Get
    End Property
#End Region

    Private Sub _LibellésEnDoublonEncoreATraiter_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs) Handles _LibellésEnDoublonEncoreATraiter.CollectionChanged
        Me.NotifyPropertyChanged(NameOf(NbLibellésEnDoublonEncoreATraiter))
    End Sub

#End Region

#Region "Libellés retenus"

#Region "LibellésRetenus"
    Private WithEvents _LibellésRetenus As ObservableCollection(Of LibelléDouvrage)
    Public ReadOnly Property LibellésRetenus() As ObservableCollection(Of LibelléDouvrage)
        Get
            Return _LibellésRetenus
        End Get
    End Property
#End Region

#Region "LibelléRetenuCourant"
    Private _LibelléRetenuCourant As LibelléDouvrage
    Public Property LibelléRetenuCourant() As LibelléDouvrage
        Get
            Return _LibelléRetenuCourant
        End Get
        Set(ByVal value As LibelléDouvrage)
            If Object.Equals(value, Me._LibelléRetenuCourant) Then Exit Property
            _LibelléRetenuCourant = value
            NotifyPropertyChanged(NameOf(LibelléRetenuCourant))
        End Set
    End Property
#End Region

#Region "NbLibellésRetenus"
    Public ReadOnly Property NbLibellésRetenus() As Integer
        Get
            Return LibellésRetenus.Count()
        End Get
    End Property
#End Region

    Private Sub _LibellésRetenus_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs) Handles _LibellésRetenus.CollectionChanged
        Me.NotifyPropertyChanged(NameOf(NbLibellésRetenus))
    End Sub

#End Region

#Region "AvancementInfo"
    Private NbOpérationsATraiter As Integer?
    Private NbOpérationsTraitées As Integer?
    Private NbClasseursATraiter As Integer?
    Private NbClasseursTraités As Integer?
    Private NbFeuillesATraiter As Integer?
    Private NbFeuillesTraitées As Integer?
    Private NbcellsATraiter As Integer?
    Private NbCellsTraitées As Integer?

    Public ReadOnly Property AvancementInfo() As String
        Get
            Dim PourcentageCells = Format(NbCellsTraitées / NbcellsATraiter, "0 %")
            Dim r = $"Opération {If(NbOpérationsTraitées, "?")}/{If(NbOpérationsATraiter, "?")}, Classeur {If(NbClasseursTraités, "?")}/{If(NbClasseursATraiter, "?")}, feuille  {If(NbFeuillesTraitées, "?")}/{If(NbFeuillesATraiter, "?")}, cellule {If(Format(NbCellsTraitées, "# ### 000"), "?")}/{If(Format(NbcellsATraiter, "# ### 000"), "?")} soit {PourcentageCells}"
            Return r
        End Get
    End Property

    Private Sub IncrémenteAvancementOpération(Optional NbATraiter As Integer? = Nothing)
        If NbATraiter IsNot Nothing Then
            NbOpérationsATraiter = NbATraiter
            NbOpérationsTraitées = 0
        Else
            NbOpérationsTraitées += 1
        End If
        NbClasseursATraiter = Nothing
        NbClasseursTraités = Nothing
        NbFeuillesATraiter = Nothing
        NbFeuillesTraitées = Nothing
        NbcellsATraiter = Nothing
        NbCellsTraitées = Nothing
        NotifyPropertyChanged(NameOf(AvancementInfo))
        XL.StatusBar = AvancementInfo
    End Sub

    Private Sub IncrémenteAvancementClasseur(Optional NbATraiter As Integer? = Nothing)
        If NbATraiter IsNot Nothing Then
            NbClasseursATraiter = NbATraiter
            NbClasseursTraités = 0
        Else
            NbClasseursTraités += 1
        End If
        NbFeuillesATraiter = Nothing
        NbFeuillesTraitées = Nothing
        NbcellsATraiter = Nothing
        NbCellsTraitées = Nothing
        NotifyPropertyChanged(NameOf(AvancementInfo))
        XL.StatusBar = AvancementInfo
    End Sub

    Private Sub IncrémenteAvancementFeuille(Optional NbATraiter As Integer? = Nothing)
        If NbATraiter IsNot Nothing Then
            NbFeuillesATraiter = NbATraiter
            NbFeuillesTraitées = 0
        Else
            NbFeuillesTraitées += 1
        End If
        NbcellsATraiter = Nothing
        NbCellsTraitées = Nothing
        NotifyPropertyChanged(NameOf(AvancementInfo))
        XL.StatusBar = AvancementInfo
    End Sub

    Private Sub IncrémenteAvancementCellule(Optional NbATraiter As Integer? = Nothing)
        If NbATraiter IsNot Nothing Then
            NbcellsATraiter = NbATraiter
            NbCellsTraitées = 0
        Else
            NbCellsTraitées += 1
        End If
        NotifyPropertyChanged(NameOf(AvancementInfo))

        Dim Pas = If(NbcellsATraiter > 100000, 10000, If(NbcellsATraiter > 10000, 1000, 10))
        If (NbCellsTraitées Mod Pas = 0) Then XL.StatusBar = AvancementInfo
    End Sub

#End Region

#End Region

#Region "Méthodes"

#Region "RécupérerLesLibellésDOuvrages"

    Public Sub Purger()
        Me._TousLesLibellés.Clear()
        Me.LibellésEnDoublonEncoreATraiter.Clear()
        Me.LibellésRetenus.Clear()
    End Sub

    Public Sub RécupérerLesLibellésDOuvrages()
        Try
            IncrémenteAvancementOpération(2)
            RécupérerTousLesLibellésDOuvrages()
            RépartirLesLibellésDOuvrages()
        Catch ex As Exception
            ManageErreur(ex, , True, False)
        End Try
    End Sub

    Public Sub RécupérerTousLesLibellésDOuvrages()
        Dim Ec = WS.EtudeCourante
        Dim OffsetChampUnité As Short

        IncrémenteAvancementClasseur(Ec.ClasseursExcel.Count)

        For Each c In Ec.ClasseursExcel
            IncrémenteAvancementFeuille(c.Bordereaux.Count)
            For Each b In c.Bordereaux
                OffsetChampUnité = GetOffsetChampsUnité(b)
                RécupérerTousLesLibellésDOuvrages(b, OffsetChampUnité)
                IncrémenteAvancementFeuille()
            Next
            IncrémenteAvancementClasseur()
        Next

        IncrémenteAvancementOpération()

    End Sub

    Private Sub RécupérerTousLesLibellésDOuvrages(b As Bordereau, OffsetChampsUnité As Short)
        Dim ws As Worksheet = b.Worksheet
        Try
            Dim AdresseLibellés = b.Paramètres.AdresseRangeLibelleOuvrage
            Dim RangeLibellés As Excel.Range = ws.Range(AdresseLibellés)

            Dim NbcellsATraiter As Integer
            Integer.TryParse(RangeLibellés.Cells.CountLarge(), NbcellsATraiter)
            IncrémenteAvancementCellule(NbcellsATraiter)
            For Each cell As Excel.Range In RangeLibellés.Cells
                If EstUneCelluleDeLibellé(cell, OffsetChampsUnité) Then
                    Me._TousLesLibellés.Add(cell)
                End If
                IncrémenteAvancementCellule()
            Next
        Catch ex As Exception
            Throw New Exception($"Echec de la récupération des libellés sur la feuille {ws?.Name}.{vbCr}
Vérifier que l'adresse de plage définie par le bordereau correspondant est correcte.", ex)
        End Try
    End Sub

    Private Function GetOffsetChampsUnité(b As Bordereau) As Short
        Dim ws As Worksheet = b.Worksheet
        Dim RLibellés As Excel.Range = ws.Range(b.Paramètres.AdresseRangeLibelleOuvrage)
        Dim RUnités As Excel.Range = ws.Range(b.Paramètres.AdresseRangeUnité)
        Dim r = RUnités.Column - RLibellés.Column
        Return r
    End Function

    Private Function EstUneCelluleDeLibellé(cell As Range, OffsetChampsUnité As Short) As Boolean
        Dim CellUnité As Excel.Range = cell.Offset(0, OffsetChampsUnité)
        Dim UnitéOK = Not (String.IsNullOrEmpty(CellUnité.Value) OrElse CStr(CellUnité.Value).ToUpper = "PM")
        Dim EstPasVide = Not String.IsNullOrEmpty(cell.Value)
        Dim r = UnitéOK And EstPasVide
        Return r
    End Function

#Region "RépartirLesLibellésDOuvrages"

    ''' <summary>Selon qu'ils sont doublons ou pas.</summary>
    Public Sub RépartirLesLibellésDOuvrages()
        Dim Ec = WS.EtudeCourante
        For Each c In Ec.ClasseursExcel
            For Each b In c.Bordereaux
                RépartirLesLibellésDOuvrages(b)
            Next
        Next

    End Sub

    Public Sub RépartirLesLibellésDOuvrages(b As Bordereau)
        Try
            Dim Ls As IEnumerable(Of LibelléDouvrage) = From c In TousLesLibellés Select ToLibellé(b, c)
            Me.LibellésEnDoublonEncoreATraiter.AddRange(Ls)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function ToLibellé(b As Bordereau, rng As Range) As LibelléDouvrage
        Dim r As New LibelléDouvrage(b, rng, 10)
        Return r
    End Function

#End Region

#End Region

#Region "Divers"

    Private Sub NotifyPropertyChanged(v As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(v))
    End Sub

#End Region

#End Region

#Region "Events"

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

#End Region

#Region "Tests et debuggage"


#End Region

End Class
