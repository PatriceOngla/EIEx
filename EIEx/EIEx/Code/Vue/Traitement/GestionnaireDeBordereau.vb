Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports EIEx
Imports Microsoft.Office.Interop.Excel
Imports Model
Imports Utils
Imports MoreLinq

Public Class GestionnaireDeBordereaux
    Implements INotifyPropertyChanged

#Region "Constructeurs"

    Public Sub New()
        _TousLesRangesDeLibellés = New List(Of Excel.Range)
        _TousLesLibellés = New LibelléDouvrageCollection
        _LibellésEnDoublonEncoreATraiter = New ObservableCollection(Of LibelléDouvrage)
        _LibellésEnTransit = New ObservableCollection(Of LibelléDouvrage)
        _LibellésRetenus = New ObservableCollection(Of LibelléDouvrage)
    End Sub

#End Region

#Region "Types"

    Private Class LibelléDouvrageCollection
        Inherits KeyedCollection(Of String, LibelléDouvrage)

        Protected Overrides Function GetKeyForItem(item As LibelléDouvrage) As String
            Return item.Libellé
        End Function

        Public Function AsReadOnly() As IEnumerable(Of LibelléDouvrage)
            Return New ReadOnlyCollection(Of LibelléDouvrage)(Me)
        End Function
    End Class

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

#Region "Ref"
    Public ReadOnly Property Ref() As Référentiel
        Get
            Return Référentiel.Instance
        End Get
    End Property
#End Region
#Region "Plages sources"

#Region "PlageDeRechercheDesLibellés"
    'Private _PlageDeRechercheDesLibellés As Excel.Range

    ''' <summary>La plage de recherche des libellés dans les bordereaux sources. </summary>
    Public ReadOnly Property PlageDeRechercheDesLibellés(b As Bordereau) As Excel.Range
        Get
            'TODO: Bordereux multiples
            'If _PlageDeRechercheDesLibellés Is Nothing then 
            Dim r As Excel.Range
            Dim ws As Worksheet = b.Worksheet
            Dim AdresseLibellés = b.Paramètres.AdresseRangeLibelleOuvrage
            r = ws.Range(AdresseLibellés)
            r = r.LimitedRange()

            'End If
            'Return _PlageDeRechercheDesLibellés
            Return r
        End Get
    End Property
#End Region

#End Region

#Region "Collections de libellés"

#Region "TousLesRangesDeLibellés"
    Private _TousLesRangesDeLibellés As List(Of Excel.Range)
    Public ReadOnly Property TousLesRangesDeLibellés() As List(Of Excel.Range)
        Get
            Return _TousLesRangesDeLibellés
        End Get
    End Property
#End Region

#Region "TousLesLibellés"
    Private _TousLesLibellés As LibelléDouvrageCollection
    Public ReadOnly Property TousLesLibellés() As IEnumerable(Of LibelléDouvrage)
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
            SélectionnerLesRangesAssociés(value)
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

#Region "LibellésEnTransit"

#Region "LibellésEnTransit"
    Private _LibellésEnTransit As ObservableCollection(Of LibelléDouvrage)
    Public ReadOnly Property LibellésEnTransit() As ObservableCollection(Of LibelléDouvrage)
        Get
            Return _LibellésEnTransit
        End Get
    End Property
#End Region

#Region "LibelléEnTransitCourant"
    Private _LibelléEnTransitCourant As LibelléDouvrage
    Public Property LibelléEnTransitCourant() As LibelléDouvrage
        Get
            Return _LibelléEnTransitCourant
        End Get
        Set(ByVal value As LibelléDouvrage)
            If Object.Equals(value, Me._LibelléEnTransitCourant) Then Exit Property
            _LibelléEnTransitCourant = value
            NotifyPropertyChanged(NameOf(LibelléEnTransitCourant))
            SélectionnerLesRangesAssociés(value)
        End Set
    End Property
#End Region

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
            SélectionnerLesRangesAssociés(value)
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

#End Region

#Region "Décomptes"

#Region "NbLignesLibelléDétéctées"
    Public ReadOnly Property NbLignesLibelléDétéctées() As Integer
        Get
            Return Me.TousLesRangesDeLibellés.Count
        End Get
    End Property
#End Region

#Region "NbLibellésUniques"
    Public ReadOnly Property NbLibellésUniques() As Integer
        Get
            Dim r = Me.TousLesLibellés.Count
            Dim LesUniques = (From rng In Me.TousLesRangesDeLibellés).DistinctBy(Of String)(Function(rng2 As Range)
                                                                                                Return rng2.Value
                                                                                            End Function)

            If LesUniques.Count <> r Then MsgBox("On a un problème.")
            Return r
        End Get
    End Property
#End Region

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
        _TousLesRangesDeLibellés.Clear()
        _TousLesLibellés.Clear()
        _LibellésEnDoublonEncoreATraiter.Clear()
        _LibellésEnTransit.Clear()
        _LibellésRetenus.Clear()
    End Sub

    Public Sub RécupérerLesLibellésDOuvrages()
        Try
            IncrémenteAvancementOpération(2)

            RécupérerTousLesLibellésDOuvrages()
            IncrémenteAvancementOpération()

            RépartirLesLibellésDOuvrages()
            IncrémenteAvancementOpération()

            Me.NotifyPropertyChanged(NameOf(NbLignesLibelléDétéctées))
            Me.NotifyPropertyChanged(NameOf(NbLibellésUniques))

        Catch ex As Exception
            ManageErreur(ex, , True, False)
        End Try
    End Sub

#Region "RécupérerTousLesLibellésDOuvrages"

    Public Sub RécupérerTousLesLibellésDOuvrages()
        Dim Ec = WS.EtudeCourante
        Dim OffsetChampUnité As Short

        Me.Purger()

        IncrémenteAvancementClasseur(Ec.ClasseursExcel.Count)

        For Each c In Ec.ClasseursExcel
            IncrémenteAvancementFeuille(c.Bordereaux.Count)
            For Each b In c.Bordereaux
                CheckPrérequis(b)
                OffsetChampUnité = GetOffsetChampsUnité(b)
                RécupérerTousLesLibellésDOuvrages(b, OffsetChampUnité)
                IncrémenteAvancementFeuille()
            Next
            IncrémenteAvancementClasseur()
        Next

        IncrémenteAvancementOpération()

    End Sub

    Private Sub CheckPrérequis(b As Bordereau)

        Dim IncomplétudeParamètres = String.IsNullOrEmpty(b.Paramètres?.AdresseRangeLibelleOuvrage) OrElse
        String.IsNullOrEmpty(b.Paramètres?.AdresseRangePrixUnitaire) OrElse String.IsNullOrEmpty(b.Paramètres?.AdresseRangeUnité)

        If IncomplétudeParamètres Then Throw New Exception("Les paramètres sont incomplets (champs adresse).")
    End Sub

    Private Sub RécupérerTousLesLibellésDOuvrages(b As Bordereau, OffsetChampsUnité As Short)
        Dim ws As Worksheet = b.Worksheet
        Try
            Dim LibelléCandidatALajout As String
            Dim Libellé As LibelléDouvrage
            Dim NbcellsATraiter As Integer

            Dim PlageRechercheDesLibellés = PlageDeRechercheDesLibellés(b)
            Integer.TryParse(PlageRechercheDesLibellés.Cells.CountLarge(), NbcellsATraiter)
            IncrémenteAvancementCellule(NbcellsATraiter)

            For Each cell As Excel.Range In PlageRechercheDesLibellés.Cells
                If EstUneCelluleDeLibellé(cell, OffsetChampsUnité) Then
                    _TousLesRangesDeLibellés.Add(cell)
                    LibelléCandidatALajout = cell.Value
                    If _TousLesLibellés.Contains(LibelléCandidatALajout) Then
                        Libellé = _TousLesLibellés(LibelléCandidatALajout)
                        Libellé.AjouterRange(cell)
                    Else
                        Libellé = ToLibellé(b, cell)
                        Me._TousLesLibellés.Add(Libellé)
                    End If
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

#Region "EstUneCelluleDeLibellé"

    Private Function EstUneCelluleDeLibellé(cell As Range, OffsetChampsUnité As Short) As Boolean
        Dim CellUnité As Excel.Range = cell.Offset(0, OffsetChampsUnité)
        Dim UnitéOK = EstUnitéValide(CellUnité.Value)
        Dim EstPasVide = Not String.IsNullOrEmpty(cell.Value)
        Dim r = UnitéOK And EstPasVide
        Return r
    End Function

    Private Function EstUnitéValide(Unité As String) As Boolean
        Dim U As Unités
        Dim r = [Enum].TryParse(Unité, True, U)
        Return r
    End Function

#End Region

#End Region

#Region "RépartirLesLibellésDOuvrages"

    ''' <summary>Selon qu'ils sont doublons ou pas.</summary>
    Public Sub RépartirLesLibellésDOuvrages()
        Dim Ec = WS.EtudeCourante
        IncrémenteAvancementClasseur(Ec.ClasseursExcel.Count)
        For Each c In Ec.ClasseursExcel
            IncrémenteAvancementFeuille(c.Bordereaux.Count)
            For Each b In c.Bordereaux
                RépartirLesLibellésDOuvrages(b)
                IncrémenteAvancementFeuille()
            Next
            IncrémenteAvancementClasseur()
        Next
    End Sub

    Public Sub RépartirLesLibellésDOuvrages(b As Bordereau)
        Try
            Dim LesDoublons As IEnumerable(Of LibelléDouvrage) = From l In TousLesLibellés Where l.Bordereau Is b AndAlso l.NbOccurrences > 1 Order By l.NbOccurrences Descending
            Dim LesPasDoublons As IEnumerable(Of LibelléDouvrage) = From l In TousLesLibellés Where l.Bordereau Is b AndAlso l.NbOccurrences = 1 Order By l.PremierRange.Row Ascending

            Me.LibellésEnDoublonEncoreATraiter.AddRange(LesDoublons)
            Me.LibellésRetenus.AddRange(LesPasDoublons)


        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function ToLibellé(b As Bordereau, rng As Range) As LibelléDouvrage
        Dim r As New LibelléDouvrage(b, rng)
        Return r
    End Function

#End Region

#End Region

#Region "GérerLaQualificationDuDoublon"

    Friend Sub GérerLaQualificationDuDoublon(VraiDoublon As Boolean, L As LibelléDouvrage)
        Me.LibellésEnDoublonEncoreATraiter.Remove(L)
        If VraiDoublon Then
            Me.LibellésRetenus.Add(L)
        Else
            GénérerLesAvatars(L)
        End If
    End Sub

    Private Sub GénérerLesAvatars(L As LibelléDouvrage)
        Dim Avatar As LibelléDouvrage
        For Each rng In L.Ranges
            Avatar = New LibelléDouvrage(L.Bordereau, rng)
            Me.LibellésEnTransit.Add(Avatar)
        Next
    End Sub

    Public Sub PurgerLeTransit()
        For Each L In Me.LibellésEnTransit
            Me.LibellésRetenus.Add(L)
        Next
        Me.LibellésEnTransit.Clear()
    End Sub

#End Region

#Region "Créer les ouvrages"

    Public Sub CréerLesOuvrages(ByRef NbOK As Integer, ByRef NbKO As Integer)
        Dim NewOuvrage As Ouvrage
        NbOK = 0 : NbKO = 0
        Dim NbTraité As Integer = 0, NbATraiter = Me.NbLibellésRetenus

        Dim LibellésDesOuvragesACréer As New List(Of LibelléDouvrage)(Me.LibellésRetenus)
        For Each L In LibellésDesOuvragesACréer
            Try
                NewOuvrage = Ref.GetNewOuvrage()
                NewOuvrage.Nom = L.Libellé
                NewOuvrage.ComplémentDeNom = L.ComplémentDeNom
                NbOK += 1
                Me.LibellésRetenus.Remove(L)
            Catch ex As Exception
                NbKO += 1
                Dim Msg = ex.GetType.Name & " - " & ex.Message
                L.Message = Msg
                ManageErreur(ex, "Erreur en création d'ouvrage", False, False)
            Finally
                NbTraité += 1
                XL.StatusBar = $"Création des ouvrages en cours... {NbTraité}/{NbATraiter} ()"
            End Try
        Next
        XL.StatusBar = ""
    End Sub

#End Region

#Region "Divers"

    Private Sub NotifyPropertyChanged(v As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(v))
    End Sub

    ''' <summary>Sélectionne le <see cref="LibelléDouvrage"/> associé au <param name="Range"/> dans la liste où il se trouve.</summary>
    Friend Sub Sélectionner(Range As Range)

        'TODO : ne traiter que si on est dans la plage libellé de l'un des bordereau (ou du courant)

        Dim V As String = TryCast(Range.Value, String)

        If V Is Nothing Then Exit Sub

        Dim EstAssociéAuRange = Function(LDo As LibelléDouvrage) As Boolean
                                    'Return LDo.EstAssociéA(Range)
                                    Return Object.Equals(LDo.LibelléSource, V)
                                End Function
        Dim LibelléAssociéAuRange = (Me.LibellésRetenus.Where(EstAssociéAuRange)).FirstOrDefault
        If LibelléAssociéAuRange IsNot Nothing Then
            Me.LibelléRetenuCourant = LibelléAssociéAuRange
        Else
            LibelléAssociéAuRange = (Me.LibellésEnDoublonEncoreATraiter.Where(EstAssociéAuRange)).FirstOrDefault
            If LibelléAssociéAuRange IsNot Nothing Then
                Me.LibelléEnDoublonCourant = LibelléAssociéAuRange
            Else
                Message($"Le libellé ""{Range.Value}"" est introuvable. Cette situation est anormale.", MsgBoxStyle.Exclamation)
            End If
        End If
    End Sub

#End Region

#Region "Pilotage Excel"

    Private Sub SélectionnerLesRangesAssociés(L As LibelléDouvrage)
        Try
            Dim OriginWS As Excel.Worksheet = XL.ActiveSheet
            If L Is Nothing Then Exit Sub
            Dim previousWs As Excel.Worksheet = Nothing '= Selection.Worksheet
            'TODO: Provisoire. On fera autrement que précédemment (sélection multiple), ça plante quand on est sur plusieurs classeurs. 
            For Each rng As Excel.Range In L.Ranges
                'XL.Union(Selection, rng)
                If rng.Worksheet IsNot previousWs Then
                    rng.Worksheet.Activate()
                    rng.Select()
                    previousWs = rng.Worksheet
                End If
            Next
            OriginWS.Activate()
        Catch ex As Exception
            ManageErreur(ex, "Echec de la sélection des cellules Excel.")
        End Try
    End Sub


#End Region

#End Region

#Region "Events"

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

#End Region

#Region "Tests et debuggage"


#End Region

End Class
