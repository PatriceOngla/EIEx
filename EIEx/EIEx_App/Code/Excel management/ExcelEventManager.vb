Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Model

Friend Class ExcelEventManager

#Region "Properties"

#Region "XL"
    Private Shared WithEvents _XL As Excel.Application
    Public Shared ReadOnly Property XL() As Excel.Application
        Get
            If _XL Is Nothing Then _XL = ExcelCommander.XL
            Return _XL
        End Get
    End Property
#End Region

#Region "TargetWindow"
    'Private Shared _TargetWindow As Excel.Window
    Public Shared ReadOnly Property TargetWindow() As Excel.Window
        Get
            'If _TargetWindow Is Nothing Then _TargetWindow = XL.ActiveWindow
            'Return _TargetWindow
            Return XL.ActiveWindow
        End Get
        'Set(ByVal value As Excel.Window)
        '    _TargetWindow = value
        'End Set
    End Property
#End Region

#Region "TargetSheet"
    Private Shared WithEvents _TargetSheet As Excel.Worksheet
    Public Shared Property TargetSheet() As Excel.Worksheet
        Get
            If _TargetSheet Is Nothing Then
                _TargetSheet = XL.ActiveSheet
            End If
            Return _TargetSheet
        End Get
        Set(ByVal value As Excel.Worksheet)
            _TargetSheet = value
        End Set
    End Property

#End Region

#Region "UCSC"

    Private Shared _UCSC As Win_Main
    Public Shared Property UCSC() As Win_Main
        Get
            Return _UCSC
        End Get
        Set(ByVal value As Win_Main)
            _UCSC = value
        End Set
    End Property

#End Region

#Region "WS"
    Public Shared ReadOnly Property WS As WorkSpace
        Get
            Return WorkSpace.Instance
        End Get
    End Property
#End Region

#End Region

#Region "Methods"

#Region "Filtrage des événements Excel"

    Private Shared Sub _XL_SheetSelectionChange(Sh As Object, Target As Excel.Range) Handles _XL.SheetSelectionChange
        Try
            UCSC.SelectedRange = Target.Address
            If LaFeuilleActiveEstCelleDeLUnDesBordereauxDuClasseurCourant() Then
                RaiseEvent TargetSelectedRangeChanged(Target)
            End If
        Catch ex As ArgumentException
            ManageErreur(ex, , True, False)
        End Try

    End Sub

    'Private Shared Function LaFeuilleAffichéeEstCelleDuBordereauCourant() As Boolean

    '    Dim awb = XL.ActiveWorkbook
    '    Dim ash As Excel.Worksheet = XL.ActiveSheet
    '    Dim awbPath = awb.FullName.Replace("""", "")
    '    Dim ashName = ash.Name
    '    Dim NomFicherBordereauCourant = WS?.ClasseurExcelCourant?.CheminFichier
    '    Dim NomFeuilleCourante = WS?.BordereauCourant?.NomFeuille
    '    Dim r As Boolean
    '    If (String.IsNullOrEmpty(NomFicherBordereauCourant)) Or String.IsNullOrEmpty(NomFeuilleCourante) Then
    '        r = False
    '    Else
    '        NomFicherBordereauCourant = NomFicherBordereauCourant.Replace("""", "")
    '        Dim OKFichier = Object.Equals(awbPath, NomFicherBordereauCourant)
    '        Dim OKFeuille = Object.Equals(ashName, NomFeuilleCourante)
    '        r = OKFichier AndAlso OKFeuille
    '    End If
    '    Return r
    'End Function

    Private Shared Function LaFeuilleActiveEstCelleDeLUnDesBordereauxDuClasseurCourant() As Boolean

        Dim awb = XL.ActiveWorkbook
        Dim ash As Excel.Worksheet = XL.ActiveSheet
        Dim cc = WS?.ClasseurExcelCourant

        Dim r = (From b In cc.Bordereaux Where b.Worksheet Is ash).Any()
        Return r

    End Function

#End Region

    Public Shared Sub CleanUp()
        _XL = Nothing
    End Sub

#End Region

#Region "Events"

    ''' <summary>Modification de la sélection dans l'une des feuille cibles.</summary>
    Public Shared Event TargetSelectedRangeChanged(newSelectedRange As Excel.Range)


#End Region


#Region "Debug et tests"

    Private Shared Sub _XL_SheetChange(Sh As Object, Target As Excel.Range) Handles _XL.SheetChange

    End Sub


#End Region

End Class
