Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office
Imports System.Runtime.InteropServices
Imports Utils

''' <summary>
''' From https://code.msdn.microsoft.com/Basics-of-using-Excel-4453945d#content.
''' Used to obtain both worksheet and named range names from a valid Excel file.
''' </summary>
Public Class ExcelInfoBag
    Implements IDisposable

#Region "Constructeurs"

    Public Sub New()
        _ExcelInstance = TryCast(Marshal.GetActiveObject("Excel.application"), Excel.Application)
        If _ExcelInstance Is Nothing Then _ExcelInstance = New Excel.Application
    End Sub

    ''' <summary>File to get information from.</summary>
    ''' <param name="FileName"></param>
    ''' <remarks>
    ''' The caller is responsible to ensure the file exists.
    ''' </remarks>
    Public Sub New(ByVal FileName As String)
        Me.FileName = FileName
    End Sub

#End Region

#Region "Propriétés"

#Region "ExcelInstance"
    Private Shared _ExcelInstance As Excel.Application
    Public Shared ReadOnly Property ExcelInstance() As Excel.Application
        Get
            Return _ExcelInstance
        End Get
    End Property
#End Region

#Region "LastException"
    Public Property LastException As Exception
#End Region

#Region "FileName"
    Private Extensions As String() = {".xls", ".xlsx"}
    Private _FileName As String
    ''' <summary>Valid/existing Excel file name to work with.</summary>
    Public Property FileName() As String
        Get
            Return _FileName
        End Get
        Set(ByVal value As String)
            If Not Extensions.Contains(IO.Path.GetExtension(value.ToLower)) Then
                Throw New Exception("Nom de fichier incorrect.")
            End If
            _FileName = value
        End Set
    End Property
#End Region

#Region "NameRanges"
    Private _NameRanges As New List(Of String)
    ''' <summary>List of named ranges in current file.</summary>
    Public ReadOnly Property NameRanges() As List(Of String)
        Get
            Return _NameRanges
        End Get
    End Property
#End Region

#Region "WorkBooks"
    Private _WorkBooks As List(Of Excel.Workbook)
    Public ReadOnly Property WorkBooks() As List(Of Excel.Workbook)
        Get
            Return _WorkBooks
        End Get
    End Property
#End Region

#Region "Sheets"
    Private _Sheets As New List(Of String)
    ''' <summary>List of work sheets in current file.</summary>
    Public ReadOnly Property Sheets() As List(Of String)
        Get
            Return _Sheets
        End Get
    End Property
#End Region

#Region "SheetsData "
    Private _SheetsData As New Dictionary(Of Int32, String)
    Public ReadOnly Property SheetsData As Dictionary(Of Int32, String)
        Get
            Return _SheetsData
        End Get
    End Property
#End Region

#End Region

#Region "Méthodes"

#Region "GetInformation"

    ''' <summary>Retrieve worksheet and name range names.</summary>
    Public Function GetInformation() As Boolean

#Region "Préparation"
        CheckFileName()
        ResetInfos()
#End Region

#Region "Déclarations"
        Dim Success As Boolean = True
        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkBooks As Excel.Workbooks = Nothing
        Dim xlWorkBook As Excel.Workbook = Nothing
        Dim xlActiveRanges As Excel.Workbook = Nothing
        Dim xlNames As Excel.Names = Nothing
        Dim xlWorkSheets As Excel.Sheets = Nothing
#End Region

#Region "Information retrieving"

        Try
            xlApp = TryCast(Marshal.GetActiveObject("Excel.application"), Excel.Application)
            If xlApp Is Nothing Then xlApp = New Excel.Application
            xlApp.DisplayAlerts = False
            xlWorkBooks = xlApp.Workbooks
            xlWorkBook = xlWorkBooks.Open(FileName)

            xlActiveRanges = xlApp.ActiveWorkbook
            xlNames = xlActiveRanges.Names

            For x As Integer = 1 To xlNames.Count
                Dim xlName As Excel.Name = xlNames.Item(x)
                _NameRanges.Add(xlName.Name)
                Runtime.InteropServices.Marshal.FinalReleaseComObject(xlName)
                xlName = Nothing
            Next

            xlWorkSheets = xlWorkBook.Sheets

            For x As Integer = 1 To xlWorkSheets.Count
                Dim Sheet1 As Excel.Worksheet = CType(xlWorkSheets(x), Excel.Worksheet)
                _Sheets.Add(Sheet1.Name)
                _SheetsData.Add(x, Sheet1.Name)
                Runtime.InteropServices.Marshal.FinalReleaseComObject(Sheet1)
                Sheet1 = Nothing
            Next

            xlWorkBook.Close()
            xlApp.UserControl = True
            xlApp.Quit()

        Catch ex As Exception
            _LastException = ex
            Success = False
        Finally

#End Region

#Region "Free memory"

            If Not xlWorkSheets Is Nothing Then
                Marshal.FinalReleaseComObject(xlWorkSheets)
                xlWorkSheets = Nothing
            End If

            If Not xlNames Is Nothing Then
                Marshal.FinalReleaseComObject(xlNames)
                xlNames = Nothing
            End If

            If Not xlActiveRanges Is Nothing Then
                Runtime.InteropServices.Marshal.FinalReleaseComObject(xlActiveRanges)
                xlActiveRanges = Nothing
            End If
            If Not xlActiveRanges Is Nothing Then
                Runtime.InteropServices.Marshal.FinalReleaseComObject(xlActiveRanges)
                xlActiveRanges = Nothing
            End If

            If Not xlWorkBook Is Nothing Then
                Marshal.FinalReleaseComObject(xlWorkBook)
                xlWorkBook = Nothing
            End If

            If Not xlWorkBooks Is Nothing Then
                Marshal.FinalReleaseComObject(xlWorkBooks)
                xlWorkBooks = Nothing
            End If

            If Not xlApp Is Nothing Then
                Marshal.FinalReleaseComObject(xlApp)
                xlApp = Nothing
            End If
        End Try

#End Region

        Return Success

    End Function

    Private Sub CheckFileName()
        If Not IO.File.Exists(FileName) Then
            Dim ex As New Exception($"Le fichier ""{FileName}"" est introuvable.")
            _LastException = ex
            Throw ex
        End If
    End Sub

    Private Sub ResetInfos()

        _Sheets.Clear()
        _NameRanges.Clear()
        _SheetsData.Clear()

    End Sub

#End Region

#Region "IDisposable Support"

    Private disposedValue As Boolean ' Pour détecter les appels redondants

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                '  TODO: supprimer l'état managé (objets managés).
            End If

            Try
                Dim FinalReleaseComObject = Sub(ByRef o As Object)
                                                'TODO : typer o en COM_Object
                                                If o IsNot Nothing Then
                                                    Marshal.FinalReleaseComObject(o)
                                                    o = Nothing
                                                End If
                                            End Sub

                FinalReleaseComObject(_ExcelInstance)

                Me.WorkBooks.DoForAll(Sub(wb As Excel.Workbook)
                                          FinalReleaseComObject(wb)
                                      End Sub)
                Me._WorkBooks = Nothing

                FinalReleaseComObject(_Sheets)
                _Sheets = Nothing

            Catch ex As Exception
                Throw ex
            End Try

            ' TODO: définir les champs de grande taille avec la valeur Null.
        End If

        disposedValue = True

    End Sub

    ' TODO: remplacer Finalize() seulement si la fonction Dispose(disposing As Boolean) ci-dessus a du code pour libérer les ressources non managées.
    Protected Overrides Sub Finalize()
        ' Ne modifiez pas ce code. Placez le code de nettoyage dans Dispose(disposing As Boolean) ci-dessus.
        Dispose(False)
        MyBase.Finalize()
    End Sub

    ' Ce code est ajouté par Visual Basic pour implémenter correctement le modèle supprimable.
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

#End Region

#End Region

#Region "Tests et debuggage"


#End Region


End Class
