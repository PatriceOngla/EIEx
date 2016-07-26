﻿Imports System.Diagnostics
Imports System.Runtime.CompilerServices
Imports EIEx_DAO
Imports Microsoft.Office.Interop.Excel
Imports Model
Imports Utils

Module ExcelCommander

#Region "XL"

    Private WithEvents _XL As Excel.Application
    Function XL() As Excel.Application
        If _XL Is Nothing Then _XL = EIExAddin.Application
        Return _XL
    End Function

    Private Ref As Référentiel = Référentiel.Instance

#End Region

#Region "Import"

    Sub ImporterProduitsDepuisExcel()
        XL.ScreenUpdating = False
        Dim NbTraités As Integer = 0
        Dim NbErreurs As Integer = 0
        Dim NbProduitsImportés As Integer = 0
        Dim NbProduitsMisAJour As Integer = 0
        Dim Résultat As RésultatImport
        Try
            With XL()
                Dim NomDuTableau = InputBox("Nom du tableau à importer : ", , "Tableau1")

                If String.IsNullOrEmpty(NomDuTableau) Then Exit Sub

                Dim Tableau As ListObject = .ActiveSheet.ListObjects(NomDuTableau)
                Dim NbATraiter = Tableau.ListRows.Count
                For Each lr As ListRow In Tableau.ListRows
                    Résultat = ImporteOUMAJProduit(lr)
                    Select Case Résultat
                        Case RésultatImport.Création
                            NbProduitsImportés += 1
                        Case RésultatImport.MAJ
                            NbProduitsMisAJour += 1
                        Case RésultatImport.Echec
                            NbErreurs += 1
                    End Select
                    NbTraités += 1
                    XL.StatusBar = $"{NbTraités}/{NbATraiter}"
                Next
            End With
            Dim Enregistrer = Message($"Import terminé. {NbProduitsImportés} produit(s) importé(s), {NbProduitsMisAJour} produit(s) mis à jour, {NbErreurs} erreur(s).{vbCr}Voulez-vous energistrer le référentiel ?", vbYesNo)

            If Enregistrer = MsgBoxResult.Yes Then
                PersistancyManager.EnregistrerLeRéférentiel()
                Message("Enregistrement effectué.")
            End If
        Catch ex As Exception
            ManageErreur(ex, $"Echec de l'import. {NbProduitsImportés} produits importés avant incident.", True, False)
        Finally
            XL.ScreenUpdating = True
            XL.StatusBar = ""
        End Try
    End Sub

    Private Function ImporteOUMAJProduit(lr As ListRow) As RésultatImport
        Dim Rg = lr.Range

        Dim CodeLydic, RefFournisseur, RefProduit As String
        Dim MotsClés As String, TabMotsClés() As String

        Dim U As String, U2 As Unités

        Dim F1 = Ref.GetFamilleById(1)
        Dim F2 = Ref.GetFamilleById(2)
        Dim IdFamille As Integer

        CodeLydic = Rg.Cells(4).value
        RefFournisseur = Rg.Cells(5).value

        RefProduit = Produit.GetRéférenceProduit(CodeLydic, RefFournisseur)
        Dim ProduitExistant As Boolean

        Try

            If String.IsNullOrEmpty(CodeLydic) OrElse String.IsNullOrEmpty(RefFournisseur) Then
                Throw New Exception("Pas de référence produit valide.")
            End If


            Dim P As Produit = Ref.GetProduitByRefFournisseur(CodeLydic, RefFournisseur)
            If P Is Nothing Then
                P = Ref.GetNewProduit()
            Else
                ProduitExistant = True
            End If

            With P

                'If Ref.LaRéfProduitExisteDéjà(RefProduit) Then Throw New Exception($"La référence produit existe déjà dans le référentiel.")

                .CodeLydic = CodeLydic
                .RéférenceFournisseur = RefFournisseur

                .Nom = Rg.Cells(6).value

                U = Rg.Cells(7).value
                U2 = [Enum].Parse(GetType(Unités), U)
                .Unité = U2

                .Prix = Rg.Cells(8).value
                .TempsDePauseUnitaire = Rg.Cells(9).value

                MotsClés = Rg.Cells(10).value : TabMotsClés = MotsClés?.Split(" ")
                If TabMotsClés IsNot Nothing Then .MotsClés.AddRange(TabMotsClés)

                IdFamille = Rg.Cells(11).value
                .Famille = If(IdFamille = "1", F1, If(IdFamille = 2, F2, Nothing))
            End With

            MarquerLigneOK(Rg, True)

            Return If(ProduitExistant, RésultatImport.MAJ, RésultatImport.Création)

        Catch ex As Exception
            MarquerLigneOK(Rg, False, ex.Message())
            Return RésultatImport.Echec
        End Try
    End Function

    Private Sub MarquerLigneOK(r As Range, Result As Boolean, Optional Msg As String = Nothing)
        Dim CellCible As Range = r.Cells(r.Cells.Count).Offset(0, 2)
        CellCible.Value = If(Result, "OK", "KO")
        If Not Result Then
            CellCible.Offset(0, 1).Value = Msg
        End If
    End Sub

    Private Enum RésultatImport
        Création
        MAJ
        Echec
    End Enum

#End Region

#Region "Divers"

    <Extension>
    Public Function ClasseurRéel(C As ClasseurExcel) As Excel.Workbook
        Dim r = (From wb As Workbook In XL.Workbooks Where wb.FullName.Equals(C.CheminFichier)).FirstOrDefault()
        'Debug.Print(XL.Workbooks.Count())
        Return r
    End Function

    <Extension>
    Public Function Worksheet(B As Bordereau) As Excel.Worksheet
        Try
            Dim cp = B.Parent
            Dim wb = cp.ClasseurRéel
            If wb Is Nothing Then
                Throw New Exception($"Le classeur ""{cp.CheminFichier}"" n'est pas ouvert.")
            End If
            If String.IsNullOrEmpty(B.NomFeuille) Then
                Throw New Exception($"Le nom de la feuille Excel n'est pas renseigné pour le bordereau ""{B.Nom}"".")
            End If
            Dim r As Excel.Worksheet = wb.Worksheets(B.NomFeuille)
            Return r
        Catch ex As Exception
            Throw New Exception($"Impossible de récupérer la feuille Excel pour le bordereau ""{B.Nom}"". 
Vérifier que le nom de la feuille est défini par le bordereau correspond à un nom de feuille existante dans le fichier Excel associé.", ex)
        End Try
    End Function

    ''' <summary>
    ''' Retourne le <paramref name="rng"/> limité à la dernière cellule utilisée de sa feuille.
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <returns></returns>
    <Extension>
    Public Function LimitedRange(rng As Excel.Range) As Excel.Range
        Dim lastCell As Excel.Range = rng.SpecialCells(XlCellType.xlCellTypeLastCell)
        Dim sh As Excel.Worksheet = rng.Parent
        Dim LastRow = rng.Row + rng.Rows.Count - 1
        Dim LastColumn = rng.Column + rng.Columns.Count - 1

        LastRow = Math.Min(LastRow, lastCell.Row)
        LastColumn = Math.Min(LastColumn, lastCell.Column)
        Dim r As Excel.Range = sh.Range(rng.Cells(1), sh.Cells(LastRow, LastColumn))
        Return r
    End Function

    '#Region "Gestion des classeurs assoxciés à l'étude courante"

    '    Private WS As WorkSpace = WorkSpace.Instance

    '    Private Function CheckEtudeCourante() As Boolean
    '        If WS.EtudeCourante Is Nothing Then
    '            Message("Aucune étude n'est sélectionnée.")
    '            Return False
    '        Else
    '            Return True
    '        End If
    '    End Function

    '#Region "OuvrirLesClasseurDeLEtudeCourante"

    '    Public Sub OuvrirLesClasseurDeLEtudeCourante()

    '        If Not CheckEtudeCourante() Then Exit Sub

    '        Dim EC = WS.EtudeCourante

    '        If EC Is Nothing Then
    '            Message("Aucune étude n'est sélectionnée.")
    '            Exit Sub
    '        End If

    '        Try
    '            With EC
    '                For Each c In .ClasseursExcel
    '                    Try
    '                        If IO.File.Exists(c.CheminFichier) Then
    '                            XL.Workbooks.Open(c.CheminFichier)
    '                        Else
    '                            Message($"Le classeur ""{c.CheminFichier}"" est introuvable.)", MsgBoxStyle.Exclamation)
    '                        End If
    '                    Catch ex As Exception
    '                        ManageErreur(ex, $"Echec de la tentative d'ouverture du classeur ""{c.CheminFichier}""", True, False)
    '                    End Try
    '                Next
    '            End With
    '        Catch ex As Exception
    '            ManageErreur(ex, , True, False)
    '        End Try
    '    End Sub

    '#End Region

    '#Region "InitiliaserLesClasseursDeLEtudeCourante"

    '    Private Sub InitiliaserLesClasseursDeLEtudeCourante()

    '        If Not CheckEtudeCourante() Then Exit Sub

    '        Dim EC = WS.EtudeCourante

    '        Try
    '            Dim NewC As ClasseurExcel

    '            If EC Is Nothing Then Throw New Exception("Pas d'étude courante.")
    '            With EC
    '                For Each wb As Excel.Workbook In XL.Workbooks
    '                    If Not ContientLeClasseur(EC, wb.FullName) Then
    '                        NewC = .AjouterNouveauClasseur()
    '                        NewC.CheminFichier = wb.FullName
    '                        AjouterLesFeuilles(NewC)
    '                        NewC.Nom = wb.Name
    '                    End If
    '                Next

    '                If Me.ClasseurExcelCourant Is Nothing Then
    '                    Me.ClasseurExcelCourant = EC.ClasseursExcel.FirstOrDefault
    '                End If
    '            End With
    '        Catch ex As Exception
    '            ManageErreur(ex, , True, False)
    '        End Try
    '    End Sub

    '    Private Sub AjouterLesFeuilles(C As ClasseurExcel)
    '        Dim WShts As Microsoft.Office.Interop.Excel.Sheets = C.ClasseurRéel?.Worksheets
    '        'Dim WShts2 As Microsoft.Office.Tools.Excel.Worksheet.Worksheets = C.ClasseurRéel?.Worksheets
    '        'Dim WShts As Excel.Worksheets = C.ClasseurRéel?.Worksheets
    '        Dim B As Bordereau
    '        Debug.Print(WShts.Count())
    '        If WShts IsNot Nothing Then
    '            For Each Wsht In WShts
    '                B = C.AjouterNouveauBordereau()
    '                B.Nom = Wsht.Name
    '                B.NomFeuille = B.Nom
    '            Next
    '        End If

    '    End Sub

    '    Private Function ContientLeClasseur(EC As Etude, Chemin As String) As Boolean
    '        Dim r As Boolean
    '        r = (From c In EC.ClasseursExcel Where Object.Equals(c.CheminFichier, Chemin)).Any()
    '        Return r
    '    End Function

    '#End Region

    '#End Region

#End Region

End Module
