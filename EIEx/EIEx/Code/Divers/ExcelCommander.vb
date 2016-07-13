Imports System.Diagnostics
Imports System.Runtime.CompilerServices
Imports Microsoft.Office.Interop.Excel
Imports Model

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
        Try
            With XL()
                Dim NomDuTableau = InputBox("Nom du tableau à importer : ", , "Tableau1")
                Dim Tableau As ListObject = .ActiveSheet.ListObjects(NomDuTableau)
                Dim NbATraiter = Tableau.ListRows.Count
                For Each lr As ListRow In Tableau.ListRows
                    If AjouteProduit(lr) Then
                        NbProduitsImportés += 1
                    Else
                        NbErreurs += 1
                    End If
                    NbTraités += 1
                    XL.StatusBar = $"{NbTraités}/{NbATraiter}"
                Next
            End With
            Message($"Import terminé. {NbProduitsImportés} produit(s) importé(s), {NbErreurs} erreur(s).")
        Catch ex As Exception
            ManageErreur(ex, $"Echec de l'import. {NbProduitsImportés} produits importés avant incident.", True, False)
        Finally
            XL.ScreenUpdating = True
            XL.StatusBar = ""
        End Try
    End Sub

    Private Function AjouteProduit(lr As ListRow) As Boolean
        Dim Rg = lr.Range

        Dim CodeLydic, RefFournisseur, RefProduit As String
        Dim MotsClés As String, TabMotsClés() As String

        Dim U As String, U2 As Unités

        Dim F1 = Ref.GetFamilleById(1)
        Dim F2 = Ref.GetFamilleById(2)
        Dim IdFamille As Integer

        Try
            Dim NewP = Ref.GetNewProduit()

            With NewP
                CodeLydic = Rg.Cells(4).value
                RefFournisseur = Rg.Cells(5).value
                RefProduit = Produit.GetRéférenceProduit(CodeLydic, RefFournisseur)

                If Ref.LaRéfProduitExisteDéjà(RefProduit) Then Throw New Exception($"La référence produit existe déjà dans le référentiel.")

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

            Return True

        Catch ex As Exception
            MarquerLigneOK(Rg, False, ex.Message())
            Return False
        End Try
    End Function

    Private Sub MarquerLigneOK(r As Range, Result As Boolean, Optional Msg As String = Nothing)
        Dim CellCible As Range = r.Cells(r.Cells.Count).Offset(0, 2)
        CellCible.Value = If(Result, "OK", "KO")
        If Not Result Then
            CellCible.Offset(0, 1).Value = Msg
        End If
    End Sub

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

#End Region

End Module
