Imports Microsoft.Office.Interop.Excel
Imports Model

Module ExcelComander

    Private WithEvents _XL As Excel.Application
    Function XL() As Excel.Application
        If _XL Is Nothing Then _XL = EIExAddin.Application
        Return _XL
    End Function

    Private Ref As Référentiel = Référentiel.Instance

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

End Module
