﻿Imports System.Diagnostics

Module Common

    Public EIExAddin As ThisAddIn = Globals.ThisAddIn

    Public Sub ManageErreur(ex As Exception, Optional Msg As String = Nothing, Optional DlgBox As Boolean = True)
        Dim SubName As String = ""
        Try
            Dim stackTab = ex.StackTrace.Split(vbCrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            SubName = stackTab(0)
            SubName = (SubName.Replace("à ", "")).Trim()
        Catch ex2 As Exception
            SubName = "##! Impossible de déterminer le nom de la procédure source ##"
        End Try

        Utils.ManageError(ex, SubName, Msg, DlgBox, ThisAddIn.Nom)
    End Sub

    Public Function Message(Msg As String, Optional Style As MsgBoxStyle = MsgBoxStyle.Exclamation) As MsgBoxResult
        Return MsgBox(Msg, Style, ThisAddIn.Nom)
    End Function

End Module
