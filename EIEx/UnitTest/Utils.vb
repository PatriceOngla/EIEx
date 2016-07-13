Module Utils

    Public Sub CopierLeFichier(CheminSource As String)
        Try
            Dim CheminDossier = IO.Path.GetDirectoryName(CheminSource)
            Dim Ext = IO.Path.GetExtension(CheminSource)
            Dim NomFichier = IO.Path.GetFileNameWithoutExtension(CheminSource) & " - Copie" & Ext
            Dim CheminCopie = IO.Path.Combine(CheminDossier, NomFichier)
            IO.File.Copy(CheminSource, CheminCopie, True)
        Catch ex As Exception
            Debug.Print("ùkjmlkhiuy - " & ex.Message)
        End Try
    End Sub

End Module
