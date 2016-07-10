Imports Model

Public Class ThisAddIn

#Region "Constructeurs"


#End Region

#Region "Propriétés"

#Region "Nom"
    Public Shared ReadOnly Property Nom() As String
        Get
            Return My.Application.Info.Title
        End Get
    End Property
#End Region

#End Region

#Region "Méthodes"

#Region "Gestion des évennements"

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Try
            ChargerLesDonnées()
        Catch ex As ArgumentException
            ManageErreur(ex, NameOf(ThisAddIn_Startup))
        End Try
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        Try
            EnregistrerLesDonnées()
            ExcelEventManager.CleanUp()
            EIExAddin = Nothing
        Catch ex As ArgumentException
            ManageErreur(ex, NameOf(ThisAddIn_Shutdown))
        End Try
    End Sub

#End Region

    Public Sub ChargerLesDonnées()
        Try
            EIExData.ChargerLeWorkspace()
            EIExData.ChargerLeRéférentiel()
        Catch ex As Exception
            ManageErreur(ex, $"Echec du chargement des données. L'Addin {ThisAddIn.Nom} n'est pas utilisable.")
        End Try

    End Sub

    Public Sub EnregistrerLesDonnées()
        Try
            EIExData.EnregistrerLeWorkspace()
            EIExData.EnregistrerLeRéférentiel()
        Catch ex As Exception
            ManageErreur(ex, $"Echec du chargement des données. L'Addin {ThisAddIn.Nom} n'est pas utilisable.")
        End Try

    End Sub

#End Region

#Region "Tests et debuggage"


#End Region

End Class
