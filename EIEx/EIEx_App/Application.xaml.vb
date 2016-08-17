Imports System.Windows
Imports System.Windows.Threading
Imports EIEx_DAO
Imports Model

Class Application

    ' Les événements de niveau application, par exemple Startup, Exit et DispatcherUnhandledException
    ' peuvent être gérés dans ce fichier.

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

#Region "ThisAddIn_Startup"

    Private Sub Application_Startup() Handles Me.Startup
        Try
            ChargerLesDonnées()
            DémarrerGestionGlobaleDesException()
        Catch ex As ArgumentException
            ManageErreur(ex, NameOf(Application_Startup))
        End Try
    End Sub

#Region "Centralisation de la gestion des erreurs du modèle"

    Private Sub Application_DispatcherUnhandledException(sender As Object, e As DispatcherUnhandledExceptionEventArgs) Handles Me.DispatcherUnhandledException
        ManageErreur(e.Exception)
    End Sub

    Private Sub DémarrerGestionGlobaleDesException()
        'TODO: désormais plus indispensable. Remplacer par une exception spécifique et un handler sur DispatcherUnhandledException

        AddHandler Système.ExceptionRaised, Sub(e As Exception, S As Système, Attendue As Boolean)
                                                TraiterLesExceptionsDesSysèmes(e, S, Attendue)
                                            End Sub
    End Sub

    Private Sub TraiterLesExceptionsDesSysèmes(e As Exception, S As Système, Attendue As Boolean)
        ManageErreur(e, $"Erreur dans ""{S.Nom}"".", AffichageSimple:=Attendue)
    End Sub

#End Region

#End Region

#Region "Application_Exit"

    Private Sub Application_Exit() Handles Me.Exit
        Try
            EnregistrerLesDonnées()
            ExcelEventManager.CleanUp()
        Catch ex As ArgumentException
            ManageErreur(ex, NameOf(Application_Exit))
        End Try
    End Sub

#End Region

#End Region

#Region "Gestion des données"

    Public Sub ChargerLesDonnées()
        Try
            PersistancyManager.ChargerLeWorkspace()
            PersistancyManager.ChargerLeRéférentiel()
        Catch ex As Exception
            ManageErreur(ex, $"Echec du chargement des données.")
        End Try

    End Sub

    Public Sub EnregistrerLesDonnées()
        Try
            PersistancyManager.EnregistrerLeWorkspace()
            PersistancyManager.EnregistrerLeRéférentiel()
        Catch ex As Exception
            ManageErreur(ex, $"Echec de l'enregistrement des données. Attention, votre travail risque d'être perdu à l'arrêt d'Excel.")
        End Try

    End Sub

#End Region

#End Region

#Region "Tests et debuggage"

#If DEBUG Then
    Public Sub Test()
        'Dim w = New Window1
        'w.ShowDialog()

        Win_SélecteurDeProduit.Cherche()

        Win_SélecteurDOuvrage.Cherche()

    End Sub
#End If

#End Region

End Class
