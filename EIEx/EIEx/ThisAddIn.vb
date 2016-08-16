Imports System.Windows
Imports EIEx_DAO
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

    '#Region "UC_Container"
    '    Private Shared _UC_Container As UC_Container
    '    Public Shared ReadOnly Property UC_Container() As UC_Container
    '        Get
    '            Return _UC_Container
    '        End Get
    '    End Property
    '#End Region

#End Region

#Region "Méthodes"

#Region "Gestion des évennements"

#Region "ThisAddIn_Startup"

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Try
            '_UC_Container = New UC_Container()
            ChargerLesDonnées()
            DémarrerGestionGlobaleDesException()
        Catch ex As ArgumentException
            ManageErreur(ex, NameOf(ThisAddIn_Startup))
        End Try
    End Sub

#Region "Centralisation de la gestion des erreurs du modèle"

    Private Sub DémarrerGestionGlobaleDesException()
        'AddHandler AppDomain.CurrentDomain.UnhandledException, Sub(sender As Object, args As UnhandledExceptionEventArgs)
        '                                                           ManageErreur(args.ExceptionObject, "", True)
        '                                                       End Sub
        AddHandler Système.ExceptionRaised, Sub(e As Exception, S As Système, Attendue As Boolean)
                                                TraiterLesExceptionsDesSysèmes(e, S, Attendue)
                                            End Sub
    End Sub

    Private Sub TraiterLesExceptionsDesSysèmes(e As Exception, S As Système, Attendue As Boolean)
        ManageErreur(e, $"Erreur dans ""{S.Nom}"".", AffichageSimple:=Attendue)
    End Sub

#End Region

#End Region

#Region "ThisAddIn_Shutdown"

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

#End Region

#Region "Gestion des données"

    Public Sub ChargerLesDonnées()
        Try
            PersistancyManager.ChargerLeWorkspace()
            PersistancyManager.ChargerLeRéférentiel()
#If DEBUG Then
            Test()
#End If
        Catch ex As Exception
            ManageErreur(ex, $"Echec du chargement des données. L'Addin {ThisAddIn.Nom} n'est pas utilisable.")
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

        'Dim SP As New UC_SélecteurDeProduit
        'SP.Show()
        UC_SélecteurDeProduit.Show2()


        'Dim SO As New UC_SélecteurDOuvrage
        'SO.Show()
    End Sub
#End If

#End Region

End Class
