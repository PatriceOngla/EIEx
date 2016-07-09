Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Windows
Imports System.Xml.Serialization

Public Class EIExData
    Inherits EIExObject

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Protected Overrides Sub Init()
        InitialiserLeWorkspace()
        InitialiserLeRéférentiel()
    End Sub

    Private Sub InitialiserLeRéférentiel()
        If Not (IO.Directory.Exists(CheminDossierRéférentiel)) Then
            IO.Directory.CreateDirectory(CheminDossierRéférentiel)
        End If
        If IO.File.Exists(CheminRéférentiel) Then
            ChargerLeRéférentiel()
        End If
    End Sub

    Private Sub InitialiserLeWorkspace()
        If Not (IO.Directory.Exists(CheminDossierWorkspace)) Then
            IO.Directory.CreateDirectory(CheminDossierWorkspace)
        End If
        If IO.File.Exists(CheminWorkspace) Then
            ChargerLeWorkspace()
        End If
    End Sub

#End Region

#Region "Propriétés"

#Region "Référentiel"

#Region "CheminDossierRéférentiel (String)"
    Public Shared ReadOnly Property CheminDossierRéférentiel() As String
        Get
            Dim AppFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            Dim r = IO.Path.Combine(AppFolder, "EIEx")
            Return r
        End Get
    End Property
#End Region

#Region "CheminRéférentiel (String)"
    Public Shared ReadOnly Property CheminRéférentiel() As String
        Get
            Dim r = IO.Path.Combine(CheminDossierRéférentiel, "EIExRef.xml")
            Return r
        End Get
    End Property
#End Region

    Public Shared ReadOnly Property Référentiel() As Référentiel
        Get
            Return Référentiel.Instance
        End Get
    End Property

#End Region

#Region "Workspace"

#Region "CheminDossierWorkspace (String)"
    Public Shared ReadOnly Property CheminDossierWorkspace() As String
        Get
            Dim r = CheminDossierRéférentiel
            Return r
        End Get
    End Property
#End Region

#Region "CheminWorkspace (String)"
    Public Shared ReadOnly Property CheminWorkspace() As String
        Get
            Dim r = IO.Path.Combine(CheminDossierWorkspace, "EIExWorkSpace.xml")
            Return r
        End Get
    End Property
#End Region

    Public Shared ReadOnly Property Workspace() As WorkSpace
        Get
            Return WorkSpace.Instance
        End Get
    End Property

#End Region

#End Region

#Region "Méthodes"

#Region "Persistance"

#Region "Référentiel"

    Public Shared Sub ChargerLeRéférentiel()
        Référentiel.Charger(CheminRéférentiel)
    End Sub

    Public Shared Sub EnregistrerLeRéférentiel()
        Référentiel.Enregistrer(CheminRéférentiel)
    End Sub

#End Region

#Region "Workspace"

    Public Shared Sub ChargerLeWorkspace()
        WorkSpace.Charger(CheminWorkspace)
    End Sub

    Public Shared Sub EnregistrerLeWorkspace()
        Workspace.Enregistrer(CheminWorkspace)
    End Sub

#End Region

#End Region

#End Region

#Region "Tests et debuggage"


#End Region

End Class
