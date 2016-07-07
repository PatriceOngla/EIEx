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

    Private _Référentiel As Référentiel
    Public ReadOnly Property Référentiel() As Référentiel
        Get
            Return Référentiel.Instance
        End Get
    End Property

#End Region

#Region "Bordereaux"
    Private _Bordereaux As ObservableCollection(Of Bordereau)
    Public ReadOnly Property Bordereaux() As ObservableCollection(Of Bordereau)
        Get
            Return _Bordereaux
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

    Public Shared Sub EnregistrerRéférentiel()
        Référentiel.Enregistrer(CheminRéférentiel)
    End Sub

#End Region

#Region "Bordereau"

    Public Sub ChargerBordereau(Chemin As String)
        Me._Bordereaux = Utils.DéSérialisation(Of ObservableCollection(Of Bordereau))(Chemin)

    End Sub

    Public Sub EnregistrerBordereaul(Chemin As String)
        Utils.Sérialiser(Me.Bordereaux, Chemin)
    End Sub

#End Region

#End Region

#End Region

#Region "Tests et debuggage"


#End Region

End Class
