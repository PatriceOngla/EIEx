Imports System.Collections.ObjectModel

Public Class Etude
    Inherits AgregateRoot(Of Etude)

#Region "Constructeurs"
    Public Sub New(Id As Integer)
        MyBase.New(Id)
    End Sub

    Protected Overrides Sub Init()
        Me.Nom = "Nouvelle étude"
        _Bordereaux = New ObservableCollection(Of Bordereau)
    End Sub

#End Region

#Region "Propriétés"

#Region "Système"
    Public Overrides ReadOnly Property Système As Système
        Get
            Return WorkSpace.Instance
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


#End Region

#Region "Tests et debuggage"


#End Region

End Class
