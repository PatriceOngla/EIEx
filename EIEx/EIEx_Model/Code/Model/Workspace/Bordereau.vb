Imports System.ComponentModel
Imports System.Windows
Imports Model

Public Class Bordereau
    Inherits EntitéDuWorkSpace

#Region "Constructeurs"

    Friend Sub New()
    End Sub

    Protected Overrides Sub Init()
        _Paramètres = New Paramètres
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

#Region "Parent"
    Private _Parent As ClasseurExcel
    Public Property Parent() As ClasseurExcel
        Get
            Return _Parent
        End Get
        Friend Set(value As ClasseurExcel)
            _Parent = value
        End Set
    End Property
#End Region

#Region "NomFeuille"
    Private _NomFeuille As String
    Public Property NomFeuille() As String
        Get
            Return _NomFeuille
        End Get
        Set(ByVal value As String)
            If Object.Equals(value, Me._NomFeuille) Then Exit Property
            _NomFeuille = value
            NotifyPropertyChanged(NameOf(NomFeuille))
        End Set
    End Property
#End Region

#Region "Paramètres (Paramètres)"
    Private WithEvents _Paramètres As Paramètres
    Public ReadOnly Property Paramètres() As Paramètres
        Get
            Return _Paramètres
        End Get
    End Property
#End Region

#End Region

#Region "Méthodes"

    Private Sub _Paramètres_PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Handles _Paramètres.PropertyChanged
        Me.NotifyPropertyChanged(NameOf(Paramètres))
    End Sub

#End Region

#Region "Tests et debuggage"


#End Region

End Class
