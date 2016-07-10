Imports System.Windows
Imports Model

Public Class Bordereau
    Inherits EntitéDuWorkSpace

#Region "Constructeurs"

    Public Sub New()

    End Sub

    Protected Overrides Sub Init()
        _Paramètres = New Paramètres
    End Sub

    'Protected Overrides Sub SetId()
    '    Me._Id = Système.GetNewId(Of Bordereau)
    'End Sub

    'Protected Overrides Sub SEnregistrerDansLeRéférentiel()
    '    Système.EnregistrerRoot(Me)
    'End Sub

#End Region

#Region "Propriétés"

#Region "Système"
    Public Overrides ReadOnly Property Système As Système
        Get
            Return WorkSpace.Instance
        End Get
    End Property

#End Region

#Region "CheminFichier"
    Private _CheminFichier As String
    Public Property CheminFichier() As String
        Get
            Return _CheminFichier
        End Get
        Set(ByVal value As String)
            If Object.Equals(value, Me._CheminFichier) Then Exit Property
            _CheminFichier = value
            NotifyPropertyChanged(NameOf(CheminFichier))
        End Set
    End Property
#End Region

#Region "Paramètres (Paramètres)"
    Private _Paramètres As Paramètres
    Public ReadOnly Property Paramètres() As Paramètres
        Get
            Return _Paramètres
        End Get
    End Property
#End Region

#End Region

#Region "Méthodes"


#End Region

#Region "Tests et debuggage"


#End Region

End Class
