Imports System.Windows

Public Class Bordereau
    Inherits AgregateRoot

#Region "Constructeurs"

    Public Sub New()

    End Sub

    Public Sub New(Id As Integer)
        MyBase.New(Id)
    End Sub
    Protected Overrides Sub Init()
        _Paramètres = New Paramètres
    End Sub

    Protected Overrides Sub SetId()
        Me._Id = Réf.GetNewId(Of Bordereau)
    End Sub

    Protected Overrides Sub SEnregistrerDansLeRéférentiel()
        Réf.EnregistrerRoot(Me)
    End Sub

#End Region

#Region "Propriétés"

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
