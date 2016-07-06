Imports System.Windows

Public Class Bordereau
    Inherits EIExObject

#Region "Constructeurs"


#End Region

#Region "Propriétés"

#Region "Nom (String)"

#Region "Déclaration et registration de NomProperty"

    Private Shared FlagsMDNom As FrameworkPropertyMetadataOptions = 0
    Private Shared MDNom As New FrameworkPropertyMetadata(0, FlagsMDNom)
    Public Shared NomProperty As DependencyProperty = DependencyProperty.Register(NameOf(Bordereau.Nom), GetType(String), GetType(Bordereau), MDNom)

#End Region

#Region "Wrapper CLR de NomProperty"
    Public Property Nom() As String
        Get
            Return GetValue(NomProperty)
        End Get
        Set(ByVal value As String)
            SetValue(NomProperty, value)
        End Set
    End Property

#End Region

#Region "Gestion évennementielle de la mise à jour de NomProperty"

    Public Event NomChangedEvent()

#End Region

#End Region

#Region "Paramètres (Paramètres)"
    Private _Paramètres = New Paramètres
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
