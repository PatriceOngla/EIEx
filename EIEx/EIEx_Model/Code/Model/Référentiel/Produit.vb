Imports System.Collections.ObjectModel

Public Class Produit
    Inherits AgregateRootDuRéférentiel(Of Produit)

#Region "Constructeurs"

    Public Sub New(Id As Integer)
        MyBase.New(Id)
        _MotsClés = New List(Of String)
    End Sub

    Protected Overrides Sub Init()
        _MotsClés = New List(Of String)
    End Sub

#End Region

#Region "Propriétés"

#Region "Unité"
    Private _Unité As Unités
    Public Property Unité() As Unités
        Get
            Return _Unité
        End Get
        Set(ByVal value As Unités)
            If Object.Equals(value, Me._Unité) Then Exit Property
            _Unité = value
            NotifyPropertyChanged(NameOf(Unité))
        End Set
    End Property
#End Region

#Region "Prix (Single)"
    Private _Prix As Single
    Public Property Prix() As Single
        Get
            Return _Prix
        End Get
        Set(ByVal value As Single)
            If Object.Equals(value, Me._Prix) Then Exit Property
            _Prix = value
            NotifyPropertyChanged(NameOf(Prix))
        End Set
    End Property
#End Region

#Region "Référence produit"

#Region "CodeLydic"
    Private _CodeLydic As String
    Public Property CodeLydic() As String
        Get
            Return _CodeLydic
        End Get
        Set(ByVal value As String)
            If Object.Equals(value, Me._CodeLydic) Then Exit Property
            Dim TestUnicité = CheckUnicitéRefProduit(value, Me.RéférenceFournisseur)
            If TestUnicité Then
                _CodeLydic = value
                NotifyPropertyChanged(NameOf(CodeLydic))
                NotifyPropertyChanged(NameOf(RéférenceProduit))
            End If
        End Set
    End Property
#End Region

#Region "RéférenceFournisseur (String)"
    Private _RéférenceFournisseur As String
    Public Property RéférenceFournisseur() As String
        Get
            Return _RéférenceFournisseur
        End Get
        Set(ByVal value As String)
            If Object.Equals(value, Me._RéférenceFournisseur) Then Exit Property
            Dim TestUnicité = CheckUnicitéRefProduit(Me.CodeLydic, value)
            If TestUnicité Then
                _RéférenceFournisseur = value
                NotifyPropertyChanged(NameOf(RéférenceFournisseur))
                NotifyPropertyChanged(NameOf(RéférenceProduit))
            End If
        End Set
    End Property
#End Region

#Region "RéférenceProduit"

    Public ReadOnly Property RéférenceProduit() As String
        Get
            Return GetRéférenceProduit(Me.CodeLydic, Me.RéférenceFournisseur)
        End Get
    End Property

    Public Shared Function GetRéférenceProduit(CodeLydic As String, ReférenceFournisseur As String) As String
        Return If(CodeLydic, "?") & "-" & If(ReférenceFournisseur, "?")
    End Function

#End Region

    Private Function CheckUnicitéRefProduit(CodeLydic As String, RéférenceFournisseur As String) As Boolean
        If Not (String.IsNullOrEmpty(CodeLydic) OrElse String.IsNullOrEmpty(RéférenceFournisseur)) Then
            Dim NewRef = GetRéférenceProduit(CodeLydic, RéférenceFournisseur)
            Dim r = Me.Ref.CheckUnicityRefProduit(NewRef)
            Return r
        Else
            Return True
        End If
    End Function

#End Region

#Region "TempsDePauseUnitaire (Integer)"
    Private _TempsDePauseUnitaire As Integer

    ''' <summary>Le temps de pause en minutes.</summary>
    Public Property TempsDePauseUnitaire() As Integer
        Get
            Return _TempsDePauseUnitaire
        End Get
        Set(ByVal value As Integer)
            If Object.Equals(value, Me._TempsDePauseUnitaire) Then Exit Property
            _TempsDePauseUnitaire = value
            NotifyPropertyChanged(NameOf(TempsDePauseUnitaire))
        End Set
    End Property
#End Region

#Region "MotsClés (ObservableCollection(of String))"
    Private _MotsClés As List(Of String)
    Public Property MotsClés() As List(Of String)
        Get
            Return _MotsClés
        End Get
        Set(value As List(Of String))
            _MotsClés = value
            NotifyPropertyChanged(NameOf(MotsClés))
        End Set
    End Property
#End Region

#Region "Famille"
    Private _Famille As FamilleDeProduit
    Public Property Famille() As FamilleDeProduit
        Get
            Return _Famille
        End Get
        Set(ByVal value As FamilleDeProduit)
            If Object.Equals(value, Me._Famille) Then Exit Property
            _Famille = value
            NotifyPropertyChanged(NameOf(Famille))
        End Set
    End Property
#End Region

#End Region

#Region "Méthodes"


#End Region

#Region "Tests et debuggage"


#End Region

End Class

Public Enum Unités
    ''' <summary>Ensemble</summary>
    Ens

    ''' <summary>?</summary>
    PM

    ''' <summary>?</summary>
    ML

    ''' <summary>Nombre</summary>
    U
End Enum
