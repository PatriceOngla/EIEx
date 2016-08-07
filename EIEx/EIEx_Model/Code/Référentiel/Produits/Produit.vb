Imports System.Collections.ObjectModel
Imports Model
Imports Utils

Public Class Produit
    Inherits Entité
    Implements IAgregateRoot, IEntitéDuRéférentiel

#Region "Constructeurs"

    Public Sub New(Id As Integer)
        Me.Id = Id
    End Sub

    Protected Overrides Sub Init()
        Me.Nom = "Nouveau produit"
        _MotsClés = New List(Of String)
    End Sub

#End Region

#Region "Propriétés"

#Region "Système"

    Public ReadOnly Property Ref As Référentiel Implements IEntitéDuRéférentiel.Ref
        Get
            Return Référentiel.Instance
        End Get
    End Property

    Public Overrides ReadOnly Property Système As Système
        Get
            Return Ref
        End Get
    End Property

#End Region

#Region "Id"
    Public ReadOnly Property Id() As Integer? Implements IAgregateRoot.Id
#End Region

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

#Region "TempsDePauseUnitaire (Single)"
    Private _TempsDePauseUnitaire As Single

    ''' <summary>Le temps de pause en minutes.</summary>
    Public Property TempsDePauseUnitaire() As Single
        Get
            Return _TempsDePauseUnitaire
        End Get
        Set(ByVal value As Single)
            If Object.Equals(value, Me._TempsDePauseUnitaire) Then Exit Property
            _TempsDePauseUnitaire = value
            NotifyPropertyChanged(NameOf(TempsDePauseUnitaire))
        End Set
    End Property
#End Region

#Region "MotsClés (ObservableCollection(of String))"
    Private WithEvents _MotsClés As List(Of String)
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

#Region "Propriétés pour recherche"

#Region "Mots"
    Private _Mots As List(Of String)

    ''' <summary>Les mots du <see cref="Nom"/> + ceux des <see cref="MotsClés"/> (pour les recherches). 
    ''' Attention, il doivent être mis à jour avant toute recherche avec la méthode <see cref="SetMotsPourTousLesProduits()"/>.</summary>
    Public ReadOnly Property Mots() As IEnumerable(Of String)
        Get
            'If _Mots Is Nothing Then SetMots()
            Return _Mots
        End Get
    End Property

    Public Sub SetMots()
        _Mots = New List(Of String)(MotsClés)
        _Mots.AddRange(Nom.Split({" "c, "'"c}))
    End Sub

    Public Shared Sub SetMotsPourTousLesProduits()
        Référentiel.Instance.Produits.DoForAll(Sub(p As Produit) p.SetMots())
    End Sub

#End Region

#Region "ToString pour affichage en list (colonnage fixe)"

    Public ReadOnly Property ToStringForListDisplay() As String
        Get
            Dim r = DisplayWithFixedColumn(Me.Id, Me.RéférenceProduit, Me.Nom, Me.MotsClés)
            Return r
        End Get
    End Property

    Public Shared ReadOnly Property ProductsListHeader() As String
        Get
            Dim r = DisplayWithFixedColumn("Id", "Référence", "Nom", "Mots-clés")
            Return r
        End Get
    End Property

    Private Shared Function DisplayWithFixedColumn(Id As String, référenceProduit As String, nom As String, motsClés As List(Of String)) As String
        Dim r = DisplayWithFixedColumn(Id, référenceProduit, nom, Join(motsClés.ToArray(), ", "))
        Return r
    End Function

    Private Shared Function DisplayWithFixedColumn(Id As String, référenceProduit As String, nom As String, motsClés As String) As String
        Dim r As String = FormateForColumn(Id, 5) & FormateForColumn(référenceProduit, 10) & FormateForColumn(nom, 100) & FormateForColumn(motsClés, 25, False)
        Return r
    End Function

#End Region

#End Region

#End Region

#Region "Méthodes"

    Public Overrides Function ToString() As String
        Return Me.ToStringForAgregateRoot(MyBase.ToString())
    End Function

#End Region

#Region "Tests et debuggage"


#End Region

End Class

''' <summary>
''' On ignore CIS (= compris), PM (= pour mémoire) et hl (= ?)
''' </summary>
Public Enum Unités
    ''' <summary>Ensemble</summary>
    Ens

    ''' <summary>Mètres linéaires</summary>
    ML

    ''' <summary>Nombre</summary>
    U
End Enum
