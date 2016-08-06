Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports Model
Imports Utils

''' <summary>
''' On distingue <see cref="PatronDOuvrage"/> et Ouvrage. Les ouvrages sont les entrées des bordereau et sont associé à des <see cref="PatronDOuvrage"/> afin de calculer leur prix sur la base du <see cref="PatronDOuvrage.PrixUnitaire"/>. 
''' </summary>
Public Class PatronDOuvrage
    Inherits Ouvrage_Base
    Implements IAgregateRoot, IEntitéDuRéférentiel

#Region "Constructeurs"

    Public Sub New(Id As Integer)
        Me.Id = Id
    End Sub

    Protected Overrides Sub Init()
        MyBase.Init()
    End Sub

#End Region

#Region "Propriétés"

#Region "EstRoot"
    Public Overrides ReadOnly Property EstRoot() As Boolean
        Get
            Return True
        End Get
    End Property
#End Region

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

    '#Region "Gestion du nommage"

    '#Region "Nom (String)"
    '    Private _Nom As String
    '    Public Overrides Property Nom() As String
    '        Get
    '            Return _Nom
    '        End Get
    '        Set(ByVal value As String)
    '            If Object.Equals(value, Me._Nom) Then Exit Property
    '            _Nom = value
    '            NotifyPropertyChanged(NameOf(Nom))
    '            If Not (Me.Libellés.Contains(value)) Then Me.Libellés.Add(value)
    '        End Set
    '    End Property
    '#End Region

    '#Region "ComplémentDeNom"
    '    Private _ComplémentDeNom As String
    '    Public Property ComplémentDeNom() As String
    '        Get
    '            Return _ComplémentDeNom
    '        End Get
    '        Set(ByVal value As String)
    '            If Object.Equals(value, Me._ComplémentDeNom) Then Exit Property
    '            _ComplémentDeNom = value
    '            NotifyPropertyChanged(NameOf(ComplémentDeNom))
    '        End Set
    '    End Property
    '#End Region

    '#Region "NomComplet"
    '    ''' <summary>Le nom saisi + le complément de nom s'il y a en a un. </summary>
    '    Public ReadOnly Property NomComplet() As String
    '        Get
    '            Return Me.Nom & If(String.IsNullOrEmpty(ComplémentDeNom), "", " - " & ComplémentDeNom)
    '        End Get
    '    End Property
    '#End Region

    '#Region "Libellés"

    '    Private WithEvents _Libellés As ObservableCollection(Of String)
    '    Public ReadOnly Property Libellés() As ObservableCollection(Of String)
    '        Get
    '            Return _Libellés
    '        End Get
    '    End Property

    '    Public ReadOnly Property NbLibellés() As Integer
    '        Get
    '            Return Me.Libellés?.Count
    '        End Get
    '    End Property

    '    Private Sub _Libellés_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs) Handles _Libellés.CollectionChanged
    '        ForcerLaCohérenceEntreLibelléPrincipalEtLaCollection()
    '        Me.NotifyPropertyChanged(NameOf(NbLibellés))

    '    End Sub

    '    Private Sub ForcerLaCohérenceEntreLibelléPrincipalEtLaCollection()
    '        If Not (String.IsNullOrEmpty(Me.Nom) OrElse Me.Libellés.Contains(Me.Nom)) Then Me.Libellés.Add(Me.Nom)
    '    End Sub

    '#End Region

    '#End Region

#End Region

#Region "Méthodes"

    Protected Overrides Sub Copier_Ex(Modèle As Ouvrage_Base)
        Me.Nom = Modèle.Nom & " (copie)"
    End Sub

#Region "ToString"
    Public Overrides Function ToString() As String
        Return Me.ToStringForAgregateRoot(MyBase.ToString())
    End Function
#End Region

#End Region

#Region "Tests et debuggage"


#End Region

End Class
