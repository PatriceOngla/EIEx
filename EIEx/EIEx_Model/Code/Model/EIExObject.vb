Imports System.ComponentModel
Imports System.Windows

Public MustInherit Class EIExObject
    Implements INotifyPropertyChanged

#Region "Constructeurs"

    Public Sub New()
        Init()
    End Sub

    ''' <summary>
    ''' Intialiser l'objet. A la charge des sous-classes. 
    ''' </summary>
    Protected MustOverride Sub Init()

#End Region

#Region "Propriétés"

#Region "Réf (Référentiel)"
    Public ReadOnly Property Réf() As Référentiel
        Get
            Return Référentiel.Instance
        End Get
    End Property
#End Region

#Region "Nom (String)"
    Private _Nom As String
    Public Overridable Property Nom() As String
        Get
            Return _Nom
        End Get
        Set(ByVal value As String)
            If Object.Equals(value, Me._Nom) Then Exit Property
            _Nom = value
            NotifyPropertyChanged(NameOf(Nom))
        End Set
    End Property
#End Region

#End Region

#Region "Méthodes"

#Region "NotifyPropertyChanged"

    Protected Sub NotifyPropertyChanged(PropertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(PropertyName))
    End Sub

#End Region

#Region "TosTring"
    Public Overrides Function ToString() As String
        Try
            Dim r = $"{Me.GetType.Name} {If(String.IsNullOrEmpty(Me.Nom), "", "'" & Me.Nom & "'")}"
            Return r
        Catch ex As Exception
            Utils.ManageError(ex, NameOf(ToString))
            Return MyBase.ToString()
        End Try
    End Function
#End Region

#End Region

#Region "Evennements"

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

#End Region

#Region "Tests et debuggage"


#End Region

End Class
