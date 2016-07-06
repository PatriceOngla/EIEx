Imports System.ComponentModel
Imports System.Windows

Public Class EIExObject
    Implements INotifyPropertyChanged

#Region "Constructeurs"

    Public Sub New()

    End Sub

    Public Sub New(Id As Integer)
        Me.Id = Id
    End Sub

#End Region

#Region "Propriétés"

#Region "Id (Integer)"
    Public ReadOnly Property Id() As Integer?
#End Region

#Region "Nom (String)"
    Private _Nom As String
    Public Property Nom() As String
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

#End Region

#Region "Evennements"

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

#End Region

#Region "Tests et debuggage"


#End Region

End Class
