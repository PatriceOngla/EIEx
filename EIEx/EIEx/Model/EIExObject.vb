Imports System.ComponentModel
Imports System.Windows

Public Class EIExObject
    Inherits DependencyObject
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Protected Sub NotifyPropertyChanged(PropertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(PropertyName))
    End Sub

End Class
