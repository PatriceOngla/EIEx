Imports System.Windows.Data
<ValueConversion(GetType(List(Of String)), GetType(String))>
Public Class ListString_String_Converter
    Implements IValueConverter

    Public Sub New()
        MyBase.New()
    End Sub

#Region " Propriétés "

#Region "TypeUI"
    Private _TypeUI As Type = GetType(String)
    Public ReadOnly Property TypeUI() As Type
        Get
            Return _TypeUI
        End Get
    End Property
#End Region

#Region "TypeData"
    Private _TypeData As Type = GetType(List(Of String))
    Public ReadOnly Property TypeData() As Type
        Get
            Return _TypeData
        End Get
    End Property
#End Region

#End Region

#Region " Méthodes "

    Private Function MsgErreurType() As String
        Return "Ce converter convertit seulement des " & Me.TypeData.Name & " en " & Me.TypeUI.Name & "."
    End Function

    Public Function Convert(ByVal value As Object, ByVal targetType As System.Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.Convert

        If value IsNot Nothing AndAlso Not Me._TypeData.IsAssignableFrom(value.GetType) Then Throw New ArgumentException(MsgErreurType)

        Dim VDonnee As List(Of String) = value
        Dim vui As String = Nothing

        If VDonnee IsNot Nothing Then
            vui = String.Join(";", VDonnee.ToArray())
        End If
        Return vui

    End Function

    Public Function ConvertBack(ByVal value As Object, ByVal targetType As System.Type, ByVal parameter As Object, ByVal culture As System.Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack

        If value IsNot Nothing AndAlso Not Me.TypeUI.IsAssignableFrom(value.GetType) Then Throw New ArgumentException(MsgErreurType)

        Dim VUI As String = value
        Dim VDonnee = New List(Of String)(VUI.Split(";"))

        Return VDonnee

    End Function

#End Region

End Class