﻿Imports System.Xml.Serialization

<Serializable>
Public Class FamilleDeProduit_DAO
    Inherits AgregateRoot_DAO(Of FamilleDeProduit)

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(F As FamilleDeProduit)
        MyBase.New(F)
        Me.Marge = F.Marge
    End Sub

#End Region

#Region "Propriétés"

#Region "Sys"
    Private Ref As Référentiel = Référentiel.Instance
    <XmlIgnore>
    Protected Overrides ReadOnly Property Sys As Système
        Get
            Return Ref
        End Get
    End Property
#End Region

#Region "Données"

    Public Property Marge() As Single?

#End Region

#End Region

#Region "Méthodes"

    Protected Overrides Function UnSerialized_Ex_Ex() As FamilleDeProduit
        Dim r = Ref.GetNewFamilleDeProduit(Me.Id)
        If r Is Nothing Then
            MsgBox("Qué passa ?")
        End If
        'r = If(r, New FamilleDeProduit(Me.Id))
        r.Marge = Marge
        Return r
    End Function

#End Region

#Region "Tests et debuggage"


#End Region

End Class
