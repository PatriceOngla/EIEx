﻿Imports System.Xml.Serialization
Imports Model
Imports Utils

Public Class Workspace_DAO
    Inherits Système_DAO(Of WorkSpace)

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(Modèle As WorkSpace)
        MyBase.New(Modèle)
        Dim Etudes = From e In Modèle.Etudes Select New Etude_DAO(e)
        Me.Etudes = New List(Of Etude_DAO)(Etudes)
    End Sub

#End Region

#Region "Propriétés"

#Region "Sys"
    Private WS As WorkSpace = WorkSpace.Instance
    <XmlIgnore>
    Protected Overrides ReadOnly Property Sys As Système
        Get
            Return WS
        End Get
    End Property
#End Region

#Region "Données"

    Public Property Etudes() As List(Of Etude_DAO)

#End Region

#End Region

#Region "Méthodes"

    Protected Overrides Sub Reload_Ex(NewT As WorkSpace)
        Me.Etudes.DoForAll(Sub(e)
                               Dim NewEtude = WS.GetNewEtude(e.Id)
                               e.UnSerialized(NewEtude)
                           End Sub)
    End Sub


#End Region

End Class
