﻿Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Windows
Imports Model
Imports Utils

Public Class Ouvrage
    Inherits Ouvrage_Base
    Implements IEntitéDuWorkSpace

#Region "Constructeurs"

    Public Sub New(Bordereau As Bordereau)
        Me.BordereauParent = Bordereau
    End Sub

    Protected Overrides Sub Init()
        MyBase.Init()
    End Sub

#End Region

#Region "Propriétés"

#Region "Système"

    Public ReadOnly Property WS As WorkSpace Implements IEntitéDuWorkSpace.WS
        Get
            Return WorkSpace.Instance
        End Get
    End Property

    Public Overrides ReadOnly Property Système As Système
        Get
            Return Me.WS
        End Get
    End Property
#End Region

#Region "BordereauParent"
    Public ReadOnly Property BordereauParent() As Bordereau
#End Region

#End Region

#Region "Méthodes"

#End Region

#Region "Tests et debuggage"


#End Region

End Class
