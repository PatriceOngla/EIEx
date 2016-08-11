Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Windows
Imports Model
Imports Utils

Public Class Ouvrage
    Inherits Ouvrage_Base
    Implements IEntitéDuWorkSpace

#Region "Constructeurs"

    Public Sub New(Bordereau As Bordereau, NumLignePlageExcel As Integer, modèle As PatronDOuvrage)
        Me.New(Bordereau, NumLignePlageExcel)
        Me.Modèle = modèle
    End Sub
    Public Sub New(Bordereau As Bordereau, NumLignePlageExcel As Integer)
        Me.BordereauParent = Bordereau
        Me.NuméroLignePlageExcel = NumLignePlageExcel
    End Sub

    Protected Overrides Sub Init()
        MyBase.Init()
    End Sub

#End Region

#Region "Propriétés"

#Region "EstRoot"
    Public Overrides ReadOnly Property EstRoot() As Boolean
        Get
            Return False
        End Get
    End Property
#End Region

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

#Region "NuméroLignePlageExcel"
    Public ReadOnly Property NuméroLignePlageExcel() As Integer
#End Region

#Region "Modèle"
    Public ReadOnly Property Modèle() As PatronDOuvrage
#End Region

#End Region

#Region "Méthodes"

    Protected Overrides Sub Copier_Ex(Modèle As Ouvrage_Base)
        Me.ComplémentDeNom = "?"
    End Sub

#End Region

#Region "Tests et debuggage"


#End Region

End Class
