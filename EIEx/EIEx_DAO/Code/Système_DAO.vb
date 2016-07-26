Imports System.Runtime.Serialization
Imports System.Xml.Serialization
Imports Model

Public MustInherit Class Système_DAO(Of T As Système)
    Inherits EIEx_Object_DAO(Of T)
    Implements ISystèmeDAO

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(Modèle As T)
        MyBase.New(Modèle)
        Me.DateModif = Modèle.DateModif
    End Sub

#End Region

#Region "Propriétés"

    Public Property DateModif() As Date Implements ISystèmeDAO.DateModif

#End Region

#Region "Méthodes"

    Public Sub ReLoad(NewT As T)
        NewT.Purger()
        NewT.DateModif = Me.DateModif
        NewT.Nom = Me.Nom
        NewT.Commentaires = Me.Commentaires
        Reload_Ex(NewT)
    End Sub

    Protected MustOverride Sub Reload_Ex(NewT As T)

    Public Sub UnSerialize2(NewT As Système) Implements ISystèmeDAO.UnSerialize
        ReLoad(NewT)
    End Sub

#Region "GetDAO"
    Private Shared Function GetConvenientDAO(S As Système) As ISystèmeDAO
        Dim TypeSystème = S.GetType()
        Select Case TypeSystème
            Case GetType(WorkSpace)
                Return New Workspace_DAO(S)
            Case GetType(Référentiel)
                Return New Référentiel_DAO(S)
            Case Else
                Throw New NotSupportedException($"Aucun DAO n'est défini pour le système ""{TypeSystème.Name}"".")
        End Select
    End Function
#End Region

    Public Sub Enregistrer(Chemin As String)
        Utils.Sérialiser(Me, Chemin)
    End Sub

#End Region

End Class
