Imports System.Xml.Serialization

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

    Protected Overrides Sub UnSerialized_Ex_Ex(nouvelleEntité As FamilleDeProduit)
        'TODO: la double extension ex_ex est-elle toujours utile ? 
        'Dim r = Ref.GetNewFamilleDeProduit(Me.Id)

        'TODO: vérifier pourquoi on faisait ce test. 
        'If (nouvelleEntité Is Nothing) Then
        '    nouvelleEntité = New FamilleDeProduit(Me.Id)
        'End If
        nouvelleEntité.Marge = Marge
    End Sub


#End Region

#Region "Tests et debuggage"


#End Region

End Class
