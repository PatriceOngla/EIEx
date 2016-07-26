Imports System.Xml.Serialization
Imports Model
Imports Utils

<Serializable>
Public Class Etude_DAO
    Inherits AgregateRoot_DAO(Of Etude)

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(E As Etude)
        MyBase.New(E)

        Me.EstOuverte = E.EstOuverte

        Dim ClasseursExcel = From c In E.ClasseursExcel Select New ClasseurExcel_DAO(c)
        Me.ClasseursExcel = New List(Of ClasseurExcel_DAO)(ClasseursExcel)

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

    <XmlAttribute>
    Public Property EstOuverte() As Boolean

    Public Property ClasseursExcel() As List(Of ClasseurExcel_DAO)


#End Region

#End Region

#Region "Méthodes"

    Protected Overrides Sub UnSerialized_Ex_Ex(NouvelleEtude As Etude)
        'Dim r = WS.GetNewEtude(Me.Id)

        NouvelleEtude.EstOuverte = Me.EstOuverte

        'Dim Classeurs = (From c In Me.ClasseursExcel Select c.UnSerialized())
        'r.ClasseursExcel.AddRange(Classeurs)

        Me.ClasseursExcel.DoForAll(Sub(CDAO As ClasseurExcel_DAO)
                                       Dim NouveauClasseur = NouvelleEtude.AjouterNouveauClasseur(CDAO.CheminFichier)
                                       CDAO.UnSerialized(NouveauClasseur)
                                   End Sub)
        'Return r
    End Sub

#End Region

#Region "Tests et debuggage"


#End Region

End Class