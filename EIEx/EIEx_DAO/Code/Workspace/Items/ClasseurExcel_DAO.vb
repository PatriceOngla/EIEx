Imports System.Xml.Serialization
Imports Model
Imports Utils

<Serializable>
Public Class ClasseurExcel_DAO
    Inherits SystèmesItems_DAO(Of ClasseurExcel)

#Region "Constructeurs"

    Public Sub New()
    End Sub

    Public Sub New(C As ClasseurExcel)
        MyBase.New(C)

        Me.CheminFichier = C.CheminFichier

        Dim SBordereaux = From b In C.Bordereaux Select New Bordereau_DAO(b)
        Me.Bordereaux = New List(Of Bordereau_DAO)(SBordereaux)
        Me.MêmeStructurePourTousLesBordereaux = C.MêmeStructurePourTousLesBordereaux

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
    Public Property CheminFichier() As String

    Public Property Bordereaux() As List(Of Bordereau_DAO)

    <XmlAttribute>
    Public Property MêmeStructurePourTousLesBordereaux As Boolean

#End Region

#End Region

#Region "Méthodes"

    Protected Overrides Sub UnSerialized_Ex(NouveauClasseur As ClasseurExcel)

        'Dim r = Parent.AjouterNouveauClasseur(Me.CheminFichier)
        NouveauClasseur.MêmeStructurePourTousLesBordereaux = Me.MêmeStructurePourTousLesBordereaux

        Me.Bordereaux.DoForAll(Sub(B As Bordereau_DAO)
                                   Dim NvoBordereau = NouveauClasseur.AjouterNouveauBordereau()
                                   B.UnSerialized(NvoBordereau)
                               End Sub)
    End Sub

#End Region

#Region "Tests et debuggage"


#End Region

End Class