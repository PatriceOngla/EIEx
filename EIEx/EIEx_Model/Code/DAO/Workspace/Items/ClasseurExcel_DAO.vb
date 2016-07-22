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

    Public Property MêmeStructurePourTousLesBordereaux As Boolean

#End Region

#End Region

#Region "Méthodes"

    Protected Overrides Function UnSerialized_Ex() As ClasseurExcel

        Dim r = New ClasseurExcel(Me.CheminFichier)

        r.MêmeStructurePourTousLesBordereaux = Me.MêmeStructurePourTousLesBordereaux

        Dim Bdx = (From b In Me.Bordereaux Select NewBordereau(r, b))
        r.Bordereaux.AddRange(Bdx)

        Return r

    End Function

    Private Function NewBordereau(c As ClasseurExcel, b As Bordereau_DAO) As Bordereau
        Dim r = b.UnSerialized()
        r.Parent = c
        Return r
    End Function

#End Region

#Region "Tests et debuggage"


#End Region

End Class