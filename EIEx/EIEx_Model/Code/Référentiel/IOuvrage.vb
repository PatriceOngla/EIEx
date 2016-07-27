Imports System.Collections.ObjectModel

Public Interface IOuvrage

#Region "Propriétés"



    Property Nom() As String

    ReadOnly Property UsagesDeProduit As ObservableCollection(Of UsageDeProduit)

    Property MotsClés() As List(Of String)

    Property TempsDePauseUnitaire() As Integer?

    Property PrixUnitaire() As Single?

#End Region

#Region "Méthodes"

    Function AjouterProduit(P As Produit, Nombre As Short) As UsageDeProduit

    Function UtiliseProduit(p As Produit) As Boolean

#End Region


End Interface
