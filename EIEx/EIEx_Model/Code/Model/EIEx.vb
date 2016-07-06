Imports System.Collections.ObjectModel
Imports System.Windows

Public Class EIEx
    Inherits EIExObject

#Region "Constructeurs"


#End Region

#Region "Propriétés"

#Region "Referentiel (Référentiel)"
    Private _Referentiel As Référentiel
    Public ReadOnly Property Referentiel() As Référentiel
        Get
            Return _Referentiel
        End Get
    End Property
#End Region

#Region "Bordereaux"
    Private _Bordereaux As ObservableCollection(Of Bordereau_DAO)
    Public ReadOnly Property Bordereaux() As ObservableCollection(Of Bordereau_DAO)
        Get
            Return _Bordereaux
        End Get
    End Property
#End Region

#End Region

#Region "Méthodes"


#End Region

#Region "Tests et debuggage"


#End Region

End Class
