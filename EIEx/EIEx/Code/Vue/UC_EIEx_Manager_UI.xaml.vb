Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports Model

Public Class UC_EIEx_Manager_UI

#Region "Constructeurs"

    Public Sub New()
        ' Cet appel est requis par le concepteur.
        InitializeComponent()
        ' Ajoutez une initialisation quelconque après l'appel InitializeComponent().
        ExcelEventManager.UCSC = Me
    End Sub

#End Region

#Region "Properties"

#Region "Ref"
    Public ReadOnly Property Ref() As Référentiel
        Get
            Return Référentiel.Instance
        End Get
    End Property
#End Region

#Region "WS"
    Public ReadOnly Property WS() As WorkSpace
        Get
            Return WorkSpace.Instance
        End Get
    End Property
#End Region

#Region "EtudeCourante"
    Public ReadOnly Property EtudeCourante() As Etude
        Get
            Return Me.UC_Etude.EtudeCourante
        End Get
    End Property
#End Region

#End Region

#Region "Methods"

#Region "UI event handlers"

#Region "Navigation"

    Private Sub UC_EIEx_Manager_UI_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseDoubleClick
        Try
            TraiterDemandeDeNavigation(e.OriginalSource)
        Catch ex As Exception
            ManageErreur(ex, NameOf(UC_EIEx_Manager_UI_MouseDoubleClick))
        End Try
    End Sub

    Private Sub TraiterDemandeDeNavigation(source As Object)
        Dim FESource = TryCast(source, FrameworkElement)
        If FESource IsNot Nothing Then
            Dim DonnéeNavigationCible = TryCast(FESource.DataContext, IAgregateRoot)
            If DonnéeNavigationCible Is Nothing Then
                Dim CBxParent = GetPrentCombobox(FESource)

                If CBxParent IsNot Nothing Then
                    DonnéeNavigationCible = TryCast(CBxParent.SelectedItem, Entité)
                End If
            End If
            If DonnéeNavigationCible IsNot Nothing Then
                NaviguerVers(DonnéeNavigationCible)
            End If
        End If

    End Sub

    Private Function GetPrentCombobox(fe As FrameworkElement) As ComboBox
        Dim Parent = VisualTreeHelper.GetParent(fe)
        Do While Parent IsNot Nothing AndAlso TypeOf Parent IsNot ComboBox
            Parent = VisualTreeHelper.GetParent(Parent)
        Loop
        Return Parent
    End Function

    Private Sub NaviguerVers(donnéeNavigationCible As Entité)
        If TypeOf donnéeNavigationCible Is UsageDeProduit Then
            donnéeNavigationCible = TryCast(donnéeNavigationCible, UsageDeProduit).Produit
        End If
        If donnéeNavigationCible IsNot Nothing Then
            AccéderALaVueRéférentiel(donnéeNavigationCible)
        End If
    End Sub

    Private Sub AccéderALaVueRéférentiel(Cible As IAgregateRoot)
        Me.TBt_Référentiel.IsSelected = True
        Dim TypeCible = Cible.GetType
        Select Case TypeCible
            Case GetType(Produit)
                AccéderAuProduit(Cible)
            Case Else
                MsgBox($"En route pour {Cible.ToString()}.")
        End Select
    End Sub

    Private Sub AccéderAuProduit(P As Produit)
        With Me.UC_RéférentielView
            .TBt_Produits.IsSelected = True
            .UC_ProduitsView.ProduitCourant = P
        End With

    End Sub

#End Region

#End Region

#End Region

#Region "Tests and debug"

#Region "SelectedRange"

    Public Property SelectedRange As String ' Excel.Range
        Get
            Return GetValue(SelectedRangeProperty)
        End Get

        Set(ByVal value As String)
            SetValue(SelectedRangeProperty, value)
        End Set
    End Property

    Public Shared ReadOnly SelectedRangeProperty As DependencyProperty =
                           DependencyProperty.Register("SelectedRange",
                           GetType(String), GetType(UC_EIEx_Manager_UI),
                           New PropertyMetadata(Nothing))

#End Region

    Private Sub UC_SubContainer_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseRightButtonUp

    End Sub

#End Region

End Class
