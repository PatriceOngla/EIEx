Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports Model

'Attention : pas proprement implémenté comme un singleton mais une seule instance doit être créée (accessible via la propriété shared <seealso cref="Win_EIEx_Manager_UI.Instance"/>).
Public Class Win_Main

#Region "Constructeurs"

    Public Sub New()
        ' Cet appel est requis par le concepteur.
        _Instance = Me

        InitializeComponent()

        ExcelEventManager.UCSC = Me
    End Sub

#End Region

#Region "Properties"

#Region "Instance"
    Private Shared _Instance As Win_Main
    Public Shared ReadOnly Property Instance() As Win_Main
        Get
            Return _Instance
        End Get
    End Property
#End Region

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
                Dim CBxParent = GetParentCombobox(FESource)

                If CBxParent IsNot Nothing Then
                    DonnéeNavigationCible = TryCast(CBxParent.SelectedItem, Entité)
                End If
            End If
            If DonnéeNavigationCible IsNot Nothing Then
                NaviguerVers(DonnéeNavigationCible)
            End If
        End If

    End Sub

    Private Function GetParentCombobox(fe As FrameworkElement) As ComboBox
        'Dim Parent = VisualTreeHelper.GetParent(fe)
        'Do While Parent IsNot Nothing AndAlso TypeOf Parent IsNot ComboBox
        '    Parent = VisualTreeHelper.GetParent(Parent)
        'Loop
        'Return Parent
        Return GetParentControl(Of ComboBox)(fe)
    End Function

    Friend Sub NaviguerVers(donnéeNavigationCible As Entité)
        If TypeOf donnéeNavigationCible Is UsageDeProduit Then
            donnéeNavigationCible = TryCast(donnéeNavigationCible, UsageDeProduit).Produit
        End If
        If TypeOf donnéeNavigationCible Is IAgregateRoot Then
            AccéderALaVueRéférentiel(donnéeNavigationCible)
        End If
    End Sub

    Private Sub AccéderALaVueRéférentiel(Cible As IAgregateRoot)
        Me.TBt_Référentiel.IsSelected = True
        Dim TypeCible = Cible.GetType
        Select Case TypeCible
            Case GetType(Produit)
                AccéderAuProduit(Cible)
            Case GetType(PatronDOuvrage)
                AccéderAuPatronDOuvrage(Cible)
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

    Private Sub AccéderAuPatronDOuvrage(PO As PatronDOuvrage)
        With Me.UC_RéférentielView
            .TBt_PatronsDOuvrage.IsSelected = True
            .UC_OuvragesView.OuvrageCourant = PO
        End With
    End Sub

#End Region

#Region "Gestion des recherches"

    Private Sub UC_EIEx_Manager_UI_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        Try
            If e.KeyboardDevice.Modifiers = ModifierKeys.Control Then
                If e.Key = Key.F OrElse Keyboard.IsKeyDown(Key.F) Then
                    If e.Key = Key.P OrElse Keyboard.IsKeyDown(Key.P) Then
                        Win_SélecteurDeProduit.Cherche()
                    ElseIf e.Key = Key.O OrElse Keyboard.IsKeyDown(Key.O) Then
                        Win_SélecteurDOuvrage.Cherche()
                    End If
                End If
            End If
        Catch ex As Exception
            ManageErreur(ex)
        End Try
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
                           GetType(String), GetType(Win_Main),
                           New PropertyMetadata(Nothing))

#End Region

    Private Sub UC_SubContainer_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseRightButtonUp

    End Sub

#End Region

End Class
