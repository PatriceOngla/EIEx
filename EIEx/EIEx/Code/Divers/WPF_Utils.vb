Imports System.Windows
Imports System.Windows.Media

Module WPF_Utils


    Public Class AlwaysOnTopBehavior
        '      Inherits Behavior(Of Window)

        '      Protected Overrides Sub OnAttached()
        '          MyBase.OnAttached()

        '          AssociatedObject.LostFocus += (s, e) => AssociatedObject.Topmost = true

        'End Sub

    End Class

    Public Function GetParentControl(Of T As FrameworkElement)(fe As FrameworkElement) As T
        Dim Parent = VisualTreeHelper.GetParent(fe)
        Do While Parent IsNot Nothing AndAlso TypeOf Parent IsNot T
            Parent = VisualTreeHelper.GetParent(Parent)
        Loop
        Return Parent
    End Function

End Module
