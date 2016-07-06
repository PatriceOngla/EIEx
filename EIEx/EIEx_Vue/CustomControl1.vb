Public Class CustomControl1
    Inherits Control

    Shared Sub New()
        'Cet appel OverrideMetadata indique au système que cet élément souhaite apporter un style différent de celui de sa classe de base.
        'Ce style est défini dans Themes\Generic.xaml
        DefaultStyleKeyProperty.OverrideMetadata(GetType(CustomControl1), New FrameworkPropertyMetadata(GetType(CustomControl1)))
    End Sub

End Class
