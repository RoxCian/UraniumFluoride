Imports System.Windows.Input

Public Class WattingWindow
    Public ReadOnly Property Info As String
        Get
            Return Me.GetValue(InfoProperty)
        End Get
    End Property
    Public Shared ReadOnly InfoProperty As Windows.DependencyProperty = Windows.DependencyProperty.Register("Info", GetType(String), GetType(WattingWindow), New Windows.PropertyMetadata("s"))
    Public Sub New()

        InitializeComponent()

    End Sub
    Public Sub New(info As String)

        InitializeComponent()

        Me.SetValue(InfoProperty, info)
    End Sub

End Class
