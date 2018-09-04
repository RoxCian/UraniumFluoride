Imports System.Windows.Input

Public Class WattingWindow
    Public ReadOnly Property Canceling As Boolean = False
    Public ReadOnly Property Info As String
        Get
            Return Me.GetValue(InfoProperty)
        End Get
    End Property
    Public Shared ReadOnly InfoProperty As Windows.DependencyProperty = Windows.DependencyProperty.Register("Info", GetType(String), GetType(WattingWindow), New Windows.PropertyMetadata("s"))
    Public Sub New()

        ' 此调用是设计器所必需的。
        InitializeComponent()

        ' 在 InitializeComponent() 调用之后添加任何初始化。

    End Sub
    Public Sub New(info As String)

        ' 此调用是设计器所必需的。
        InitializeComponent()

        ' 在 InitializeComponent() 调用之后添加任何初始化。
        Me.SetValue(InfoProperty, info)
    End Sub

    Private Sub WattingWindow_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        If e.Key = Key.Escape Then
            Me._Canceling = True
            Me.Dispatcher.InvokeShutdown()
            Me.Close()
        End If
    End Sub
End Class
