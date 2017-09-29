Public Class frmBaseWindow
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBaseWindow))
        Me.SuspendLayout()
        '
        'frmBaseWindow
        '
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(292, 273)
        Me.Name = "frmBaseWindow"
        Me.Text = "frmBaseWindow"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub TextEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        DirectCast(sender, TextBox).BackColor = Color.Wheat
    End Sub

    Private Sub TextLeave(ByVal sender As Object, ByVal e As System.EventArgs)
        DirectCast(sender, TextBox).BackColor = Color.FromKnownColor(KnownColor.Window)
    End Sub

    Private Sub SetEventHandlers(ByVal ctrlContainer As Control)

        For Each ctrl As Control In ctrlContainer.Controls
            If TypeOf ctrl Is TextBox Then
                AddHandler ctrl.Enter, _
                   AddressOf TextEnter
                AddHandler ctrl.Leave, _
                   AddressOf TextLeave
            End If

            'recursive call
            If ctrl.HasChildren Then
                SetEventHandlers(ctrl)
            End If
        Next
    End Sub

    Private Sub frmBaseWindow_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.SetEventHandlers(Me)

    End Sub
End Class
