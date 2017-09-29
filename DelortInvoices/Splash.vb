Imports System
Imports System.Drawing

Public Class Splash

    Inherits System.Windows.Forms.Form
    Private loaded As Boolean = False

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
    Friend WithEvents grpSplash As System.Windows.Forms.GroupBox
    Friend WithEvents lblWelcome As System.Windows.Forms.Label
    Friend WithEvents UltraPictureBox1 As Infragistics.Win.UltraWinEditors.UltraPictureBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Splash))
        Me.grpSplash = New System.Windows.Forms.GroupBox
        Me.lblWelcome = New System.Windows.Forms.Label
        Me.UltraPictureBox1 = New Infragistics.Win.UltraWinEditors.UltraPictureBox
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.grpSplash.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpSplash
        '
        Me.grpSplash.BackColor = System.Drawing.Color.White
        Me.grpSplash.Controls.Add(Me.lblWelcome)
        Me.grpSplash.Controls.Add(Me.UltraPictureBox1)
        Me.grpSplash.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.grpSplash.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpSplash.ForeColor = System.Drawing.Color.MidnightBlue
        Me.grpSplash.Location = New System.Drawing.Point(16, 11)
        Me.grpSplash.Name = "grpSplash"
        Me.grpSplash.Size = New System.Drawing.Size(369, 301)
        Me.grpSplash.TabIndex = 2
        Me.grpSplash.TabStop = False
        Me.grpSplash.Text = "Homeowner Services Business Software"
        '
        'lblWelcome
        '
        Me.lblWelcome.BackColor = System.Drawing.Color.Transparent
        Me.lblWelcome.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWelcome.ForeColor = System.Drawing.Color.MidnightBlue
        Me.lblWelcome.Location = New System.Drawing.Point(84, 271)
        Me.lblWelcome.Name = "lblWelcome"
        Me.lblWelcome.Size = New System.Drawing.Size(211, 23)
        Me.lblWelcome.TabIndex = 3
        Me.lblWelcome.Text = "John Delort and Sons"
        '
        'UltraPictureBox1
        '
        Me.UltraPictureBox1.BackColor = System.Drawing.Color.White
        Me.UltraPictureBox1.BorderShadowColor = System.Drawing.Color.Empty
        Me.UltraPictureBox1.Image = CType(resources.GetObject("UltraPictureBox1.Image"), Object)
        Me.UltraPictureBox1.Location = New System.Drawing.Point(74, 25)
        Me.UltraPictureBox1.Name = "UltraPictureBox1"
        Me.UltraPictureBox1.Size = New System.Drawing.Size(217, 232)
        Me.UltraPictureBox1.TabIndex = 2
        '
        'Timer1
        '
        '
        'Splash
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(400, 328)
        Me.Controls.Add(Me.grpSplash)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Splash"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Splash"
        Me.grpSplash.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Splash_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Me.Opacity = 0.0
            Me.Timer1.Interval = 50
            Me.Timer1.Start()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        If loaded Then
            'The form is being faded out.
            Me.Opacity -= 0.05

            Me.Timer1.Interval = 50 'In case we have just finished pausing.

            If Me.Opacity <= 0.0 Then
                Me.Close()
            End If
        Else
            'The form is being faded in.
            Me.Opacity += 0.05

            If Me.Opacity >= 1.0 Then
                Me.loaded = True

                'Pause for 10 seconds.
                Me.Timer1.Interval = 3000
            End If
        End If

    End Sub
End Class
