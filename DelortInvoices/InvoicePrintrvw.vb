Option Strict On
Public Class ActRptPrvw
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
    Public WithEvents ARVeiwer As DataDynamics.ActiveReports.Viewer.Viewer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ARVeiwer = New DataDynamics.ActiveReports.Viewer.Viewer
        Me.SuspendLayout()
        '
        'ARVeiwer
        '
        Me.ARVeiwer.BackColor = System.Drawing.SystemColors.Control
        Me.ARVeiwer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ARVeiwer.Location = New System.Drawing.Point(0, 0)
        Me.ARVeiwer.Name = "ARVeiwer"
        Me.ARVeiwer.ReportViewer.CurrentPage = 0
        Me.ARVeiwer.ReportViewer.MultiplePageCols = 3
        Me.ARVeiwer.ReportViewer.MultiplePageRows = 1
        Me.ARVeiwer.Size = New System.Drawing.Size(692, 473)
        Me.ARVeiwer.TabIndex = 6
        Me.ARVeiwer.TableOfContents.Enabled = False
        Me.ARVeiwer.TableOfContents.Text = "Contents"
        Me.ARVeiwer.TableOfContents.Width = 200
        Me.ARVeiwer.Toolbar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'ActRptPrvw
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(692, 473)
        Me.Controls.Add(Me.ARVeiwer)
        Me.Name = "ActRptPrvw"
        Me.Text = "Invoice Print Preview"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub ActRptPrvw_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.MinimizeBox = False

    End Sub
End Class
