Option Strict On

Public Class frmOpenOneInvoice
    Inherits frmBaseWindow

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
    Friend WithEvents grpMain As System.Windows.Forms.GroupBox
    Friend WithEvents txtInvNo As System.Windows.Forms.TextBox
    Friend WithEvents lblEnterInv As System.Windows.Forms.Label
    Friend WithEvents grpFooter As System.Windows.Forms.GroupBox
    Friend WithEvents btnGo As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOpenOneInvoice))
        Me.grpMain = New System.Windows.Forms.GroupBox()
        Me.lblEnterInv = New System.Windows.Forms.Label()
        Me.txtInvNo = New System.Windows.Forms.TextBox()
        Me.grpFooter = New System.Windows.Forms.GroupBox()
        Me.btnGo = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.grpMain.SuspendLayout()
        Me.grpFooter.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpMain
        '
        Me.grpMain.BackColor = System.Drawing.Color.Transparent
        Me.grpMain.Controls.Add(Me.lblEnterInv)
        Me.grpMain.Controls.Add(Me.txtInvNo)
        Me.grpMain.Location = New System.Drawing.Point(13, 7)
        Me.grpMain.Name = "grpMain"
        Me.grpMain.Size = New System.Drawing.Size(182, 41)
        Me.grpMain.TabIndex = 1
        Me.grpMain.TabStop = False
        '
        'lblEnterInv
        '
        Me.lblEnterInv.BackColor = System.Drawing.Color.Transparent
        Me.lblEnterInv.Location = New System.Drawing.Point(16, 16)
        Me.lblEnterInv.Name = "lblEnterInv"
        Me.lblEnterInv.Size = New System.Drawing.Size(86, 16)
        Me.lblEnterInv.TabIndex = 3
        Me.lblEnterInv.Text = "Invoice No.:"
        '
        'txtInvNo
        '
        Me.txtInvNo.Location = New System.Drawing.Point(108, 14)
        Me.txtInvNo.Name = "txtInvNo"
        Me.txtInvNo.Size = New System.Drawing.Size(56, 22)
        Me.txtInvNo.TabIndex = 2
        '
        'grpFooter
        '
        Me.grpFooter.BackColor = System.Drawing.Color.Transparent
        Me.grpFooter.Controls.Add(Me.btnGo)
        Me.grpFooter.Controls.Add(Me.btnExit)
        Me.grpFooter.Location = New System.Drawing.Point(13, 47)
        Me.grpFooter.Name = "grpFooter"
        Me.grpFooter.Size = New System.Drawing.Size(182, 41)
        Me.grpFooter.TabIndex = 2
        Me.grpFooter.TabStop = False
        '
        'btnGo
        '
        Me.btnGo.BackColor = System.Drawing.Color.AliceBlue
        Me.btnGo.Enabled = False
        Me.btnGo.Image = CType(resources.GetObject("btnGo.Image"), System.Drawing.Image)
        Me.btnGo.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnGo.Location = New System.Drawing.Point(107, 12)
        Me.btnGo.Name = "btnGo"
        Me.btnGo.Size = New System.Drawing.Size(56, 24)
        Me.btnGo.TabIndex = 1
        Me.btnGo.Text = "&Go"
        Me.btnGo.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnGo.UseVisualStyleBackColor = False
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.AliceBlue
        Me.btnExit.Image = CType(resources.GetObject("btnExit.Image"), System.Drawing.Image)
        Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnExit.Location = New System.Drawing.Point(18, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(56, 21)
        Me.btnExit.TabIndex = 0
        Me.btnExit.Text = "E&xit"
        Me.btnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'frmOpenOneInvoice
        '
        Me.AcceptButton = Me.btnGo
        Me.AccessibleName = ""
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 15)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(206, 100)
        Me.Controls.Add(Me.grpFooter)
        Me.Controls.Add(Me.grpMain)
        Me.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmOpenOneInvoice"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Open an Invoice"
        Me.grpMain.ResumeLayout(False)
        Me.grpMain.PerformLayout()
        Me.grpFooter.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub txtInvNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtInvNo.TextChanged

        If IsNumeric(Me.txtInvNo.Text) Then
            Me.btnGo.Enabled = True
        End If

    End Sub

    Private Sub txtInvNo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtInvNo.KeyPress

        e.Handled = Utilities.NumericOnly(e, Me.txtInvNo)

    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try

            Me.Close()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub btnGo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGo.Click

        Try
            If IsNumeric(Me.txtInvNo.Text) Then

                If Utilities.OpenInvoice(Convert.ToInt32(Me.txtInvNo.Text), Me.ParentForm) = False Then
                    MessageBox.Show("Invoice Does Not Exist", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If
            End If

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub
End Class
