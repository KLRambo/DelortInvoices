Option Strict On

Imports DelortBusObjects
Imports System.Configuration.ConfigurationSettings


Public Class frmFindOneCustomer

    Inherits frmBaseWindow

    Private Customer As New Customer(g_settings)

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
    Friend WithEvents lblLastName As System.Windows.Forms.Label
    Friend WithEvents lblFirstName As System.Windows.Forms.Label
    Friend WithEvents txtLastName As System.Windows.Forms.TextBox
    Friend WithEvents txtFirstName As System.Windows.Forms.TextBox
    Friend WithEvents grpFooter As System.Windows.Forms.GroupBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnGo As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFindOneCustomer))
        Me.grpMain = New System.Windows.Forms.GroupBox()
        Me.lblFirstName = New System.Windows.Forms.Label()
        Me.lblLastName = New System.Windows.Forms.Label()
        Me.txtFirstName = New System.Windows.Forms.TextBox()
        Me.txtLastName = New System.Windows.Forms.TextBox()
        Me.grpFooter = New System.Windows.Forms.GroupBox()
        Me.btnGo = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.grpMain.SuspendLayout()
        Me.grpFooter.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpMain
        '
        Me.grpMain.BackColor = System.Drawing.Color.Transparent
        Me.grpMain.Controls.Add(Me.lblFirstName)
        Me.grpMain.Controls.Add(Me.lblLastName)
        Me.grpMain.Controls.Add(Me.txtFirstName)
        Me.grpMain.Controls.Add(Me.txtLastName)
        Me.grpMain.Location = New System.Drawing.Point(8, 4)
        Me.grpMain.Name = "grpMain"
        Me.grpMain.Size = New System.Drawing.Size(256, 96)
        Me.grpMain.TabIndex = 0
        Me.grpMain.TabStop = False
        Me.grpMain.Text = "Enter Exact Match For Both Names"
        '
        'lblFirstName
        '
        Me.lblFirstName.Location = New System.Drawing.Point(9, 62)
        Me.lblFirstName.Name = "lblFirstName"
        Me.lblFirstName.Size = New System.Drawing.Size(79, 16)
        Me.lblFirstName.TabIndex = 3
        Me.lblFirstName.Text = "First Name:"
        '
        'lblLastName
        '
        Me.lblLastName.Location = New System.Drawing.Point(9, 22)
        Me.lblLastName.Name = "lblLastName"
        Me.lblLastName.Size = New System.Drawing.Size(79, 16)
        Me.lblLastName.TabIndex = 2
        Me.lblLastName.Text = "Last Name:"
        '
        'txtFirstName
        '
        Me.txtFirstName.Location = New System.Drawing.Point(96, 62)
        Me.txtFirstName.Name = "txtFirstName"
        Me.txtFirstName.Size = New System.Drawing.Size(152, 22)
        Me.txtFirstName.TabIndex = 1
        '
        'txtLastName
        '
        Me.txtLastName.Location = New System.Drawing.Point(96, 22)
        Me.txtLastName.Name = "txtLastName"
        Me.txtLastName.Size = New System.Drawing.Size(152, 22)
        Me.txtLastName.TabIndex = 0
        '
        'grpFooter
        '
        Me.grpFooter.BackColor = System.Drawing.Color.Transparent
        Me.grpFooter.Controls.Add(Me.btnGo)
        Me.grpFooter.Controls.Add(Me.btnCancel)
        Me.grpFooter.Location = New System.Drawing.Point(8, 104)
        Me.grpFooter.Name = "grpFooter"
        Me.grpFooter.Size = New System.Drawing.Size(256, 40)
        Me.grpFooter.TabIndex = 3
        Me.grpFooter.TabStop = False
        '
        'btnGo
        '
        Me.btnGo.BackColor = System.Drawing.Color.AliceBlue
        Me.btnGo.Image = CType(resources.GetObject("btnGo.Image"), System.Drawing.Image)
        Me.btnGo.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnGo.Location = New System.Drawing.Point(171, 10)
        Me.btnGo.Name = "btnGo"
        Me.btnGo.Size = New System.Drawing.Size(72, 24)
        Me.btnGo.TabIndex = 3
        Me.btnGo.Text = "Go"
        Me.btnGo.UseVisualStyleBackColor = False
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.Color.AliceBlue
        Me.btnCancel.Image = CType(resources.GetObject("btnCancel.Image"), System.Drawing.Image)
        Me.btnCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCancel.Location = New System.Drawing.Point(8, 10)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(72, 24)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "E&xit"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'frmFindOneCustomer
        '
        Me.AcceptButton = Me.btnGo
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 15)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(274, 152)
        Me.Controls.Add(Me.grpFooter)
        Me.Controls.Add(Me.grpMain)
        Me.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmFindOneCustomer"
        Me.Text = "Find Customer"
        Me.grpMain.ResumeLayout(False)
        Me.grpMain.PerformLayout()
        Me.grpFooter.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnGo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGo.Click
        Try
            Dim cusObj As New Customer
            cusObj = Me.Customer.GetCustomer

            If cusObj.CustID = 0 Then
                MessageBox.Show("Customer not found.", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                Utilities.ViewAddEditCustomer(Me.MdiParent, cusObj, Nothing)

            End If

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click

        Try

            Me.Close()

        Catch ex As Exception

            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub txtLastName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLastName.TextChanged
        Me.Customer.LastName = Me.txtLastName.Text.Trim

    End Sub

    Private Sub txtFirstName_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFirstName.TextChanged
        Me.Customer.FirstName = Me.txtFirstName.Text.Trim

    End Sub


End Class
