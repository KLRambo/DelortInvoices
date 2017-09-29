Option Strict On
Imports DelortBusObjects
Imports System.Configuration.ConfigurationSettings

Public Class frmPrintCustomers

    Inherits System.Windows.Forms.Form
    Public Customer As New Customer(g_Settings)

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
    Friend WithEvents grpFooter As System.Windows.Forms.GroupBox
    Friend WithEvents rdoAllCustomers As System.Windows.Forms.RadioButton
    Friend WithEvents rdoWindowsCust As System.Windows.Forms.RadioButton
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents grpSettings As System.Windows.Forms.GroupBox
    Friend WithEvents rdoSingleSpace As System.Windows.Forms.RadioButton
    Friend WithEvents rdoDoubleSpacing As System.Windows.Forms.RadioButton
    Friend WithEvents rdoNonWindows As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rdoNotePad As System.Windows.Forms.RadioButton
    Friend WithEvents rdoStandard As System.Windows.Forms.RadioButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPrintCustomers))
        Me.grpMain = New System.Windows.Forms.GroupBox
        Me.rdoNonWindows = New System.Windows.Forms.RadioButton
        Me.rdoWindowsCust = New System.Windows.Forms.RadioButton
        Me.rdoAllCustomers = New System.Windows.Forms.RadioButton
        Me.grpFooter = New System.Windows.Forms.GroupBox
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.grpSettings = New System.Windows.Forms.GroupBox
        Me.rdoDoubleSpacing = New System.Windows.Forms.RadioButton
        Me.rdoSingleSpace = New System.Windows.Forms.RadioButton
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.rdoNotePad = New System.Windows.Forms.RadioButton
        Me.rdoStandard = New System.Windows.Forms.RadioButton
        Me.grpMain.SuspendLayout()
        Me.grpFooter.SuspendLayout()
        Me.grpSettings.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpMain
        '
        Me.grpMain.Controls.Add(Me.rdoNonWindows)
        Me.grpMain.Controls.Add(Me.rdoWindowsCust)
        Me.grpMain.Controls.Add(Me.rdoAllCustomers)
        Me.grpMain.Location = New System.Drawing.Point(4, 10)
        Me.grpMain.Name = "grpMain"
        Me.grpMain.Size = New System.Drawing.Size(202, 118)
        Me.grpMain.TabIndex = 0
        Me.grpMain.TabStop = False
        Me.grpMain.Text = "Report"
        '
        'rdoNonWindows
        '
        Me.rdoNonWindows.Location = New System.Drawing.Point(16, 88)
        Me.rdoNonWindows.Name = "rdoNonWindows"
        Me.rdoNonWindows.Size = New System.Drawing.Size(176, 16)
        Me.rdoNonWindows.TabIndex = 2
        Me.rdoNonWindows.Text = "Non Window Customers"
        '
        'rdoWindowsCust
        '
        Me.rdoWindowsCust.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoWindowsCust.ForeColor = System.Drawing.SystemColors.ControlText
        Me.rdoWindowsCust.Location = New System.Drawing.Point(16, 56)
        Me.rdoWindowsCust.Name = "rdoWindowsCust"
        Me.rdoWindowsCust.Size = New System.Drawing.Size(176, 16)
        Me.rdoWindowsCust.TabIndex = 1
        Me.rdoWindowsCust.Text = "Window Customers"
        '
        'rdoAllCustomers
        '
        Me.rdoAllCustomers.Checked = True
        Me.rdoAllCustomers.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoAllCustomers.ForeColor = System.Drawing.SystemColors.ControlText
        Me.rdoAllCustomers.Location = New System.Drawing.Point(16, 24)
        Me.rdoAllCustomers.Name = "rdoAllCustomers"
        Me.rdoAllCustomers.Size = New System.Drawing.Size(128, 16)
        Me.rdoAllCustomers.TabIndex = 0
        Me.rdoAllCustomers.TabStop = True
        Me.rdoAllCustomers.Text = "All Customers"
        '
        'grpFooter
        '
        Me.grpFooter.Controls.Add(Me.btnPrint)
        Me.grpFooter.Controls.Add(Me.btnExit)
        Me.grpFooter.Location = New System.Drawing.Point(4, 232)
        Me.grpFooter.Name = "grpFooter"
        Me.grpFooter.Size = New System.Drawing.Size(202, 48)
        Me.grpFooter.TabIndex = 1
        Me.grpFooter.TabStop = False
        '
        'btnPrint
        '
        Me.btnPrint.BackColor = System.Drawing.Color.Azure
        Me.btnPrint.Image = CType(resources.GetObject("btnPrint.Image"), System.Drawing.Image)
        Me.btnPrint.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnPrint.Location = New System.Drawing.Point(123, 17)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(64, 22)
        Me.btnPrint.TabIndex = 1
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.Azure
        Me.btnExit.Image = CType(resources.GetObject("btnExit.Image"), System.Drawing.Image)
        Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnExit.Location = New System.Drawing.Point(11, 17)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(64, 22)
        Me.btnExit.TabIndex = 0
        Me.btnExit.Text = "E&xit"
        Me.btnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'grpSettings
        '
        Me.grpSettings.Controls.Add(Me.rdoDoubleSpacing)
        Me.grpSettings.Controls.Add(Me.rdoSingleSpace)
        Me.grpSettings.Location = New System.Drawing.Point(6, 131)
        Me.grpSettings.Name = "grpSettings"
        Me.grpSettings.Size = New System.Drawing.Size(96, 98)
        Me.grpSettings.TabIndex = 2
        Me.grpSettings.TabStop = False
        Me.grpSettings.Text = "Spacing"
        '
        'rdoDoubleSpacing
        '
        Me.rdoDoubleSpacing.Location = New System.Drawing.Point(8, 64)
        Me.rdoDoubleSpacing.Name = "rdoDoubleSpacing"
        Me.rdoDoubleSpacing.Size = New System.Drawing.Size(72, 16)
        Me.rdoDoubleSpacing.TabIndex = 1
        Me.rdoDoubleSpacing.Text = "Double"
        '
        'rdoSingleSpace
        '
        Me.rdoSingleSpace.Checked = True
        Me.rdoSingleSpace.Location = New System.Drawing.Point(8, 24)
        Me.rdoSingleSpace.Name = "rdoSingleSpace"
        Me.rdoSingleSpace.Size = New System.Drawing.Size(64, 16)
        Me.rdoSingleSpace.TabIndex = 0
        Me.rdoSingleSpace.TabStop = True
        Me.rdoSingleSpace.Text = "Single"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rdoNotePad)
        Me.GroupBox1.Controls.Add(Me.rdoStandard)
        Me.GroupBox1.Location = New System.Drawing.Point(112, 131)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(96, 98)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Output"
        '
        'rdoNotePad
        '
        Me.rdoNotePad.Location = New System.Drawing.Point(11, 64)
        Me.rdoNotePad.Name = "rdoNotePad"
        Me.rdoNotePad.Size = New System.Drawing.Size(77, 16)
        Me.rdoNotePad.TabIndex = 1
        Me.rdoNotePad.Text = "Notepad"
        '
        'rdoStandard
        '
        Me.rdoStandard.Checked = True
        Me.rdoStandard.Location = New System.Drawing.Point(11, 24)
        Me.rdoStandard.Name = "rdoStandard"
        Me.rdoStandard.Size = New System.Drawing.Size(77, 16)
        Me.rdoStandard.TabIndex = 0
        Me.rdoStandard.TabStop = True
        Me.rdoStandard.Text = "Standard"
        '
        'frmPrintCustomers
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.BackColor = System.Drawing.Color.LightSteelBlue
        Me.ClientSize = New System.Drawing.Size(210, 288)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.grpSettings)
        Me.Controls.Add(Me.grpFooter)
        Me.Controls.Add(Me.grpMain)
        Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmPrintCustomers"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Print Customers"
        Me.grpMain.ResumeLayout(False)
        Me.grpFooter.ResumeLayout(False)
        Me.grpSettings.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click

        Cursor = Cursors.WaitCursor

        Dim Spacing As String
        If Me.rdoSingleSpace.Checked Then
            Spacing = "Single"
        Else
            Spacing = "Double"
        End If

        Try

            If Me.rdoNotePad.Checked Then
                If Me.rdoAllCustomers.Checked Then
                    Utilities.PrintCustomers(Customer.GetAllCustomers, "AllCustomers.txt", Spacing, Utilities.ReportType.All)
                ElseIf Me.rdoWindowsCust.Checked Then
                    Utilities.PrintCustomers(Customer.GetWindowCustomers, "WindowCustomers.txt", Spacing, Utilities.ReportType.Windows)
                ElseIf Me.rdoNonWindows.Checked Then
                    Utilities.PrintCustomers(Customer.GetNonWindowCustomers, "NonWindowsCustomers.txt", Spacing, Utilities.ReportType.NonWindows)
                End If

            Else
                If Me.rdoAllCustomers.Checked Then
                    Utilities.PrintCustomersActiveReports(Customer.GetAllCustomers, "AllCustomers.txt", Spacing, Utilities.ReportType.All)
                ElseIf Me.rdoWindowsCust.Checked Then
                    Utilities.PrintCustomersActiveReports(Customer.GetWindowCustomers, "WindowCustomers.txt", Spacing, Utilities.ReportType.Windows)
                ElseIf Me.rdoNonWindows.Checked Then
                    Utilities.PrintCustomersActiveReports(Customer.GetNonWindowCustomers, "NonWindowsCustomers.txt", Spacing, Utilities.ReportType.NonWindows)
                End If

            End If

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)


        Finally

            Cursor = Cursors.Default

        End Try

    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click

        If Not IsNothing(Me.Customer) Then
            Me.Customer = Nothing
        End If

        Me.Dispose()

    End Sub
End Class
