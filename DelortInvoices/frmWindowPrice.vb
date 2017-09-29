Imports DelortBusObjects
Imports System.Text.RegularExpressions

Public Class frmWindowPrice
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
    Friend WithEvents rdoIncrease As System.Windows.Forms.RadioButton
    Friend WithEvents rdoDecrease As System.Windows.Forms.RadioButton
    Friend WithEvents grpAction As System.Windows.Forms.GroupBox
    Friend WithEvents grpType As System.Windows.Forms.GroupBox
    Friend WithEvents lblAmount As System.Windows.Forms.Label
    Friend WithEvents txtAmount As System.Windows.Forms.TextBox
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnDoIt As System.Windows.Forms.Button
    Friend WithEvents rdoSetAll As System.Windows.Forms.RadioButton
    Friend WithEvents rdoPercent As System.Windows.Forms.RadioButton
    Friend WithEvents rdoDollar As System.Windows.Forms.RadioButton
    Friend WithEvents ErrorProvider As System.Windows.Forms.ErrorProvider
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmWindowPrice))
        Me.grpMain = New System.Windows.Forms.GroupBox()
        Me.rdoDecrease = New System.Windows.Forms.RadioButton()
        Me.rdoIncrease = New System.Windows.Forms.RadioButton()
        Me.grpAction = New System.Windows.Forms.GroupBox()
        Me.btnDoIt = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.grpType = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.rdoSetAll = New System.Windows.Forms.RadioButton()
        Me.lblAmount = New System.Windows.Forms.Label()
        Me.txtAmount = New System.Windows.Forms.TextBox()
        Me.rdoPercent = New System.Windows.Forms.RadioButton()
        Me.rdoDollar = New System.Windows.Forms.RadioButton()
        Me.ErrorProvider = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.grpMain.SuspendLayout()
        Me.grpAction.SuspendLayout()
        Me.grpType.SuspendLayout()
        CType(Me.ErrorProvider, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grpMain
        '
        Me.grpMain.BackColor = System.Drawing.Color.Transparent
        Me.grpMain.Controls.Add(Me.rdoDecrease)
        Me.grpMain.Controls.Add(Me.rdoIncrease)
        Me.grpMain.Location = New System.Drawing.Point(8, 16)
        Me.grpMain.Name = "grpMain"
        Me.grpMain.Size = New System.Drawing.Size(192, 56)
        Me.grpMain.TabIndex = 0
        Me.grpMain.TabStop = False
        Me.grpMain.Text = "Direction:"
        '
        'rdoDecrease
        '
        Me.rdoDecrease.Location = New System.Drawing.Point(96, 20)
        Me.rdoDecrease.Name = "rdoDecrease"
        Me.rdoDecrease.Size = New System.Drawing.Size(80, 17)
        Me.rdoDecrease.TabIndex = 1
        Me.rdoDecrease.Text = "Decrease"
        '
        'rdoIncrease
        '
        Me.rdoIncrease.Checked = True
        Me.rdoIncrease.Location = New System.Drawing.Point(7, 20)
        Me.rdoIncrease.Name = "rdoIncrease"
        Me.rdoIncrease.Size = New System.Drawing.Size(73, 17)
        Me.rdoIncrease.TabIndex = 0
        Me.rdoIncrease.TabStop = True
        Me.rdoIncrease.Text = "Increase"
        '
        'grpAction
        '
        Me.grpAction.BackColor = System.Drawing.Color.Transparent
        Me.grpAction.Controls.Add(Me.btnDoIt)
        Me.grpAction.Controls.Add(Me.btnExit)
        Me.grpAction.Location = New System.Drawing.Point(8, 240)
        Me.grpAction.Name = "grpAction"
        Me.grpAction.Size = New System.Drawing.Size(192, 48)
        Me.grpAction.TabIndex = 1
        Me.grpAction.TabStop = False
        Me.grpAction.Text = "Action:"
        '
        'btnDoIt
        '
        Me.btnDoIt.Enabled = False
        Me.btnDoIt.Image = CType(resources.GetObject("btnDoIt.Image"), System.Drawing.Image)
        Me.btnDoIt.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDoIt.Location = New System.Drawing.Point(117, 16)
        Me.btnDoIt.Name = "btnDoIt"
        Me.btnDoIt.Size = New System.Drawing.Size(56, 24)
        Me.btnDoIt.TabIndex = 1
        Me.btnDoIt.Text = "Do It"
        Me.btnDoIt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnExit
        '
        Me.btnExit.Image = CType(resources.GetObject("btnExit.Image"), System.Drawing.Image)
        Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnExit.Location = New System.Drawing.Point(16, 16)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(56, 24)
        Me.btnExit.TabIndex = 0
        Me.btnExit.Text = "E&xit"
        Me.btnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'grpType
        '
        Me.grpType.BackColor = System.Drawing.Color.Transparent
        Me.grpType.Controls.Add(Me.Label1)
        Me.grpType.Controls.Add(Me.rdoSetAll)
        Me.grpType.Controls.Add(Me.lblAmount)
        Me.grpType.Controls.Add(Me.txtAmount)
        Me.grpType.Controls.Add(Me.rdoPercent)
        Me.grpType.Controls.Add(Me.rdoDollar)
        Me.grpType.Location = New System.Drawing.Point(8, 80)
        Me.grpType.Name = "grpType"
        Me.grpType.Size = New System.Drawing.Size(192, 152)
        Me.grpType.TabIndex = 3
        Me.grpType.TabStop = False
        Me.grpType.Text = "Type:"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(55, 126)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(120, 16)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "(Format $ as XX.XX)"
        '
        'rdoSetAll
        '
        Me.rdoSetAll.Location = New System.Drawing.Point(8, 61)
        Me.rdoSetAll.Name = "rdoSetAll"
        Me.rdoSetAll.Size = New System.Drawing.Size(160, 17)
        Me.rdoSetAll.TabIndex = 6
        Me.rdoSetAll.Text = "Set All To Amount Below"
        '
        'lblAmount
        '
        Me.lblAmount.Location = New System.Drawing.Point(8, 102)
        Me.lblAmount.Name = "lblAmount"
        Me.lblAmount.Size = New System.Drawing.Size(64, 20)
        Me.lblAmount.TabIndex = 5
        Me.lblAmount.Text = "Amount:"
        '
        'txtAmount
        '
        Me.txtAmount.Location = New System.Drawing.Point(88, 102)
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.Size = New System.Drawing.Size(80, 21)
        Me.txtAmount.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.txtAmount, "Enter doll amounts as 99.99")
        '
        'rdoPercent
        '
        Me.rdoPercent.Location = New System.Drawing.Point(88, 26)
        Me.rdoPercent.Name = "rdoPercent"
        Me.rdoPercent.Size = New System.Drawing.Size(72, 17)
        Me.rdoPercent.TabIndex = 3
        Me.rdoPercent.Text = "Percent"
        '
        'rdoDollar
        '
        Me.rdoDollar.Checked = True
        Me.rdoDollar.Location = New System.Drawing.Point(8, 26)
        Me.rdoDollar.Name = "rdoDollar"
        Me.rdoDollar.Size = New System.Drawing.Size(72, 17)
        Me.rdoDollar.TabIndex = 2
        Me.rdoDollar.TabStop = True
        Me.rdoDollar.Text = "Dollar "
        '
        'ErrorProvider
        '
        Me.ErrorProvider.ContainerControl = Me
        '
        'frmWindowPrice
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.ClientSize = New System.Drawing.Size(210, 304)
        Me.Controls.Add(Me.grpType)
        Me.Controls.Add(Me.grpAction)
        Me.Controls.Add(Me.grpMain)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmWindowPrice"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Change Window Pricing"
        Me.grpMain.ResumeLayout(False)
        Me.grpAction.ResumeLayout(False)
        Me.grpType.ResumeLayout(False)
        Me.grpType.PerformLayout()
        CType(Me.ErrorProvider, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public WithEvents Customer As New Customer(g_Settings)

    Private SQlPart1 As String = "UPDATE Customers SET WindowPrice = WindowPrice "
    Private SQLPlusMinus As String 'WILL BE EITHER + OR -
    Private SQLType As String 'will be either ( WindowPrice * .25) or a number

    Private SQLLastpart As String = " WHERE WindowsCustomer = '1' AND IsActive = '1'"



    Private Sub txtAmount_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAmount.KeyPress
        e.Handled = Utilities.NumericOnly(e, Me.txtAmount)

    End Sub


    Private Sub btnDoIt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDoIt.Click

        Try
            Dim Sql As String
            Dim Msg As String = "Are you sure you want to "

            If SQLPlusMinus.Trim = "+" Then
                Msg += "increase ALL Customer's window pricing "

            Else
                Msg += "decrease ALL Customer's window pricing "

            End If


            If Convert.ToDouble(Me.txtAmount.Text) <= 0 Then
                MessageBox.Show("Must type in an amount greater than 0", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Sub
            Else

                If Me.rdoSetAll.Checked Then
                    Msg = "Are you sure you want to set ALL customer window's price to $" & Me.txtAmount.Text & "?"

                ElseIf Me.rdoDollar.Checked Then
                    Msg += "by " & Me.txtAmount.Text & " dollars?"
                Else
                    Msg += "by " & Me.txtAmount.Text & " percent?"

                End If

                If MessageBox.Show(Msg, ProgramName, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.No Then
                    Exit Sub

                End If
            End If


            If Me.rdoSetAll.Checked = True Then
                Sql = "UPDATE Customers SET WindowPrice = "
                Sql += Me.txtAmount.Text
                Sql += " WHERE WindowsCustomer = '1' AND IsActive = '1'"
            Else

                If Me.rdoDollar.Checked Then
                    SQLType = Me.txtAmount.Text
                Else
                    SQLType = " (WindowPrice * " & (Convert.ToDouble(Me.txtAmount.Text) * 0.01) & ")"
                End If

                Sql = SQlPart1 & SQLPlusMinus & SQLType & SQLLastpart

            End If


            Customer.UpdateWindowPricing(Sql)

            MessageBox.Show("Prices updated successfully", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try



    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()

    End Sub


    Private Sub rdoDecrease_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdoDecrease.Click, rdoIncrease.Click

        Me.txtAmount.Text = ""

        Select Case DirectCast(sender, RadioButton).Name

            Case "rdoDecrease"
                SQLPlusMinus = "- "
            Case "rdoIncrease"
                SQLPlusMinus = "+ "
        End Select

    End Sub

    Private Sub txtAmount_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAmount.TextChanged
        Try

            If Me.rdoPercent.Checked Then
                If Not Regex.IsMatch(Me.txtAmount.Text.Trim, "\d$") Then
                    Me.btnDoIt.Enabled = False
                    Me.ErrorProvider.SetError(Me.txtAmount, "For dollar amounts use decimal, just a number for %")
                Else
                    Me.btnDoIt.Enabled = True
                    Me.ErrorProvider.SetError(Me.txtAmount, "")

                End If
            Else
                If Not Regex.IsMatch(Me.txtAmount.Text.Trim, "^[0-9]+(\.[0-9][0-9])") Then
                    Me.btnDoIt.Enabled = False
                    Me.ErrorProvider.SetError(Me.txtAmount, "For dollar amounts use decimal, just a number for %")

                Else
                    Me.btnDoIt.Enabled = True
                    Me.ErrorProvider.SetError(Me.txtAmount, "")

                End If
            End If


        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)


        End Try

    End Sub

    Private Sub rdoDollar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdoDollar.Click, rdoPercent.Click, rdoSetAll.Click
        Me.txtAmount.Text = ""

    End Sub
End Class
