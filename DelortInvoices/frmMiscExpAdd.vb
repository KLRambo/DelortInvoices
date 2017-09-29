Option Strict On

Imports System.Configuration.ConfigurationSettings

Public Class frmMiscExpAdd

    Inherits frmBaseWindow

    Private WithEvents Expenses As New DelortBusObjects.MiscExp(g_settings)
    Private Loading As Boolean = False
    Public CallingForm As frmMiscExpRpt
    Public dsExpItem As DataSet

    Public Enum Mode
        Add = 0
        Edit = 1
    End Enum

    Public FormMode As Mode

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal Action As Mode)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Me.FormMode = Action

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
    Friend WithEvents lblDate As System.Windows.Forms.Label
    Friend WithEvents lblDesc As System.Windows.Forms.Label
    Friend WithEvents lblNotes As System.Windows.Forms.Label
    Friend WithEvents dtpMiscExp As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtNotes As System.Windows.Forms.TextBox
    Friend WithEvents grpFooter As System.Windows.Forms.GroupBox
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents lblCost As System.Windows.Forms.Label
    Friend WithEvents txtCost As System.Windows.Forms.TextBox
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents cboType As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMiscExpAdd))
        Me.grpMain = New System.Windows.Forms.GroupBox()
        Me.cboType = New System.Windows.Forms.ComboBox()
        Me.lblType = New System.Windows.Forms.Label()
        Me.txtCost = New System.Windows.Forms.TextBox()
        Me.lblCost = New System.Windows.Forms.Label()
        Me.txtNotes = New System.Windows.Forms.TextBox()
        Me.lblNotes = New System.Windows.Forms.Label()
        Me.txtDesc = New System.Windows.Forms.TextBox()
        Me.lblDesc = New System.Windows.Forms.Label()
        Me.dtpMiscExp = New System.Windows.Forms.DateTimePicker()
        Me.lblDate = New System.Windows.Forms.Label()
        Me.grpFooter = New System.Windows.Forms.GroupBox()
        Me.btnView = New System.Windows.Forms.Button()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.grpMain.SuspendLayout()
        Me.grpFooter.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpMain
        '
        Me.grpMain.BackColor = System.Drawing.Color.Transparent
        Me.grpMain.Controls.Add(Me.cboType)
        Me.grpMain.Controls.Add(Me.lblType)
        Me.grpMain.Controls.Add(Me.txtCost)
        Me.grpMain.Controls.Add(Me.lblCost)
        Me.grpMain.Controls.Add(Me.txtNotes)
        Me.grpMain.Controls.Add(Me.lblNotes)
        Me.grpMain.Controls.Add(Me.txtDesc)
        Me.grpMain.Controls.Add(Me.lblDesc)
        Me.grpMain.Controls.Add(Me.dtpMiscExp)
        Me.grpMain.Controls.Add(Me.lblDate)
        Me.grpMain.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpMain.Location = New System.Drawing.Point(8, 4)
        Me.grpMain.Name = "grpMain"
        Me.grpMain.Size = New System.Drawing.Size(312, 260)
        Me.grpMain.TabIndex = 0
        Me.grpMain.TabStop = False
        '
        'cboType
        '
        Me.cboType.Location = New System.Drawing.Point(99, 50)
        Me.cboType.Name = "cboType"
        Me.cboType.Size = New System.Drawing.Size(197, 21)
        Me.cboType.TabIndex = 1
        '
        'lblType
        '
        Me.lblType.Location = New System.Drawing.Point(17, 50)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(74, 16)
        Me.lblType.TabIndex = 7
        Me.lblType.Text = "Category:"
        '
        'txtCost
        '
        Me.txtCost.Location = New System.Drawing.Point(99, 154)
        Me.txtCost.Name = "txtCost"
        Me.txtCost.Size = New System.Drawing.Size(104, 21)
        Me.txtCost.TabIndex = 3
        '
        'lblCost
        '
        Me.lblCost.Location = New System.Drawing.Point(38, 154)
        Me.lblCost.Name = "lblCost"
        Me.lblCost.Size = New System.Drawing.Size(53, 16)
        Me.lblCost.TabIndex = 6
        Me.lblCost.Text = "Cost:"
        '
        'txtNotes
        '
        Me.txtNotes.Location = New System.Drawing.Point(99, 189)
        Me.txtNotes.Multiline = True
        Me.txtNotes.Name = "txtNotes"
        Me.txtNotes.Size = New System.Drawing.Size(197, 56)
        Me.txtNotes.TabIndex = 4
        '
        'lblNotes
        '
        Me.lblNotes.Location = New System.Drawing.Point(38, 189)
        Me.lblNotes.Name = "lblNotes"
        Me.lblNotes.Size = New System.Drawing.Size(53, 16)
        Me.lblNotes.TabIndex = 4
        Me.lblNotes.Text = "Notes:"
        '
        'txtDesc
        '
        Me.txtDesc.Location = New System.Drawing.Point(99, 85)
        Me.txtDesc.Multiline = True
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.Size = New System.Drawing.Size(197, 56)
        Me.txtDesc.TabIndex = 2
        '
        'lblDesc
        '
        Me.lblDesc.Location = New System.Drawing.Point(11, 85)
        Me.lblDesc.Name = "lblDesc"
        Me.lblDesc.Size = New System.Drawing.Size(80, 16)
        Me.lblDesc.TabIndex = 2
        Me.lblDesc.Text = "Description:"
        '
        'dtpMiscExp
        '
        Me.dtpMiscExp.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpMiscExp.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpMiscExp.Location = New System.Drawing.Point(99, 15)
        Me.dtpMiscExp.Name = "dtpMiscExp"
        Me.dtpMiscExp.Size = New System.Drawing.Size(104, 21)
        Me.dtpMiscExp.TabIndex = 0
        '
        'lblDate
        '
        Me.lblDate.Location = New System.Drawing.Point(43, 15)
        Me.lblDate.Name = "lblDate"
        Me.lblDate.Size = New System.Drawing.Size(48, 16)
        Me.lblDate.TabIndex = 0
        Me.lblDate.Text = "Date:"
        '
        'grpFooter
        '
        Me.grpFooter.BackColor = System.Drawing.Color.Transparent
        Me.grpFooter.Controls.Add(Me.btnView)
        Me.grpFooter.Controls.Add(Me.btnAdd)
        Me.grpFooter.Controls.Add(Me.btnExit)
        Me.grpFooter.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpFooter.Location = New System.Drawing.Point(8, 264)
        Me.grpFooter.Name = "grpFooter"
        Me.grpFooter.Size = New System.Drawing.Size(312, 49)
        Me.grpFooter.TabIndex = 2
        Me.grpFooter.TabStop = False
        '
        'btnView
        '
        Me.btnView.BackColor = System.Drawing.Color.AliceBlue
        Me.btnView.Image = CType(resources.GetObject("btnView.Image"), System.Drawing.Image)
        Me.btnView.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnView.Location = New System.Drawing.Point(113, 14)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(88, 24)
        Me.btnView.TabIndex = 6
        Me.btnView.Text = "&View All"
        Me.btnView.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnView.UseVisualStyleBackColor = False
        '
        'btnAdd
        '
        Me.btnAdd.BackColor = System.Drawing.Color.AliceBlue
        Me.btnAdd.Enabled = False
        Me.btnAdd.Image = CType(resources.GetObject("btnAdd.Image"), System.Drawing.Image)
        Me.btnAdd.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnAdd.Location = New System.Drawing.Point(213, 14)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(88, 24)
        Me.btnAdd.TabIndex = 7
        Me.btnAdd.Text = "   &Save"
        Me.btnAdd.UseVisualStyleBackColor = False
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.AliceBlue
        Me.btnExit.Image = CType(resources.GetObject("btnExit.Image"), System.Drawing.Image)
        Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnExit.Location = New System.Drawing.Point(13, 14)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(88, 24)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'lblStatus
        '
        Me.lblStatus.BackColor = System.Drawing.Color.Transparent
        Me.lblStatus.Location = New System.Drawing.Point(8, 317)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(312, 21)
        Me.lblStatus.TabIndex = 3
        '
        'frmMiscExpAdd
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 15)
        Me.ClientSize = New System.Drawing.Size(330, 344)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.grpFooter)
        Me.Controls.Add(Me.grpMain)
        Me.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmMiscExpAdd"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Misc Expenses"
        Me.grpMain.ResumeLayout(False)
        Me.grpMain.PerformLayout()
        Me.grpFooter.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub txtCost_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCost.TextChanged

        If Me.Loading = True Then Exit Sub

        If IsNumeric(Me.txtCost.Text) Then
            Me.Expenses.Cost = Convert.ToDouble(Me.txtCost.Text)
        Else
            Me.Expenses.Cost = 0
        End If

        If Me.FormMode = Mode.Add Then
            Me.lblStatus.Text = "Adding Misc Expenses..."
        Else
            Me.lblStatus.Text = "Editing Misc Expenses..."
        End If
    End Sub

    Private Sub txtCost_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCost.KeyPress

        e.Handled = Utilities.NumericOnly(e, Me.txtCost)

    End Sub

    Private Sub dtpMiscExp_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpMiscExp.ValueChanged

        Try

            Me.Expenses.ExpDate = Convert.ToDateTime(Me.dtpMiscExp.Value)

        Catch ex As Exception

            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub txtDesc_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDesc.TextChanged
        If Me.Loading = True Then Exit Sub

        Me.Expenses.Description = Me.txtDesc.Text.Trim
        Me.lblStatus.ForeColor = System.Drawing.Color.Black

        If Me.FormMode = Mode.Add Then
            Me.lblStatus.Text = "Adding Misc Expense..."
        Else
            Me.lblStatus.Text = "Editing Misc Expense..."
        End If

    End Sub

    Private Sub Expenses_Valid(ByVal IsValid As Boolean) Handles Expenses.Valid

        Me.btnAdd.Enabled = IsValid

    End Sub

    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click

        Try

            Me.Close()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub txtNotes_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNotes.TextChanged

        If Me.Loading = True Then Exit Sub

        Me.Expenses.Notes = Me.txtNotes.Text

        Me.lblStatus.ForeColor = System.Drawing.Color.Black

        If Me.FormMode = Mode.Add Then
            Me.lblStatus.Text = "Adding Misc Expense..."
        Else
            Me.lblStatus.Text = "Editing Misc Expense..."
        End If

    End Sub

    Private Sub btnAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAdd.Click

        Try

            Me.Expenses.InsertUpdateDelete()

            If Me.FormMode = Mode.Add Then
                ResetForm()
            End If

            If Not Me.CallingForm Is Nothing Then
                Me.CallingForm.LoadData()

            End If

            Me.lblStatus.ForeColor = System.Drawing.Color.Red
            Me.lblStatus.Text = "Expense Updated!"
            Me.Expenses.IsDirty = False
            Me.btnAdd.Enabled = False

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
    Private Sub ResetForm()

        Me.txtCost.Text = vbNullString
        Me.txtDesc.Text = vbNullString
        Me.dtpMiscExp.Value = Now
        Me.txtNotes.Text = vbNullString

    End Sub

    Private Sub btnView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnView.Click

        Try
            Utilities.ViewMiscExpReport(Me.ParentForm)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub frmMiscExpAdd_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try

            If Me.FormMode = Mode.Edit Then
                Utilities.LoadMiscExpTypes(Me.cboType, "All")
                LoadData()
                Me.btnAdd.Enabled = False
                Me.Expenses.IsValid = True
            Else
                Utilities.LoadMiscExpTypes(Me.cboType, "Active")
            End If


            Me.Loading = False
            Me.Expenses.IsDirty = False

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
    Private Sub LoadData()

        With Me.dsExpItem.Tables(0).Rows(0)
            If IsDate(.Item("DateOfExp")) Then
                Me.dtpMiscExp.Value = Convert.ToDateTime(.Item("DateOfExp"))
            End If
            Me.txtCost.Text = .Item("Cost").ToString
            Me.txtDesc.Text = .Item("Description").ToString
            Me.txtNotes.Text = .Item("Notes").ToString

            'mark as Update mode
            Me.Expenses.CrudMode = DelortBusObjects.MiscExp.DBMode.Update
            Me.Expenses.ItemNo = Convert.ToInt32(.Item(0))
            Me.setTypeCombo(Convert.ToInt32(.Item("CategoryID")))
            Me.Expenses.Category = Convert.ToInt32(.Item("CategoryID"))
        End With

        Me.Loading = False

    End Sub

    Private Sub frmMiscExpAdd_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        Try
            If Me.Expenses.IsDirty Then

                If MessageBox.Show("Changes have been made and not save, Click OK to discard, CANCEL to go back and update", ProgramName, MessageBoxButtons.OKCancel) = DialogResult.OK Then
                    e.Cancel = False
                Else
                    e.Cancel = True
                End If

            End If

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub Expenses_Dirty() Handles Expenses.Dirty

        If Me.Expenses.IsValid Then
            Me.btnAdd.Enabled = True
        End If

    End Sub

    Private Sub Expenses_DisableView() Handles Expenses.DisableView

        Me.btnView.Enabled = False

    End Sub

    Private Sub cboType_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboType.SelectionChangeCommitted
        Me.Expenses.Category = Convert.ToInt32(Me.cboType.SelectedValue)

    End Sub

    Private Sub setTypeCombo(ByVal TypeID As Int32)

        Dim Index As Integer = 0

        For Each Item As DataRowView In Me.cboType.Items

            If CInt(Item.Item(0)) = TypeID Then
                cboType.SelectedIndex = Index
                Exit Sub
            End If

            Index += 1

        Next

    End Sub
End Class
