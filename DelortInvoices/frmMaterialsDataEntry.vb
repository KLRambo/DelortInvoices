Option Strict On

Imports System.Configuration.ConfigurationSettings

Public Class frmMaterialsDataEntry
    Inherits frmBaseWindow

    Private WithEvents Materials As New DelortBusObjects.Materials(g_settings)
    Public dsMaterialItem As DataSet
    Private Loading As Boolean = False
    Public CallingForm As frmViewMaterials

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
    Friend WithEvents txtMaterialDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtMaterialCost As System.Windows.Forms.TextBox
    Friend WithEvents lblDesc As System.Windows.Forms.Label
    Friend WithEvents lblCost As System.Windows.Forms.Label
    Friend WithEvents lblNotes As System.Windows.Forms.Label
    Friend WithEvents grpFooter As System.Windows.Forms.GroupBox
    Friend WithEvents lblDate As System.Windows.Forms.Label
    Friend WithEvents dtpMaterial As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents txtNotes As System.Windows.Forms.TextBox
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMaterialsDataEntry))
        Me.grpMain = New System.Windows.Forms.GroupBox()
        Me.dtpMaterial = New System.Windows.Forms.DateTimePicker()
        Me.lblDate = New System.Windows.Forms.Label()
        Me.lblNotes = New System.Windows.Forms.Label()
        Me.lblCost = New System.Windows.Forms.Label()
        Me.lblDesc = New System.Windows.Forms.Label()
        Me.txtNotes = New System.Windows.Forms.TextBox()
        Me.txtMaterialCost = New System.Windows.Forms.TextBox()
        Me.txtMaterialDesc = New System.Windows.Forms.TextBox()
        Me.grpFooter = New System.Windows.Forms.GroupBox()
        Me.btnView = New System.Windows.Forms.Button()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.grpMain.SuspendLayout()
        Me.grpFooter.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpMain
        '
        Me.grpMain.BackColor = System.Drawing.Color.Transparent
        Me.grpMain.Controls.Add(Me.dtpMaterial)
        Me.grpMain.Controls.Add(Me.lblDate)
        Me.grpMain.Controls.Add(Me.lblNotes)
        Me.grpMain.Controls.Add(Me.lblCost)
        Me.grpMain.Controls.Add(Me.lblDesc)
        Me.grpMain.Controls.Add(Me.txtNotes)
        Me.grpMain.Controls.Add(Me.txtMaterialCost)
        Me.grpMain.Controls.Add(Me.txtMaterialDesc)
        Me.grpMain.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpMain.Location = New System.Drawing.Point(8, 8)
        Me.grpMain.Name = "grpMain"
        Me.grpMain.Size = New System.Drawing.Size(336, 200)
        Me.grpMain.TabIndex = 0
        Me.grpMain.TabStop = False
        '
        'dtpMaterial
        '
        Me.dtpMaterial.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpMaterial.Location = New System.Drawing.Point(102, 18)
        Me.dtpMaterial.Name = "dtpMaterial"
        Me.dtpMaterial.Size = New System.Drawing.Size(98, 21)
        Me.dtpMaterial.TabIndex = 7
        '
        'lblDate
        '
        Me.lblDate.Location = New System.Drawing.Point(16, 18)
        Me.lblDate.Name = "lblDate"
        Me.lblDate.Size = New System.Drawing.Size(48, 16)
        Me.lblDate.TabIndex = 6
        Me.lblDate.Text = "Date:"
        '
        'lblNotes
        '
        Me.lblNotes.Location = New System.Drawing.Point(16, 140)
        Me.lblNotes.Name = "lblNotes"
        Me.lblNotes.Size = New System.Drawing.Size(48, 16)
        Me.lblNotes.TabIndex = 5
        Me.lblNotes.Text = "Notes:"
        '
        'lblCost
        '
        Me.lblCost.Location = New System.Drawing.Point(16, 107)
        Me.lblCost.Name = "lblCost"
        Me.lblCost.Size = New System.Drawing.Size(40, 16)
        Me.lblCost.TabIndex = 4
        Me.lblCost.Text = "Cost:"
        '
        'lblDesc
        '
        Me.lblDesc.Location = New System.Drawing.Point(16, 51)
        Me.lblDesc.Name = "lblDesc"
        Me.lblDesc.Size = New System.Drawing.Size(80, 16)
        Me.lblDesc.TabIndex = 3
        Me.lblDesc.Text = "Description:"
        '
        'txtNotes
        '
        Me.txtNotes.Location = New System.Drawing.Point(102, 140)
        Me.txtNotes.Multiline = True
        Me.txtNotes.Name = "txtNotes"
        Me.txtNotes.Size = New System.Drawing.Size(218, 45)
        Me.txtNotes.TabIndex = 2
        '
        'txtMaterialCost
        '
        Me.txtMaterialCost.Location = New System.Drawing.Point(102, 107)
        Me.txtMaterialCost.Name = "txtMaterialCost"
        Me.txtMaterialCost.Size = New System.Drawing.Size(52, 21)
        Me.txtMaterialCost.TabIndex = 1
        '
        'txtMaterialDesc
        '
        Me.txtMaterialDesc.Location = New System.Drawing.Point(102, 51)
        Me.txtMaterialDesc.Multiline = True
        Me.txtMaterialDesc.Name = "txtMaterialDesc"
        Me.txtMaterialDesc.Size = New System.Drawing.Size(218, 45)
        Me.txtMaterialDesc.TabIndex = 0
        '
        'grpFooter
        '
        Me.grpFooter.BackColor = System.Drawing.Color.Transparent
        Me.grpFooter.Controls.Add(Me.btnView)
        Me.grpFooter.Controls.Add(Me.btnAdd)
        Me.grpFooter.Controls.Add(Me.btnExit)
        Me.grpFooter.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpFooter.Location = New System.Drawing.Point(8, 208)
        Me.grpFooter.Name = "grpFooter"
        Me.grpFooter.Size = New System.Drawing.Size(336, 48)
        Me.grpFooter.TabIndex = 1
        Me.grpFooter.TabStop = False
        '
        'btnView
        '
        Me.btnView.BackColor = System.Drawing.Color.AliceBlue
        Me.btnView.Image = CType(resources.GetObject("btnView.Image"), System.Drawing.Image)
        Me.btnView.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnView.Location = New System.Drawing.Point(123, 14)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(88, 24)
        Me.btnView.TabIndex = 7
        Me.btnView.Text = "&View All"
        Me.btnView.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.btnView, "View Materials Report")
        Me.btnView.UseVisualStyleBackColor = False
        '
        'btnAdd
        '
        Me.btnAdd.BackColor = System.Drawing.Color.AliceBlue
        Me.btnAdd.Enabled = False
        Me.btnAdd.Image = CType(resources.GetObject("btnAdd.Image"), System.Drawing.Image)
        Me.btnAdd.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnAdd.Location = New System.Drawing.Point(232, 14)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(88, 24)
        Me.btnAdd.TabIndex = 6
        Me.btnAdd.Text = "   &Save"
        Me.ToolTip1.SetToolTip(Me.btnAdd, "Save Changes")
        Me.btnAdd.UseVisualStyleBackColor = False
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.AliceBlue
        Me.btnExit.Image = CType(resources.GetObject("btnExit.Image"), System.Drawing.Image)
        Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnExit.Location = New System.Drawing.Point(14, 14)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(88, 24)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.ToolTip1.SetToolTip(Me.btnExit, "Exit Form")
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'lblStatus
        '
        Me.lblStatus.BackColor = System.Drawing.Color.Transparent
        Me.lblStatus.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblStatus.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus.Location = New System.Drawing.Point(8, 264)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(336, 24)
        Me.lblStatus.TabIndex = 2
        Me.lblStatus.Text = "Materials Data Entry"
        '
        'frmMaterialsDataEntry
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 15)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(350, 292)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.grpFooter)
        Me.Controls.Add(Me.grpMain)
        Me.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmMaterialsDataEntry"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Data Entry For Materials"
        Me.grpMain.ResumeLayout(False)
        Me.grpMain.PerformLayout()
        Me.grpFooter.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub txtMaterialCost_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMaterialCost.TextChanged
        If Me.Loading = True Then Exit Sub

        If IsNumeric(Me.txtMaterialCost.Text) Then
            Me.Materials.MaterialCost = Convert.ToDouble(Me.txtMaterialCost.Text)
        Else
            Me.Materials.MaterialCost = 0
        End If

        If Me.FormMode = Mode.Add Then
            Me.lblStatus.Text = "Adding Materials..."
        Else
            Me.lblStatus.Text = "Editing Materials..."
        End If

    End Sub

    Private Sub txtMaterialCost_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaterialCost.KeyPress

        e.Handled = Utilities.NumericOnly(e, Me.txtMaterialCost)

    End Sub


    Private Sub txtMaterialDesc_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMaterialDesc.TextChanged

        If Me.Loading = True Then Exit Sub

        Me.Materials.MaterialDesc = Me.txtMaterialDesc.Text.Trim
        Me.lblStatus.ForeColor = System.Drawing.Color.Black
        If Me.FormMode = Mode.Add Then
            Me.lblStatus.Text = "Adding Materials..."
        Else
            Me.lblStatus.Text = "Editing Materials..."
        End If


    End Sub

    Private Sub Materials_Valid(ByVal IsValid As Boolean) Handles Materials.Valid

        Me.btnAdd.Enabled = IsValid

    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try

            Me.Close()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub dtpMaterial_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpMaterial.ValueChanged

        Try

            Me.Materials.PurchaseDate = Convert.ToDateTime(Me.dtpMaterial.Value)

        Catch ex As Exception

            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub txtNotes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNotes.TextChanged

        If Me.Loading = True Then Exit Sub

        Me.Materials.Notes = Me.txtNotes.Text

        Me.lblStatus.ForeColor = System.Drawing.Color.Black

        If Me.FormMode = Mode.Add Then
            Me.lblStatus.Text = "Adding Materials..."
        Else
            Me.lblStatus.Text = "Editing Materials..."
        End If

    End Sub

    Private Sub btnAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAdd.Click

        Try
            Me.Materials.InsertUpdateDelete()
            If Me.FormMode = Mode.Add Then
                ResetForm()
            End If
            If Not Me.CallingForm Is Nothing Then
                Me.CallingForm.LoadData()

            End If

            Me.lblStatus.ForeColor = System.Drawing.Color.Red
            Me.lblStatus.Text = "Materials Updated!"
            Me.Materials.IsDirty = False
            Me.btnAdd.Enabled = False

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
    Private Sub ResetForm()

        Me.txtMaterialCost.Text = vbNullString
        Me.txtMaterialDesc.Text = vbNullString
        Me.dtpMaterial.Value = Now
        Me.txtNotes.Text = vbNullString

    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click

        Try
            Utilities.ViewMaterialReport(Me.ParentForm)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub frmMaterialsDataEntry_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If Me.FormMode = Mode.Edit Then
                LoadData()
                Me.btnAdd.Enabled = False
                Me.Materials.IsValid = True
            End If

            Me.Loading = False
            Me.Materials.IsDirty = False

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub
    Private Sub LoadData()

        With Me.dsMaterialItem.Tables(0).Rows(0)
            Me.dtpMaterial.Value = Convert.ToDateTime(.Item("DatePurchased"))
            Me.txtMaterialCost.Text = .Item("MaterialCost").ToString
            Me.txtMaterialDesc.Text = .Item("MaterialDesc").ToString
            Me.txtNotes.Text = .Item("Notes").ToString
            'mark as Update mode
            Me.Materials.CrudMode = DelortBusObjects.Materials.DBMode.Update
            Me.Materials.MaterialNo = Convert.ToInt32(.Item(0))

        End With

        Me.Loading = False

    End Sub

    Private Sub Materials_DisableView() Handles Materials.DisableView

        Me.btnView.Enabled = False

    End Sub


    Private Sub frmMaterialsDataEntry_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Try
            If Me.Materials.IsDirty Then

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

    Private Sub Materials_Dirty1() Handles Materials.Dirty
        If Me.Materials.IsValid Then
            Me.btnAdd.Enabled = True
        End If
    End Sub
End Class
