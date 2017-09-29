
#Region " Options "

Option Explicit On
Option Strict On

#End Region

Public Class frmEditMiscCat

    Inherits System.Windows.Forms.Form
    Private IsDirty As Boolean = False
    Private IsLoading As Boolean = True

    Public MiscCat As DelortBusObjects.MiscCategory


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
    Friend WithEvents grpActionButtons As System.Windows.Forms.GroupBox
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents txtCategoryName As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmEditMiscCat))
        Me.grpActionButtons = New System.Windows.Forms.GroupBox
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.txtCategoryName = New System.Windows.Forms.TextBox
        Me.grpActionButtons.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpActionButtons
        '
        Me.grpActionButtons.Controls.Add(Me.btnSave)
        Me.grpActionButtons.Controls.Add(Me.btnExit)
        Me.grpActionButtons.Location = New System.Drawing.Point(12, 48)
        Me.grpActionButtons.Name = "grpActionButtons"
        Me.grpActionButtons.Size = New System.Drawing.Size(212, 48)
        Me.grpActionButtons.TabIndex = 3
        Me.grpActionButtons.TabStop = False
        '
        'btnSave
        '
        Me.btnSave.BackColor = System.Drawing.Color.AliceBlue
        Me.btnSave.Enabled = False
        Me.btnSave.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Image = CType(resources.GetObject("btnSave.Image"), System.Drawing.Image)
        Me.btnSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnSave.Location = New System.Drawing.Point(16, 16)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(69, 22)
        Me.btnSave.TabIndex = 7
        Me.btnSave.Text = "&Save"
        Me.btnSave.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.AliceBlue
        Me.btnExit.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.Image = CType(resources.GetObject("btnExit.Image"), System.Drawing.Image)
        Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnExit.Location = New System.Drawing.Point(124, 16)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(69, 22)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "E&xit"
        Me.btnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtCategoryName
        '
        Me.txtCategoryName.Location = New System.Drawing.Point(12, 16)
        Me.txtCategoryName.Name = "txtCategoryName"
        Me.txtCategoryName.Size = New System.Drawing.Size(212, 22)
        Me.txtCategoryName.TabIndex = 4
        Me.txtCategoryName.Text = ""
        '
        'frmEditMiscCat
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 15)
        Me.ClientSize = New System.Drawing.Size(230, 108)
        Me.Controls.Add(Me.txtCategoryName)
        Me.Controls.Add(Me.grpActionButtons)
        Me.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmEditMiscCat"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Edit/Add Misc Category"
        Me.grpActionButtons.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmEditMiscCat_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If Me.MiscCat.CrudMode = DelortBusObjects.MiscCategory.DBMode.Update Then
            Me.txtCategoryName.Text = MiscCat.Description
            Me.IsDirty = False
        End If

        Me.IsLoading = False

    End Sub

    Private Sub txtCategoryName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCategoryName.TextChanged

        If Me.txtCategoryName.Text.Trim.Length > 1 And Me.IsLoading = False Then
            Me.btnSave.Enabled = True
            Me.MiscCat.Description = Me.txtCategoryName.Text.Trim
            Me.IsDirty = True
        Else
            Me.btnSave.Enabled = False
        End If

    End Sub

    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click

        Try
            Me.MiscCat.InsertUpdateDelete()
            IsDirty = False
            Me.Close()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.IsDirty = False

        End Try

    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click

        Try
            If Me.IsDirty = True Then
                If MessageBox.Show("Changes have been made and not saved, are you sure you want to exit?", ProgramName, MessageBoxButtons.YesNo, MessageBoxIcon.Information) = DialogResult.Yes Then
                    Me.Close()
                End If
            Else
                Me.Close()
            End If

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

End Class
