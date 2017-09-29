
#Region " Options "

Option Explicit On
Option Strict On

#End Region

Imports DelortBusObjects
Imports System.Configuration.ConfigurationSettings

Public Class frmNotes

    Inherits frmBaseWindow

    Public InvoiceNo As Integer
    Private WithEvents InvoiceNote As InvoiceNote
    Private Loading As Boolean = True

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal invNo As Integer)

        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Me.InvoiceNo = invNo

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
    Friend WithEvents grpHeader As System.Windows.Forms.GroupBox
    Friend WithEvents grpFooter As System.Windows.Forms.GroupBox
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents txtNote As System.Windows.Forms.TextBox
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmNotes))
        Me.grpHeader = New System.Windows.Forms.GroupBox()
        Me.txtNote = New System.Windows.Forms.TextBox()
        Me.grpFooter = New System.Windows.Forms.GroupBox()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.btnView = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.grpHeader.SuspendLayout()
        Me.grpFooter.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpHeader
        '
        Me.grpHeader.BackColor = System.Drawing.Color.Transparent
        Me.grpHeader.Controls.Add(Me.txtNote)
        Me.grpHeader.Location = New System.Drawing.Point(2, 8)
        Me.grpHeader.Name = "grpHeader"
        Me.grpHeader.Size = New System.Drawing.Size(400, 120)
        Me.grpHeader.TabIndex = 0
        Me.grpHeader.TabStop = False
        Me.grpHeader.Text = "Type Notes Here"
        '
        'txtNote
        '
        Me.txtNote.Location = New System.Drawing.Point(8, 16)
        Me.txtNote.Multiline = True
        Me.txtNote.Name = "txtNote"
        Me.txtNote.Size = New System.Drawing.Size(384, 96)
        Me.txtNote.TabIndex = 0
        '
        'grpFooter
        '
        Me.grpFooter.BackColor = System.Drawing.Color.Transparent
        Me.grpFooter.Controls.Add(Me.btnAdd)
        Me.grpFooter.Controls.Add(Me.btnView)
        Me.grpFooter.Controls.Add(Me.btnExit)
        Me.grpFooter.Location = New System.Drawing.Point(0, 133)
        Me.grpFooter.Name = "grpFooter"
        Me.grpFooter.Size = New System.Drawing.Size(400, 48)
        Me.grpFooter.TabIndex = 1
        Me.grpFooter.TabStop = False
        '
        'btnAdd
        '
        Me.btnAdd.BackColor = System.Drawing.Color.AliceBlue
        Me.btnAdd.Enabled = False
        Me.btnAdd.Image = CType(resources.GetObject("btnAdd.Image"), System.Drawing.Image)
        Me.btnAdd.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnAdd.Location = New System.Drawing.Point(295, 16)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(83, 24)
        Me.btnAdd.TabIndex = 2
        Me.btnAdd.Text = "&Save"
        Me.btnAdd.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnAdd.UseVisualStyleBackColor = False
        '
        'btnView
        '
        Me.btnView.BackColor = System.Drawing.Color.AliceBlue
        Me.btnView.Image = CType(resources.GetObject("btnView.Image"), System.Drawing.Image)
        Me.btnView.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnView.Location = New System.Drawing.Point(153, 16)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(95, 24)
        Me.btnView.TabIndex = 1
        Me.btnView.Text = "&View All "
        Me.btnView.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnView.UseVisualStyleBackColor = False
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.AliceBlue
        Me.btnExit.Image = CType(resources.GetObject("btnExit.Image"), System.Drawing.Image)
        Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnExit.Location = New System.Drawing.Point(23, 17)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(83, 24)
        Me.btnExit.TabIndex = 0
        Me.btnExit.Text = "E&xit"
        Me.btnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'lblStatus
        '
        Me.lblStatus.BackColor = System.Drawing.Color.Transparent
        Me.lblStatus.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblStatus.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus.Location = New System.Drawing.Point(0, 188)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(400, 24)
        Me.lblStatus.TabIndex = 2
        Me.lblStatus.Text = " Notes Data Entry "
        '
        'frmNotes
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 15)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(406, 220)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.grpFooter)
        Me.Controls.Add(Me.grpHeader)
        Me.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmNotes"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add Notes"
        Me.grpHeader.ResumeLayout(False)
        Me.grpHeader.PerformLayout()
        Me.grpFooter.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub txtNote_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNote.TextChanged

        If Loading = True Then Exit Sub

        Me.InvoiceNote.Note = Me.txtNote.Text.Trim

        Me.lblStatus.ForeColor = System.Drawing.Color.Black
        Me.lblStatus.Text = "Editing Notes"

    End Sub

    Private Sub InvoiceNote_Valid(ByVal IsValid As Boolean) Handles InvoiceNote.Valid

        Me.btnAdd.Enabled = IsValid

    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click

        Try
            Me.InvoiceNote.InsertUpdateDelete()
            Loading = True
            Me.txtNote.Text = vbNullString
            Me.lblStatus.ForeColor = System.Drawing.Color.Red
            Me.lblStatus.Text = "Note Added"
            Loading = False
            Me.InvoiceNote.IsDirty = False

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub frmNotes_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.InvoiceNote = New InvoiceNote(g_settings, InvoiceNo)
            Me.Text = "Add Notes For Invoice# " & Me.InvoiceNo.ToString

            Me.InvoiceNote.IsDirty = False
            Me.Loading = False

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click


        Try
            Dim ds As DataSet = Me.InvoiceNote.GetAllInvoiceNotes(Me.InvoiceNo)

            If ds.Tables(0).Rows.Count = 0 Then
                Me.lblStatus.ForeColor = System.Drawing.Color.Black
                Me.lblStatus.Text = "No Notes for This Invoice"
            Else

                Dim frmViewAll As New frmViewNotes(ds, InvoiceNo)

                frmViewAll.ShowDialog()

                frmViewAll.Dispose()
            End If

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try

            Me.Close()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub frmNotes_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        Try
            If Me.InvoiceNote.IsDirty Then
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

End Class
