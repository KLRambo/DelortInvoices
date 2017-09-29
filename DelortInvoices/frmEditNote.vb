
#Region " Options "

Option Explicit On
Option Strict On

#End Region

Imports DelortBusObjects
Imports System.Configuration.ConfigurationSettings

Public Class frmEditNote

    Inherits frmBaseWindow

    Private NoteNumber As Integer
    Private WithEvents Note As New DelortBusObjects.InvoiceNote(g_Settings, 0)
    Public CallingForm As frmViewNotes
    Private Loading As Boolean = True

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal NoteNo As Integer)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        NoteNumber = NoteNo

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
    Friend WithEvents grpFooter As System.Windows.Forms.GroupBox
    Friend WithEvents grpHeader As System.Windows.Forms.GroupBox
    Friend WithEvents txtNote As System.Windows.Forms.TextBox
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents lblNotes As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEditNote))
        Me.grpFooter = New System.Windows.Forms.GroupBox()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.grpHeader = New System.Windows.Forms.GroupBox()
        Me.txtNote = New System.Windows.Forms.TextBox()
        Me.lblNotes = New System.Windows.Forms.Label()
        Me.grpFooter.SuspendLayout()
        Me.grpHeader.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpFooter
        '
        Me.grpFooter.BackColor = System.Drawing.Color.Transparent
        Me.grpFooter.Controls.Add(Me.btnUpdate)
        Me.grpFooter.Controls.Add(Me.btnExit)
        Me.grpFooter.Location = New System.Drawing.Point(0, 127)
        Me.grpFooter.Name = "grpFooter"
        Me.grpFooter.Size = New System.Drawing.Size(408, 40)
        Me.grpFooter.TabIndex = 4
        Me.grpFooter.TabStop = False
        '
        'btnUpdate
        '
        Me.btnUpdate.BackColor = System.Drawing.Color.AliceBlue
        Me.btnUpdate.Enabled = False
        Me.btnUpdate.Image = CType(resources.GetObject("btnUpdate.Image"), System.Drawing.Image)
        Me.btnUpdate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnUpdate.Location = New System.Drawing.Point(326, 10)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(72, 24)
        Me.btnUpdate.TabIndex = 4
        Me.btnUpdate.Text = "&Save"
        Me.btnUpdate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnUpdate.UseVisualStyleBackColor = False
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.AliceBlue
        Me.btnExit.Image = CType(resources.GetObject("btnExit.Image"), System.Drawing.Image)
        Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnExit.Location = New System.Drawing.Point(17, 10)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(71, 24)
        Me.btnExit.TabIndex = 3
        Me.btnExit.Text = "E&xit"
        Me.btnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'grpHeader
        '
        Me.grpHeader.BackColor = System.Drawing.Color.Transparent
        Me.grpHeader.Controls.Add(Me.txtNote)
        Me.grpHeader.Location = New System.Drawing.Point(0, 4)
        Me.grpHeader.Name = "grpHeader"
        Me.grpHeader.Size = New System.Drawing.Size(408, 120)
        Me.grpHeader.TabIndex = 3
        Me.grpHeader.TabStop = False
        Me.grpHeader.Text = "Type Notes Here"
        '
        'txtNote
        '
        Me.txtNote.Location = New System.Drawing.Point(10, 14)
        Me.txtNote.Multiline = True
        Me.txtNote.Name = "txtNote"
        Me.txtNote.Size = New System.Drawing.Size(390, 96)
        Me.txtNote.TabIndex = 2
        '
        'lblNotes
        '
        Me.lblNotes.BackColor = System.Drawing.Color.Transparent
        Me.lblNotes.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNotes.Location = New System.Drawing.Point(0, 176)
        Me.lblNotes.Name = "lblNotes"
        Me.lblNotes.Size = New System.Drawing.Size(408, 32)
        Me.lblNotes.TabIndex = 5
        Me.lblNotes.Text = "Notes Data Entry"
        '
        'frmEditNote
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 15)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(416, 204)
        Me.Controls.Add(Me.lblNotes)
        Me.Controls.Add(Me.grpFooter)
        Me.Controls.Add(Me.grpHeader)
        Me.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmEditNote"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Edit Note"
        Me.grpFooter.ResumeLayout(False)
        Me.grpHeader.ResumeLayout(False)
        Me.grpHeader.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click

        Try

            Me.Note.InsertUpdateDelete()
            If Not IsNothing(Me.CallingForm) Then
                Me.CallingForm.RefreshData()
            End If

            Me.btnUpdate.Enabled = False
            Me.lblNotes.ForeColor = System.Drawing.Color.Red
            Me.lblNotes.Text = "Note Updated"

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub frmEditNote_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Note.CrudMode = InvoiceNote.DBMode.Update
            Me.Note.NoteNo = Me.NoteNumber
            Me.txtNote.Text = Me.Note.GetNote(Me.NoteNumber)

            Me.Loading = False
            Me.Note.IsDirty = False

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


    End Sub

    Private Sub txtNote_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNote.TextChanged

        Me.Note.Note = Me.txtNote.Text.Trim
        Me.lblNotes.ForeColor = System.Drawing.Color.Black
        Me.lblNotes.Text = "Editing Notes..."

    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click

        Try

            Me.Close()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub Note_Valid(ByVal IsValid As Boolean) Handles Note.Valid
        If Me.Loading = True Then Exit Sub

        Me.btnUpdate.Enabled = IsValid

    End Sub

    Private Sub frmEditNote_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        If Me.Note.IsDirty Then
            If MessageBox.Show("Changes have been made and not save, Click OK to discard, CANCEL to go back and update", ProgramName, MessageBoxButtons.OKCancel) = DialogResult.OK Then
                e.Cancel = False
            Else
                e.Cancel = True

            End If
        End If

    End Sub
End Class
