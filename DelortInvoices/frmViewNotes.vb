Option Strict On
Imports System.Configuration.ConfigurationSettings

Public Class frmViewNotes
    Inherits frmBaseWindow

    Private InvoiceNo As Integer
    Private InvoiceNote As DelortBusObjects.InvoiceNote
    Public dsNotes As DataSet

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal ds As DataSet, ByVal InvoiceNum As Integer)

        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Me.InvoiceNo = InvoiceNUm
        Me.dsNotes = ds



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
    Friend WithEvents dgNotes As System.Windows.Forms.DataGrid
    Friend WithEvents cmMain As System.Windows.Forms.ContextMenu
    Friend WithEvents cmEdit As System.Windows.Forms.MenuItem
    Friend WithEvents cmDelete As System.Windows.Forms.MenuItem
    Friend WithEvents grpFooter As System.Windows.Forms.GroupBox
    Friend WithEvents btnOpen As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmViewNotes))
        Me.grpMain = New System.Windows.Forms.GroupBox()
        Me.dgNotes = New System.Windows.Forms.DataGrid()
        Me.cmMain = New System.Windows.Forms.ContextMenu()
        Me.cmEdit = New System.Windows.Forms.MenuItem()
        Me.cmDelete = New System.Windows.Forms.MenuItem()
        Me.grpFooter = New System.Windows.Forms.GroupBox()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnOpen = New System.Windows.Forms.Button()
        Me.grpMain.SuspendLayout()
        CType(Me.dgNotes, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpFooter.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpMain
        '
        Me.grpMain.BackColor = System.Drawing.Color.Transparent
        Me.grpMain.Controls.Add(Me.dgNotes)
        Me.grpMain.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpMain.Location = New System.Drawing.Point(8, 8)
        Me.grpMain.Name = "grpMain"
        Me.grpMain.Size = New System.Drawing.Size(480, 240)
        Me.grpMain.TabIndex = 0
        Me.grpMain.TabStop = False
        '
        'dgNotes
        '
        Me.dgNotes.AlternatingBackColor = System.Drawing.Color.Lavender
        Me.dgNotes.BackColor = System.Drawing.Color.WhiteSmoke
        Me.dgNotes.BackgroundColor = System.Drawing.Color.LightGray
        Me.dgNotes.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.dgNotes.CaptionBackColor = System.Drawing.Color.LightSteelBlue
        Me.dgNotes.CaptionFont = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.dgNotes.CaptionForeColor = System.Drawing.Color.MidnightBlue
        Me.dgNotes.ContextMenu = Me.cmMain
        Me.dgNotes.DataMember = ""
        Me.dgNotes.FlatMode = True
        Me.dgNotes.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dgNotes.ForeColor = System.Drawing.Color.MidnightBlue
        Me.dgNotes.GridLineColor = System.Drawing.Color.Gainsboro
        Me.dgNotes.GridLineStyle = System.Windows.Forms.DataGridLineStyle.None
        Me.dgNotes.HeaderBackColor = System.Drawing.Color.MidnightBlue
        Me.dgNotes.HeaderFont = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.dgNotes.HeaderForeColor = System.Drawing.Color.WhiteSmoke
        Me.dgNotes.LinkColor = System.Drawing.Color.Teal
        Me.dgNotes.Location = New System.Drawing.Point(11, 16)
        Me.dgNotes.Name = "dgNotes"
        Me.dgNotes.ParentRowsBackColor = System.Drawing.Color.Gainsboro
        Me.dgNotes.ParentRowsForeColor = System.Drawing.Color.MidnightBlue
        Me.dgNotes.SelectionBackColor = System.Drawing.Color.CadetBlue
        Me.dgNotes.SelectionForeColor = System.Drawing.Color.WhiteSmoke
        Me.dgNotes.Size = New System.Drawing.Size(456, 216)
        Me.dgNotes.TabIndex = 0
        '
        'cmMain
        '
        Me.cmMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.cmEdit, Me.cmDelete})
        '
        'cmEdit
        '
        Me.cmEdit.Index = 0
        Me.cmEdit.Text = "Edit Note"
        '
        'cmDelete
        '
        Me.cmDelete.Index = 1
        Me.cmDelete.Text = "Delete Note"
        '
        'grpFooter
        '
        Me.grpFooter.BackColor = System.Drawing.Color.Transparent
        Me.grpFooter.Controls.Add(Me.btnDelete)
        Me.grpFooter.Controls.Add(Me.btnExit)
        Me.grpFooter.Controls.Add(Me.btnOpen)
        Me.grpFooter.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpFooter.Location = New System.Drawing.Point(8, 248)
        Me.grpFooter.Name = "grpFooter"
        Me.grpFooter.Size = New System.Drawing.Size(480, 48)
        Me.grpFooter.TabIndex = 1
        Me.grpFooter.TabStop = False
        '
        'btnDelete
        '
        Me.btnDelete.BackColor = System.Drawing.Color.AliceBlue
        Me.btnDelete.Image = CType(resources.GetObject("btnDelete.Image"), System.Drawing.Image)
        Me.btnDelete.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDelete.Location = New System.Drawing.Point(184, 18)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 24)
        Me.btnDelete.TabIndex = 2
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnDelete.UseVisualStyleBackColor = False
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.AliceBlue
        Me.btnExit.Image = CType(resources.GetObject("btnExit.Image"), System.Drawing.Image)
        Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnExit.Location = New System.Drawing.Point(24, 18)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 24)
        Me.btnExit.TabIndex = 1
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'btnOpen
        '
        Me.btnOpen.BackColor = System.Drawing.Color.AliceBlue
        Me.btnOpen.Image = CType(resources.GetObject("btnOpen.Image"), System.Drawing.Image)
        Me.btnOpen.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnOpen.Location = New System.Drawing.Point(344, 18)
        Me.btnOpen.Name = "btnOpen"
        Me.btnOpen.Size = New System.Drawing.Size(80, 24)
        Me.btnOpen.TabIndex = 0
        Me.btnOpen.Text = "&Open"
        Me.btnOpen.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnOpen.UseVisualStyleBackColor = False
        '
        'frmViewNotes
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(502, 300)
        Me.Controls.Add(Me.grpFooter)
        Me.Controls.Add(Me.grpMain)
        Me.Font = New System.Drawing.Font("MS Outlook", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmViewNotes"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "View Notes"
        Me.grpMain.ResumeLayout(False)
        CType(Me.dgNotes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpFooter.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Sub SetUpDataGrid()

        With Me.dgNotes
            .ReadOnly = True
            .Text = "Notes"
        End With

        Me.dgNotes.TableStyles.Clear()

        Dim ts As New DataGridTableStyle
        Dim col As New DataGridTextBoxColumn


        With ts
            .MappingName = "Notes"
            .AllowSorting = True
            .GridLineColor = Color.DarkBlue
            .HeaderBackColor = Color.DarkBlue
            .HeaderForeColor = Color.White
        End With



        With col
            .MappingName = "NoteNo"
            .Width = 0
            ts.GridColumnStyles.Add(col)
        End With


        col = New DataGridTextBoxColumn
        With col
            .MappingName = "Notes"
            .HeaderText = "Notes"
            .Width = 450
            .Alignment = HorizontalAlignment.Left
            ts.GridColumnStyles.Add(col)
        End With

        Me.dgNotes.TableStyles.Add(ts)

    End Sub
    Private Sub LoadData()

        'Me.InvoiceNote = New DelortBusObjects.InvoiceNote(g_settings, InvoiceNo)
        'Dim ds As DataSet = Me.InvoiceNote.GetAllInvoiceNotes(Me.InvoiceNo)
        'If ds.Tables(0).Rows.Count = 0 Then
        '    Throw New Exception("No Notes for this Invoice")
        'End If
        Me.dgNotes.DataSource = Me.dsNotes.Tables(0)

        Me.SetUpDataGrid()

    End Sub

    Private Sub cmEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmEdit.Click
        Try

            Me.EditNote()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try


    End Sub
    Private Sub EditNote()

        If Me.dgNotes.VisibleRowCount > 0 Then

            Dim NoteNo As Integer
            NoteNo = Convert.ToInt32(Me.dgNotes(Me.dgNotes.CurrentRowIndex, 0))

            Dim frmEditNote As New frmEditNote(NoteNo)
            frmEditNote.CallingForm = Me

            frmEditNote.ShowDialog()

            frmEditNote.Dispose()

            Me.LoadData()

        End If

    End Sub

    Private Sub frmViewNotes_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.InvoiceNote = New DelortBusObjects.InvoiceNote(g_settings, InvoiceNo)
            Me.LoadData()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub btnOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpen.Click

        Try
            EditNote()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub cmDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmDelete.Click
        Try
            DeleteNote()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub
    Public Sub RefreshData()

        Me.dsNotes = InvoiceNote.GetAllInvoiceNotes(Me.InvoiceNo)
        Me.LoadData()

    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub
    Private Sub DeleteNote()

        If Me.dgNotes.VisibleRowCount > 0 Then

            Dim NoteNo As Integer

            NoteNo = Convert.ToInt32(Me.dgNotes(Me.dgNotes.CurrentRowIndex, 0))

            If MessageBox.Show("Are you sure?", ProgramName, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.No Then
                Exit Sub
            End If

            Me.InvoiceNote.DeleteNoteByNoteNo(NoteNo)

            Me.RefreshData()

        End If

    End Sub

    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            DeleteNote()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub dgNotes_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgNotes.DoubleClick
        Try

            Me.EditNote()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
End Class
