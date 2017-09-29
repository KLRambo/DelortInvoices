
#Region " Options "

Option Explicit On
Option Strict On

#End Region

Imports System.Configuration.ConfigurationSettings

Public Class frmMiscExpRpt

    Inherits frmBaseWindow

    Private _Ds As DataSet
    Private _Date1 As Date = Nothing
    Private _Date2 As Date = Nothing
    Private _ReportType As ReportType = 0

    Friend WithEvents toolStrip As System.Windows.Forms.ToolStrip
    Friend WithEvents tsRun As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsOpen As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsDelete As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsPrint As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents tsCategory As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsMonthly As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsAddNew As System.Windows.Forms.ToolStripButton
    Private WithEvents Expenses As New DelortBusObjects.MiscExp(g_Settings)

    Enum ReportType
        CATEGORY = 0
        MONTHLY = 1
    End Enum

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
    Friend WithEvents grpHeader As System.Windows.Forms.GroupBox
    Friend WithEvents lblTo As System.Windows.Forms.Label
    Friend WithEvents lblFrom As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmnuDelete As System.Windows.Forms.MenuItem
    Friend WithEvents cmnuMain As System.Windows.Forms.ContextMenu
    Friend WithEvents cmnuOpen As System.Windows.Forms.MenuItem
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents grpFooter2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtTotalCost As System.Windows.Forms.TextBox
    Friend WithEvents txtItems As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblTotalItems As System.Windows.Forms.Label
    Friend WithEvents ErrProv As System.Windows.Forms.ErrorProvider
    Friend WithEvents grpMain As System.Windows.Forms.GroupBox
    Friend WithEvents dgExps As System.Windows.Forms.DataGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMiscExpRpt))
        Me.cmnuDelete = New System.Windows.Forms.MenuItem()
        Me.cmnuMain = New System.Windows.Forms.ContextMenu()
        Me.cmnuOpen = New System.Windows.Forms.MenuItem()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.toolStrip = New System.Windows.Forms.ToolStrip()
        Me.tsRun = New System.Windows.Forms.ToolStripButton()
        Me.tsOpen = New System.Windows.Forms.ToolStripButton()
        Me.tsAddNew = New System.Windows.Forms.ToolStripButton()
        Me.tsDelete = New System.Windows.Forms.ToolStripButton()
        Me.tsPrint = New System.Windows.Forms.ToolStripDropDownButton()
        Me.tsCategory = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsMonthly = New System.Windows.Forms.ToolStripMenuItem()
        Me.grpMain = New System.Windows.Forms.GroupBox()
        Me.dgExps = New System.Windows.Forms.DataGrid()
        Me.grpHeader = New System.Windows.Forms.GroupBox()
        Me.lblTo = New System.Windows.Forms.Label()
        Me.lblFrom = New System.Windows.Forms.Label()
        Me.dtpTo = New System.Windows.Forms.DateTimePicker()
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker()
        Me.grpFooter2 = New System.Windows.Forms.GroupBox()
        Me.txtTotalCost = New System.Windows.Forms.TextBox()
        Me.txtItems = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblTotalItems = New System.Windows.Forms.Label()
        Me.ErrProv = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.toolStrip.SuspendLayout()
        Me.grpMain.SuspendLayout()
        CType(Me.dgExps, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpHeader.SuspendLayout()
        Me.grpFooter2.SuspendLayout()
        CType(Me.ErrProv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmnuDelete
        '
        Me.cmnuDelete.Index = 1
        Me.cmnuDelete.Text = "Delete Item"
        '
        'cmnuMain
        '
        Me.cmnuMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.cmnuOpen, Me.cmnuDelete})
        '
        'cmnuOpen
        '
        Me.cmnuOpen.Index = 0
        Me.cmnuOpen.Text = "Open Item"
        '
        'toolStrip
        '
        Me.toolStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsRun, Me.tsOpen, Me.tsAddNew, Me.tsDelete, Me.tsPrint})
        Me.toolStrip.Location = New System.Drawing.Point(0, 0)
        Me.toolStrip.Name = "toolStrip"
        Me.toolStrip.Size = New System.Drawing.Size(520, 25)
        Me.toolStrip.TabIndex = 14
        Me.toolStrip.Text = "ToolStrip1"
        '
        'tsRun
        '
        Me.tsRun.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsRun.Image = CType(resources.GetObject("tsRun.Image"), System.Drawing.Image)
        Me.tsRun.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsRun.Name = "tsRun"
        Me.tsRun.Size = New System.Drawing.Size(23, 22)
        Me.tsRun.ToolTipText = "Run"
        '
        'tsOpen
        '
        Me.tsOpen.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsOpen.Image = CType(resources.GetObject("tsOpen.Image"), System.Drawing.Image)
        Me.tsOpen.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsOpen.Name = "tsOpen"
        Me.tsOpen.Size = New System.Drawing.Size(23, 22)
        Me.tsOpen.ToolTipText = "Open Item"
        '
        'tsAddNew
        '
        Me.tsAddNew.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsAddNew.Image = CType(resources.GetObject("tsAddNew.Image"), System.Drawing.Image)
        Me.tsAddNew.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsAddNew.Name = "tsAddNew"
        Me.tsAddNew.Size = New System.Drawing.Size(23, 22)
        '
        'tsDelete
        '
        Me.tsDelete.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsDelete.Image = CType(resources.GetObject("tsDelete.Image"), System.Drawing.Image)
        Me.tsDelete.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsDelete.Name = "tsDelete"
        Me.tsDelete.Size = New System.Drawing.Size(23, 22)
        Me.tsDelete.ToolTipText = "Delete Item"
        '
        'tsPrint
        '
        Me.tsPrint.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsPrint.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsCategory, Me.tsMonthly})
        Me.tsPrint.Image = CType(resources.GetObject("tsPrint.Image"), System.Drawing.Image)
        Me.tsPrint.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsPrint.Name = "tsPrint"
        Me.tsPrint.Size = New System.Drawing.Size(29, 22)
        Me.tsPrint.ToolTipText = "Report Type"
        '
        'tsCategory
        '
        Me.tsCategory.Name = "tsCategory"
        Me.tsCategory.Size = New System.Drawing.Size(155, 22)
        Me.tsCategory.Text = "By Category"
        '
        'tsMonthly
        '
        Me.tsMonthly.Name = "tsMonthly"
        Me.tsMonthly.Size = New System.Drawing.Size(155, 22)
        Me.tsMonthly.Text = "Year By Month"
        '
        'grpMain
        '
        Me.grpMain.BackColor = System.Drawing.Color.Transparent
        Me.grpMain.Controls.Add(Me.dgExps)
        Me.grpMain.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpMain.Location = New System.Drawing.Point(9, 84)
        Me.grpMain.Name = "grpMain"
        Me.grpMain.Size = New System.Drawing.Size(506, 434)
        Me.grpMain.TabIndex = 9
        Me.grpMain.TabStop = False
        Me.grpMain.Text = "Expenses"
        '
        'dgExps
        '
        Me.dgExps.AlternatingBackColor = System.Drawing.Color.LightGray
        Me.dgExps.BackColor = System.Drawing.Color.Gainsboro
        Me.dgExps.BackgroundColor = System.Drawing.Color.Silver
        Me.dgExps.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.dgExps.CaptionBackColor = System.Drawing.Color.LightSteelBlue
        Me.dgExps.CaptionFont = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.dgExps.CaptionForeColor = System.Drawing.Color.MidnightBlue
        Me.dgExps.CaptionVisible = False
        Me.dgExps.ContextMenu = Me.cmnuMain
        Me.dgExps.DataMember = ""
        Me.dgExps.FlatMode = True
        Me.dgExps.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dgExps.ForeColor = System.Drawing.Color.Black
        Me.dgExps.GridLineColor = System.Drawing.Color.DimGray
        Me.dgExps.GridLineStyle = System.Windows.Forms.DataGridLineStyle.None
        Me.dgExps.HeaderBackColor = System.Drawing.Color.MidnightBlue
        Me.dgExps.HeaderFont = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.dgExps.HeaderForeColor = System.Drawing.Color.White
        Me.dgExps.LinkColor = System.Drawing.Color.MidnightBlue
        Me.dgExps.Location = New System.Drawing.Point(10, 16)
        Me.dgExps.Name = "dgExps"
        Me.dgExps.ParentRowsBackColor = System.Drawing.Color.DarkGray
        Me.dgExps.ParentRowsForeColor = System.Drawing.Color.Black
        Me.dgExps.SelectionBackColor = System.Drawing.Color.CadetBlue
        Me.dgExps.SelectionForeColor = System.Drawing.Color.White
        Me.dgExps.Size = New System.Drawing.Size(488, 411)
        Me.dgExps.TabIndex = 0
        '
        'grpHeader
        '
        Me.grpHeader.BackColor = System.Drawing.Color.Transparent
        Me.grpHeader.Controls.Add(Me.lblTo)
        Me.grpHeader.Controls.Add(Me.lblFrom)
        Me.grpHeader.Controls.Add(Me.dtpTo)
        Me.grpHeader.Controls.Add(Me.dtpFrom)
        Me.grpHeader.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpHeader.Location = New System.Drawing.Point(9, 28)
        Me.grpHeader.Name = "grpHeader"
        Me.grpHeader.Size = New System.Drawing.Size(506, 51)
        Me.grpHeader.TabIndex = 10
        Me.grpHeader.TabStop = False
        Me.grpHeader.Text = "Select Criteria"
        '
        'lblTo
        '
        Me.lblTo.Location = New System.Drawing.Point(295, 21)
        Me.lblTo.Name = "lblTo"
        Me.lblTo.Size = New System.Drawing.Size(62, 16)
        Me.lblTo.TabIndex = 5
        Me.lblTo.Text = "To Date:"
        '
        'lblFrom
        '
        Me.lblFrom.Location = New System.Drawing.Point(15, 21)
        Me.lblFrom.Name = "lblFrom"
        Me.lblFrom.Size = New System.Drawing.Size(93, 16)
        Me.lblFrom.TabIndex = 4
        Me.lblFrom.Text = "From Date:"
        '
        'dtpTo
        '
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpTo.Location = New System.Drawing.Point(382, 21)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(107, 22)
        Me.dtpTo.TabIndex = 3
        '
        'dtpFrom
        '
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFrom.Location = New System.Drawing.Point(114, 21)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(107, 22)
        Me.dtpFrom.TabIndex = 2
        '
        'grpFooter2
        '
        Me.grpFooter2.BackColor = System.Drawing.Color.Transparent
        Me.grpFooter2.Controls.Add(Me.txtTotalCost)
        Me.grpFooter2.Controls.Add(Me.txtItems)
        Me.grpFooter2.Controls.Add(Me.Label1)
        Me.grpFooter2.Controls.Add(Me.lblTotalItems)
        Me.grpFooter2.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpFooter2.Location = New System.Drawing.Point(9, 517)
        Me.grpFooter2.Name = "grpFooter2"
        Me.grpFooter2.Size = New System.Drawing.Size(506, 42)
        Me.grpFooter2.TabIndex = 12
        Me.grpFooter2.TabStop = False
        Me.grpFooter2.Text = "Total"
        '
        'txtTotalCost
        '
        Me.txtTotalCost.Location = New System.Drawing.Point(394, 15)
        Me.txtTotalCost.Name = "txtTotalCost"
        Me.txtTotalCost.ReadOnly = True
        Me.txtTotalCost.Size = New System.Drawing.Size(104, 22)
        Me.txtTotalCost.TabIndex = 3
        Me.txtTotalCost.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtItems
        '
        Me.txtItems.Location = New System.Drawing.Point(142, 15)
        Me.txtItems.Name = "txtItems"
        Me.txtItems.ReadOnly = True
        Me.txtItems.Size = New System.Drawing.Size(53, 22)
        Me.txtItems.TabIndex = 2
        Me.txtItems.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(311, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 14)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Total Cost:"
        '
        'lblTotalItems
        '
        Me.lblTotalItems.Location = New System.Drawing.Point(15, 15)
        Me.lblTotalItems.Name = "lblTotalItems"
        Me.lblTotalItems.Size = New System.Drawing.Size(120, 14)
        Me.lblTotalItems.TabIndex = 0
        Me.lblTotalItems.Text = "Number of Items:"
        '
        'ErrProv
        '
        Me.ErrProv.ContainerControl = Me
        '
        'frmMiscExpRpt
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 15)
        Me.ClientSize = New System.Drawing.Size(520, 565)
        Me.Controls.Add(Me.toolStrip)
        Me.Controls.Add(Me.grpMain)
        Me.Controls.Add(Me.grpHeader)
        Me.Controls.Add(Me.grpFooter2)
        Me.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmMiscExpRpt"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Misc Expenses Report"
        Me.toolStrip.ResumeLayout(False)
        Me.toolStrip.PerformLayout()
        Me.grpMain.ResumeLayout(False)
        CType(Me.dgExps, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpHeader.ResumeLayout(False)
        Me.grpFooter2.ResumeLayout(False)
        Me.grpFooter2.PerformLayout()
        CType(Me.ErrProv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

#Region " Public Properties "
    Public Property Date1() As Date
        Get
            Return Me._Date1

        End Get

        Set(ByVal Value As Date)
            Me._Date1 = Value

        End Set

    End Property
    Public Property Date2() As Date

        Get
            Return Me._Date2
        End Get

        Set(ByVal Value As Date)
            Me._Date2 = Value
        End Set

    End Property
#End Region

    Public Sub LoadData()

        Me._Ds = Expenses.GetMiscExps(Me.Date1.ToShortDateString, Me.Date2.ToShortDateString)

        SetUpDataGrid()

        Me.dgExps.DataSource = Me._Ds.Tables(0)
        TallyReport()

    End Sub

    Private Sub SetUpDataGrid()

        With Me.dgExps
            .ReadOnly = True
            .Text = "Misc Expenses"
        End With

        Me.dgExps.TableStyles.Clear()

        Dim ts As New DataGridTableStyle
        Dim col As New DataGridTextBoxColumn

        With ts
            .MappingName = "MiscExp"
            .GridLineColor = Color.DarkBlue
            .HeaderBackColor = Color.DarkBlue
            .HeaderForeColor = Color.White
        End With

        With col
            .MappingName = "ItemNo"
            .Width = 0
            ts.GridColumnStyles.Add(col)
        End With


        col = New DataGridTextBoxColumn
        With col
            .MappingName = "DateOfExp"
            .HeaderText = "Date"
            .Width = 83
            .Format = "d"
            .Alignment = HorizontalAlignment.Left
            ts.GridColumnStyles.Add(col)
        End With

        Me.dgExps.TableStyles.Add(ts)

        col = New DataGridTextBoxColumn
        With col
            .MappingName = "Category"
            .HeaderText = "Category"
            .Width = 150
            .Alignment = HorizontalAlignment.Left
            ts.GridColumnStyles.Add(col)
        End With

        Me.dgExps.TableStyles.Add(ts)

        col = New DataGridTextBoxColumn
        With col
            .MappingName = "Description"
            .HeaderText = "Description"
            .Width = 300
            .Alignment = HorizontalAlignment.Left
            ts.GridColumnStyles.Add(col)
        End With

        Me.dgExps.TableStyles.Add(ts)

        col = New DataGridTextBoxColumn
        With col
            .MappingName = "Cost"
            .HeaderText = "Cost"
            .Width = 100
            .Format = "c"
            .Alignment = HorizontalAlignment.Left
            ts.GridColumnStyles.Add(col)
        End With
        Me.dgExps.TableStyles.Add(ts)

        col = New DataGridTextBoxColumn
        With col
            .MappingName = "Notes"
            .HeaderText = "Notes"
            .Width = 350
            .Alignment = HorizontalAlignment.Left
            ts.GridColumnStyles.Add(col)
        End With
        Me.dgExps.TableStyles.Add(ts)


    End Sub

    Private Sub TallyReport()

        Dim dt As DataTable = CType(Me.dgExps.DataSource, DataTable)
        Dim TotalCost As Double

        If IsNothing(dt) Then
            Me.txtItems.Text = "0"
            TotalCost = 0
        Else
            Me.txtItems.Text = dt.Rows.Count.ToString

            For Each row As DataRow In dt.Rows
                TotalCost += Convert.ToDouble(row.Item("Cost"))
            Next

        End If

        Me.txtTotalCost.Text = TotalCost.ToString("c")

    End Sub

    Private Sub OpenExpItem()

        If Me.dgExps.CurrentRowIndex > -1 Then
            Dim ItemNo As Integer = Convert.ToInt32(Me.dgExps(Me.dgExps.CurrentRowIndex, 0))

            Dim ds As DataSet = Me.Expenses.GetMiscExpItem(ItemNo)

            If ds.Tables(0).Rows.Count = 1 Then
                Utilities.ViewMiscExpItem(Me.MdiParent, Me, ds)
            End If
        End If

    End Sub

    Private Sub DeleteExpItem()

        'get the item date and description
        Dim itemDesc As String = CStr(Me.dgExps(Me.dgExps.CurrentRowIndex, 1)) &
            " - " & CStr(Me.dgExps(Me.dgExps.CurrentRowIndex, 3))

        If Me._Ds.Tables(0).Rows.Count > 0 Then
            If Me.dgExps.VisibleRowCount > 0 Then
                If MessageBox.Show("Are you sure you want to delete: " & itemDesc & "?", ProgramName, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.Yes Then
                    Dim ItemNo As Integer = Convert.ToInt32(Me.dgExps(Me.dgExps.CurrentRowIndex, 0))
                    Me.Expenses.CrudMode = DelortBusObjects.MiscExp.DBMode.Delete
                    Me.Expenses.ItemNo = ItemNo
                    Me.Expenses.InsertUpdateDelete()
                    'refresh the display asfter delete
                    Me.LoadData()
                End If
            End If
        End If

    End Sub

    Private Sub dtpFrom_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpFrom.ValueChanged

        Try
            Me.dgExps.DataSource = Nothing
            Me.Date1 = Convert.ToDateTime(dtpFrom.Value)
            If Date1 > Me.Date2 Then
                Me.ErrProv.SetError(Me.dtpFrom, "To date must be equal to or less than from date")
            Else
                Me.ErrProv.SetError(Me.dtpFrom, "")
                Me.ErrProv.SetError(Me.dtpTo, "")
            End If

            Me.dtpTo.MinDate = dtpFrom.Value

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub dtpTo_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpTo.ValueChanged

        Try
            Me.dgExps.DataSource = Nothing
            Me.Date2 = Convert.ToDateTime(Me.dtpTo.Value)

            If Date1 > Me.Date2 Then
                Me.ErrProv.SetError(Me.dtpTo, "To date must be equal to or later than from date")
            Else
                Me.ErrProv.SetError(Me.dtpTo, "")
                Me.ErrProv.SetError(Me.dtpFrom, "")
            End If

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub dgExps_DataSourceChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgExps.DataSourceChanged

        Try

            If Not IsNothing(Me.dgExps.DataSource) AndAlso Me._Ds.Tables(0).Rows.Count > 0 Then
                Me.tsDelete.Enabled = True
                Me.tsOpen.Enabled = True
            Else
                Me.tsDelete.Enabled = False
                Me.tsOpen.Enabled = False
            End If

            Me.TallyReport()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub btnOpenItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try

            OpenExpItem()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        Try
            DeleteExpItem()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub dgExps_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgExps.DoubleClick

        Try
            Me.OpenExpItem()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub cmnuOpen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmnuOpen.Click

        Try
            Me.OpenExpItem()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub cmnuDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmnuDelete.Click

        Try
            DeleteExpItem()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub btnRefresh_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        Try

            If Me.dgExps.VisibleColumnCount > 0 Then
                Me.LoadData()
            End If

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.Close()

    End Sub

    Private Sub tsRun_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsRun.Click

        Try
            LoadData()
            Me.SetUpDataGrid()
            TallyReport()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub tsOpen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsOpen.Click

        Try

            OpenExpItem()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub tsDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsDelete.Click

        Try
            DeleteExpItem()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub Print(ByVal Year As String)

        Dim ds As DataSet

        Try
            If Me._ReportType = ReportType.CATEGORY Then
                ds = Me.Expenses.GetGroupSummaryReport(CStr(Me.Date1), CStr(Me.Date2))
                Utilities.PrintMiscExpByCategory(ds, Date1.ToShortDateString, Date2.ToShortDateString)
            ElseIf Me._ReportType = ReportType.MONTHLY Then
                ds = Me.Expenses.GetMoSummaryReport(Year)
                Utilities.PrintMiscExpByMonth(ds, Year)
            End If

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub tsMonthly_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsMonthly.Click, tsCategory.Click

        Select Case DirectCast(sender, ToolStripMenuItem).Name
            Case Me.tsCategory.Name
                Me._ReportType = ReportType.CATEGORY
                Me.Print("")
            Case tsMonthly.Name
                Me._ReportType = ReportType.MONTHLY
        End Select

    End Sub

    Private Sub SetUpUI()

        LoadPrintSummaryYearButtons()

        Me.dtpTo.Value = Today
        Me.dtpFrom.Value = DateSerial(Year(Now()), Month(Now()), 1)

        LoadData()
        Me.SetUpDataGrid()

    End Sub

    Private Sub frmMiscExpRpt_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetUpUI()

    End Sub

    Private Sub AddMiscItem()
        Utilities.ViewMiscAdd(Me.ParentForm)
    End Sub

    Private Sub tsAddNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsAddNew.Click
        Me.AddMiscItem()
    End Sub

    Private Sub LoadPrintSummaryYearButtons()

        Dim sql As String = "SELECT DISTINCT Year(DateOfExp)  as Years FROM MiscExps;"
        Dim ds As DataSet
        Dim tsm As ToolStripMenuItem

        ds = AccessUtils.Utils.GetDataSet(My.Settings.DBConn, sql, Nothing)

        With Me.tsMonthly
            For Each row As DataRow In ds.Tables(0).Rows
                tsm = New ToolStripMenuItem(row.Item(0).ToString, Nothing, AddressOf ToolStripMenuItem_Click)
                .DropDownItems.Add(tsm)
            Next

        End With
    End Sub

    Private Sub ToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        Try

            Cursor = Cursors.WaitCursor
            Dim Year As String = DirectCast(sender, ToolStripMenuItem).Text

            Me.Date1 = CDate("01/01/" & Year)
            Me.Date2 = CDate("12/31/" & Year)

            Me.Print(Year)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            Cursor = Cursors.Default
        End Try

    End Sub

End Class
