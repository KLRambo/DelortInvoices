Option Strict On

Imports System.Configuration.ConfigurationSettings

Public Class frmViewMaterials
    Inherits System.Windows.Forms.Form

    Private _Ds As DataSet
    Private WithEvents Materials As New DelortBusObjects.Materials(g_settings)
    Private _Date1 As Date = Nothing
    Private _Date2 As Date = Nothing
    Friend WithEvents tsMaterials As System.Windows.Forms.ToolStrip
    Friend WithEvents tsPrint As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents tsPrintItemized As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsPrintSummary As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsRun As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsOpen As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsDelete As System.Windows.Forms.ToolStripButton
    Friend WithEvents StatusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusCount As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusSumm As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tsSpacer As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tsAdd As System.Windows.Forms.ToolStripButton

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
    Friend WithEvents grpHeader As System.Windows.Forms.GroupBox
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblFrom As System.Windows.Forms.Label
    Friend WithEvents lblTo As System.Windows.Forms.Label
    Friend WithEvents dgMaterials As System.Windows.Forms.DataGrid
    Friend WithEvents cmnuMain As System.Windows.Forms.ContextMenu
    Friend WithEvents cmnuOpen As System.Windows.Forms.MenuItem
    Friend WithEvents cmnuDelete As System.Windows.Forms.MenuItem
    Friend WithEvents ErrProv As System.Windows.Forms.ErrorProvider
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmViewMaterials))
        Me.grpMain = New System.Windows.Forms.GroupBox()
        Me.dgMaterials = New System.Windows.Forms.DataGrid()
        Me.cmnuMain = New System.Windows.Forms.ContextMenu()
        Me.cmnuOpen = New System.Windows.Forms.MenuItem()
        Me.cmnuDelete = New System.Windows.Forms.MenuItem()
        Me.grpHeader = New System.Windows.Forms.GroupBox()
        Me.lblTo = New System.Windows.Forms.Label()
        Me.lblFrom = New System.Windows.Forms.Label()
        Me.dtpTo = New System.Windows.Forms.DateTimePicker()
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker()
        Me.ErrProv = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.tsMaterials = New System.Windows.Forms.ToolStrip()
        Me.tsRun = New System.Windows.Forms.ToolStripButton()
        Me.tsOpen = New System.Windows.Forms.ToolStripButton()
        Me.tsAdd = New System.Windows.Forms.ToolStripButton()
        Me.tsDelete = New System.Windows.Forms.ToolStripButton()
        Me.tsPrint = New System.Windows.Forms.ToolStripDropDownButton()
        Me.tsPrintItemized = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsPrintSummary = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStrip = New System.Windows.Forms.StatusStrip()
        Me.tsSpacer = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusCount = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusSumm = New System.Windows.Forms.ToolStripStatusLabel()
        Me.grpMain.SuspendLayout()
        CType(Me.dgMaterials, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpHeader.SuspendLayout()
        CType(Me.ErrProv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tsMaterials.SuspendLayout()
        Me.StatusStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpMain
        '
        Me.grpMain.BackColor = System.Drawing.Color.Transparent
        Me.grpMain.Controls.Add(Me.dgMaterials)
        Me.grpMain.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpMain.Location = New System.Drawing.Point(9, 81)
        Me.grpMain.Name = "grpMain"
        Me.grpMain.Size = New System.Drawing.Size(503, 452)
        Me.grpMain.TabIndex = 1
        Me.grpMain.TabStop = False
        Me.grpMain.Text = "Materials"
        '
        'dgMaterials
        '
        Me.dgMaterials.AlternatingBackColor = System.Drawing.Color.LightGray
        Me.dgMaterials.BackColor = System.Drawing.Color.Gainsboro
        Me.dgMaterials.BackgroundColor = System.Drawing.Color.Silver
        Me.dgMaterials.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.dgMaterials.CaptionBackColor = System.Drawing.Color.LightSteelBlue
        Me.dgMaterials.CaptionFont = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.dgMaterials.CaptionForeColor = System.Drawing.Color.MidnightBlue
        Me.dgMaterials.CaptionVisible = False
        Me.dgMaterials.ContextMenu = Me.cmnuMain
        Me.dgMaterials.DataMember = ""
        Me.dgMaterials.FlatMode = True
        Me.dgMaterials.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dgMaterials.ForeColor = System.Drawing.Color.Black
        Me.dgMaterials.GridLineColor = System.Drawing.Color.DimGray
        Me.dgMaterials.GridLineStyle = System.Windows.Forms.DataGridLineStyle.None
        Me.dgMaterials.HeaderBackColor = System.Drawing.Color.RoyalBlue
        Me.dgMaterials.HeaderFont = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.dgMaterials.HeaderForeColor = System.Drawing.Color.White
        Me.dgMaterials.LinkColor = System.Drawing.Color.MidnightBlue
        Me.dgMaterials.Location = New System.Drawing.Point(6, 17)
        Me.dgMaterials.Name = "dgMaterials"
        Me.dgMaterials.ParentRowsBackColor = System.Drawing.Color.DarkGray
        Me.dgMaterials.ParentRowsForeColor = System.Drawing.Color.Black
        Me.dgMaterials.SelectionBackColor = System.Drawing.Color.CadetBlue
        Me.dgMaterials.SelectionForeColor = System.Drawing.Color.White
        Me.dgMaterials.Size = New System.Drawing.Size(488, 425)
        Me.dgMaterials.TabIndex = 0
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
        'cmnuDelete
        '
        Me.cmnuDelete.Index = 1
        Me.cmnuDelete.Text = "Delete Item"
        '
        'grpHeader
        '
        Me.grpHeader.BackColor = System.Drawing.Color.Transparent
        Me.grpHeader.Controls.Add(Me.lblTo)
        Me.grpHeader.Controls.Add(Me.lblFrom)
        Me.grpHeader.Controls.Add(Me.dtpTo)
        Me.grpHeader.Controls.Add(Me.dtpFrom)
        Me.grpHeader.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpHeader.Location = New System.Drawing.Point(9, 27)
        Me.grpHeader.Name = "grpHeader"
        Me.grpHeader.Size = New System.Drawing.Size(503, 49)
        Me.grpHeader.TabIndex = 2
        Me.grpHeader.TabStop = False
        Me.grpHeader.Text = "Select Dates:"
        '
        'lblTo
        '
        Me.lblTo.Location = New System.Drawing.Point(269, 18)
        Me.lblTo.Name = "lblTo"
        Me.lblTo.Size = New System.Drawing.Size(65, 16)
        Me.lblTo.TabIndex = 5
        Me.lblTo.Text = "To Date:"
        '
        'lblFrom
        '
        Me.lblFrom.Location = New System.Drawing.Point(8, 18)
        Me.lblFrom.Name = "lblFrom"
        Me.lblFrom.Size = New System.Drawing.Size(78, 16)
        Me.lblFrom.TabIndex = 4
        Me.lblFrom.Text = "From Date:"
        '
        'dtpTo
        '
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpTo.Location = New System.Drawing.Point(340, 18)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(147, 22)
        Me.dtpTo.TabIndex = 3
        '
        'dtpFrom
        '
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFrom.Location = New System.Drawing.Point(92, 18)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(147, 22)
        Me.dtpFrom.TabIndex = 2
        '
        'ErrProv
        '
        Me.ErrProv.ContainerControl = Me
        '
        'tsMaterials
        '
        Me.tsMaterials.BackColor = System.Drawing.Color.Transparent
        Me.tsMaterials.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsRun, Me.tsOpen, Me.tsAdd, Me.tsDelete, Me.tsPrint})
        Me.tsMaterials.Location = New System.Drawing.Point(0, 0)
        Me.tsMaterials.Name = "tsMaterials"
        Me.tsMaterials.Size = New System.Drawing.Size(522, 25)
        Me.tsMaterials.TabIndex = 9
        Me.tsMaterials.Text = "ToolStrip1"
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
        Me.tsOpen.ToolTipText = "Open Record"
        '
        'tsAdd
        '
        Me.tsAdd.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsAdd.Image = CType(resources.GetObject("tsAdd.Image"), System.Drawing.Image)
        Me.tsAdd.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsAdd.Name = "tsAdd"
        Me.tsAdd.Size = New System.Drawing.Size(23, 22)
        Me.tsAdd.ToolTipText = "Add New"
        '
        'tsDelete
        '
        Me.tsDelete.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsDelete.Image = CType(resources.GetObject("tsDelete.Image"), System.Drawing.Image)
        Me.tsDelete.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsDelete.Name = "tsDelete"
        Me.tsDelete.Size = New System.Drawing.Size(23, 22)
        Me.tsDelete.ToolTipText = "Delete"
        '
        'tsPrint
        '
        Me.tsPrint.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsPrint.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsPrintItemized, Me.tsPrintSummary})
        Me.tsPrint.Image = CType(resources.GetObject("tsPrint.Image"), System.Drawing.Image)
        Me.tsPrint.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsPrint.Name = "tsPrint"
        Me.tsPrint.Size = New System.Drawing.Size(29, 22)
        Me.tsPrint.ToolTipText = "Print Reports"
        '
        'tsPrintItemized
        '
        Me.tsPrintItemized.Image = CType(resources.GetObject("tsPrintItemized.Image"), System.Drawing.Image)
        Me.tsPrintItemized.Name = "tsPrintItemized"
        Me.tsPrintItemized.Size = New System.Drawing.Size(165, 22)
        Me.tsPrintItemized.Text = "Itemized Report"
        '
        'tsPrintSummary
        '
        Me.tsPrintSummary.Image = CType(resources.GetObject("tsPrintSummary.Image"), System.Drawing.Image)
        Me.tsPrintSummary.Name = "tsPrintSummary"
        Me.tsPrintSummary.Size = New System.Drawing.Size(165, 22)
        Me.tsPrintSummary.Text = "Summary Report"
        '
        'StatusStrip
        '
        Me.StatusStrip.BackColor = System.Drawing.Color.Transparent
        Me.StatusStrip.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsSpacer, Me.StatusCount, Me.StatusSumm})
        Me.StatusStrip.Location = New System.Drawing.Point(0, 543)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.Size = New System.Drawing.Size(522, 22)
        Me.StatusStrip.TabIndex = 10
        '
        'tsSpacer
        '
        Me.tsSpacer.Name = "tsSpacer"
        Me.tsSpacer.Size = New System.Drawing.Size(0, 17)
        Me.tsSpacer.Visible = False
        '
        'StatusCount
        '
        Me.StatusCount.BackColor = System.Drawing.Color.Transparent
        Me.StatusCount.ForeColor = System.Drawing.Color.MediumBlue
        Me.StatusCount.Name = "StatusCount"
        Me.StatusCount.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.StatusCount.Size = New System.Drawing.Size(253, 17)
        Me.StatusCount.Spring = True
        '
        'StatusSumm
        '
        Me.StatusSumm.BackColor = System.Drawing.Color.Transparent
        Me.StatusSumm.ForeColor = System.Drawing.Color.MediumBlue
        Me.StatusSumm.Name = "StatusSumm"
        Me.StatusSumm.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.StatusSumm.Size = New System.Drawing.Size(253, 17)
        Me.StatusSumm.Spring = True
        '
        'frmViewMaterials
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 15)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(522, 565)
        Me.Controls.Add(Me.StatusStrip)
        Me.Controls.Add(Me.tsMaterials)
        Me.Controls.Add(Me.grpHeader)
        Me.Controls.Add(Me.grpMain)
        Me.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmViewMaterials"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "View Materials Report"
        Me.grpMain.ResumeLayout(False)
        CType(Me.dgMaterials, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpHeader.ResumeLayout(False)
        CType(Me.ErrProv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tsMaterials.ResumeLayout(False)
        Me.tsMaterials.PerformLayout()
        Me.StatusStrip.ResumeLayout(False)
        Me.StatusStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region


    Public Sub LoadData()


        Me._Ds = Me.Materials.GetMaterials(Me.Date1, Me.Date2)


        Me.dgMaterials.DataSource = Me._Ds.Tables(0)
        TallyReport()

    End Sub

    Private Sub SetUpDataGrid()

        With Me.dgMaterials
            .ReadOnly = True
            .Text = "Materials"
        End With

        Me.dgMaterials.TableStyles.Clear()

        Dim ts As New DataGridTableStyle
        Dim col As New DataGridTextBoxColumn

        With ts
            .MappingName = "Materials"
            .GridLineColor = Color.DarkBlue
            .HeaderBackColor = Color.DarkBlue
            .HeaderForeColor = Color.White
        End With

        With col
            .MappingName = "MaterialNo"
            .Width = 0
            ts.GridColumnStyles.Add(col)
        End With


        col = New DataGridTextBoxColumn
        With col
            .MappingName = "DatePurchased"
            .HeaderText = "Date"
            .Width = 100
            .Format = "d"
            .Alignment = HorizontalAlignment.Left
            ts.GridColumnStyles.Add(col)
        End With

        Me.dgMaterials.TableStyles.Add(ts)

        col = New DataGridTextBoxColumn
        With col
            .MappingName = "MaterialDesc"
            .HeaderText = "Description"
            .Width = 350
            .Alignment = HorizontalAlignment.Left
            ts.GridColumnStyles.Add(col)
        End With

        Me.dgMaterials.TableStyles.Add(ts)

        col = New DataGridTextBoxColumn
        With col
            .MappingName = "MaterialCost"
            .HeaderText = "Cost"
            .Width = 100
            .Format = "c"
            .Alignment = HorizontalAlignment.Left
            ts.GridColumnStyles.Add(col)
        End With
        Me.dgMaterials.TableStyles.Add(ts)

        col = New DataGridTextBoxColumn
        With col
            .MappingName = "Notes"
            .HeaderText = "Notes"
            .Width = 450
            .Alignment = HorizontalAlignment.Left
            ts.GridColumnStyles.Add(col)
        End With

        Me.dgMaterials.TableStyles.Add(ts)

    End Sub

    Private Sub TallyReport()

        Dim TotalCost As Double


        Me.StatusCount.Text = "Count: " & Me._Ds.Tables(0).Rows.Count.ToString

        For Each row As DataRow In Me._Ds.Tables(0).Rows
            TotalCost += Convert.ToDouble(row.Item("MaterialCost"))
        Next

        Me.StatusSumm.Text = "Total: " & TotalCost.ToString("c")

    End Sub

    Private Sub OpenMaterialItem()

        If Me.dgMaterials.CurrentRowIndex > -1 Then

            Dim MaterialNo As Integer = Convert.ToInt32(Me.dgMaterials(Me.dgMaterials.CurrentRowIndex, 0))

            Dim ds As DataSet = Me.Materials.GetMaterialItem(MaterialNo)

            If ds.Tables(0).Rows.Count = 1 Then

                Utilities.ViewEditMaterials(Me.MdiParent, Me, ds)

            End If
        End If


    End Sub

    Protected Overridable Sub btnOpenItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try

            Me.OpenMaterialItem()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub dtpFrom_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpFrom.ValueChanged

        Try
            Me.dgMaterials.DataSource = Nothing
            Me.Date1 = Convert.ToDateTime(dtpFrom.Value)
            If Date1 > Me.Date2 Then
                Me.ErrProv.SetError(Me.dtpFrom, "To date must be equal to or later than from date")
            Else
                Me.ErrProv.SetError(Me.dtpFrom, "")
                Me.ErrProv.SetError(Me.dtpTo, "")
            End If

            Me.dtpTo.MinDate = Me.dtpFrom.Value

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub dtpTo_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpTo.ValueChanged

        Try
            Me.dgMaterials.DataSource = Nothing
            Me.Date2 = Convert.ToDateTime(Me.dtpTo.Value)

            If Date1 > Me.Date2 Then
                Me.ErrProv.SetError(Me.dtpTo, "To date must be equal to or later than from date")
                Me.tsRun.Enabled = False
            Else
                Me.ErrProv.SetError(Me.dtpTo, "")
                Me.ErrProv.SetError(Me.dtpFrom, "")
                Me.tsRun.Enabled = True
            End If

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub dgMaterials_DataSourceChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgMaterials.DataSourceChanged

        Try
            Me.StatusSumm.Text = ""
            Me.StatusCount.Text = ""
            If Not IsNothing(Me.dgMaterials.DataSource) AndAlso Me._Ds.Tables(0).Rows.Count > 0 Then
                Me.tsDelete.Enabled = True
                Me.tsOpen.Enabled = True
            Else
                Me.tsDelete.Enabled = False
                Me.tsOpen.Enabled = False
            End If

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub DeleteMaterialItem()

        'get the item date and description
        Dim itemDesc As String = CStr(Me.dgMaterials(Me.dgMaterials.CurrentRowIndex, 1)) &
            " - " & CStr(Me.dgMaterials(Me.dgMaterials.CurrentRowIndex, 2))

        If Me._Ds.Tables(0).Rows.Count > 0 Then
            If Me.dgMaterials.VisibleRowCount > 0 Then
                If MessageBox.Show("Are you sure you want to delete: " & itemDesc & "?", ProgramName, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.Yes Then
                    Dim MaterialNo As Integer = Convert.ToInt32(Me.dgMaterials(Me.dgMaterials.CurrentRowIndex, 0))
                    Me.Materials.CrudMode = DelortBusObjects.Materials.DBMode.Delete
                    Me.Materials.MaterialNo = MaterialNo
                    Me.Materials.InsertUpdateDelete()
                    'refresh the display after delete
                    Me.LoadData()
                End If

            End If
        End If



    End Sub

    Private Sub cmnuOpen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmnuOpen.Click

        Try

            Me.OpenMaterialItem()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub cmnuDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmnuDelete.Click

        Try
            DeleteMaterialItem()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub dgMaterials_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgMaterials.DoubleClick

        Try
            OpenMaterialItem()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub frmViewMaterials_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        If Not IsNothing(_Ds) Then
            _Ds.Dispose()
            _Ds = Nothing
        End If

    End Sub

    Private Sub tsPrintItemized_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsPrintItemized.Click

        Try

            Cursor = Cursors.WaitCursor

            Dim MaterialsRpt As New MaterialsRpt

            With MaterialsRpt
                .FromDate = Me.dtpFrom.Value.ToShortDateString
                .ToDate = Me.dtpTo.Value.ToShortDateString
            End With

            MaterialsRpt.Show()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            Cursor = Cursors.Default
        End Try

    End Sub

    Private Sub frmViewMaterials_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Me.SetUpUI()

    End Sub

    Private Sub ToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        Try

            Cursor = Cursors.WaitCursor
            Dim Year As String = DirectCast(sender, ToolStripMenuItem).Text

            Dim strDate1 As String = "01/01/" & Year
            Dim strDate2 As String = "12/31/" & Year

            Utilities.PrintMaterials(Materials.GetMoSummary(CDate(strDate1), CDate(strDate2)), Year)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            Cursor = Cursors.Default
        End Try

    End Sub

    Private Sub LoadPrintSummaryYearButtons()

        Me.tsPrintItemized.DropDownItems.Clear()

        Dim sql As String = "SELECT DISTINCT Year(DatePurchased)  as Years FROM Materials;"
        Dim ds As DataSet
        Dim tsm As ToolStripMenuItem

        ds = AccessUtils.Utils.GetDataSet(My.Settings.DBConn, sql, Nothing)

        With Me.tsPrintSummary
            For Each row As DataRow In ds.Tables(0).Rows
                tsm = New ToolStripMenuItem(row.Item(0).ToString, Nothing, AddressOf ToolStripMenuItem_Click)
                .DropDownItems.Add(tsm)
            Next

        End With
    End Sub

    Private Sub tsExit_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try

            Me.Close()
        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub tsOpen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsOpen.Click
        Try

            Me.OpenMaterialItem()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub tsDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsDelete.Click

        Try
            DeleteMaterialItem()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

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

    Private Sub tsAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsAdd.Click

        Try
            Utilities.ViewAddMaterials(Me.MdiParent)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub SetUpUI()

        LoadPrintSummaryYearButtons()

        Me.dtpTo.Value = Today
        Me.dtpFrom.Value = DateSerial(Year(Now()), Month(Now()), 1)
        Me.Date2 = Today
        Me.Date1 = DateSerial(Year(Now()), Month(Now()), 1)
        SetUpDataGrid()
        LoadData()

    End Sub

End Class
