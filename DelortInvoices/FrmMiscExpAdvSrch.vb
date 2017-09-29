Option Strict On

Imports DelortBusObjects
Imports System.Configuration.ConfigurationSettings

Public Class frmMiscExpAdvSrch

    Inherits frmBaseWindow

    Private WithEvents m_AdvSrch As New MiscExpAdvSrch(g_settings)
    Private Expenses As New MiscExp(g_settings)

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
    Friend WithEvents grpDates As System.Windows.Forms.GroupBox
    Friend WithEvents lblTo As System.Windows.Forms.Label
    Friend WithEvents lblFrom As System.Windows.Forms.Label
    Friend WithEvents dtpToDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpFromDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents grpAction As System.Windows.Forms.GroupBox
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnOpen As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents grpSearchText As System.Windows.Forms.GroupBox
    Friend WithEvents txtSearchData As System.Windows.Forms.TextBox
    Friend WithEvents grpType As System.Windows.Forms.GroupBox
    Friend WithEvents rdoNotes As System.Windows.Forms.RadioButton
    Friend WithEvents rdoDesc As System.Windows.Forms.RadioButton
    Friend WithEvents grpInvoices As System.Windows.Forms.GroupBox
    Friend WithEvents grdList As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents grpCost As System.Windows.Forms.GroupBox
    Friend WithEvents txtPriceFrom As System.Windows.Forms.TextBox
    Friend WithEvents txtPriceTo As System.Windows.Forms.TextBox
    Friend WithEvents lblPriceTo As System.Windows.Forms.Label
    Friend WithEvents lblRange As System.Windows.Forms.Label
    Friend WithEvents rdoPriceRange As System.Windows.Forms.RadioButton
    Friend WithEvents ErrorProvider As System.Windows.Forms.ErrorProvider
    Friend WithEvents ToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents CmMain As System.Windows.Forms.ContextMenu
    Friend WithEvents cmOpen As System.Windows.Forms.MenuItem
    Friend WithEvents cmDelete As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMiscExpAdvSrch))
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.grpDates = New System.Windows.Forms.GroupBox()
        Me.lblTo = New System.Windows.Forms.Label()
        Me.lblFrom = New System.Windows.Forms.Label()
        Me.dtpToDate = New System.Windows.Forms.DateTimePicker()
        Me.dtpFromDate = New System.Windows.Forms.DateTimePicker()
        Me.grpAction = New System.Windows.Forms.GroupBox()
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnOpen = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.grpSearchText = New System.Windows.Forms.GroupBox()
        Me.txtSearchData = New System.Windows.Forms.TextBox()
        Me.grpType = New System.Windows.Forms.GroupBox()
        Me.rdoPriceRange = New System.Windows.Forms.RadioButton()
        Me.rdoNotes = New System.Windows.Forms.RadioButton()
        Me.rdoDesc = New System.Windows.Forms.RadioButton()
        Me.grpInvoices = New System.Windows.Forms.GroupBox()
        Me.grdList = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.CmMain = New System.Windows.Forms.ContextMenu()
        Me.cmOpen = New System.Windows.Forms.MenuItem()
        Me.cmDelete = New System.Windows.Forms.MenuItem()
        Me.grpCost = New System.Windows.Forms.GroupBox()
        Me.lblRange = New System.Windows.Forms.Label()
        Me.lblPriceTo = New System.Windows.Forms.Label()
        Me.txtPriceTo = New System.Windows.Forms.TextBox()
        Me.txtPriceFrom = New System.Windows.Forms.TextBox()
        Me.ErrorProvider = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.grpDates.SuspendLayout()
        Me.grpAction.SuspendLayout()
        Me.grpSearchText.SuspendLayout()
        Me.grpType.SuspendLayout()
        Me.grpInvoices.SuspendLayout()
        CType(Me.grdList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpCost.SuspendLayout()
        CType(Me.ErrorProvider, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grpDates
        '
        Me.grpDates.BackColor = System.Drawing.Color.Transparent
        Me.grpDates.Controls.Add(Me.lblTo)
        Me.grpDates.Controls.Add(Me.lblFrom)
        Me.grpDates.Controls.Add(Me.dtpToDate)
        Me.grpDates.Controls.Add(Me.dtpFromDate)
        Me.grpDates.Location = New System.Drawing.Point(176, 8)
        Me.grpDates.Name = "grpDates"
        Me.grpDates.Size = New System.Drawing.Size(184, 72)
        Me.grpDates.TabIndex = 8
        Me.grpDates.TabStop = False
        Me.grpDates.Text = "Select Dates:"
        '
        'lblTo
        '
        Me.lblTo.Location = New System.Drawing.Point(19, 44)
        Me.lblTo.Name = "lblTo"
        Me.lblTo.Size = New System.Drawing.Size(28, 16)
        Me.lblTo.TabIndex = 3
        Me.lblTo.Text = "To:"
        '
        'lblFrom
        '
        Me.lblFrom.Location = New System.Drawing.Point(15, 15)
        Me.lblFrom.Name = "lblFrom"
        Me.lblFrom.Size = New System.Drawing.Size(45, 16)
        Me.lblFrom.TabIndex = 2
        Me.lblFrom.Text = "From:"
        '
        'dtpToDate
        '
        Me.dtpToDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpToDate.Location = New System.Drawing.Point(61, 42)
        Me.dtpToDate.Name = "dtpToDate"
        Me.dtpToDate.Size = New System.Drawing.Size(102, 21)
        Me.dtpToDate.TabIndex = 0
        '
        'dtpFromDate
        '
        Me.dtpFromDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFromDate.Location = New System.Drawing.Point(61, 15)
        Me.dtpFromDate.Name = "dtpFromDate"
        Me.dtpFromDate.Size = New System.Drawing.Size(102, 21)
        Me.dtpFromDate.TabIndex = 0
        '
        'grpAction
        '
        Me.grpAction.BackColor = System.Drawing.Color.Transparent
        Me.grpAction.Controls.Add(Me.btnRefresh)
        Me.grpAction.Controls.Add(Me.btnDelete)
        Me.grpAction.Controls.Add(Me.btnOpen)
        Me.grpAction.Controls.Add(Me.btnExit)
        Me.grpAction.Location = New System.Drawing.Point(176, 136)
        Me.grpAction.Name = "grpAction"
        Me.grpAction.Size = New System.Drawing.Size(184, 80)
        Me.grpAction.TabIndex = 7
        Me.grpAction.TabStop = False
        Me.grpAction.Text = "Action"
        '
        'btnRefresh
        '
        Me.btnRefresh.BackColor = System.Drawing.Color.AliceBlue
        Me.btnRefresh.Enabled = False
        Me.btnRefresh.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRefresh.Image = CType(resources.GetObject("btnRefresh.Image"), System.Drawing.Image)
        Me.btnRefresh.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnRefresh.Location = New System.Drawing.Point(8, 48)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(72, 22)
        Me.btnRefresh.TabIndex = 0
        Me.btnRefresh.Text = "&Run"
        Me.btnRefresh.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip.SetToolTip(Me.btnRefresh, "Run/Refresh Report")
        Me.btnRefresh.UseVisualStyleBackColor = False
        '
        'btnDelete
        '
        Me.btnDelete.BackColor = System.Drawing.Color.AliceBlue
        Me.btnDelete.Enabled = False
        Me.btnDelete.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.Image = CType(resources.GetObject("btnDelete.Image"), System.Drawing.Image)
        Me.btnDelete.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDelete.Location = New System.Drawing.Point(96, 17)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(72, 22)
        Me.btnDelete.TabIndex = 0
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip.SetToolTip(Me.btnDelete, "Delete Selected Item")
        Me.btnDelete.UseVisualStyleBackColor = False
        '
        'btnOpen
        '
        Me.btnOpen.BackColor = System.Drawing.Color.AliceBlue
        Me.btnOpen.Enabled = False
        Me.btnOpen.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOpen.Image = CType(resources.GetObject("btnOpen.Image"), System.Drawing.Image)
        Me.btnOpen.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnOpen.Location = New System.Drawing.Point(8, 17)
        Me.btnOpen.Name = "btnOpen"
        Me.btnOpen.Size = New System.Drawing.Size(72, 22)
        Me.btnOpen.TabIndex = 0
        Me.btnOpen.Text = "&Open"
        Me.btnOpen.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip.SetToolTip(Me.btnOpen, "Open Selected Item")
        Me.btnOpen.UseVisualStyleBackColor = False
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.AliceBlue
        Me.btnExit.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.Image = CType(resources.GetObject("btnExit.Image"), System.Drawing.Image)
        Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnExit.Location = New System.Drawing.Point(96, 48)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(72, 22)
        Me.btnExit.TabIndex = 0
        Me.btnExit.Text = "  E&xit"
        Me.ToolTip.SetToolTip(Me.btnExit, "Exit form")
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'grpSearchText
        '
        Me.grpSearchText.BackColor = System.Drawing.Color.Transparent
        Me.grpSearchText.Controls.Add(Me.txtSearchData)
        Me.grpSearchText.Location = New System.Drawing.Point(176, 80)
        Me.grpSearchText.Name = "grpSearchText"
        Me.grpSearchText.Size = New System.Drawing.Size(184, 56)
        Me.grpSearchText.TabIndex = 6
        Me.grpSearchText.TabStop = False
        Me.grpSearchText.Text = "Enter Data To Search For"
        '
        'txtSearchData
        '
        Me.txtSearchData.Location = New System.Drawing.Point(10, 24)
        Me.txtSearchData.Multiline = True
        Me.txtSearchData.Name = "txtSearchData"
        Me.txtSearchData.Size = New System.Drawing.Size(155, 19)
        Me.txtSearchData.TabIndex = 0
        Me.ToolTip.SetToolTip(Me.txtSearchData, "Enter Text to Search for in Selected Field")
        '
        'grpType
        '
        Me.grpType.BackColor = System.Drawing.Color.Transparent
        Me.grpType.Controls.Add(Me.rdoPriceRange)
        Me.grpType.Controls.Add(Me.rdoNotes)
        Me.grpType.Controls.Add(Me.rdoDesc)
        Me.grpType.Location = New System.Drawing.Point(8, 8)
        Me.grpType.Name = "grpType"
        Me.grpType.Size = New System.Drawing.Size(160, 128)
        Me.grpType.TabIndex = 5
        Me.grpType.TabStop = False
        Me.grpType.Text = "Select Criteria"
        '
        'rdoPriceRange
        '
        Me.rdoPriceRange.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoPriceRange.Location = New System.Drawing.Point(16, 99)
        Me.rdoPriceRange.Name = "rdoPriceRange"
        Me.rdoPriceRange.Size = New System.Drawing.Size(96, 16)
        Me.rdoPriceRange.TabIndex = 0
        Me.rdoPriceRange.Text = "Price Range"
        '
        'rdoNotes
        '
        Me.rdoNotes.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoNotes.Location = New System.Drawing.Point(16, 61)
        Me.rdoNotes.Name = "rdoNotes"
        Me.rdoNotes.Size = New System.Drawing.Size(72, 16)
        Me.rdoNotes.TabIndex = 0
        Me.rdoNotes.Text = "Notes"
        '
        'rdoDesc
        '
        Me.rdoDesc.Checked = True
        Me.rdoDesc.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoDesc.Location = New System.Drawing.Point(16, 23)
        Me.rdoDesc.Name = "rdoDesc"
        Me.rdoDesc.Size = New System.Drawing.Size(104, 16)
        Me.rdoDesc.TabIndex = 0
        Me.rdoDesc.TabStop = True
        Me.rdoDesc.Text = "Description"
        '
        'grpInvoices
        '
        Me.grpInvoices.BackColor = System.Drawing.Color.Transparent
        Me.grpInvoices.Controls.Add(Me.grdList)
        Me.grpInvoices.Location = New System.Drawing.Point(8, 223)
        Me.grpInvoices.Name = "grpInvoices"
        Me.grpInvoices.Size = New System.Drawing.Size(352, 331)
        Me.grpInvoices.TabIndex = 9
        Me.grpInvoices.TabStop = False
        Me.grpInvoices.Text = "Search Results"
        '
        'grdList
        '
        Me.grdList.ContextMenu = Me.CmMain
        Appearance1.ForeColor = System.Drawing.SystemColors.ControlText
        Appearance1.TextHAlignAsString = "Left"
        Me.grdList.DisplayLayout.Appearance = Appearance1
        Appearance2.ForeColor = System.Drawing.Color.Blue
        Appearance2.TextHAlignAsString = "Center"
        Me.grdList.DisplayLayout.CaptionAppearance = Appearance2
        Me.grdList.DisplayLayout.GroupByBox.Hidden = True
        Me.grdList.DisplayLayout.MaxColScrollRegions = 1
        Me.grdList.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdList.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdList.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdList.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdList.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdList.DisplayLayout.Override.CellPadding = 0
        Me.grdList.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Appearance3.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdList.DisplayLayout.Override.SelectedRowAppearance = Appearance3
        Me.grdList.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdList.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdList.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdList.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grdList.Location = New System.Drawing.Point(8, 18)
        Me.grdList.Name = "grdList"
        Me.grdList.Size = New System.Drawing.Size(336, 306)
        Me.grdList.TabIndex = 3
        Me.grdList.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChange
        '
        'CmMain
        '
        Me.CmMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.cmOpen, Me.cmDelete})
        '
        'cmOpen
        '
        Me.cmOpen.Index = 0
        Me.cmOpen.Text = "&Open"
        '
        'cmDelete
        '
        Me.cmDelete.Index = 1
        Me.cmDelete.Text = "&Delete"
        '
        'grpCost
        '
        Me.grpCost.BackColor = System.Drawing.Color.Transparent
        Me.grpCost.Controls.Add(Me.lblRange)
        Me.grpCost.Controls.Add(Me.lblPriceTo)
        Me.grpCost.Controls.Add(Me.txtPriceTo)
        Me.grpCost.Controls.Add(Me.txtPriceFrom)
        Me.grpCost.Location = New System.Drawing.Point(8, 136)
        Me.grpCost.Name = "grpCost"
        Me.grpCost.Size = New System.Drawing.Size(160, 80)
        Me.grpCost.TabIndex = 10
        Me.grpCost.TabStop = False
        Me.grpCost.Text = "Cost"
        '
        'lblRange
        '
        Me.lblRange.Location = New System.Drawing.Point(9, 19)
        Me.lblRange.Name = "lblRange"
        Me.lblRange.Size = New System.Drawing.Size(119, 16)
        Me.lblRange.TabIndex = 4
        Me.lblRange.Text = "Enter Price Range:"
        '
        'lblPriceTo
        '
        Me.lblPriceTo.Location = New System.Drawing.Point(57, 48)
        Me.lblPriceTo.Name = "lblPriceTo"
        Me.lblPriceTo.Size = New System.Drawing.Size(24, 16)
        Me.lblPriceTo.TabIndex = 3
        Me.lblPriceTo.Text = "To:"
        '
        'txtPriceTo
        '
        Me.txtPriceTo.Location = New System.Drawing.Point(86, 48)
        Me.txtPriceTo.Name = "txtPriceTo"
        Me.txtPriceTo.Size = New System.Drawing.Size(48, 21)
        Me.txtPriceTo.TabIndex = 0
        Me.txtPriceTo.Text = "100"
        '
        'txtPriceFrom
        '
        Me.txtPriceFrom.Location = New System.Drawing.Point(6, 48)
        Me.txtPriceFrom.Name = "txtPriceFrom"
        Me.txtPriceFrom.Size = New System.Drawing.Size(48, 21)
        Me.txtPriceFrom.TabIndex = 0
        Me.txtPriceFrom.Text = "0"
        '
        'ErrorProvider
        '
        Me.ErrorProvider.ContainerControl = Me
        '
        'frmMiscExpAdvSrch
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.ClientSize = New System.Drawing.Size(368, 561)
        Me.Controls.Add(Me.grpCost)
        Me.Controls.Add(Me.grpInvoices)
        Me.Controls.Add(Me.grpDates)
        Me.Controls.Add(Me.grpAction)
        Me.Controls.Add(Me.grpSearchText)
        Me.Controls.Add(Me.grpType)
        Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "frmMiscExpAdvSrch"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Advanced Misc Expense Search"
        Me.grpDates.ResumeLayout(False)
        Me.grpAction.ResumeLayout(False)
        Me.grpSearchText.ResumeLayout(False)
        Me.grpSearchText.PerformLayout()
        Me.grpType.ResumeLayout(False)
        Me.grpInvoices.ResumeLayout(False)
        CType(Me.grdList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpCost.ResumeLayout(False)
        Me.grpCost.PerformLayout()
        CType(Me.ErrorProvider, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub dtpFromDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpFromDate.ValueChanged, dtpToDate.ValueChanged

        Try

            Me.m_AdvSrch.FromDate = Convert.ToDateTime(Me.dtpFromDate.Value)
            Me.m_AdvSrch.ToDate = Convert.ToDateTime(Me.dtpToDate.Value)

            Me.dtpToDate.MinDate = dtpFromDate.Value

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub m_MatAdvSrch_HasData(ByVal Data As Boolean) Handles m_AdvSrch.HasData

        Me.btnOpen.Enabled = Data
        Me.btnDelete.Enabled = Data

    End Sub

    Private Sub txtSearchData_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSearchData.TextChanged

        Try

            Me.m_AdvSrch.SearchText = Me.txtSearchData.Text.Trim

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub rdoDesc_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdoDesc.CheckedChanged, rdoNotes.CheckedChanged, rdoPriceRange.CheckedChanged

        Try

            With Me.m_AdvSrch

                If Me.rdoDesc.Checked Then
                    .ReportType = MiscExpAdvSrch.SearchType.Desc
                ElseIf Me.rdoNotes.Checked Then
                    .ReportType = MiscExpAdvSrch.SearchType.Notes
                ElseIf Me.rdoPriceRange.Checked Then
                    .ReportType = MiscExpAdvSrch.SearchType.PriceRange
                End If

            End With

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub frmMatAdvSearch_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'set default dates
        Me.dtpFromDate.Value = DateSerial(Year(Now()), Month(Now()), 1)

        'Load Defaults for business object
        With m_AdvSrch
            .FromDate = Me.dtpFromDate.Value
            .ToDate = Now
        End With


    End Sub

    Protected Overridable Sub btnRefresh_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRefresh.Click

        Me.RefreshData()


    End Sub

    Private Sub txtPriceFrom_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPriceFrom.KeyPress
        e.Handled = Utilities.NumericOnly(e, Me.txtPriceFrom)

    End Sub

    Private Sub txtPriceTo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPriceTo.KeyPress

        e.Handled = Utilities.NumericOnly(e, Me.txtPriceTo)

    End Sub

    Private Sub txtPriceFrom_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPriceFrom.TextChanged

        If IsNumeric(Me.txtPriceFrom.Text) Then
            Me.m_AdvSrch.PriceFrom = Convert.ToDouble(Me.txtPriceFrom.Text)
        Else
            Me.m_AdvSrch.PriceFrom = 0
        End If


    End Sub

    Private Sub txtPriceTo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPriceTo.TextChanged

        If IsNumeric(Me.txtPriceTo.Text) Then
            Me.m_AdvSrch.PriceTo = Convert.ToDouble(Me.txtPriceTo.Text)
        Else
            Me.m_AdvSrch.PriceTo = 0
        End If

    End Sub

    Private Sub m_MatAdvSrch_IsPriceRange(ByVal bolPriceRange As Boolean) Handles m_AdvSrch.IsPriceRange

        Me.txtSearchData.Enabled = Not (bolPriceRange)
        Me.grpCost.Enabled = bolPriceRange

    End Sub

    Private Sub m_MatAdvSrch_ReadyToSearch(ByVal Valid As Boolean) Handles m_AdvSrch.ReadyToSearch

        Me.btnRefresh.Enabled = Valid

    End Sub

    Private Sub grdList_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdList.InitializeLayout

        With e.Layout.Bands(0)
            .Columns(0).Hidden = True
            .Columns(1).Header.Caption = "    Date"
            .Columns(1).Width = 85

            .Columns(2).Header.Caption = "Category"
            .Columns(2).Width = 110

            .Columns(3).Header.Caption = "Description"
            .Columns(3).Width = 110

            .Columns(4).Header.Caption = "Cost"
            .Columns(4).Width = 80
            .Columns(4).Format = "c"

            .Columns(5).Header.Caption = "Notes"
            .Columns(5).Width = 110

            .Columns(6).Hidden = True

        End With


    End Sub

    Private Sub m_MatAdvSrch_CostRangeValid(ByVal bolValid As Boolean) Handles m_AdvSrch.CostRangeValid

        If bolValid = False Then
            Me.ErrorProvider.SetError(Me.txtPriceTo, "'Price One' can't be greater than 'Price Two'")

        Else
            Me.ErrorProvider.SetError(Me.txtPriceTo, "")

        End If
    End Sub

    Private Sub OpenItem()


        If Me.grdList.Rows.VisibleRowCount > 0 Then

            Dim ItemNo As Integer = Convert.ToInt32(Me.grdList.ActiveRow.Cells(0).Value)

            Dim ds As DataSet = Me.Expenses.GetMiscExpItem(ItemNo)

            If ds.Tables(0).Rows.Count = 1 Then
                Utilities.ViewMiscExpItem(Me.MdiParent, Nothing, ds)

            End If
        End If


    End Sub

    Private Sub btnOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpen.Click

        OpenItem()

    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click

        Me.Close()

    End Sub
    Private Sub DeleteItem()

        If Me.grdList.Rows.VisibleRowCount > 0 Then
            If MessageBox.Show("Are you sure you want to delete this item?", ProgramName, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.Yes Then
                Dim ItemNo As Integer = Convert.ToInt32(Me.grdList.ActiveRow.Cells(0).Value)
                Me.Expenses.CrudMode = DelortBusObjects.MiscExp.DBMode.Delete
                Me.Expenses.ItemNo = ItemNo
                Me.Expenses.InsertUpdateDelete()
            End If

        End If

        Me.grdList.DataSource = Me.m_AdvSrch.GetData

    End Sub

    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click

        Me.DeleteItem()

    End Sub

    Private Sub cmOpen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmOpen.Click
        Me.OpenItem()

    End Sub

    Private Sub cmDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmDelete.Click

        Me.DeleteItem()

    End Sub

    Private Sub m_MatAdvSrch_TextChanged() Handles m_AdvSrch.TextChanged

        Me.grdList.Text = Me.m_AdvSrch.ReportText

    End Sub
    Protected Overridable Sub RefreshData()

        Try

            Cursor = Cursors.WaitCursor

            Me.grdList.DataSource = m_AdvSrch.GetData

        Catch ex As Exception

            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Cursor = Cursors.Default

        End Try

    End Sub

    Private Sub txtPriceFrom_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPriceFrom.Validated

        If Me.txtPriceFrom.Text.Trim.Length = 0 Then
            Me.txtPriceFrom.Text = "0"
        End If

    End Sub

    Private Sub txtPriceTo_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPriceTo.Validated

        If Me.txtPriceTo.Text.Trim.Length = 0 Then
            Me.txtPriceTo.Text = "0"
        End If

    End Sub

    Private Sub m_AdvSrch_SearchDataRequired(ByVal bolRequired As Boolean) Handles m_AdvSrch.SearchDataRequired

        If bolRequired = True Then
            Me.ErrorProvider.SetError(Me.txtSearchData, "Value Required")
        Else
            Me.ErrorProvider.SetError(Me.txtSearchData, "")
        End If

    End Sub

    Private Sub m_AdvSrch_DateRangeValid(ByVal isValid As Boolean) Handles m_AdvSrch.DateRangeValid

        If isValid = True Then
            Me.ErrorProvider.SetError(Me.dtpFromDate, "")
        Else
            Me.ErrorProvider.SetError(Me.dtpFromDate, "'From Date' Can't be later than 'To Date'")

        End If

    End Sub

    Private Sub m_AdvSrch_DataSourceChanged() Handles m_AdvSrch.DataSourceChanged

        Me.grdList.DataSource = Me.m_AdvSrch.DataSource

    End Sub
End Class
