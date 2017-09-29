Option Strict On

Imports DelortBusObjects
Imports System.Configuration.ConfigurationSettings

Public Class frmInvAdvancedSearch

    Inherits frmBaseWindow

    Private WithEvents m_AdvSearch As New InvAdvancedSearch(g_settings)


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
    Friend WithEvents rdoLaborDesc As System.Windows.Forms.RadioButton
    Friend WithEvents rdoMatDesc As System.Windows.Forms.RadioButton
    Friend WithEvents rdoInvNotes As System.Windows.Forms.RadioButton
    Friend WithEvents grpInvoices As System.Windows.Forms.GroupBox
    Friend WithEvents grdInvoices As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents grpAction As System.Windows.Forms.GroupBox
    Friend WithEvents txtSearchData As System.Windows.Forms.TextBox
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnOpen As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents cmMain As System.Windows.Forms.ContextMenu
    Friend WithEvents cmOpen As System.Windows.Forms.MenuItem
    Friend WithEvents cmDelete As System.Windows.Forms.MenuItem
    Friend WithEvents grpType As System.Windows.Forms.GroupBox
    Friend WithEvents grpSearchText As System.Windows.Forms.GroupBox
    Friend WithEvents grpDates As System.Windows.Forms.GroupBox
    Friend WithEvents dtpFromDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpToDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblFrom As System.Windows.Forms.Label
    Friend WithEvents lblTo As System.Windows.Forms.Label
    Friend WithEvents ErrorProvider As System.Windows.Forms.ErrorProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmInvAdvancedSearch))
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.cmMain = New System.Windows.Forms.ContextMenu()
        Me.cmOpen = New System.Windows.Forms.MenuItem()
        Me.cmDelete = New System.Windows.Forms.MenuItem()
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
        Me.grpInvoices = New System.Windows.Forms.GroupBox()
        Me.grdInvoices = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.grpType = New System.Windows.Forms.GroupBox()
        Me.rdoInvNotes = New System.Windows.Forms.RadioButton()
        Me.rdoMatDesc = New System.Windows.Forms.RadioButton()
        Me.rdoLaborDesc = New System.Windows.Forms.RadioButton()
        Me.ErrorProvider = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.grpDates.SuspendLayout()
        Me.grpAction.SuspendLayout()
        Me.grpSearchText.SuspendLayout()
        Me.grpInvoices.SuspendLayout()
        CType(Me.grdInvoices, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpType.SuspendLayout()
        CType(Me.ErrorProvider, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmMain
        '
        Me.cmMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.cmOpen, Me.cmDelete})
        '
        'cmOpen
        '
        Me.cmOpen.Index = 0
        Me.cmOpen.Text = "Open"
        '
        'cmDelete
        '
        Me.cmDelete.Index = 1
        Me.cmDelete.Text = "Delete"
        '
        'grpDates
        '
        Me.grpDates.BackColor = System.Drawing.Color.Transparent
        Me.grpDates.Controls.Add(Me.lblTo)
        Me.grpDates.Controls.Add(Me.lblFrom)
        Me.grpDates.Controls.Add(Me.dtpToDate)
        Me.grpDates.Controls.Add(Me.dtpFromDate)
        Me.grpDates.Location = New System.Drawing.Point(177, 8)
        Me.grpDates.Name = "grpDates"
        Me.grpDates.Size = New System.Drawing.Size(184, 72)
        Me.grpDates.TabIndex = 4
        Me.grpDates.TabStop = False
        Me.grpDates.Text = "Select Dates:"
        '
        'lblTo
        '
        Me.lblTo.Location = New System.Drawing.Point(12, 46)
        Me.lblTo.Name = "lblTo"
        Me.lblTo.Size = New System.Drawing.Size(28, 16)
        Me.lblTo.TabIndex = 3
        Me.lblTo.Text = "To:"
        '
        'lblFrom
        '
        Me.lblFrom.Location = New System.Drawing.Point(12, 15)
        Me.lblFrom.Name = "lblFrom"
        Me.lblFrom.Size = New System.Drawing.Size(45, 16)
        Me.lblFrom.TabIndex = 2
        Me.lblFrom.Text = "From:"
        '
        'dtpToDate
        '
        Me.dtpToDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpToDate.Location = New System.Drawing.Point(58, 46)
        Me.dtpToDate.Name = "dtpToDate"
        Me.dtpToDate.Size = New System.Drawing.Size(96, 21)
        Me.dtpToDate.TabIndex = 1
        '
        'dtpFromDate
        '
        Me.dtpFromDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFromDate.Location = New System.Drawing.Point(58, 15)
        Me.dtpFromDate.Name = "dtpFromDate"
        Me.dtpFromDate.Size = New System.Drawing.Size(96, 21)
        Me.dtpFromDate.TabIndex = 0
        '
        'grpAction
        '
        Me.grpAction.BackColor = System.Drawing.Color.Transparent
        Me.grpAction.Controls.Add(Me.btnRefresh)
        Me.grpAction.Controls.Add(Me.btnDelete)
        Me.grpAction.Controls.Add(Me.btnOpen)
        Me.grpAction.Controls.Add(Me.btnExit)
        Me.grpAction.Location = New System.Drawing.Point(372, 8)
        Me.grpAction.Name = "grpAction"
        Me.grpAction.Size = New System.Drawing.Size(96, 136)
        Me.grpAction.TabIndex = 3
        Me.grpAction.TabStop = False
        Me.grpAction.Text = "Action:"
        '
        'btnRefresh
        '
        Me.btnRefresh.BackColor = System.Drawing.Color.AliceBlue
        Me.btnRefresh.Enabled = False
        Me.btnRefresh.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRefresh.Image = CType(resources.GetObject("btnRefresh.Image"), System.Drawing.Image)
        Me.btnRefresh.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnRefresh.Location = New System.Drawing.Point(8, 80)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(80, 22)
        Me.btnRefresh.TabIndex = 13
        Me.btnRefresh.Text = "&Run"
        Me.btnRefresh.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnRefresh.UseVisualStyleBackColor = False
        '
        'btnDelete
        '
        Me.btnDelete.BackColor = System.Drawing.Color.AliceBlue
        Me.btnDelete.Enabled = False
        Me.btnDelete.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.Image = CType(resources.GetObject("btnDelete.Image"), System.Drawing.Image)
        Me.btnDelete.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDelete.Location = New System.Drawing.Point(9, 47)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 22)
        Me.btnDelete.TabIndex = 12
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnDelete.UseVisualStyleBackColor = False
        '
        'btnOpen
        '
        Me.btnOpen.BackColor = System.Drawing.Color.AliceBlue
        Me.btnOpen.Enabled = False
        Me.btnOpen.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOpen.Image = CType(resources.GetObject("btnOpen.Image"), System.Drawing.Image)
        Me.btnOpen.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnOpen.Location = New System.Drawing.Point(8, 17)
        Me.btnOpen.Name = "btnOpen"
        Me.btnOpen.Size = New System.Drawing.Size(80, 22)
        Me.btnOpen.TabIndex = 11
        Me.btnOpen.Text = "&Open"
        Me.btnOpen.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnOpen.UseVisualStyleBackColor = False
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.AliceBlue
        Me.btnExit.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.Image = CType(resources.GetObject("btnExit.Image"), System.Drawing.Image)
        Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnExit.Location = New System.Drawing.Point(8, 109)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 22)
        Me.btnExit.TabIndex = 10
        Me.btnExit.Text = "  E&xit"
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'grpSearchText
        '
        Me.grpSearchText.BackColor = System.Drawing.Color.Transparent
        Me.grpSearchText.Controls.Add(Me.txtSearchData)
        Me.grpSearchText.Location = New System.Drawing.Point(177, 88)
        Me.grpSearchText.Name = "grpSearchText"
        Me.grpSearchText.Size = New System.Drawing.Size(184, 56)
        Me.grpSearchText.TabIndex = 2
        Me.grpSearchText.TabStop = False
        Me.grpSearchText.Text = "Enter Data To Search For:"
        '
        'txtSearchData
        '
        Me.txtSearchData.Location = New System.Drawing.Point(10, 21)
        Me.txtSearchData.Multiline = True
        Me.txtSearchData.Name = "txtSearchData"
        Me.txtSearchData.Size = New System.Drawing.Size(155, 19)
        Me.txtSearchData.TabIndex = 0
        '
        'grpInvoices
        '
        Me.grpInvoices.BackColor = System.Drawing.Color.Transparent
        Me.grpInvoices.Controls.Add(Me.grdInvoices)
        Me.grpInvoices.Location = New System.Drawing.Point(8, 152)
        Me.grpInvoices.Name = "grpInvoices"
        Me.grpInvoices.Size = New System.Drawing.Size(460, 403)
        Me.grpInvoices.TabIndex = 1
        Me.grpInvoices.TabStop = False
        Me.grpInvoices.Text = "Search Results"
        '
        'grdInvoices
        '
        Me.grdInvoices.ContextMenu = Me.cmMain
        Appearance1.ForeColor = System.Drawing.SystemColors.ControlText
        Appearance1.TextHAlignAsString = "Left"
        Me.grdInvoices.DisplayLayout.Appearance = Appearance1
        Appearance2.ForeColor = System.Drawing.Color.Blue
        Appearance2.TextHAlignAsString = "Center"
        Me.grdInvoices.DisplayLayout.CaptionAppearance = Appearance2
        Me.grdInvoices.DisplayLayout.GroupByBox.Hidden = True
        Me.grdInvoices.DisplayLayout.MaxColScrollRegions = 1
        Me.grdInvoices.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdInvoices.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdInvoices.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdInvoices.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdInvoices.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdInvoices.DisplayLayout.Override.CellPadding = 0
        Me.grdInvoices.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Appearance3.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdInvoices.DisplayLayout.Override.SelectedRowAppearance = Appearance3
        Me.grdInvoices.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdInvoices.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdInvoices.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdInvoices.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grdInvoices.Location = New System.Drawing.Point(10, 19)
        Me.grdInvoices.Name = "grdInvoices"
        Me.grdInvoices.Size = New System.Drawing.Size(443, 375)
        Me.grdInvoices.TabIndex = 3
        Me.grdInvoices.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChange
        '
        'grpType
        '
        Me.grpType.BackColor = System.Drawing.Color.Transparent
        Me.grpType.Controls.Add(Me.rdoInvNotes)
        Me.grpType.Controls.Add(Me.rdoMatDesc)
        Me.grpType.Controls.Add(Me.rdoLaborDesc)
        Me.grpType.Location = New System.Drawing.Point(8, 8)
        Me.grpType.Name = "grpType"
        Me.grpType.Size = New System.Drawing.Size(158, 136)
        Me.grpType.TabIndex = 0
        Me.grpType.TabStop = False
        Me.grpType.Text = "Select Criteria:"
        '
        'rdoInvNotes
        '
        Me.rdoInvNotes.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoInvNotes.Location = New System.Drawing.Point(24, 102)
        Me.rdoInvNotes.Name = "rdoInvNotes"
        Me.rdoInvNotes.Size = New System.Drawing.Size(104, 16)
        Me.rdoInvNotes.TabIndex = 2
        Me.rdoInvNotes.Text = "Invoice Notes"
        '
        'rdoMatDesc
        '
        Me.rdoMatDesc.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoMatDesc.Location = New System.Drawing.Point(24, 62)
        Me.rdoMatDesc.Name = "rdoMatDesc"
        Me.rdoMatDesc.Size = New System.Drawing.Size(136, 16)
        Me.rdoMatDesc.TabIndex = 1
        Me.rdoMatDesc.Text = "Material Description"
        '
        'rdoLaborDesc
        '
        Me.rdoLaborDesc.Checked = True
        Me.rdoLaborDesc.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoLaborDesc.Location = New System.Drawing.Point(24, 22)
        Me.rdoLaborDesc.Name = "rdoLaborDesc"
        Me.rdoLaborDesc.Size = New System.Drawing.Size(128, 16)
        Me.rdoLaborDesc.TabIndex = 0
        Me.rdoLaborDesc.TabStop = True
        Me.rdoLaborDesc.Text = "Labor Description"
        '
        'ErrorProvider
        '
        Me.ErrorProvider.ContainerControl = Me
        '
        'frmInvAdvancedSearch
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.ClientSize = New System.Drawing.Size(473, 561)
        Me.Controls.Add(Me.grpDates)
        Me.Controls.Add(Me.grpAction)
        Me.Controls.Add(Me.grpSearchText)
        Me.Controls.Add(Me.grpInvoices)
        Me.Controls.Add(Me.grpType)
        Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "frmInvAdvancedSearch"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Advanced Invoice Search"
        Me.grpDates.ResumeLayout(False)
        Me.grpAction.ResumeLayout(False)
        Me.grpSearchText.ResumeLayout(False)
        Me.grpSearchText.PerformLayout()
        Me.grpInvoices.ResumeLayout(False)
        CType(Me.grdInvoices, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpType.ResumeLayout(False)
        CType(Me.ErrorProvider, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub txtSearchData_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSearchData.TextChanged

        Try

            Me.m_AdvSearch.SearchText = Me.txtSearchData.Text.Trim
            Me.grdInvoices.DataSource = Me.m_AdvSearch.DataSource

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub rdoInvNotes_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdoInvNotes.CheckedChanged, rdoLaborDesc.CheckedChanged, rdoMatDesc.CheckedChanged

        Try
            If Me.rdoInvNotes.Checked Then
                Me.m_AdvSearch.ReportType = InvAdvancedSearch.SearchType.InvoiceNotes
            ElseIf Me.rdoLaborDesc.Checked Then
                Me.m_AdvSearch.ReportType = InvAdvancedSearch.SearchType.LaborDesc
            ElseIf Me.rdoMatDesc.Checked Then
                Me.m_AdvSearch.ReportType = InvAdvancedSearch.SearchType.MaterialDesc
            End If

            Me.grdInvoices.DataSource = Me.m_AdvSearch.DataSource

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click

        Try

            Cursor = Cursors.WaitCursor
            Me.grdInvoices.DataSource = Me.m_AdvSearch.GetData()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            Cursor = Cursors.Default

        End Try

    End Sub

    Private Sub m_AdvSearch_HasData(ByVal Data As Boolean) Handles m_AdvSearch.HasData

        Me.btnDelete.Enabled = Data
        Me.btnOpen.Enabled = Data

    End Sub

    Private Sub m_AdvSearch_ReadyToSearch(ByVal Valid As Boolean) Handles m_AdvSearch.ReadyToSearch

        Me.btnRefresh.Enabled = Valid

    End Sub
    Private Sub OpenInvoice()

        If Me.grdInvoices.Rows.VisibleRowCount > 0 Then

            Dim InvNo As Integer = Convert.ToInt32(Me.grdInvoices.ActiveRow.Cells("InvoiceNo").Value)

            Utilities.OpenInvoice(InvNo, Nothing, Me.ParentForm)

        End If

    End Sub
    Private Sub DeleteInvoice()

        Dim invNum As Integer = Convert.ToInt32(Me.grdInvoices.ActiveRow.Cells("InvoiceNo").Value)

        If MessageBox.Show("Are You Sure You Want To Delete Invoice# " & invNum.ToString & "?", ProgramName, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.Yes Then

            Dim inv As New DelortBusObjects.Invoice(g_settings)
            inv.Delete(Convert.ToInt32(invNum))

            'RefreshReport
            Me.grdInvoices.DataSource = Me.m_AdvSearch.GetData()

        End If

    End Sub

    Private Sub btnOpen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOpen.Click

        Try

            OpenInvoice()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click

        Try

            Me.DeleteInvoice()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


    End Sub

    Private Sub grdInvoices_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdInvoices.DoubleClickRow
        Me.OpenInvoice()
    End Sub

    Private Sub grdInvoices_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdInvoices.InitializeLayout

        With e.Layout.Bands(0)
            .Columns("LastName").Header.Caption = "    Last Name"
            .Columns("LastName").Width = 110
            .Columns("FirstName").Header.Caption = "First Name"
            .Columns("FirstName").Width = 110

            .Columns("InvoiceNo").Header.Caption = "Invoice No."
            .Columns("InvoiceNo").Width = 85

            .Columns("InvoiceDate").Header.Caption = "Invoice Date"
            .Columns("InvoiceDate").Width = 110

        End With

    End Sub

    Private Sub cmOpen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmOpen.Click

        Try

            Me.OpenInvoice()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub cmDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmDelete.Click

        Try

            Me.DeleteInvoice()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click

        Try

            Me.Close()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub dtpFromDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpFromDate.ValueChanged, dtpToDate.ValueChanged

        Try

            Me.m_AdvSearch.FromDate = Convert.ToDateTime(Me.dtpFromDate.Value)
            Me.m_AdvSearch.ToDate = Convert.ToDateTime(Me.dtpToDate.Value)

            If Me.dtpFromDate.Value > Me.dtpToDate.Value Then
                Me.ErrorProvider.SetError(Me.dtpFromDate, "'From Date' Can't be later than 'To Date'")
            Else
                Me.ErrorProvider.SetError(Me.dtpFromDate, "")
            End If

            Me.grdInvoices.DataSource = Me.m_AdvSearch.DataSource

            Me.dtpToDate.MinDate = dtpFromDate.Value

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub frmInvAdvancedSearch_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            'set default dates
            Me.dtpFromDate.Value = DateSerial(Year(Now()), Month(Now()), 1)
            Me.m_AdvSearch.FromDate = Me.dtpFromDate.Value
            Me.m_AdvSearch.ToDate = Now

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


    End Sub


    Private Sub m_AdvSearch_ReportTextChanged() Handles m_AdvSearch.ReportTextChanged

        Me.grdInvoices.Text = Me.m_AdvSearch.ReportText

    End Sub

    Private Sub m_AdvSearch_SearchDataRequired(ByVal bolRequired As Boolean) Handles m_AdvSearch.SearchDataRequired

        If bolRequired Then
            Me.ErrorProvider.SetError(Me.txtSearchData, "Value required")

        Else
            Me.ErrorProvider.SetError(Me.txtSearchData, "")

        End If
    End Sub
End Class
