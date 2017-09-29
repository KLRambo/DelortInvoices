Option Strict On

Imports DelortBusObjects
Imports System.Configuration.ConfigurationSettings

Public Class frmCustomerInvoices
    Inherits frmBaseWindow

    Private WithEvents Customers As New Customer(g_settings)
    Private m_CusInvoices As New DelortBusObjects.CustomerReport(g_settings)
    Private m_DS As DataSet
    Private m_IsLoading As Boolean = True

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
    Friend WithEvents grdInvoices As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents grpAction As System.Windows.Forms.GroupBox
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnOpen As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents grpDate As System.Windows.Forms.GroupBox
    Friend WithEvents lblTo As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblFromDate As System.Windows.Forms.Label
    Friend WithEvents grpType As System.Windows.Forms.GroupBox
    Friend WithEvents rdoWithInvoices As System.Windows.Forms.RadioButton
    Friend WithEvents rdoNoInvs As System.Windows.Forms.RadioButton
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents cmMain As System.Windows.Forms.ContextMenu
    Friend WithEvents cmOpen As System.Windows.Forms.MenuItem
    Friend WithEvents cmDelete As System.Windows.Forms.MenuItem
    Friend WithEvents ToolTip As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCustomerInvoices))
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.cmMain = New System.Windows.Forms.ContextMenu()
        Me.cmOpen = New System.Windows.Forms.MenuItem()
        Me.cmDelete = New System.Windows.Forms.MenuItem()
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnOpen = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.grdInvoices = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.grpType = New System.Windows.Forms.GroupBox()
        Me.rdoNoInvs = New System.Windows.Forms.RadioButton()
        Me.rdoWithInvoices = New System.Windows.Forms.RadioButton()
        Me.grpDate = New System.Windows.Forms.GroupBox()
        Me.lblTo = New System.Windows.Forms.Label()
        Me.dtpTo = New System.Windows.Forms.DateTimePicker()
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker()
        Me.lblFromDate = New System.Windows.Forms.Label()
        Me.grpAction = New System.Windows.Forms.GroupBox()
        Me.grpMain = New System.Windows.Forms.GroupBox()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        CType(Me.grdInvoices, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpType.SuspendLayout()
        Me.grpDate.SuspendLayout()
        Me.grpAction.SuspendLayout()
        Me.grpMain.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
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
        'btnRefresh
        '
        Me.btnRefresh.BackColor = System.Drawing.Color.AliceBlue
        Me.btnRefresh.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRefresh.Image = CType(resources.GetObject("btnRefresh.Image"), System.Drawing.Image)
        Me.btnRefresh.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnRefresh.Location = New System.Drawing.Point(105, 16)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(78, 22)
        Me.btnRefresh.TabIndex = 9
        Me.btnRefresh.Text = "  &Run"
        Me.ToolTip.SetToolTip(Me.btnRefresh, "Run or Rerun Report")
        Me.btnRefresh.UseVisualStyleBackColor = False
        '
        'btnDelete
        '
        Me.btnDelete.BackColor = System.Drawing.Color.AliceBlue
        Me.btnDelete.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.Image = CType(resources.GetObject("btnDelete.Image"), System.Drawing.Image)
        Me.btnDelete.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDelete.Location = New System.Drawing.Point(10, 45)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(77, 22)
        Me.btnDelete.TabIndex = 7
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip.SetToolTip(Me.btnDelete, "Delete Customer or Invoice")
        Me.btnDelete.UseVisualStyleBackColor = False
        '
        'btnOpen
        '
        Me.btnOpen.BackColor = System.Drawing.Color.AliceBlue
        Me.btnOpen.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOpen.Image = CType(resources.GetObject("btnOpen.Image"), System.Drawing.Image)
        Me.btnOpen.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnOpen.Location = New System.Drawing.Point(10, 16)
        Me.btnOpen.Name = "btnOpen"
        Me.btnOpen.Size = New System.Drawing.Size(77, 22)
        Me.btnOpen.TabIndex = 6
        Me.btnOpen.Text = "&Open"
        Me.btnOpen.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip.SetToolTip(Me.btnOpen, "Open Customer or Invoice")
        Me.btnOpen.UseVisualStyleBackColor = False
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.AliceBlue
        Me.btnExit.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.Image = CType(resources.GetObject("btnExit.Image"), System.Drawing.Image)
        Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnExit.Location = New System.Drawing.Point(105, 45)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(78, 22)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "  E&xit"
        Me.ToolTip.SetToolTip(Me.btnExit, "Exit Form")
        Me.btnExit.UseVisualStyleBackColor = False
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
        Me.grdInvoices.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grdInvoices.Location = New System.Drawing.Point(5, 14)
        Me.grdInvoices.Name = "grdInvoices"
        Me.grdInvoices.Size = New System.Drawing.Size(556, 449)
        Me.grdInvoices.TabIndex = 2
        Me.grdInvoices.Text = "Customers"
        Me.ToolTip.SetToolTip(Me.grdInvoices, "Double Click to Open")
        Me.grdInvoices.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChange
        '
        'grpType
        '
        Me.grpType.BackColor = System.Drawing.Color.Transparent
        Me.grpType.Controls.Add(Me.rdoNoInvs)
        Me.grpType.Controls.Add(Me.rdoWithInvoices)
        Me.grpType.Location = New System.Drawing.Point(8, 8)
        Me.grpType.Name = "grpType"
        Me.grpType.Size = New System.Drawing.Size(201, 72)
        Me.grpType.TabIndex = 8
        Me.grpType.TabStop = False
        Me.grpType.Text = "Report Type"
        '
        'rdoNoInvs
        '
        Me.rdoNoInvs.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoNoInvs.Location = New System.Drawing.Point(8, 43)
        Me.rdoNoInvs.Name = "rdoNoInvs"
        Me.rdoNoInvs.Size = New System.Drawing.Size(188, 16)
        Me.rdoNoInvs.TabIndex = 13
        Me.rdoNoInvs.Text = "Customers With No Invoices"
        '
        'rdoWithInvoices
        '
        Me.rdoWithInvoices.Checked = True
        Me.rdoWithInvoices.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoWithInvoices.Location = New System.Drawing.Point(8, 16)
        Me.rdoWithInvoices.Name = "rdoWithInvoices"
        Me.rdoWithInvoices.Size = New System.Drawing.Size(168, 16)
        Me.rdoWithInvoices.TabIndex = 12
        Me.rdoWithInvoices.TabStop = True
        Me.rdoWithInvoices.Text = "Customers With Invoices"
        '
        'grpDate
        '
        Me.grpDate.BackColor = System.Drawing.Color.Transparent
        Me.grpDate.Controls.Add(Me.lblTo)
        Me.grpDate.Controls.Add(Me.dtpTo)
        Me.grpDate.Controls.Add(Me.dtpFrom)
        Me.grpDate.Controls.Add(Me.lblFromDate)
        Me.grpDate.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpDate.Location = New System.Drawing.Point(214, 8)
        Me.grpDate.Name = "grpDate"
        Me.grpDate.Size = New System.Drawing.Size(163, 72)
        Me.grpDate.TabIndex = 6
        Me.grpDate.TabStop = False
        Me.grpDate.Text = "Date Range"
        '
        'lblTo
        '
        Me.lblTo.Location = New System.Drawing.Point(5, 41)
        Me.lblTo.Name = "lblTo"
        Me.lblTo.Size = New System.Drawing.Size(32, 16)
        Me.lblTo.TabIndex = 18
        Me.lblTo.Text = "To:"
        '
        'dtpTo
        '
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpTo.Location = New System.Drawing.Point(42, 41)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(99, 21)
        Me.dtpTo.TabIndex = 15
        '
        'dtpFrom
        '
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFrom.Location = New System.Drawing.Point(42, 19)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(99, 21)
        Me.dtpFrom.TabIndex = 14
        '
        'lblFromDate
        '
        Me.lblFromDate.Location = New System.Drawing.Point(5, 19)
        Me.lblFromDate.Name = "lblFromDate"
        Me.lblFromDate.Size = New System.Drawing.Size(40, 16)
        Me.lblFromDate.TabIndex = 12
        Me.lblFromDate.Text = "Fr:"
        '
        'grpAction
        '
        Me.grpAction.BackColor = System.Drawing.Color.Transparent
        Me.grpAction.Controls.Add(Me.btnRefresh)
        Me.grpAction.Controls.Add(Me.btnDelete)
        Me.grpAction.Controls.Add(Me.btnOpen)
        Me.grpAction.Controls.Add(Me.btnExit)
        Me.grpAction.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpAction.Location = New System.Drawing.Point(383, 8)
        Me.grpAction.Name = "grpAction"
        Me.grpAction.Size = New System.Drawing.Size(191, 72)
        Me.grpAction.TabIndex = 5
        Me.grpAction.TabStop = False
        Me.grpAction.Text = "Action Buttons"
        '
        'grpMain
        '
        Me.grpMain.BackColor = System.Drawing.Color.Transparent
        Me.grpMain.Controls.Add(Me.grdInvoices)
        Me.grpMain.Location = New System.Drawing.Point(8, 86)
        Me.grpMain.Name = "grpMain"
        Me.grpMain.Size = New System.Drawing.Size(566, 472)
        Me.grpMain.TabIndex = 0
        Me.grpMain.TabStop = False
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'frmCustomerInvoices
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(581, 561)
        Me.Controls.Add(Me.grpType)
        Me.Controls.Add(Me.grpDate)
        Me.Controls.Add(Me.grpAction)
        Me.Controls.Add(Me.grpMain)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "frmCustomerInvoices"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Customer-Invoice Report"
        CType(Me.grdInvoices, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpType.ResumeLayout(False)
        Me.grpDate.ResumeLayout(False)
        Me.grpAction.ResumeLayout(False)
        Me.grpMain.ResumeLayout(False)
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub grdInvoices_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdInvoices.DoubleClickRow
        Me.OpenEntity()
    End Sub

    Private Sub grdInvoices_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdInvoices.InitializeLayout

        If Me.m_DS.Tables.Count = 2 Then

            With e.Layout.Bands(0)
                .Columns(0).Hidden = True
                .Columns(1).Header.Caption = "     Last Name"
                .Columns(1).Width = 150
                .Columns(2).Header.Caption = "First Name"
                .Columns(1).Width = 150
                .Columns(3).Header.Caption = "Invoice Count"
                .Columns(3).Width = 140
                .Columns(4).Header.Caption = "Total"
                .Columns(4).Width = 100
                .Columns(4).Format = ("c")
            End With
            With e.Layout.Bands(1)
                .Columns(0).Hidden = True
                .Columns(1).Header.Caption = "Invoice No"
                .Columns(1).Width = 100
                .Columns(2).Header.Caption = "Invoice Date"
                .Columns(1).Width = 100
                .Columns(3).Header.Caption = "Invoice Total"
                .Columns(3).Width = 90
                .Columns(3).Format = "c"
            End With
            Me.grdInvoices.Text = Me.grdInvoices.Rows.VisibleRowCount & " Customer(s) With Invoices Between " & Me.dtpFrom.Value.ToShortDateString & " and " & Me.dtpTo.Value.ToShortDateString
        Else

            With e.Layout.Bands(0)
                .Columns(0).Hidden = True
                .Columns(1).Header.Caption = "     Last Name"
                .Columns(1).Width = 100
                .Columns(2).Header.Caption = "First Name"
                .Columns(1).Width = 80
                .Columns(3).Header.Caption = "Addr1"
                .Columns(3).Width = 160
                .Columns(4).Header.Caption = "Addr2"
                .Columns(4).Width = 85
                .Columns(5).Header.Caption = "City"
                .Columns(5).Width = 100

            End With
            Me.grdInvoices.Text = Me.grdInvoices.Rows.VisibleRowCount & " Customer(s) With NO Invoices Between " & Me.dtpFrom.Value.ToShortDateString & " and " & Me.dtpTo.Value.ToShortDateString

        End If

        'Hack city
        btnDelete.Enabled = grdInvoices.Rows.VisibleRowCount <> 0
        btnOpen.Enabled = grdInvoices.Rows.VisibleRowCount <> 0

    End Sub

    Private Sub OpenInvoice()

        If Me.grdInvoices.Rows.VisibleRowCount > 0 Then
            Dim InvNo As Integer = Convert.ToInt32(Me.grdInvoices.ActiveRow.Cells("InvoiceNo").Value)

            Utilities.OpenInvoice(InvNo, Nothing, Me.ParentForm)

        End If

    End Sub

    Private Sub OpenCustomer()

        If Me.grdInvoices.Rows.VisibleRowCount > 0 Then
            Dim CustNo As Integer = Convert.ToInt32(Me.grdInvoices.ActiveRow.Cells("CustomerID").Value)
            Dim CusObj As DelortBusObjects.Customer = Me.Customers.GetCustomer(CustNo)
            Utilities.ViewAddEditCustomer(Me.ParentForm, CusObj, Nothing)
        End If
    End Sub

    Private Sub btnOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpen.Click

        Try
            Me.OpenEntity()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub frmCustomerInvoices_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.dtpFrom.Value = DateSerial(Year(Now()), Month(Now()), 1)

        Me.GetCustomersInvoices()

        m_IsLoading = False

    End Sub

    Private Sub GetCustomersInvoices()

        Me.grdInvoices.DataSource = Nothing
        Me.m_DS = m_CusInvoices.GetCustomersWithInvoices(Me.dtpFrom.Value.ToShortDateString, Me.dtpTo.Value.ToShortDateString)
        Me.grdInvoices.DataSource = Me.m_DS

    End Sub

    Private Sub rdoNoInvs_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdoNoInvs.Click

        Try
            Me.GetCustomersWithNoInvoices()


        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub rdoWithInvoices_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdoWithInvoices.Click

        Try
            If Me.rdoWithInvoices.Checked Then
                If Me.m_IsLoading = False Then
                    GetCustomersInvoices()
                End If
            End If

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
    Private Sub DeleteCustomer()

        If Me.grdInvoices.Rows.VisibleRowCount > 0 Then
            Dim FirstName As String = Me.grdInvoices.ActiveRow.Cells("FirstName").Value.ToString
            Dim Lastname As String = Me.grdInvoices.ActiveRow.Cells("Lastname").Value.ToString

            If MessageBox.Show("Are you sure you want to delete " & FirstName & " " & Lastname & "?", ProgramName, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.Yes Then
                Me.Customers.DBMode = Customer.CrudMode.Delete
                Me.Customers.CustID = Convert.ToInt32(Me.grdInvoices.ActiveRow.Cells(0).Value)
                Me.Customers.InsertUpdateDelete()
            End If

        End If
    End Sub

    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click

        Try

            Me.DeleteEntity()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
    Private Sub DeleteInvoice()

        Dim invNum As Integer = Convert.ToInt32(Me.grdInvoices.ActiveRow.Cells("InvoiceNo").Value)

        If MessageBox.Show("Are You Sure You Want To Delete Invoice# " & invNum.ToString & "?", ProgramName, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.Yes Then

            Dim inv As New DelortBusObjects.Invoice(g_settings)
            inv.Delete(Convert.ToInt32(invNum))

            Me.RefreshReport()

        End If

    End Sub
    Private Sub RefreshReport()

        If Me.rdoWithInvoices.Checked Then
            GetCustomersInvoices()
        Else
            Me.GetCustomersWithNoInvoices()
        End If

    End Sub
    Private Sub GetCustomersWithNoInvoices()

        Me.grdInvoices.DataSource = Nothing

        Me.m_DS = Me.m_CusInvoices.GetCustomersWithNoInvoices(Me.dtpFrom.Value.ToShortDateString, Me.dtpTo.Value.ToShortDateString)

        Me.grdInvoices.DataSource = Me.m_DS

    End Sub

    Private Sub btnRefresh_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRefresh.Click

        Try
            Me.RefreshReport()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub dtpFrom_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpFrom.ValueChanged, dtpTo.ValueChanged

        Try

            If Convert.ToDateTime(Me.dtpTo.Value) < Convert.ToDateTime(Me.dtpFrom.Value) Then
                Me.grpAction.Enabled = False
                Me.grpType.Enabled = False
                Me.ErrorProvider1.SetError(Me.dtpFrom, "'From Date' can't be later than 'To Date'.")
            Else
                Me.grpAction.Enabled = True
                Me.grpType.Enabled = True
                Me.ErrorProvider1.SetError(Me.dtpFrom, "")
                Me.ErrorProvider1.SetError(Me.dtpTo, "")

            End If

            If Me.m_IsLoading = False Then
                Me.grdInvoices.DataSource = Nothing

            End If

            Me.dtpTo.MinDate = dtpFrom.Value

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub cmDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmDelete.Click

        Try

            DeleteEntity()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
    Private Sub DeleteEntity()

        If Me.grdInvoices.Rows.VisibleRowCount = 0 Then
            Exit Sub
        End If

        If Me.grdInvoices.ActiveRow.Cells.Exists("InvoiceNo") Then
            Me.DeleteInvoice()
        Else
            DeleteCustomer()
        End If

    End Sub
    Private Sub OpenEntity()

        If Me.grdInvoices.Rows.VisibleRowCount = 0 Then
            Exit Sub
        End If

        If Me.grdInvoices.ActiveRow.Cells.Exists("InvoiceNo") Then
            OpenInvoice()
        Else
            Me.OpenCustomer()

        End If

    End Sub

    Private Sub cmOpen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmOpen.Click

        Try

            Me.OpenEntity()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click

        Try

            Me.m_CusInvoices = Nothing
            Me.m_DS.Dispose()
            Me.m_DS = Nothing

            Me.Close()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
End Class
