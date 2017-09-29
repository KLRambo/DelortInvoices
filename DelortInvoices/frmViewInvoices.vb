Option Strict On

Imports AccessUtils
Imports DelortBusObjects
Imports System.Configuration.ConfigurationSettings

Public Class frmViewEditInvoices

    Inherits frmBaseWindow

    Private CustomerID As Integer
    Private Date1 As Date
    Private Date2 As Date
    Private dsInv As DataSet
    Private Loading As Boolean = True


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
    Friend WithEvents grpInv As System.Windows.Forms.GroupBox
    Friend WithEvents dgInv As System.Windows.Forms.DataGrid
    Friend WithEvents cmnuDG As System.Windows.Forms.ContextMenu
    Friend WithEvents cboCustomers As System.Windows.Forms.ComboBox
    Friend WithEvents rdoAllCustomers As System.Windows.Forms.RadioButton
    Friend WithEvents rdoSelectCust As System.Windows.Forms.RadioButton
    Friend WithEvents grpFooter As System.Windows.Forms.GroupBox
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnRun As System.Windows.Forms.Button
    Friend WithEvents grpCustomers As System.Windows.Forms.GroupBox
    Friend WithEvents grpDate As System.Windows.Forms.GroupBox
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmnuView As System.Windows.Forms.MenuItem
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents cmnuDelete As System.Windows.Forms.MenuItem
    Friend WithEvents grpTotals As System.Windows.Forms.GroupBox
    Friend WithEvents lblLaborDisplay As System.Windows.Forms.Label
    Friend WithEvents lblLaborTotal As System.Windows.Forms.Label
    Friend WithEvents lblMaterialDisplay As System.Windows.Forms.Label
    Friend WithEvents lblMaterialTotal As System.Windows.Forms.Label
    Friend WithEvents lblTotal As System.Windows.Forms.Label
    Friend WithEvents lblTotalDisplay As System.Windows.Forms.Label
    Friend WithEvents btnOpenInvoice As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents cmnuPrint As System.Windows.Forms.MenuItem
    Friend WithEvents ErrProv As System.Windows.Forms.ErrorProvider
    Friend WithEvents lblCusName As System.Windows.Forms.Label
    Friend WithEvents lblFromDate As System.Windows.Forms.Label
    Friend WithEvents lblTo As System.Windows.Forms.Label
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmViewEditInvoices))
        Me.grpInv = New System.Windows.Forms.GroupBox()
        Me.dgInv = New System.Windows.Forms.DataGrid()
        Me.cmnuDG = New System.Windows.Forms.ContextMenu()
        Me.cmnuView = New System.Windows.Forms.MenuItem()
        Me.cmnuDelete = New System.Windows.Forms.MenuItem()
        Me.cmnuPrint = New System.Windows.Forms.MenuItem()
        Me.grpCustomers = New System.Windows.Forms.GroupBox()
        Me.lblCusName = New System.Windows.Forms.Label()
        Me.rdoSelectCust = New System.Windows.Forms.RadioButton()
        Me.rdoAllCustomers = New System.Windows.Forms.RadioButton()
        Me.cboCustomers = New System.Windows.Forms.ComboBox()
        Me.grpFooter = New System.Windows.Forms.GroupBox()
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnOpenInvoice = New System.Windows.Forms.Button()
        Me.btnRun = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.grpDate = New System.Windows.Forms.GroupBox()
        Me.lblTo = New System.Windows.Forms.Label()
        Me.dtpTo = New System.Windows.Forms.DateTimePicker()
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker()
        Me.lblFromDate = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.grpTotals = New System.Windows.Forms.GroupBox()
        Me.lblTotal = New System.Windows.Forms.Label()
        Me.lblTotalDisplay = New System.Windows.Forms.Label()
        Me.lblMaterialTotal = New System.Windows.Forms.Label()
        Me.lblMaterialDisplay = New System.Windows.Forms.Label()
        Me.lblLaborTotal = New System.Windows.Forms.Label()
        Me.lblLaborDisplay = New System.Windows.Forms.Label()
        Me.ErrProv = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.grpInv.SuspendLayout()
        CType(Me.dgInv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpCustomers.SuspendLayout()
        Me.grpFooter.SuspendLayout()
        Me.grpDate.SuspendLayout()
        Me.grpTotals.SuspendLayout()
        CType(Me.ErrProv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grpInv
        '
        Me.grpInv.BackColor = System.Drawing.Color.Transparent
        Me.grpInv.Controls.Add(Me.dgInv)
        Me.grpInv.Location = New System.Drawing.Point(12, 121)
        Me.grpInv.Name = "grpInv"
        Me.grpInv.Size = New System.Drawing.Size(508, 387)
        Me.grpInv.TabIndex = 0
        Me.grpInv.TabStop = False
        '
        'dgInv
        '
        Me.dgInv.AlternatingBackColor = System.Drawing.Color.Lavender
        Me.dgInv.BackColor = System.Drawing.Color.WhiteSmoke
        Me.dgInv.BackgroundColor = System.Drawing.Color.LightGray
        Me.dgInv.CaptionBackColor = System.Drawing.Color.LightSteelBlue
        Me.dgInv.CaptionFont = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.dgInv.CaptionForeColor = System.Drawing.Color.MidnightBlue
        Me.dgInv.CaptionText = "Click On Column Headers To Sort"
        Me.dgInv.CaptionVisible = False
        Me.dgInv.ContextMenu = Me.cmnuDG
        Me.dgInv.DataMember = ""
        Me.dgInv.FlatMode = True
        Me.dgInv.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dgInv.ForeColor = System.Drawing.Color.MidnightBlue
        Me.dgInv.GridLineColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.dgInv.GridLineStyle = System.Windows.Forms.DataGridLineStyle.None
        Me.dgInv.HeaderBackColor = System.Drawing.Color.MidnightBlue
        Me.dgInv.HeaderFont = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.dgInv.HeaderForeColor = System.Drawing.Color.WhiteSmoke
        Me.dgInv.LinkColor = System.Drawing.Color.Teal
        Me.dgInv.Location = New System.Drawing.Point(7, 16)
        Me.dgInv.Name = "dgInv"
        Me.dgInv.ParentRowsBackColor = System.Drawing.Color.AliceBlue
        Me.dgInv.ParentRowsForeColor = System.Drawing.Color.MidnightBlue
        Me.dgInv.SelectionBackColor = System.Drawing.Color.CadetBlue
        Me.dgInv.SelectionForeColor = System.Drawing.Color.WhiteSmoke
        Me.dgInv.Size = New System.Drawing.Size(489, 366)
        Me.dgInv.TabIndex = 1
        '
        'cmnuDG
        '
        Me.cmnuDG.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.cmnuView, Me.cmnuDelete, Me.cmnuPrint})
        '
        'cmnuView
        '
        Me.cmnuView.Index = 0
        Me.cmnuView.Text = "Open Invoice"
        '
        'cmnuDelete
        '
        Me.cmnuDelete.Index = 1
        Me.cmnuDelete.Text = "Delete Invoice"
        '
        'cmnuPrint
        '
        Me.cmnuPrint.Index = 2
        Me.cmnuPrint.Text = "Print Invoice"
        '
        'grpCustomers
        '
        Me.grpCustomers.BackColor = System.Drawing.Color.Transparent
        Me.grpCustomers.Controls.Add(Me.lblCusName)
        Me.grpCustomers.Controls.Add(Me.rdoSelectCust)
        Me.grpCustomers.Controls.Add(Me.rdoAllCustomers)
        Me.grpCustomers.Controls.Add(Me.cboCustomers)
        Me.grpCustomers.Location = New System.Drawing.Point(12, 3)
        Me.grpCustomers.Name = "grpCustomers"
        Me.grpCustomers.Size = New System.Drawing.Size(256, 72)
        Me.grpCustomers.TabIndex = 1
        Me.grpCustomers.TabStop = False
        Me.grpCustomers.Text = "Customers"
        '
        'lblCusName
        '
        Me.lblCusName.Location = New System.Drawing.Point(88, 16)
        Me.lblCusName.Name = "lblCusName"
        Me.lblCusName.Size = New System.Drawing.Size(160, 16)
        Me.lblCusName.TabIndex = 3
        Me.lblCusName.Text = "Customer Name"
        '
        'rdoSelectCust
        '
        Me.rdoSelectCust.Location = New System.Drawing.Point(15, 38)
        Me.rdoSelectCust.Name = "rdoSelectCust"
        Me.rdoSelectCust.Size = New System.Drawing.Size(72, 24)
        Me.rdoSelectCust.TabIndex = 2
        Me.rdoSelectCust.Text = "Select"
        '
        'rdoAllCustomers
        '
        Me.rdoAllCustomers.Checked = True
        Me.rdoAllCustomers.Location = New System.Drawing.Point(15, 16)
        Me.rdoAllCustomers.Name = "rdoAllCustomers"
        Me.rdoAllCustomers.Size = New System.Drawing.Size(56, 24)
        Me.rdoAllCustomers.TabIndex = 1
        Me.rdoAllCustomers.TabStop = True
        Me.rdoAllCustomers.Text = "All"
        '
        'cboCustomers
        '
        Me.cboCustomers.Location = New System.Drawing.Point(88, 38)
        Me.cboCustomers.Name = "cboCustomers"
        Me.cboCustomers.Size = New System.Drawing.Size(157, 22)
        Me.cboCustomers.TabIndex = 0
        '
        'grpFooter
        '
        Me.grpFooter.BackColor = System.Drawing.Color.Transparent
        Me.grpFooter.Controls.Add(Me.btnRefresh)
        Me.grpFooter.Controls.Add(Me.btnPrint)
        Me.grpFooter.Controls.Add(Me.btnDelete)
        Me.grpFooter.Controls.Add(Me.btnOpenInvoice)
        Me.grpFooter.Controls.Add(Me.btnRun)
        Me.grpFooter.Controls.Add(Me.btnExit)
        Me.grpFooter.Location = New System.Drawing.Point(12, 77)
        Me.grpFooter.Name = "grpFooter"
        Me.grpFooter.Size = New System.Drawing.Size(508, 42)
        Me.grpFooter.TabIndex = 3
        Me.grpFooter.TabStop = False
        '
        'btnRefresh
        '
        Me.btnRefresh.BackColor = System.Drawing.Color.AliceBlue
        Me.btnRefresh.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRefresh.Image = CType(resources.GetObject("btnRefresh.Image"), System.Drawing.Image)
        Me.btnRefresh.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnRefresh.Location = New System.Drawing.Point(338, 13)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(80, 22)
        Me.btnRefresh.TabIndex = 5
        Me.btnRefresh.Text = "R&efresh"
        Me.btnRefresh.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.btnRefresh, "Refresh the Report")
        Me.btnRefresh.UseVisualStyleBackColor = False
        '
        'btnPrint
        '
        Me.btnPrint.BackColor = System.Drawing.Color.AliceBlue
        Me.btnPrint.Enabled = False
        Me.btnPrint.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrint.Image = CType(resources.GetObject("btnPrint.Image"), System.Drawing.Image)
        Me.btnPrint.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnPrint.Location = New System.Drawing.Point(258, 13)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(80, 22)
        Me.btnPrint.TabIndex = 4
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.btnPrint, "Print Selected Invoice")
        Me.btnPrint.UseVisualStyleBackColor = False
        '
        'btnDelete
        '
        Me.btnDelete.BackColor = System.Drawing.Color.AliceBlue
        Me.btnDelete.Enabled = False
        Me.btnDelete.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.Image = CType(resources.GetObject("btnDelete.Image"), System.Drawing.Image)
        Me.btnDelete.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDelete.Location = New System.Drawing.Point(178, 13)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 22)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.btnDelete, "Delete Selected Invoice")
        Me.btnDelete.UseVisualStyleBackColor = False
        '
        'btnOpenInvoice
        '
        Me.btnOpenInvoice.BackColor = System.Drawing.Color.AliceBlue
        Me.btnOpenInvoice.Enabled = False
        Me.btnOpenInvoice.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOpenInvoice.Image = CType(resources.GetObject("btnOpenInvoice.Image"), System.Drawing.Image)
        Me.btnOpenInvoice.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnOpenInvoice.Location = New System.Drawing.Point(98, 13)
        Me.btnOpenInvoice.Name = "btnOpenInvoice"
        Me.btnOpenInvoice.Size = New System.Drawing.Size(80, 22)
        Me.btnOpenInvoice.TabIndex = 2
        Me.btnOpenInvoice.Text = "&Open"
        Me.btnOpenInvoice.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.btnOpenInvoice, "Open Selected Invoice")
        Me.btnOpenInvoice.UseVisualStyleBackColor = False
        '
        'btnRun
        '
        Me.btnRun.BackColor = System.Drawing.Color.AliceBlue
        Me.btnRun.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRun.Image = CType(resources.GetObject("btnRun.Image"), System.Drawing.Image)
        Me.btnRun.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnRun.Location = New System.Drawing.Point(418, 13)
        Me.btnRun.Name = "btnRun"
        Me.btnRun.Size = New System.Drawing.Size(85, 22)
        Me.btnRun.TabIndex = 1
        Me.btnRun.Text = "   &Run"
        Me.ToolTip1.SetToolTip(Me.btnRun, "Run the Report")
        Me.btnRun.UseVisualStyleBackColor = False
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.AliceBlue
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.Image = CType(resources.GetObject("btnExit.Image"), System.Drawing.Image)
        Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnExit.Location = New System.Drawing.Point(11, 13)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(87, 22)
        Me.btnExit.TabIndex = 0
        Me.btnExit.Text = "E&xit"
        Me.ToolTip1.SetToolTip(Me.btnExit, "Close Form")
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'grpDate
        '
        Me.grpDate.BackColor = System.Drawing.Color.Transparent
        Me.grpDate.Controls.Add(Me.lblTo)
        Me.grpDate.Controls.Add(Me.dtpTo)
        Me.grpDate.Controls.Add(Me.dtpFrom)
        Me.grpDate.Controls.Add(Me.lblFromDate)
        Me.grpDate.Location = New System.Drawing.Point(270, 3)
        Me.grpDate.Name = "grpDate"
        Me.grpDate.Size = New System.Drawing.Size(250, 72)
        Me.grpDate.TabIndex = 4
        Me.grpDate.TabStop = False
        Me.grpDate.Text = "Date Range"
        '
        'lblTo
        '
        Me.lblTo.Location = New System.Drawing.Point(18, 41)
        Me.lblTo.Name = "lblTo"
        Me.lblTo.Size = New System.Drawing.Size(62, 16)
        Me.lblTo.TabIndex = 18
        Me.lblTo.Text = "To Date:"
        '
        'dtpTo
        '
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpTo.Location = New System.Drawing.Point(118, 41)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(113, 22)
        Me.dtpTo.TabIndex = 15
        '
        'dtpFrom
        '
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFrom.Location = New System.Drawing.Point(118, 19)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(113, 22)
        Me.dtpFrom.TabIndex = 14
        '
        'lblFromDate
        '
        Me.lblFromDate.Location = New System.Drawing.Point(18, 19)
        Me.lblFromDate.Name = "lblFromDate"
        Me.lblFromDate.Size = New System.Drawing.Size(94, 16)
        Me.lblFromDate.TabIndex = 12
        Me.lblFromDate.Text = "From Date:"
        '
        'grpTotals
        '
        Me.grpTotals.BackColor = System.Drawing.Color.Transparent
        Me.grpTotals.Controls.Add(Me.lblTotal)
        Me.grpTotals.Controls.Add(Me.lblTotalDisplay)
        Me.grpTotals.Controls.Add(Me.lblMaterialTotal)
        Me.grpTotals.Controls.Add(Me.lblMaterialDisplay)
        Me.grpTotals.Controls.Add(Me.lblLaborTotal)
        Me.grpTotals.Controls.Add(Me.lblLaborDisplay)
        Me.grpTotals.Location = New System.Drawing.Point(12, 510)
        Me.grpTotals.Name = "grpTotals"
        Me.grpTotals.Size = New System.Drawing.Size(508, 48)
        Me.grpTotals.TabIndex = 5
        Me.grpTotals.TabStop = False
        Me.grpTotals.Text = "Report Totals"
        '
        'lblTotal
        '
        Me.lblTotal.BackColor = System.Drawing.Color.White
        Me.lblTotal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTotal.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotal.Location = New System.Drawing.Point(394, 19)
        Me.lblTotal.Name = "lblTotal"
        Me.lblTotal.Size = New System.Drawing.Size(85, 16)
        Me.lblTotal.TabIndex = 5
        Me.lblTotal.Text = "$0.00"
        Me.lblTotal.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblTotalDisplay
        '
        Me.lblTotalDisplay.Location = New System.Drawing.Point(336, 19)
        Me.lblTotalDisplay.Name = "lblTotalDisplay"
        Me.lblTotalDisplay.Size = New System.Drawing.Size(48, 16)
        Me.lblTotalDisplay.TabIndex = 4
        Me.lblTotalDisplay.Text = "Total:"
        '
        'lblMaterialTotal
        '
        Me.lblMaterialTotal.BackColor = System.Drawing.Color.White
        Me.lblMaterialTotal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblMaterialTotal.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMaterialTotal.Location = New System.Drawing.Point(243, 19)
        Me.lblMaterialTotal.Name = "lblMaterialTotal"
        Me.lblMaterialTotal.Size = New System.Drawing.Size(85, 16)
        Me.lblMaterialTotal.TabIndex = 3
        Me.lblMaterialTotal.Text = "$0.00"
        Me.lblMaterialTotal.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblMaterialDisplay
        '
        Me.lblMaterialDisplay.Location = New System.Drawing.Point(170, 19)
        Me.lblMaterialDisplay.Name = "lblMaterialDisplay"
        Me.lblMaterialDisplay.Size = New System.Drawing.Size(72, 16)
        Me.lblMaterialDisplay.TabIndex = 2
        Me.lblMaterialDisplay.Text = "Materials:"
        '
        'lblLaborTotal
        '
        Me.lblLaborTotal.BackColor = System.Drawing.Color.White
        Me.lblLaborTotal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLaborTotal.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLaborTotal.Location = New System.Drawing.Point(67, 19)
        Me.lblLaborTotal.Name = "lblLaborTotal"
        Me.lblLaborTotal.Size = New System.Drawing.Size(85, 16)
        Me.lblLaborTotal.TabIndex = 1
        Me.lblLaborTotal.Text = "$0.00"
        Me.lblLaborTotal.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblLaborDisplay
        '
        Me.lblLaborDisplay.Location = New System.Drawing.Point(8, 19)
        Me.lblLaborDisplay.Name = "lblLaborDisplay"
        Me.lblLaborDisplay.Size = New System.Drawing.Size(48, 16)
        Me.lblLaborDisplay.TabIndex = 0
        Me.lblLaborDisplay.Text = "Labor:"
        '
        'ErrProv
        '
        Me.ErrProv.ContainerControl = Me
        '
        'frmViewEditInvoices
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 15)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(526, 561)
        Me.Controls.Add(Me.grpTotals)
        Me.Controls.Add(Me.grpDate)
        Me.Controls.Add(Me.grpFooter)
        Me.Controls.Add(Me.grpCustomers)
        Me.Controls.Add(Me.grpInv)
        Me.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmViewEditInvoices"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "View Invoices"
        Me.grpInv.ResumeLayout(False)
        CType(Me.dgInv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpCustomers.ResumeLayout(False)
        Me.grpFooter.ResumeLayout(False)
        Me.grpDate.ResumeLayout(False)
        Me.grpTotals.ResumeLayout(False)
        CType(Me.ErrProv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmViewEditInvoices_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Loading = True

            Me.SetUpUI()

            Loading = False

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub LoadCustomerData()

        Dim ds As New DataSet("Invoices")
        Dim inv As New DelortBusObjects.Invoice(g_Settings)

        If Me.Date1 = #1/1/1799# Then
            ds = inv.GetCustomerInvoices(Me.CustomerID)
        Else
            ds = inv.GetCustomerInvoices(Me.CustomerID, Me.Date1.ToShortDateString, Me.Date2.ToShortDateString)

        End If

        SetUpDataGrid("Customer")

        Me.dgInv.DataSource = ds.Tables("Invoices")
        TallyReportTotals(ds.Tables(0))

    End Sub

    Private Sub SetUpDataGrid(ByVal Type As String)

        With dgInv
            .ReadOnly = True
            .TableStyles.Clear()
        End With


        Dim ts As New DataGridTableStyle
        Dim col As New DataGridTextBoxColumn

        With ts
            .MappingName = "Invoices"
            .AllowSorting = True
            .GridLineColor = Color.DarkBlue
            .HeaderBackColor = Color.DarkBlue
            .HeaderForeColor = Color.White
        End With



        With col
            .MappingName = "InvoiceNo"
            .HeaderText = "Invoice No."
            If Type = "All" Then
                .Width = 80
            Else
                .Width = 130
            End If

            .Alignment = HorizontalAlignment.Left
            ts.GridColumnStyles.Add(col)
        End With

        col = New DataGridTextBoxColumn

        If Type = "All" Then
            col = New DataGridTextBoxColumn
            With col
                .MappingName = "LastName"
                .HeaderText = "Last Name"
                .Width = 90
                .Alignment = HorizontalAlignment.Left
                ts.GridColumnStyles.Add(col)
            End With

            Me.dgInv.TableStyles.Add(ts)
            col = New DataGridTextBoxColumn
            With col
                .MappingName = "FirstName"
                .HeaderText = "First Name"
                .Width = 90
                .Alignment = HorizontalAlignment.Left
                ts.GridColumnStyles.Add(col)
            End With

            Me.dgInv.TableStyles.Add(ts)


        End If

        col = New DataGridTextBoxColumn
        With col
            .MappingName = "InvoiceDate"
            .HeaderText = "Date"
            If Type = "All" Then
                .Width = 100
            Else
                .Width = 150
            End If

            .Alignment = HorizontalAlignment.Left
            ts.GridColumnStyles.Add(col)
        End With

        Me.dgInv.TableStyles.Add(ts)


        col = New DataGridTextBoxColumn

        With col
            .MappingName = "LaborTotal"
            .HeaderText = "Labor"
            If Type = "All" Then
                .Width = 90
            Else
                .Width = 100
            End If
            .Alignment = HorizontalAlignment.Left
            .Format = "c"
            ts.GridColumnStyles.Add(col)
        End With


        col = New DataGridTextBoxColumn

        With col
            .MappingName = "MaterialTotal"
            .HeaderText = "Material"
            If Type = "All" Then
                .Width = 90
            Else
                .Width = 100
            End If
            .Alignment = HorizontalAlignment.Left
            .Format = "c"
            ts.GridColumnStyles.Add(col)
        End With

        col = New DataGridTextBoxColumn

        With col
            .MappingName = "InvoiceTotal"
            .HeaderText = "Total"
            If Type = "All" Then
                .Width = 90
            Else
                .Width = 100
            End If

            .Alignment = HorizontalAlignment.Left
            .Format = "c"
            ts.GridColumnStyles.Add(col)
        End With

        Me.dgInv.TableStyles.Add(ts)


    End Sub

    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click
        Try
            Cursor = Cursors.WaitCursor

            Me.btnRun.Enabled = False
            RefreshReport()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            Cursor = Cursors.Default

        End Try

    End Sub

    Private Sub cboCustomers_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCustomers.SelectionChangeCommitted

        Try
            Me.btnRun.Enabled = True
            Me.CustomerID = Convert.ToInt32(Me.cboCustomers.SelectedValue)
            Me.dgInv.DataSource = Nothing

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

    Private Sub rdoAllCustomers_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoAllCustomers.CheckedChanged, rdoSelectCust.CheckedChanged
        Try
            Me.dgInv.DataSource = Nothing
            Me.Text = "View Invoices"
            Me.btnRun.Enabled = True

            Dim rdo As RadioButton = CType(sender, RadioButton)

            Select Case rdo.Name
                Case Me.rdoAllCustomers.Name
                    Me.cboCustomers.Enabled = False
                    Me.CustomerID = 0

                Case Me.rdoSelectCust.Name
                    Me.cboCustomers.Enabled = True
                    Me.CustomerID = Convert.ToInt32(Me.cboCustomers.SelectedValue)
            End Select

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub LoadAllInvoices()

        Dim dv As New DataView

        If Me.dsInv Is Nothing Then
            Dim inv As New DelortBusObjects.Invoice(g_Settings)
            Me.dsInv = inv.GetAllInvoices(Me.Date1.ToShortDateString, Me.Date2.ToShortDateString)
        End If


        'klr
        SetUpDataGrid("All")

        Me.dgInv.DataSource = Me.dsInv.Tables("Invoices")

        TallyReportTotals(Me.dsInv.Tables(0))

    End Sub

    Private Sub dtpFrom_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpFrom.ValueChanged

        Try

            Me.Date1 = Convert.ToDateTime(Me.dtpFrom.Value)

            If Loading = False Then
                Me.dgInv.DataSource = Nothing
            End If

            If Me.dtpFrom.Value > Me.dtpTo.Value Then
                Me.ErrProv.SetError(Me.dtpFrom, "From Date Shouldn't be later than To Date")
                Me.btnRun.Enabled = False
            Else
                Me.btnRun.Enabled = True
                Me.ErrProv.SetError(Me.dtpFrom, "")
                Me.ErrProv.SetError(Me.dtpTo, "")

            End If

            Me.dtpTo.MinDate = dtpFrom.Value

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub dtpTo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpTo.ValueChanged

        Try
            Me.Date2 = Convert.ToDateTime(Me.dtpTo.Value)

            If Me.Loading = False Then
                Me.dgInv.DataSource = Nothing
            End If

            If Me.dtpFrom.Value > Me.dtpTo.Value Then
                Me.ErrProv.SetError(Me.dtpTo, "From Date Shouldn't be later than To Date")
                Me.btnRun.Enabled = False
            Else
                Me.ErrProv.SetError(Me.dtpFrom, "")
                Me.ErrProv.SetError(Me.dtpTo, "")
                Me.btnRun.Enabled = True
            End If

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub


    Private Sub cmnuView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmnuView.Click

        Try
            OpenInvoice()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try



    End Sub

    Private Sub TallyReportTotals(ByVal dt As DataTable)

        Dim MaterialTotal As Double
        Dim LaborTotal As Double
        Dim Total As Double

        For Each row As DataRow In dt.Rows
            With row
                If Not IsDBNull(.Item("MaterialTotal")) Then
                    MaterialTotal += Convert.ToDouble(.Item("MaterialTotal"))
                End If
                If Not IsDBNull(.Item("LaborTotal")) Then
                    LaborTotal += Convert.ToDouble(.Item("LaborTotal"))
                End If

                If Not IsDBNull(.Item("InvoiceTotal")) Then
                    Total += Convert.ToDouble(.Item("InvoiceTotal"))
                End If
            End With
        Next

        Me.lblLaborTotal.Text = LaborTotal.ToString("c")
        Me.lblMaterialTotal.Text = MaterialTotal.ToString("c")
        Me.lblTotal.Text = Total.ToString("c")
        Me.Text = dt.Rows.Count.ToString & " Invoice(s)"

    End Sub

    Private Sub OpenInvoice()

        Dim invNum As Integer


        If Me.dgInv.CurrentRowIndex > -1 Then
            invNum = Convert.ToInt32(dgInv(Me.dgInv.CurrentRowIndex, 0))
            Utilities.OpenInvoice(invNum, Me, Me.ParentForm)

            Me.RefreshReport()

        End If

    End Sub

    Private Sub btnOpenInvoice_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOpenInvoice.Click

        Try
            Me.OpenInvoice()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub dgInv_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgInv.DoubleClick
        Try
            OpenInvoice()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click

        Try
            DeleteInvoice()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
    Public Sub RefreshReport()

        Me.dsInv = Nothing

        If CustomerID > 0 Then

            LoadCustomerData()
        Else

            LoadAllInvoices()
        End If
    End Sub

    Private Sub DeleteInvoice()

        Dim invNum As Integer

        If Me.dgInv.VisibleRowCount > 0 Then
            invNum = Convert.ToInt32(dgInv(Me.dgInv.CurrentRowIndex, 0))

            If MessageBox.Show("Are You Sure You Want To Delete Invoice# " & invNum.ToString & "?", ProgramName, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.Yes Then

                Dim inv As New DelortBusObjects.Invoice(g_Settings)
                inv.Delete(invNum)

                Me.RefreshReport()

            End If
        End If
    End Sub

    Private Sub cmnuDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmnuDelete.Click

        Try
            DeleteInvoice()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub


    Private Sub dgInv_DataSourceChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgInv.DataSourceChanged

        Try
            If Me.dgInv.DataSource Is Nothing Then
                Me.btnOpenInvoice.Enabled = False
                Me.btnDelete.Enabled = False
                Me.btnPrint.Enabled = False
                Me.lblLaborTotal.Text = "$0.00"
                Me.lblMaterialTotal.Text = "$0.00"
                Me.lblTotal.Text = "$0.00"
                Exit Sub
            End If

            Dim Table As DataTable = CType(Me.dgInv.DataSource, DataTable)

            If Table.Rows.Count > 0 Then
                Me.btnOpenInvoice.Enabled = True
                Me.btnDelete.Enabled = True
                Me.btnPrint.Enabled = True
            Else
                Me.btnOpenInvoice.Enabled = False
                Me.btnDelete.Enabled = False
                Me.btnPrint.Enabled = False
            End If

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click

        Cursor = Cursors.WaitCursor

        Try
            Me.PrintInvoice()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            Cursor = Cursors.Default

        End Try


    End Sub

    Private Sub cmnuPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmnuPrint.Click

        Try
            Me.PrintInvoice()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
    Private Sub PrintInvoice()

        Dim invNum As Integer

        Try

            If Me.dgInv.VisibleRowCount > 0 Then
                invNum = Convert.ToInt32(dgInv(Me.dgInv.CurrentRowIndex, 0))
                Utilities.PrintInvoice(invNum)

            Else
                Exit Sub
            End If

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click

        If Me.dgInv.VisibleColumnCount > 0 Then
            Me.RefreshReport()
        End If

    End Sub

    Private Sub SetUpUI()

        Me.Loading = True
        Utilities.LoadCustCombo(Me.cboCustomers, 0)
        Me.dtpFrom.Value = DateSerial(Year(Now()), Month(Now()), 1)
        Me.Date1 = DateSerial(Year(Now()), Month(Now()), 1)
        Me.Date2 = Today

        Me.RefreshReport()


    End Sub

End Class
