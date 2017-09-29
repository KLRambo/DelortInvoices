
#Region " Options "

Option Explicit On
Option Strict On

#End Region

Imports DelortBusObjects
Imports System.Configuration.ConfigurationSettings
Imports Infragistics.Win.UltraWinGrid


Public Class frmCustomerReport


    Inherits frmBaseWindow

    Private WithEvents Customers As New Customer(g_settings)
    Private WithEvents CustReport As New CustomerReport(g_Settings)
    Friend WithEvents StatusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents Status As System.Windows.Forms.ToolStripStatusLabel
    Private Loading As Boolean = True

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Me.CustReport.SetMasterDataSource()

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
    Friend WithEvents grpCustomers As System.Windows.Forms.GroupBox
    Friend WithEvents rdoAllCustomers As System.Windows.Forms.RadioButton
    Friend WithEvents grpInv As System.Windows.Forms.GroupBox
    Friend WithEvents grpAction As System.Windows.Forms.GroupBox
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents grpNameSearch As System.Windows.Forms.GroupBox
    Friend WithEvents txtLastName As System.Windows.Forms.TextBox
    Friend WithEvents lblLastName As System.Windows.Forms.Label
    Friend WithEvents txtFirstName As System.Windows.Forms.TextBox
    Friend WithEvents lblFirstName As System.Windows.Forms.Label
    Friend WithEvents rdoFirstContains As System.Windows.Forms.RadioButton
    Friend WithEvents grpLastName As System.Windows.Forms.GroupBox
    Friend WithEvents rdoLastStarts As System.Windows.Forms.RadioButton
    Friend WithEvents rdoLastContains As System.Windows.Forms.RadioButton
    Friend WithEvents rdoSearchByName As System.Windows.Forms.RadioButton
    Friend WithEvents btnOpen As System.Windows.Forms.Button
    Friend WithEvents grpFirstName As System.Windows.Forms.GroupBox
    Friend WithEvents cmnuMain As System.Windows.Forms.ContextMenu
    Friend WithEvents cmnuOpen As System.Windows.Forms.MenuItem
    Friend WithEvents cmnuDelete As System.Windows.Forms.MenuItem
    Friend WithEvents rdoFirstStarts As System.Windows.Forms.RadioButton
    Friend WithEvents btnAddNew As System.Windows.Forms.Button
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents ugCustomers As Infragistics.Win.UltraWinGrid.UltraGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCustomerReport))
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.cmnuMain = New System.Windows.Forms.ContextMenu()
        Me.cmnuOpen = New System.Windows.Forms.MenuItem()
        Me.cmnuDelete = New System.Windows.Forms.MenuItem()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.StatusStrip = New System.Windows.Forms.StatusStrip()
        Me.Status = New System.Windows.Forms.ToolStripStatusLabel()
        Me.grpNameSearch = New System.Windows.Forms.GroupBox()
        Me.grpFirstName = New System.Windows.Forms.GroupBox()
        Me.rdoFirstStarts = New System.Windows.Forms.RadioButton()
        Me.rdoFirstContains = New System.Windows.Forms.RadioButton()
        Me.txtFirstName = New System.Windows.Forms.TextBox()
        Me.lblFirstName = New System.Windows.Forms.Label()
        Me.txtLastName = New System.Windows.Forms.TextBox()
        Me.lblLastName = New System.Windows.Forms.Label()
        Me.grpLastName = New System.Windows.Forms.GroupBox()
        Me.rdoLastContains = New System.Windows.Forms.RadioButton()
        Me.rdoLastStarts = New System.Windows.Forms.RadioButton()
        Me.grpAction = New System.Windows.Forms.GroupBox()
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.btnAddNew = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnOpen = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.grpInv = New System.Windows.Forms.GroupBox()
        Me.ugCustomers = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.grpCustomers = New System.Windows.Forms.GroupBox()
        Me.rdoSearchByName = New System.Windows.Forms.RadioButton()
        Me.rdoAllCustomers = New System.Windows.Forms.RadioButton()
        Me.StatusStrip.SuspendLayout()
        Me.grpNameSearch.SuspendLayout()
        Me.grpFirstName.SuspendLayout()
        Me.grpLastName.SuspendLayout()
        Me.grpAction.SuspendLayout()
        Me.grpInv.SuspendLayout()
        CType(Me.ugCustomers, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpCustomers.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmnuMain
        '
        Me.cmnuMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.cmnuOpen, Me.cmnuDelete})
        '
        'cmnuOpen
        '
        Me.cmnuOpen.Index = 0
        Me.cmnuOpen.Text = "Open Customer"
        '
        'cmnuDelete
        '
        Me.cmnuDelete.Index = 1
        Me.cmnuDelete.Text = "Delete Customer"
        '
        'StatusStrip
        '
        Me.StatusStrip.BackColor = System.Drawing.Color.Transparent
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Status})
        Me.StatusStrip.Location = New System.Drawing.Point(0, 544)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.Size = New System.Drawing.Size(502, 22)
        Me.StatusStrip.TabIndex = 6
        Me.StatusStrip.Text = "StatusStrip1"
        '
        'Status
        '
        Me.Status.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Status.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Status.Name = "Status"
        Me.Status.Size = New System.Drawing.Size(0, 17)
        '
        'grpNameSearch
        '
        Me.grpNameSearch.BackColor = System.Drawing.Color.Transparent
        Me.grpNameSearch.Controls.Add(Me.grpFirstName)
        Me.grpNameSearch.Controls.Add(Me.txtFirstName)
        Me.grpNameSearch.Controls.Add(Me.lblFirstName)
        Me.grpNameSearch.Controls.Add(Me.txtLastName)
        Me.grpNameSearch.Controls.Add(Me.lblLastName)
        Me.grpNameSearch.Controls.Add(Me.grpLastName)
        Me.grpNameSearch.Location = New System.Drawing.Point(172, 8)
        Me.grpNameSearch.Name = "grpNameSearch"
        Me.grpNameSearch.Size = New System.Drawing.Size(324, 96)
        Me.grpNameSearch.TabIndex = 5
        Me.grpNameSearch.TabStop = False
        '
        'grpFirstName
        '
        Me.grpFirstName.Controls.Add(Me.rdoFirstStarts)
        Me.grpFirstName.Controls.Add(Me.rdoFirstContains)
        Me.grpFirstName.Location = New System.Drawing.Point(210, 52)
        Me.grpFirstName.Name = "grpFirstName"
        Me.grpFirstName.Size = New System.Drawing.Size(110, 40)
        Me.grpFirstName.TabIndex = 17
        Me.grpFirstName.TabStop = False
        '
        'rdoFirstStarts
        '
        Me.rdoFirstStarts.Checked = True
        Me.rdoFirstStarts.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoFirstStarts.Location = New System.Drawing.Point(8, 7)
        Me.rdoFirstStarts.Name = "rdoFirstStarts"
        Me.rdoFirstStarts.Size = New System.Drawing.Size(88, 16)
        Me.rdoFirstStarts.TabIndex = 22
        Me.rdoFirstStarts.TabStop = True
        Me.rdoFirstStarts.Text = "Starts With"
        '
        'rdoFirstContains
        '
        Me.rdoFirstContains.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoFirstContains.Location = New System.Drawing.Point(8, 23)
        Me.rdoFirstContains.Name = "rdoFirstContains"
        Me.rdoFirstContains.Size = New System.Drawing.Size(75, 14)
        Me.rdoFirstContains.TabIndex = 21
        Me.rdoFirstContains.Text = "Contains"
        '
        'txtFirstName
        '
        Me.txtFirstName.Location = New System.Drawing.Point(89, 61)
        Me.txtFirstName.Name = "txtFirstName"
        Me.txtFirstName.Size = New System.Drawing.Size(111, 21)
        Me.txtFirstName.TabIndex = 10
        '
        'lblFirstName
        '
        Me.lblFirstName.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFirstName.Location = New System.Drawing.Point(2, 64)
        Me.lblFirstName.Name = "lblFirstName"
        Me.lblFirstName.Size = New System.Drawing.Size(80, 16)
        Me.lblFirstName.TabIndex = 9
        Me.lblFirstName.Text = "First Name:"
        '
        'txtLastName
        '
        Me.txtLastName.Location = New System.Drawing.Point(89, 16)
        Me.txtLastName.Name = "txtLastName"
        Me.txtLastName.Size = New System.Drawing.Size(111, 21)
        Me.txtLastName.TabIndex = 12
        '
        'lblLastName
        '
        Me.lblLastName.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLastName.Location = New System.Drawing.Point(2, 16)
        Me.lblLastName.Name = "lblLastName"
        Me.lblLastName.Size = New System.Drawing.Size(88, 16)
        Me.lblLastName.TabIndex = 11
        Me.lblLastName.Text = "Last Name:"
        '
        'grpLastName
        '
        Me.grpLastName.Controls.Add(Me.rdoLastContains)
        Me.grpLastName.Controls.Add(Me.rdoLastStarts)
        Me.grpLastName.Location = New System.Drawing.Point(210, 7)
        Me.grpLastName.Name = "grpLastName"
        Me.grpLastName.Size = New System.Drawing.Size(110, 40)
        Me.grpLastName.TabIndex = 18
        Me.grpLastName.TabStop = False
        '
        'rdoLastContains
        '
        Me.rdoLastContains.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoLastContains.Location = New System.Drawing.Point(9, 22)
        Me.rdoLastContains.Name = "rdoLastContains"
        Me.rdoLastContains.Size = New System.Drawing.Size(87, 16)
        Me.rdoLastContains.TabIndex = 21
        Me.rdoLastContains.Text = "Contains"
        '
        'rdoLastStarts
        '
        Me.rdoLastStarts.Checked = True
        Me.rdoLastStarts.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoLastStarts.Location = New System.Drawing.Point(9, 9)
        Me.rdoLastStarts.Name = "rdoLastStarts"
        Me.rdoLastStarts.Size = New System.Drawing.Size(89, 14)
        Me.rdoLastStarts.TabIndex = 17
        Me.rdoLastStarts.TabStop = True
        Me.rdoLastStarts.Text = "Starts With"
        '
        'grpAction
        '
        Me.grpAction.BackColor = System.Drawing.Color.Transparent
        Me.grpAction.Controls.Add(Me.btnRefresh)
        Me.grpAction.Controls.Add(Me.btnAddNew)
        Me.grpAction.Controls.Add(Me.btnDelete)
        Me.grpAction.Controls.Add(Me.btnOpen)
        Me.grpAction.Controls.Add(Me.btnExit)
        Me.grpAction.Location = New System.Drawing.Point(12, 104)
        Me.grpAction.Name = "grpAction"
        Me.grpAction.Size = New System.Drawing.Size(484, 48)
        Me.grpAction.TabIndex = 4
        Me.grpAction.TabStop = False
        '
        'btnRefresh
        '
        Me.btnRefresh.BackColor = System.Drawing.Color.AliceBlue
        Me.btnRefresh.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRefresh.Image = CType(resources.GetObject("btnRefresh.Image"), System.Drawing.Image)
        Me.btnRefresh.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnRefresh.Location = New System.Drawing.Point(374, 16)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(80, 22)
        Me.btnRefresh.TabIndex = 9
        Me.btnRefresh.Text = "&Refresh"
        Me.btnRefresh.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.btnRefresh, "Refresh Data")
        Me.btnRefresh.UseVisualStyleBackColor = False
        '
        'btnAddNew
        '
        Me.btnAddNew.BackColor = System.Drawing.Color.AliceBlue
        Me.btnAddNew.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddNew.Image = CType(resources.GetObject("btnAddNew.Image"), System.Drawing.Image)
        Me.btnAddNew.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnAddNew.Location = New System.Drawing.Point(288, 16)
        Me.btnAddNew.Name = "btnAddNew"
        Me.btnAddNew.Size = New System.Drawing.Size(80, 22)
        Me.btnAddNew.TabIndex = 8
        Me.btnAddNew.Text = "   &New"
        Me.ToolTip1.SetToolTip(Me.btnAddNew, "Create a New Invoice")
        Me.btnAddNew.UseVisualStyleBackColor = False
        '
        'btnDelete
        '
        Me.btnDelete.BackColor = System.Drawing.Color.AliceBlue
        Me.btnDelete.Enabled = False
        Me.btnDelete.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.Image = CType(resources.GetObject("btnDelete.Image"), System.Drawing.Image)
        Me.btnDelete.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDelete.Location = New System.Drawing.Point(202, 16)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 22)
        Me.btnDelete.TabIndex = 7
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.btnDelete, "Delete Selected Invoice")
        Me.btnDelete.UseVisualStyleBackColor = False
        '
        'btnOpen
        '
        Me.btnOpen.BackColor = System.Drawing.Color.AliceBlue
        Me.btnOpen.Enabled = False
        Me.btnOpen.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOpen.Image = CType(resources.GetObject("btnOpen.Image"), System.Drawing.Image)
        Me.btnOpen.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnOpen.Location = New System.Drawing.Point(116, 16)
        Me.btnOpen.Name = "btnOpen"
        Me.btnOpen.Size = New System.Drawing.Size(80, 22)
        Me.btnOpen.TabIndex = 6
        Me.btnOpen.Text = "&Open"
        Me.btnOpen.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.btnOpen, "Open Selected Invoice")
        Me.btnOpen.UseVisualStyleBackColor = False
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.AliceBlue
        Me.btnExit.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.Image = CType(resources.GetObject("btnExit.Image"), System.Drawing.Image)
        Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnExit.Location = New System.Drawing.Point(30, 16)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 22)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "  E&xit"
        Me.ToolTip1.SetToolTip(Me.btnExit, "Exit Form")
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'grpInv
        '
        Me.grpInv.BackColor = System.Drawing.Color.Transparent
        Me.grpInv.Controls.Add(Me.ugCustomers)
        Me.grpInv.Location = New System.Drawing.Point(12, 153)
        Me.grpInv.Name = "grpInv"
        Me.grpInv.Size = New System.Drawing.Size(484, 383)
        Me.grpInv.TabIndex = 3
        Me.grpInv.TabStop = False
        '
        'ugCustomers
        '
        Me.ugCustomers.ContextMenu = Me.cmnuMain
        Appearance1.TextHAlignAsString = "Left"
        Me.ugCustomers.DisplayLayout.Appearance = Appearance1
        Me.ugCustomers.DisplayLayout.MaxColScrollRegions = 1
        Me.ugCustomers.DisplayLayout.MaxRowScrollRegions = 1
        Me.ugCustomers.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugCustomers.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.ugCustomers.DisplayLayout.Override.AllowGroupBy = Infragistics.Win.DefaultableBoolean.[False]
        Me.ugCustomers.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.ugCustomers.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.ugCustomers.DisplayLayout.Override.CellPadding = 0
        Me.ugCustomers.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugCustomers.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugCustomers.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.ugCustomers.Location = New System.Drawing.Point(8, 18)
        Me.ugCustomers.Name = "ugCustomers"
        Me.ugCustomers.Size = New System.Drawing.Size(464, 358)
        Me.ugCustomers.TabIndex = 0
        '
        'grpCustomers
        '
        Me.grpCustomers.BackColor = System.Drawing.Color.Transparent
        Me.grpCustomers.Controls.Add(Me.rdoSearchByName)
        Me.grpCustomers.Controls.Add(Me.rdoAllCustomers)
        Me.grpCustomers.Location = New System.Drawing.Point(12, 8)
        Me.grpCustomers.Name = "grpCustomers"
        Me.grpCustomers.Size = New System.Drawing.Size(152, 96)
        Me.grpCustomers.TabIndex = 2
        Me.grpCustomers.TabStop = False
        Me.grpCustomers.Text = "Customers"
        '
        'rdoSearchByName
        '
        Me.rdoSearchByName.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoSearchByName.Location = New System.Drawing.Point(16, 60)
        Me.rdoSearchByName.Name = "rdoSearchByName"
        Me.rdoSearchByName.Size = New System.Drawing.Size(128, 23)
        Me.rdoSearchByName.TabIndex = 3
        Me.rdoSearchByName.Text = "Search By Name"
        '
        'rdoAllCustomers
        '
        Me.rdoAllCustomers.Checked = True
        Me.rdoAllCustomers.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoAllCustomers.Location = New System.Drawing.Point(15, 16)
        Me.rdoAllCustomers.Name = "rdoAllCustomers"
        Me.rdoAllCustomers.Size = New System.Drawing.Size(56, 16)
        Me.rdoAllCustomers.TabIndex = 1
        Me.rdoAllCustomers.TabStop = True
        Me.rdoAllCustomers.Text = "All"
        '
        'frmCustomerReport
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(502, 566)
        Me.Controls.Add(Me.StatusStrip)
        Me.Controls.Add(Me.grpNameSearch)
        Me.Controls.Add(Me.grpAction)
        Me.Controls.Add(Me.grpInv)
        Me.Controls.Add(Me.grpCustomers)
        Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmCustomerReport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Search For Customers"
        Me.StatusStrip.ResumeLayout(False)
        Me.StatusStrip.PerformLayout()
        Me.grpNameSearch.ResumeLayout(False)
        Me.grpNameSearch.PerformLayout()
        Me.grpFirstName.ResumeLayout(False)
        Me.grpLastName.ResumeLayout(False)
        Me.grpAction.ResumeLayout(False)
        Me.grpInv.ResumeLayout(False)
        CType(Me.ugCustomers, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpCustomers.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try
            Me.ugCustomers.DataSource = Me.CustReport.DataView

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub


    Private Sub rdoAllCustomers_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdoAllCustomers.CheckedChanged

        Try

            If Me.rdoAllCustomers.Checked = True Then
                Me.CustReport.SearchingBy = CustomerReport.SearchBy.All
                Me.txtFirstName.Text = ""
                Me.txtLastName.Text = ""
            Else
                Me.CustReport.SearchingBy = CustomerReport.SearchBy.Name
            End If

            Me.ugCustomers.DataSource = Me.CustReport.DataView

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub txtFirstName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFirstName.TextChanged

        Me.CustReport.FirstName = Me.txtFirstName.Text.Trim


    End Sub

    Private Sub txtLastName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLastName.TextChanged, txtFirstName.TextChanged

        Me.CustReport.LastName = Me.txtLastName.Text.Trim
        Me.CustReport.FirstName = Me.txtFirstName.Text.Trim
        Me.ugCustomers.DataSource = Me.CustReport.DataView

    End Sub

    Private Sub OpenCustomer()

        If Me.ugCustomers.Rows.VisibleRowCount > 0 Then
            Dim CustNo As Integer = Convert.ToInt32(Me.ugCustomers.ActiveRow.Cells(0).Value)
            Dim CusObj As Customer = Me.Customers.GetCustomer(CustNo)
            Utilities.ViewAddEditCustomer(Me.ParentForm, CusObj, Me)
        End If

    End Sub
    Private Sub DeleteCustomer()

        If Me.ugCustomers.Rows.VisibleRowCount > 0 Then
            Dim FirstName As String = Me.ugCustomers.ActiveRow.Cells(2).Value.ToString
            Dim Lastname As String = Me.ugCustomers.ActiveRow.Cells(1).Value.ToString

            If MessageBox.Show("Are you sure you want to delete " & FirstName & " " & Lastname & "?", ProgramName, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.Yes Then
                Me.Customers.DBMode = Customer.CrudMode.Delete
                Dim CustNo As Integer = Convert.ToInt32(Me.ugCustomers.ActiveRow.Cells(0).Value)
                Me.Customers.DBMode = Customer.CrudMode.Delete
                Me.Customers.CustID = CustNo
                Me.Customers.InsertUpdateDelete()
            End If

        End If

        Me.ReloadData()



    End Sub

    Private Sub btnOpen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOpen.Click

        Try

            Me.OpenCustomer()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub cmnuOpen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmnuOpen.Click

        Try

            Me.OpenCustomer()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub cmnuDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmnuDelete.Click
        Try

            Me.DeleteCustomer()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click

        Try

            Me.DeleteCustomer()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub
    Public Sub ReloadData()

        Me.CustReport.SetMasterDataSource()
        Me.ugCustomers.DataSource = Me.CustReport.DataView

    End Sub

    Private Sub rdoFirstStarts_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdoFirstStarts.CheckedChanged

        Me.CustReport.FirstNameStartsWith = Me.rdoFirstStarts.Checked

    End Sub

    Private Sub rdoLastStarts_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdoLastStarts.CheckedChanged
        Me.CustReport.LastNameStartsWith = Me.rdoLastStarts.Checked

    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub btnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddNew.Click

        Try
            Utilities.ViewAddEditCustomer(Me.MdiParent, Nothing, Me)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click

        Try
            Me.ReloadData()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub CustReport_SearchingByAll(ByVal Value As Boolean) Handles CustReport.SearchingByAll

        Me.txtFirstName.Enabled = Not (Value)
        Me.txtLastName.Enabled = Not (Value)

    End Sub


    Private Sub CustReport_ValidRecords(ByVal Recs As Boolean) Handles CustReport.ValidRecords

        Me.btnDelete.Enabled = Recs
        Me.btnOpen.Enabled = Recs

    End Sub

    Private Sub ugCustomers_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugCustomers.InitializeLayout

        With e.Layout.Bands(0)
            .Columns(0).Hidden = True
            .Columns(1).Header.Caption = "    Last Name"
            .Columns(1).Width = 100
            .Columns(2).Header.Caption = "First Name"
            .Columns(1).Width = 100
            .Columns(3).Header.Caption = "Address1"
            .Columns(3).Width = 200
            .Columns(4).Header.Caption = "Address2"
            .Columns(4).Width = 100
            .Columns(5).Header.Caption = "City"
            .Columns(5).Width = 150
            .Columns(6).Header.Caption = "State"
            .Columns(6).Width = 50
            .Columns(7).Header.Caption = "Zip"
            .Columns(7).Width = 100
            .Columns(8).Header.Caption = "Phone"
            .Columns(8).Width = 125
            .Columns(9).Header.Caption = "Cell"
            .Columns(9).Width = 125
            .Columns(10).Header.Caption = "Notes"
            .Columns(10).Width = 400
            .Columns(11).Hidden = True
            .Columns(12).Hidden = True
            .Columns(13).Hidden = True
            .Columns(14).Hidden = True
        End With

        Me.Status.Text = "Total number of customers: " & Me.ugCustomers.Rows.Count.ToString


    End Sub

    Private Sub ugCustomers_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles ugCustomers.DoubleClickRow

        Try
            Me.OpenCustomer()

        Catch ex As NullReferenceException
            'do nothing, user just clicking on the wrong spot
        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub frmCustomerReport_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        Customers = Nothing
        CustReport = Nothing

    End Sub

    Private Sub ugCustomers_MouseEnterElement(ByVal sender As Object, ByVal e As Infragistics.Win.UIElementEventArgs) Handles ugCustomers.MouseEnterElement

        Dim Cell As Infragistics.Win.UltraWinGrid.UltraGridCell = DirectCast(e.Element.GetContext(GetType(UltraGridCell)), UltraGridCell)

        Dim strText As String

        If Not Cell Is Nothing Then
            ' tool tip over cell
            If Cell.Row.Cells("Phone").Value.ToString.Length <> 0 Then
                strText = Cell.Row.Cells("FirstName").Value.ToString & " " & Cell.Row.Cells("LastName").Value.ToString & " Phone: " & Cell.Row.Cells(8).Value.ToString
                ToolTip1.SetToolTip(Me.ugCustomers, strText)
            End If
        Else
            ToolTip1.SetToolTip(Me.ugCustomers, "")

        End If

    End Sub


End Class
