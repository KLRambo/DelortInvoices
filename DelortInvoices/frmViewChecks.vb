
#Region " Imports "

Imports Infragistics.Win.UltraWinGrid

#End Region

Public Class frmViewChecks
    Inherits System.Windows.Forms.Form

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
    Friend Shadows WithEvents ContextMenu As System.Windows.Forms.ContextMenu
    Friend WithEvents cmOpen As System.Windows.Forms.MenuItem
    Friend WithEvents cmDelete As System.Windows.Forms.MenuItem
    Friend WithEvents grpDates As System.Windows.Forms.GroupBox
    Friend WithEvents grpButtons As System.Windows.Forms.GroupBox
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblFr As System.Windows.Forms.Label
    Friend WithEvents lblTo As System.Windows.Forms.Label
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents HelpProvider As System.Windows.Forms.HelpProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmViewChecks))
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.grpMain = New System.Windows.Forms.GroupBox
        Me.grdInvoices = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.ContextMenu = New System.Windows.Forms.ContextMenu
        Me.cmOpen = New System.Windows.Forms.MenuItem
        Me.cmDelete = New System.Windows.Forms.MenuItem
        Me.grpDates = New System.Windows.Forms.GroupBox
        Me.lblTo = New System.Windows.Forms.Label
        Me.lblFr = New System.Windows.Forms.Label
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.grpButtons = New System.Windows.Forms.GroupBox
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnAdd = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnRefresh = New System.Windows.Forms.Button
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.HelpProvider = New System.Windows.Forms.HelpProvider
        Me.grpMain.SuspendLayout()
        CType(Me.grdInvoices, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpDates.SuspendLayout()
        Me.grpButtons.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpMain
        '
        Me.grpMain.BackColor = System.Drawing.Color.Transparent
        Me.grpMain.Controls.Add(Me.grdInvoices)
        Me.grpMain.Location = New System.Drawing.Point(8, 64)
        Me.grpMain.Name = "grpMain"
        Me.grpMain.Size = New System.Drawing.Size(848, 504)
        Me.grpMain.TabIndex = 1
        Me.grpMain.TabStop = False
        '
        'grdInvoices
        '
        Me.grdInvoices.ContextMenu = Me.ContextMenu
        Appearance1.ForeColor = System.Drawing.SystemColors.ControlText
        Appearance1.ImageBackground = CType(resources.GetObject("Appearance1.ImageBackground"), System.Drawing.Image)
        Appearance1.ImageBackgroundStyle = Infragistics.Win.ImageBackgroundStyle.Stretched
        Appearance1.ImageHAlign = Infragistics.Win.HAlign.Center
        Appearance1.ImageVAlign = Infragistics.Win.VAlign.Middle
        Appearance1.TextHAlign = Infragistics.Win.HAlign.Left
        Me.grdInvoices.DisplayLayout.Appearance = Appearance1
        Appearance2.ForeColor = System.Drawing.Color.Blue
        Appearance2.TextHAlign = Infragistics.Win.HAlign.Center
        Me.grdInvoices.DisplayLayout.CaptionAppearance = Appearance2
        Me.grdInvoices.DisplayLayout.GroupByBox.Hidden = True
        Me.grdInvoices.DisplayLayout.MaxColScrollRegions = 1
        Me.grdInvoices.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdInvoices.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdInvoices.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        Me.grdInvoices.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.grdInvoices.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdInvoices.DisplayLayout.Override.CellPadding = 0
        Me.grdInvoices.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Appearance3.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdInvoices.DisplayLayout.Override.SelectedRowAppearance = Appearance3
        Me.grdInvoices.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdInvoices.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdInvoices.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdInvoices.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grdInvoices.Location = New System.Drawing.Point(9, 16)
        Me.grdInvoices.Name = "grdInvoices"
        Me.grdInvoices.Size = New System.Drawing.Size(831, 480)
        Me.grdInvoices.TabIndex = 2
        Me.grdInvoices.Text = "Check Ledger"
        Me.grdInvoices.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChange
        '
        'ContextMenu
        '
        Me.ContextMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.cmOpen, Me.cmDelete})
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
        Me.grpDates.Controls.Add(Me.lblFr)
        Me.grpDates.Controls.Add(Me.dtpTo)
        Me.grpDates.Controls.Add(Me.dtpFrom)
        Me.grpDates.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpDates.Location = New System.Drawing.Point(8, 8)
        Me.grpDates.Name = "grpDates"
        Me.grpDates.Size = New System.Drawing.Size(352, 48)
        Me.grpDates.TabIndex = 2
        Me.grpDates.TabStop = False
        Me.grpDates.Text = "Select Dates"
        '
        'lblTo
        '
        Me.lblTo.Location = New System.Drawing.Point(184, 16)
        Me.lblTo.Name = "lblTo"
        Me.lblTo.Size = New System.Drawing.Size(30, 23)
        Me.lblTo.TabIndex = 3
        Me.lblTo.Text = "To:"
        '
        'lblFr
        '
        Me.lblFr.Location = New System.Drawing.Point(12, 16)
        Me.lblFr.Name = "lblFr"
        Me.lblFr.Size = New System.Drawing.Size(40, 23)
        Me.lblFr.TabIndex = 2
        Me.lblFr.Text = "From:"
        '
        'dtpTo
        '
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpTo.Location = New System.Drawing.Point(224, 16)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(114, 21)
        Me.dtpTo.TabIndex = 1
        '
        'dtpFrom
        '
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpFrom.Location = New System.Drawing.Point(62, 16)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(112, 21)
        Me.dtpFrom.TabIndex = 0
        '
        'grpButtons
        '
        Me.grpButtons.BackColor = System.Drawing.Color.Transparent
        Me.grpButtons.Controls.Add(Me.btnExit)
        Me.grpButtons.Controls.Add(Me.btnAdd)
        Me.grpButtons.Controls.Add(Me.btnEdit)
        Me.grpButtons.Controls.Add(Me.btnRefresh)
        Me.grpButtons.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpButtons.Location = New System.Drawing.Point(416, 8)
        Me.grpButtons.Name = "grpButtons"
        Me.grpButtons.Size = New System.Drawing.Size(440, 48)
        Me.grpButtons.TabIndex = 3
        Me.grpButtons.TabStop = False
        '
        'btnExit
        '
        Me.btnExit.Image = CType(resources.GetObject("btnExit.Image"), System.Drawing.Image)
        Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnExit.Location = New System.Drawing.Point(336, 16)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(88, 23)
        Me.btnExit.TabIndex = 3
        Me.btnExit.Text = "Exit"
        Me.btnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnAdd
        '
        Me.btnAdd.Image = CType(resources.GetObject("btnAdd.Image"), System.Drawing.Image)
        Me.btnAdd.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnAdd.Location = New System.Drawing.Point(224, 16)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(88, 23)
        Me.btnAdd.TabIndex = 2
        Me.btnAdd.Text = "Add New"
        Me.btnAdd.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnEdit
        '
        Me.btnEdit.Image = CType(resources.GetObject("btnEdit.Image"), System.Drawing.Image)
        Me.btnEdit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnEdit.Location = New System.Drawing.Point(120, 16)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(88, 23)
        Me.btnEdit.TabIndex = 1
        Me.btnEdit.Text = "Edit Mode"
        Me.btnEdit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.btnEdit, "Opens Selected row; Customer or Check")
        '
        'btnRefresh
        '
        Me.btnRefresh.Image = CType(resources.GetObject("btnRefresh.Image"), System.Drawing.Image)
        Me.btnRefresh.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnRefresh.Location = New System.Drawing.Point(16, 16)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(88, 23)
        Me.btnRefresh.TabIndex = 0
        Me.btnRefresh.Text = "Refresh"
        Me.btnRefresh.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ImageList1
        '
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'HelpProvider
        '
        Me.HelpProvider.HelpNamespace = "InvoicesHelp.chm"
        '
        'frmViewChecks
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(862, 572)
        Me.Controls.Add(Me.grpButtons)
        Me.Controls.Add(Me.grpDates)
        Me.Controls.Add(Me.grpMain)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.HelpProvider.SetHelpKeyword(Me, "CheckRegistry")
        Me.HelpProvider.SetHelpNavigator(Me, System.Windows.Forms.HelpNavigator.KeywordIndex)
        Me.MaximizeBox = False
        Me.Name = "frmViewChecks"
        Me.HelpProvider.SetShowHelp(Me, True)
        Me.Text = "Check Registry"
        Me.grpMain.ResumeLayout(False)
        CType(Me.grdInvoices, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpDates.ResumeLayout(False)
        Me.grpButtons.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Private Variables "

    Private WithEvents _Check As New DelortBusObjects.Check(g_Settings)
    Private Type As New Infragistics.Win.ValueList
    Private _EditMode As Boolean = False
    Private _isNew As Boolean = False
    Private isLoading As Boolean = True

#End Region

#Region " Properties "

    Public Property EditMode() As Boolean
        Get
            Return _EditMode
        End Get

        Set(ByVal value As Boolean)
            _EditMode = value


            Me.btnEdit.Enabled = Not (value) AndAlso Me.grdInvoices.Rows.Count > 0

            Me.btnAdd.Enabled = Not (value)

            If value = False Then
                Me.Text = "Check Registry " & Space(75) & "** Read-only Mode **"
                EditReadOnlyMode()
            Else
                Me.Text = "Check Registry " & Space(75) & "** Edit Mode **"
            End If

        End Set

    End Property

#End Region

#Region " Private Methods "

    Private Sub SetUpValueList()

        With Me.Type
            .Key = "Type"
            .ValueListItems.Add("Credit")
            .ValueListItems.Add("Debit")
        End With

    End Sub

    Private Sub ClickGridButton(ByVal sender As UltraGrid)

        With sender
            If .ActiveCell.Column.Key = "Delete" Then
                If MessageBox.Show("Are you sure you want to delete check # " & _
                    CStr(.ActiveRow.Cells("CodeNumber").Value) & "?", ProgramName, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    Me.DeleteRow(CStr(.ActiveRow.Cells(0).Value))
                End If
            ElseIf .ActiveCell.Column.Key = "Save" Then
                Me.SaveRow(.ActiveRow)
                Me.btnAdd.Enabled = True
            End If
        End With

    End Sub

    Private Sub GridTextChange(ByVal sender As UltraGrid)

        With sender.ActiveCell
            .IgnoreRowColActivation = True
            .Appearance.ForeColor = Drawing.Color.Red
        End With

        With sender.ActiveRow
            .Cells("CodeNumber").Appearance.ForeColor = Drawing.Color.Red
            .Cells("Save").IgnoreRowColActivation = True
            .Cells("Save").Activation = Activation.AllowEdit
        End With

    End Sub

    Private Sub AddRowToGrid()

        Me._Check.CrudMode = DelortBusObjects.Check.DBMode.Insert
        Me._isNew = True
        Me.EditMode = True

        Me.grdInvoices.Rows.Band.AddNew()

        With Me.grdInvoices.DisplayLayout.Bands(0)
            'Allow users to add
            .Override.AllowAddNew = AllowAddNew.Yes
            For Each col As UltraGridColumn In .Columns
                col.CellActivation = Activation.AllowEdit
            Next
            .Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True

        End With

        For Each row As UltraGridRow In Me.grdInvoices.DisplayLayout.Rows
            If row.IsAddRow Then
                row.Activation = Activation.AllowEdit
                For Each cell As UltraGridCell In row.Cells
                    If cell.Column.Key = "Delete" Then
                        cell.Activation = Activation.Disabled
                    End If
                Next
            Else
                row.Activation = Activation.Disabled

            End If
        Next

    End Sub

    Private Sub InitializeGridLayout(ByVal e As InitializeLayoutEventArgs)

        With e.Layout.Bands(0)
            .Override.AllowAddNew = AllowAddNew.Yes
            .Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False

            .Columns.Add("Save")
            With .Columns("Save")
                .Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button
                .Width = 50
                .Header.VisiblePosition = 0
                .ButtonDisplayStyle = ButtonDisplayStyle.Always
                .Header.Caption = "   Save"
                .CellButtonAppearance.Image = Me.ImageList1.Images(1)
                .CellButtonAppearance.ImageHAlign = Infragistics.Win.HAlign.Center
            End With

            .Columns.Add("Delete")
            With .Columns("Delete")
                .Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button
                .Width = 65
                .CellActivation = Activation.AllowEdit
                .Header.VisiblePosition = 1
                .ButtonDisplayStyle = ButtonDisplayStyle.Always
                .Header.Caption = "  Delete"
                .CellButtonAppearance.Image = Me.ImageList1.Images(0)
                .CellButtonAppearance.ImageHAlign = Infragistics.Win.HAlign.Center
            End With
            For Each col As UltraGridColumn In .Columns
                col.SortIndicator = SortIndicator.Ascending
                col.CellActivation = Activation.ActivateOnly
            Next

            'turn filtering off on selected columns
            .Columns("Delete").AllowRowFiltering = Infragistics.Win.DefaultableBoolean.False
            'turn sorting off for delete column
            .Columns("Delete").SortIndicator = SortIndicator.Disabled
            'turn filtering off on selected columns
            .Columns("Save").AllowRowFiltering = Infragistics.Win.DefaultableBoolean.False
            'turn sorting off for delete column
            .Columns("Save").SortIndicator = SortIndicator.Disabled

            'set numeric/currency columns
            .Columns("Amount").Style = ColumnStyle.Double
            .Columns("Balance").Style = ColumnStyle.Double
            .Columns("CodeNumber").Style = ColumnStyle.Integer

            .Columns("TransDate").Style = ColumnStyle.Date

            'turn filtering on for some rows
            .Columns("Type").AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True
            .Columns("TransDate").AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True

            .Columns(0).Hidden = True
            .Columns("CodeNumber").Header.Caption = "Number"
            .Columns("CodeNumber").Width = 80
            .Columns("Desc").Header.Caption = "Transaction Description"
            .Columns("Desc").Width = 250

            .Columns("TransDate").Header.Caption = "Date"
            .Columns("TransDate").Width = 100

            'assign value lists to columns
            .Columns("Type").Style = ColumnStyle.DropDownList
            .Columns("Type").ButtonDisplayStyle = ButtonDisplayStyle.Always
            .Columns("Type").Width = 80

            .Columns("Amount").Format = ("c")
            .Columns("Balance").Format = ("c")

            If (Not e.Layout.ValueLists.Exists("Type")) Then
                e.Layout.ValueLists.Add(Type)
            End If

            .Columns("Type").ValueList = e.Layout.ValueLists("Type")

        End With

    End Sub

    Private Sub SaveRow(ByVal row As UltraGridRow)
        Try

            With Me._Check
                If Not IsDBNull(row.Cells("Amount").Value) AndAlso IsNumeric(row.Cells("Amount").Value) Then
                    .Amount = CDbl(row.Cells("Amount").Value)
                Else
                    .Amount = Nothing
                End If
                If Not IsDBNull(row.Cells("Balance").Value) AndAlso IsNumeric(row.Cells("Balance").Value) Then
                    .Balance = CDbl(row.Cells("Balance").Value)
                Else
                    .Balance = Nothing
                End If
                If Not IsDBNull(row.Cells("Desc").Value) Then
                    .Description = CStr(row.Cells("Desc").Value)
                Else
                    .Description = ""
                End If
                If Not IsDBNull(row.Cells("CodeNumber").Value) Then
                    .Number = CStr(row.Cells("CodeNumber").Value)
                Else
                    .Number = ""
                End If
                If Not IsDBNull(row.Cells("TransDate").Value) Then
                    .TransDate = CDate(row.Cells("TransDate").Value)
                Else
                    .TransDate = Nothing
                End If
                If Not IsDBNull(row.Cells("Type").Value) Then
                    .Type = CStr(row.Cells("Type").Text)
                Else
                    .Type = ""
                End If
                If .CrudMode = DelortBusObjects.Check.DBMode.Update Then
                    .LedgerID = CInt(CStr(row.Cells("LedgerID").Value))
                End If
                .InsertUpdateDelete()
            End With

            Me.RefreshDisplay()

        Catch ex As Exception
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub SetToEditMode()

        Me._Check.CrudMode = DelortBusObjects.Check.DBMode.Update
        'Me.btnEdit.Enabled = False

        With Me.grdInvoices.DisplayLayout.Bands(0)
            For Each col As UltraGridColumn In .Columns
                If col.Key <> "Save" Then
                    col.CellActivation = Activation.AllowEdit
                    col.CellAppearance.ForeColor = Color.Blue
                End If
            Next
            'Allow users to edit
            .Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True
        End With
    End Sub

    Private Sub DeleteRow(ByVal LedgerID As String)

        Try

            With Me._Check
                .LedgerID = CInt(LedgerID)
                .CrudMode = DelortBusObjects.Check.DBMode.Delete
                .InsertUpdateDelete()
            End With

            Me.RefreshDisplay()

        Catch ex As Exception
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub RefreshDisplay()

        Me.grdInvoices.DataSource = Nothing
        Me.grdInvoices.DataSource = Me._Check.GetChecks(CStr(Me.dtpFrom.Value), CStr(Me.dtpTo.Value)).Tables(0)

        Me.btnEdit.Enabled = Me.grdInvoices.Rows.Count > 0

        Me.EditMode = False
        Me._isNew = False
    End Sub

    Private Sub SetUpDefaults()

        Me.dtpFrom.Value = New Date(Now.Year, 1, 1)
        Me.dtpTo.Value = New Date(Now.Year, 12, 31)
        Me.RefreshDisplay()

    End Sub

    Private Sub EditReadOnlyMode()

        With Me.grdInvoices.DisplayLayout.Bands(0)
            .Columns("Save").CellActivation = Activation.Disabled
            .Columns("Delete").CellActivation = Activation.Disabled
        End With

    End Sub

#End Region

#Region " Control Event handlers "

    Private Sub frmViewChecks_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            SetUpValueList()
            Me.SetUpDefaults()
            Me.isLoading = False

        Catch ex As Exception
            MessageBox.Show(ex.ToString, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub grdInvoices_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdInvoices.InitializeLayout
        Try
            InitializeGridLayout(e)

        Catch ex As Exception
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub btnAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAdd.Click

        Try
            Me.AddRowToGrid()

        Catch ex As Exception
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub grdInvoices_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdInvoices.CellChange

        Try
            Me.GridTextChange(DirectCast(sender, UltraGrid))

        Catch ex As Exception
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub grdInvoices_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdInvoices.ClickCellButton
        Try
            Me.ClickGridButton(CType(sender, Infragistics.Win.UltraWinGrid.UltraGrid))

        Catch ex As Exception
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click

        Try
            Me.EditMode = True
            Me.SetToEditMode()

        Catch ex As Exception
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click

        Try
            Me.RefreshDisplay()
        Catch ex As Exception
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()

    End Sub

    Private Sub dtpFrom_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpFrom.ValueChanged, dtpTo.ValueChanged

        Try
            Me.RefreshDisplay()

            Me.dtpTo.MinDate = dtpFrom.Value

        Catch ex As Exception
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub frmViewChecks_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        If Me.isLoading Then Exit Sub

        If Me.EditMode Then
            If MessageBox.Show("You are in edit mode; exit anyway?", "Please confirm.", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                e.Cancel = False
            Else
                e.Cancel = True
            End If
        End If

    End Sub

#End Region

    Private Sub grdInvoices_InitializeRowsCollection(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowsCollectionEventArgs) Handles grdInvoices.InitializeRowsCollection

    End Sub
End Class
