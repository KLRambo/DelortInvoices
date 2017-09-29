Option Strict On

'DisplayLayout-GroupBy Box-Hidden=True
Imports Infragistics
Imports Infragistics.Shared
Imports Infragistics.Win
Imports DelortBusObjects
Imports System.Configuration.ConfigurationSettings
Imports System.Data.OleDb

Public Class frmCategoriesEdit

    Inherits frmBaseWindow

    Public dsCategories As DataSet

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
    Friend WithEvents ugCats As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents grpAction As System.Windows.Forms.GroupBox
    Friend WithEvents btnOpen As System.Windows.Forms.Button
    Friend WithEvents btnAddNew As System.Windows.Forms.Button
    Friend WithEvents btnReactivate As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents lblFooter As System.Windows.Forms.Label
    Friend WithEvents cmMain As System.Windows.Forms.ContextMenu
    Friend WithEvents cmEdit As System.Windows.Forms.MenuItem
    Friend WithEvents CMDeactivate As System.Windows.Forms.MenuItem
    Friend WithEvents cmReactivate As System.Windows.Forms.MenuItem
    Friend WithEvents btnDeactivate As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents cmDelete As System.Windows.Forms.MenuItem
    Friend WithEvents ToolTip As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCategoriesEdit))
        Me.ugCats = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.cmMain = New System.Windows.Forms.ContextMenu()
        Me.cmEdit = New System.Windows.Forms.MenuItem()
        Me.CMDeactivate = New System.Windows.Forms.MenuItem()
        Me.cmReactivate = New System.Windows.Forms.MenuItem()
        Me.cmDelete = New System.Windows.Forms.MenuItem()
        Me.grpAction = New System.Windows.Forms.GroupBox()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnAddNew = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnReactivate = New System.Windows.Forms.Button()
        Me.btnDeactivate = New System.Windows.Forms.Button()
        Me.btnOpen = New System.Windows.Forms.Button()
        Me.lblFooter = New System.Windows.Forms.Label()
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        CType(Me.ugCats, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpAction.SuspendLayout()
        Me.SuspendLayout()
        '
        'ugCats
        '
        Me.ugCats.ContextMenu = Me.cmMain
        Appearance1.TextHAlignAsString = "Left"
        Me.ugCats.DisplayLayout.Appearance = Appearance1
        Me.ugCats.DisplayLayout.GroupByBox.Hidden = True
        Me.ugCats.DisplayLayout.MaxColScrollRegions = 1
        Me.ugCats.DisplayLayout.MaxRowScrollRegions = 1
        Me.ugCats.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugCats.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.ugCats.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.ugCats.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.ugCats.DisplayLayout.Override.CellPadding = 0
        Me.ugCats.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Appearance2.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.ugCats.DisplayLayout.Override.SelectedRowAppearance = Appearance2
        Me.ugCats.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugCats.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.ugCats.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.ugCats.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ugCats.Location = New System.Drawing.Point(3, 8)
        Me.ugCats.Name = "ugCats"
        Me.ugCats.Size = New System.Drawing.Size(280, 224)
        Me.ugCats.TabIndex = 1
        Me.ugCats.Text = "Misc Expenses Categories"
        Me.ugCats.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChange
        '
        'cmMain
        '
        Me.cmMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.cmEdit, Me.CMDeactivate, Me.cmReactivate, Me.cmDelete})
        '
        'cmEdit
        '
        Me.cmEdit.Index = 0
        Me.cmEdit.Text = "&Edit"
        '
        'CMDeactivate
        '
        Me.CMDeactivate.Index = 1
        Me.CMDeactivate.Text = "&Deactivate"
        '
        'cmReactivate
        '
        Me.cmReactivate.Index = 2
        Me.cmReactivate.Text = "&Reactivate"
        '
        'cmDelete
        '
        Me.cmDelete.Index = 3
        Me.cmDelete.Text = "Delete"
        '
        'grpAction
        '
        Me.grpAction.BackColor = System.Drawing.Color.Transparent
        Me.grpAction.Controls.Add(Me.btnExit)
        Me.grpAction.Controls.Add(Me.btnAddNew)
        Me.grpAction.Controls.Add(Me.btnDelete)
        Me.grpAction.Controls.Add(Me.btnReactivate)
        Me.grpAction.Controls.Add(Me.btnDeactivate)
        Me.grpAction.Controls.Add(Me.btnOpen)
        Me.grpAction.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpAction.Location = New System.Drawing.Point(3, 238)
        Me.grpAction.Name = "grpAction"
        Me.grpAction.Size = New System.Drawing.Size(280, 80)
        Me.grpAction.TabIndex = 2
        Me.grpAction.TabStop = False
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.AliceBlue
        Me.btnExit.Image = CType(resources.GetObject("btnExit.Image"), System.Drawing.Image)
        Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnExit.Location = New System.Drawing.Point(182, 46)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(87, 24)
        Me.btnExit.TabIndex = 6
        Me.btnExit.Text = "   E&xit"
        Me.ToolTip.SetToolTip(Me.btnExit, "Exit Form")
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'btnAddNew
        '
        Me.btnAddNew.BackColor = System.Drawing.Color.AliceBlue
        Me.btnAddNew.Image = CType(resources.GetObject("btnAddNew.Image"), System.Drawing.Image)
        Me.btnAddNew.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnAddNew.Location = New System.Drawing.Point(182, 14)
        Me.btnAddNew.Name = "btnAddNew"
        Me.btnAddNew.Size = New System.Drawing.Size(87, 24)
        Me.btnAddNew.TabIndex = 5
        Me.btnAddNew.Text = "&Add New"
        Me.btnAddNew.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip.SetToolTip(Me.btnAddNew, "Add a New Category")
        Me.btnAddNew.UseVisualStyleBackColor = False
        '
        'btnDelete
        '
        Me.btnDelete.BackColor = System.Drawing.Color.AliceBlue
        Me.btnDelete.Image = CType(resources.GetObject("btnDelete.Image"), System.Drawing.Image)
        Me.btnDelete.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDelete.Location = New System.Drawing.Point(95, 46)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(87, 24)
        Me.btnDelete.TabIndex = 4
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip.SetToolTip(Me.btnDelete, "Delete Selected Category")
        Me.btnDelete.UseVisualStyleBackColor = False
        '
        'btnReactivate
        '
        Me.btnReactivate.BackColor = System.Drawing.Color.AliceBlue
        Me.btnReactivate.Image = CType(resources.GetObject("btnReactivate.Image"), System.Drawing.Image)
        Me.btnReactivate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnReactivate.Location = New System.Drawing.Point(8, 46)
        Me.btnReactivate.Name = "btnReactivate"
        Me.btnReactivate.Size = New System.Drawing.Size(87, 24)
        Me.btnReactivate.TabIndex = 3
        Me.btnReactivate.Text = "&Reactivate"
        Me.btnReactivate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip.SetToolTip(Me.btnReactivate, "Reactivate Selected Category")
        Me.btnReactivate.UseVisualStyleBackColor = False
        '
        'btnDeactivate
        '
        Me.btnDeactivate.BackColor = System.Drawing.Color.AliceBlue
        Me.btnDeactivate.Image = CType(resources.GetObject("btnDeactivate.Image"), System.Drawing.Image)
        Me.btnDeactivate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDeactivate.Location = New System.Drawing.Point(95, 14)
        Me.btnDeactivate.Name = "btnDeactivate"
        Me.btnDeactivate.Size = New System.Drawing.Size(87, 24)
        Me.btnDeactivate.TabIndex = 1
        Me.btnDeactivate.Text = "&Deactivate"
        Me.btnDeactivate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip.SetToolTip(Me.btnDeactivate, "Deactivate Selected Category")
        Me.btnDeactivate.UseVisualStyleBackColor = False
        '
        'btnOpen
        '
        Me.btnOpen.BackColor = System.Drawing.Color.AliceBlue
        Me.btnOpen.Image = CType(resources.GetObject("btnOpen.Image"), System.Drawing.Image)
        Me.btnOpen.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnOpen.Location = New System.Drawing.Point(8, 14)
        Me.btnOpen.Name = "btnOpen"
        Me.btnOpen.Size = New System.Drawing.Size(87, 24)
        Me.btnOpen.TabIndex = 0
        Me.btnOpen.Text = "    &Edit"
        Me.ToolTip.SetToolTip(Me.btnOpen, "Edit Selected Category")
        Me.btnOpen.UseVisualStyleBackColor = False
        '
        'lblFooter
        '
        Me.lblFooter.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblFooter.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFooter.ForeColor = System.Drawing.Color.Red
        Me.lblFooter.Location = New System.Drawing.Point(3, 320)
        Me.lblFooter.Name = "lblFooter"
        Me.lblFooter.Size = New System.Drawing.Size(280, 24)
        Me.lblFooter.TabIndex = 3
        Me.lblFooter.Text = "Blue on White=Active  Red on Yellow=Inactive"
        '
        'frmCategoriesEdit
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(286, 350)
        Me.Controls.Add(Me.lblFooter)
        Me.Controls.Add(Me.grpAction)
        Me.Controls.Add(Me.ugCats)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmCategoriesEdit"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Misc Categories Edit Form"
        CType(Me.ugCats, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpAction.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub ugCats_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugCats.InitializeLayout

        With e.Layout.Bands(0)
            .Columns("CategoryID").Hidden = True
            .Columns("Description").Header.Caption = ""
            .Columns("Description").Width = 250
            .Columns("Active").Hidden = True
        End With

    End Sub

    Private Sub btnOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpen.Click

        Try
            OpenCategory()
            RefreshData()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
    Private Sub OpenCategory()

        Utilities.ViewEditCategory(Convert.ToInt32(ugCats.ActiveRow.Cells("CategoryID").Value), Utilities.CrudMode.UPDATE)

    End Sub

    Private Sub btnAddNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddNew.Click

        Try
            Utilities.ViewEditCategory(Convert.ToInt32(ugCats.ActiveRow.Cells("CategoryID").Value), Utilities.CrudMode.INSERT)
            RefreshData()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub btnDeactivate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeactivate.Click

        Try

            DeactivateCategory()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub


    Private Sub btnReactivate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReactivate.Click

        Try

            Reactivate()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub


    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click

        Me.Close()

    End Sub
    Private Function CreateCategoryObject() As MiscCategory

        Dim CatID As Integer = Convert.ToInt32(ugCats.ActiveRow.Cells("CategoryID").Value)
        Dim Cat As New MiscCategory(g_settings, CatID)

        Return Cat

    End Function

    Private Sub RefreshData()

        Dim MiscExp As New MiscExp((g_settings))

        Dim MiscCats As DataSet = MiscExp.GetExpenseCategories
        Me.ugCats.DataSource = MiscCats


    End Sub

    Private Sub ugCats_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles ugCats.InitializeRow

        If e.Row.Cells("Active").Value.ToString = "0" Then
            e.Row.Appearance.ForeColor = Color.Red
            e.Row.Appearance.BackColor = Color.Yellow
        Else
            e.Row.Appearance.ForeColor = Color.Blue

        End If

    End Sub


    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click

        Try
            Me.DeleteCategory()

        Catch ex As OleDbException
            Errhandler.Logging.LogError(ex)
            MessageBox.Show("Category can't be deleted beause there are records associated with it, however you may DEACTIVATE it.", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub CMDeactivate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMDeactivate.Click

        Try

            Me.DeactivateCategory()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub
    Private Sub Reactivate()

        If ugCats.ActiveRow.Cells("Active").Value.ToString = "0" Then

            Dim Cat As MiscCategory = Me.CreateCategoryObject

            With Cat
                .Active = 1
                .CrudMode = MiscCategory.DBMode.Update
                Cat.InsertUpdateDelete()
            End With

            RefreshData()

        Else
            MessageBox.Show(ugCats.ActiveRow.Cells("Description").Value.ToString & " category is already active!", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If
    End Sub
    Private Sub DeactivateCategory()

        If ugCats.ActiveRow.Cells("Active").Value.ToString = "1" Then

            If MessageBox.Show("Are you sure you want to deactivate " & ugCats.ActiveRow.Cells("Description").Value.ToString & "?", ProgramName, MessageBoxButtons.YesNo) = DialogResult.Yes Then
                Dim Cat As MiscCategory = Me.CreateCategoryObject

                With Cat
                    .Active = 0
                    .CrudMode = MiscCategory.DBMode.Update
                    .InsertUpdateDelete()
                End With
                RefreshData()
            End If
        Else
            MessageBox.Show(ugCats.ActiveRow.Cells("Description").Value.ToString & " category is already inactive!", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If

    End Sub
    Private Sub cmReactivate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmReactivate.Click

        Try

            Reactivate()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try


    End Sub

    Private Sub cmEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmEdit.Click

        Try

            OpenCategory()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub DeleteCategory()

        If MessageBox.Show("Are you sure you want to delete " & ugCats.ActiveRow.Cells("Description").Value.ToString & "?", ProgramName, MessageBoxButtons.YesNo, MessageBoxIcon.Information) = DialogResult.Yes Then
            Dim Cat As MiscCategory = Me.CreateCategoryObject
            Cat.CrudMode = MiscCategory.DBMode.Delete
            Cat.InsertUpdateDelete()
            Me.RefreshData()
        End If


    End Sub

    Private Sub cmDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmDelete.Click

        Try
            Me.DeleteCategory()

        Catch ex As OleDbException
            Errhandler.Logging.LogError(ex)
            MessageBox.Show("Category can't be deleted beause there are records associated with it, however you may DEACTIVATE it.", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
End Class
