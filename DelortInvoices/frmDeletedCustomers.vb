Imports DelortBusObjects
Imports System.Configuration.ConfigurationSettings


Public Class frmDeletedCustomers
    Inherits frmBaseWindow

    Private Customers As New Customer(g_Settings)

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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ugCust As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents cmnuMain As System.Windows.Forms.ContextMenu
    Friend WithEvents cmnuDelete As System.Windows.Forms.MenuItem
    Friend WithEvents cmnuRefresh As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.ugCust = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.cmnuMain = New System.Windows.Forms.ContextMenu()
        Me.cmnuDelete = New System.Windows.Forms.MenuItem()
        Me.cmnuRefresh = New System.Windows.Forms.MenuItem()
        Me.GroupBox1.SuspendLayout()
        CType(Me.ugCust, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.ugCust)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(8, 6)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(647, 416)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Select Row, Then Right-Click"
        '
        'ugCust
        '
        Me.ugCust.ContextMenu = Me.cmnuMain
        Appearance1.TextHAlignAsString = "Left"
        Me.ugCust.DisplayLayout.Appearance = Appearance1
        Me.ugCust.DisplayLayout.MaxColScrollRegions = 1
        Me.ugCust.DisplayLayout.MaxRowScrollRegions = 1
        Me.ugCust.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugCust.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.ugCust.DisplayLayout.Override.AllowGroupBy = Infragistics.Win.DefaultableBoolean.[False]
        Me.ugCust.DisplayLayout.Override.AllowGroupSwapping = Infragistics.Win.UltraWinGrid.AllowGroupSwapping.NotAllowed
        Me.ugCust.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.ugCust.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.ugCust.DisplayLayout.Override.CellPadding = 0
        Me.ugCust.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugCust.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugCust.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.ugCust.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ugCust.Location = New System.Drawing.Point(8, 20)
        Me.ugCust.Name = "ugCust"
        Me.ugCust.Size = New System.Drawing.Size(632, 385)
        Me.ugCust.TabIndex = 2
        '
        'cmnuMain
        '
        Me.cmnuMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.cmnuDelete, Me.cmnuRefresh})
        '
        'cmnuDelete
        '
        Me.cmnuDelete.Index = 0
        Me.cmnuDelete.Text = "Udelete Customer"
        '
        'cmnuRefresh
        '
        Me.cmnuRefresh.Index = 1
        Me.cmnuRefresh.Text = "&Refresh"
        '
        'frmDeletedCustomers
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(659, 428)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "frmDeletedCustomers"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Deleted Customers"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.ugCust, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmDeletedCustomers_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadData()

    End Sub
    Private Sub LoadData()
        Me.ugCust.DataSource = Customers.GetDeletedCustomers

    End Sub
    Private Sub ugCust_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugCust.InitializeLayout

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
            .Columns(10).Hidden = True
            .Columns(11).Hidden = True
            .Columns(12).Hidden = True
            .Columns(13).Hidden = True

        End With

    End Sub

    Private Sub UnDeleteCustomer()

        If Me.ugCust.Rows.VisibleRowCount > 0 Then

            Dim FirstName As String = Me.ugCust.ActiveRow.Cells(2).Value.ToString
            Dim Lastname As String = Me.ugCust.ActiveRow.Cells(1).Value.ToString

            If MessageBox.Show("Are you sure you want to undelete " & FirstName & " " & Lastname & "?", ProgramName, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.Yes Then
                Customers.UndeleteCustomer(CInt(Me.ugCust.ActiveRow.Cells(0).Value))
            End If

            Me.LoadData()

        End If
    End Sub

    Private Sub cmnuDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmnuDelete.Click
        Me.UnDeleteCustomer()

    End Sub

    Private Sub cmnuRefresh_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmnuRefresh.Click
        LoadData()
    End Sub
End Class
