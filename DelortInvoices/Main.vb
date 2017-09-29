Option Strict On

Imports AccessUtils
Imports System.Collections.Specialized

Public Class frmMain
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
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents tbrMain As System.Windows.Forms.ToolBar
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents mnuMain As System.Windows.Forms.MainMenu
    Friend WithEvents mnuHelp As System.Windows.Forms.MenuItem
    Friend WithEvents mnuABout As System.Windows.Forms.MenuItem
    Friend WithEvents mnuWindow As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTileH As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTileV As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCascade As System.Windows.Forms.MenuItem
    Friend WithEvents mnuArrange As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTools As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCalculator As System.Windows.Forms.MenuItem
    Friend WithEvents tbViewAllInvoices As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbAddNewInvoice As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbOpenInvoice As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbPrintInvoice As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbSep1 As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbViewAllCustomers As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbAddCustomer As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbEditCustomer As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbSep2 As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbViewAllMaterials As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbAddMaterials As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbsep3 As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbExit As System.Windows.Forms.ToolBarButton
    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuNotepad As System.Windows.Forms.MenuItem
    Friend WithEvents LargeIcons As System.Windows.Forms.ImageList
    Friend WithEvents mnuMaint As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCompact As System.Windows.Forms.MenuItem
    Friend WithEvents Sep4 As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbInet As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbInvSum As System.Windows.Forms.ToolBarButton
    Friend WithEvents MnuInvoices As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCustomer As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMaterial As System.Windows.Forms.MenuItem
    Friend WithEvents mnuViewRegualrInvRpt As System.Windows.Forms.MenuItem
    Friend WithEvents mnuViewInvMoSummRpt As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddNewInv As System.Windows.Forms.MenuItem
    Friend WithEvents mnuOPenOneInvoice As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPrintOneInvoice As System.Windows.Forms.MenuItem
    Friend WithEvents mnuViewRegularCustRpt As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddNewCustomers As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFindCustByName As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPrintCustomers As System.Windows.Forms.MenuItem
    Friend WithEvents mnuViewMaterialRpt As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddNewMaterialItem As System.Windows.Forms.MenuItem
    Friend WithEvents mnuOpen As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMiscExp As System.Windows.Forms.MenuItem
    Friend WithEvents mnuViewMiscRpt As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddMiscExp As System.Windows.Forms.MenuItem
    Friend WithEvents tbMiscReport As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbMiscAdd As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbSep5 As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbCalc As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbNotePad As System.Windows.Forms.ToolBarButton
    Friend WithEvents mnuEditCategories As System.Windows.Forms.MenuItem
    Friend WithEvents tbEditMiscCats As System.Windows.Forms.ToolBarButton
    Friend WithEvents mnuViewCustomersInvoices As System.Windows.Forms.MenuItem
    Friend WithEvents tbCusInv As System.Windows.Forms.ToolBarButton
    Friend WithEvents mnuAdvInvSearch As System.Windows.Forms.MenuItem
    Friend WithEvents tbInvAdvSearch As System.Windows.Forms.ToolBarButton
    Friend WithEvents mnuMatAdvSrch As System.Windows.Forms.MenuItem
    Friend WithEvents tbMatAdvSrch As System.Windows.Forms.ToolBarButton
    Friend WithEvents mnuMiscExpAdvSrch As System.Windows.Forms.MenuItem
    Friend WithEvents tbMiscExpAdvSrch As System.Windows.Forms.ToolBarButton
    Friend WithEvents mnuUndelCust As System.Windows.Forms.MenuItem
    Friend WithEvents mnuWindowPrice As System.Windows.Forms.MenuItem
    Friend WithEvents uLstBar As Infragistics.Win.UltraWinListBar.UltraListBar
    Friend WithEvents mnuProgHelp As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCloseAll As System.Windows.Forms.MenuItem
    Friend WithEvents HelpProvider As System.Windows.Forms.HelpProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Dim Group1 As Infragistics.Win.UltraWinListBar.Group = New Infragistics.Win.UltraWinListBar.Group(True)
        Dim Item1 As Infragistics.Win.UltraWinListBar.Item = New Infragistics.Win.UltraWinListBar.Item()
        Dim Item2 As Infragistics.Win.UltraWinListBar.Item = New Infragistics.Win.UltraWinListBar.Item()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Item3 As Infragistics.Win.UltraWinListBar.Item = New Infragistics.Win.UltraWinListBar.Item()
        Dim Item4 As Infragistics.Win.UltraWinListBar.Item = New Infragistics.Win.UltraWinListBar.Item()
        Dim Item5 As Infragistics.Win.UltraWinListBar.Item = New Infragistics.Win.UltraWinListBar.Item()
        Dim Item6 As Infragistics.Win.UltraWinListBar.Item = New Infragistics.Win.UltraWinListBar.Item()
        Dim Item7 As Infragistics.Win.UltraWinListBar.Item = New Infragistics.Win.UltraWinListBar.Item()
        Dim Group2 As Infragistics.Win.UltraWinListBar.Group = New Infragistics.Win.UltraWinListBar.Group()
        Dim Item8 As Infragistics.Win.UltraWinListBar.Item = New Infragistics.Win.UltraWinListBar.Item()
        Dim Item9 As Infragistics.Win.UltraWinListBar.Item = New Infragistics.Win.UltraWinListBar.Item()
        Dim Item10 As Infragistics.Win.UltraWinListBar.Item = New Infragistics.Win.UltraWinListBar.Item()
        Dim Group3 As Infragistics.Win.UltraWinListBar.Group = New Infragistics.Win.UltraWinListBar.Group()
        Dim Item11 As Infragistics.Win.UltraWinListBar.Item = New Infragistics.Win.UltraWinListBar.Item()
        Dim Item12 As Infragistics.Win.UltraWinListBar.Item = New Infragistics.Win.UltraWinListBar.Item()
        Dim Item13 As Infragistics.Win.UltraWinListBar.Item = New Infragistics.Win.UltraWinListBar.Item()
        Dim Item14 As Infragistics.Win.UltraWinListBar.Item = New Infragistics.Win.UltraWinListBar.Item()
        Dim Item15 As Infragistics.Win.UltraWinListBar.Item = New Infragistics.Win.UltraWinListBar.Item()
        Dim Group4 As Infragistics.Win.UltraWinListBar.Group = New Infragistics.Win.UltraWinListBar.Group()
        Dim Item16 As Infragistics.Win.UltraWinListBar.Item = New Infragistics.Win.UltraWinListBar.Item()
        Dim Item17 As Infragistics.Win.UltraWinListBar.Item = New Infragistics.Win.UltraWinListBar.Item()
        Dim Item18 As Infragistics.Win.UltraWinListBar.Item = New Infragistics.Win.UltraWinListBar.Item()
        Dim Item19 As Infragistics.Win.UltraWinListBar.Item = New Infragistics.Win.UltraWinListBar.Item()
        Dim Group5 As Infragistics.Win.UltraWinListBar.Group = New Infragistics.Win.UltraWinListBar.Group()
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Item20 As Infragistics.Win.UltraWinListBar.Item = New Infragistics.Win.UltraWinListBar.Item()
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.tbViewAllInvoices = New System.Windows.Forms.ToolBarButton()
        Me.tbAddNewInvoice = New System.Windows.Forms.ToolBarButton()
        Me.tbOpenInvoice = New System.Windows.Forms.ToolBarButton()
        Me.tbrMain = New System.Windows.Forms.ToolBar()
        Me.tbInvSum = New System.Windows.Forms.ToolBarButton()
        Me.tbCusInv = New System.Windows.Forms.ToolBarButton()
        Me.tbPrintInvoice = New System.Windows.Forms.ToolBarButton()
        Me.tbInvAdvSearch = New System.Windows.Forms.ToolBarButton()
        Me.tbSep1 = New System.Windows.Forms.ToolBarButton()
        Me.tbViewAllCustomers = New System.Windows.Forms.ToolBarButton()
        Me.tbAddCustomer = New System.Windows.Forms.ToolBarButton()
        Me.tbEditCustomer = New System.Windows.Forms.ToolBarButton()
        Me.tbSep2 = New System.Windows.Forms.ToolBarButton()
        Me.tbViewAllMaterials = New System.Windows.Forms.ToolBarButton()
        Me.tbAddMaterials = New System.Windows.Forms.ToolBarButton()
        Me.tbMatAdvSrch = New System.Windows.Forms.ToolBarButton()
        Me.tbsep3 = New System.Windows.Forms.ToolBarButton()
        Me.tbMiscReport = New System.Windows.Forms.ToolBarButton()
        Me.tbMiscAdd = New System.Windows.Forms.ToolBarButton()
        Me.tbEditMiscCats = New System.Windows.Forms.ToolBarButton()
        Me.tbMiscExpAdvSrch = New System.Windows.Forms.ToolBarButton()
        Me.Sep4 = New System.Windows.Forms.ToolBarButton()
        Me.tbInet = New System.Windows.Forms.ToolBarButton()
        Me.tbCalc = New System.Windows.Forms.ToolBarButton()
        Me.tbNotePad = New System.Windows.Forms.ToolBarButton()
        Me.tbSep5 = New System.Windows.Forms.ToolBarButton()
        Me.tbExit = New System.Windows.Forms.ToolBarButton()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.mnuMain = New System.Windows.Forms.MainMenu(Me.components)
        Me.mnuFile = New System.Windows.Forms.MenuItem()
        Me.mnuOpen = New System.Windows.Forms.MenuItem()
        Me.MnuInvoices = New System.Windows.Forms.MenuItem()
        Me.mnuViewRegualrInvRpt = New System.Windows.Forms.MenuItem()
        Me.mnuViewInvMoSummRpt = New System.Windows.Forms.MenuItem()
        Me.mnuAddNewInv = New System.Windows.Forms.MenuItem()
        Me.mnuOPenOneInvoice = New System.Windows.Forms.MenuItem()
        Me.mnuPrintOneInvoice = New System.Windows.Forms.MenuItem()
        Me.mnuViewCustomersInvoices = New System.Windows.Forms.MenuItem()
        Me.mnuAdvInvSearch = New System.Windows.Forms.MenuItem()
        Me.mnuCustomer = New System.Windows.Forms.MenuItem()
        Me.mnuViewRegularCustRpt = New System.Windows.Forms.MenuItem()
        Me.mnuAddNewCustomers = New System.Windows.Forms.MenuItem()
        Me.mnuFindCustByName = New System.Windows.Forms.MenuItem()
        Me.mnuPrintCustomers = New System.Windows.Forms.MenuItem()
        Me.mnuMaterial = New System.Windows.Forms.MenuItem()
        Me.mnuViewMaterialRpt = New System.Windows.Forms.MenuItem()
        Me.mnuAddNewMaterialItem = New System.Windows.Forms.MenuItem()
        Me.mnuMatAdvSrch = New System.Windows.Forms.MenuItem()
        Me.mnuMiscExp = New System.Windows.Forms.MenuItem()
        Me.mnuViewMiscRpt = New System.Windows.Forms.MenuItem()
        Me.mnuAddMiscExp = New System.Windows.Forms.MenuItem()
        Me.mnuEditCategories = New System.Windows.Forms.MenuItem()
        Me.mnuMiscExpAdvSrch = New System.Windows.Forms.MenuItem()
        Me.mnuExit = New System.Windows.Forms.MenuItem()
        Me.mnuTools = New System.Windows.Forms.MenuItem()
        Me.mnuCalculator = New System.Windows.Forms.MenuItem()
        Me.mnuNotepad = New System.Windows.Forms.MenuItem()
        Me.mnuMaint = New System.Windows.Forms.MenuItem()
        Me.mnuCompact = New System.Windows.Forms.MenuItem()
        Me.mnuUndelCust = New System.Windows.Forms.MenuItem()
        Me.mnuWindowPrice = New System.Windows.Forms.MenuItem()
        Me.mnuWindow = New System.Windows.Forms.MenuItem()
        Me.mnuTileH = New System.Windows.Forms.MenuItem()
        Me.mnuTileV = New System.Windows.Forms.MenuItem()
        Me.mnuCascade = New System.Windows.Forms.MenuItem()
        Me.mnuArrange = New System.Windows.Forms.MenuItem()
        Me.mnuHelp = New System.Windows.Forms.MenuItem()
        Me.mnuABout = New System.Windows.Forms.MenuItem()
        Me.mnuProgHelp = New System.Windows.Forms.MenuItem()
        Me.LargeIcons = New System.Windows.Forms.ImageList(Me.components)
        Me.uLstBar = New Infragistics.Win.UltraWinListBar.UltraListBar()
        Me.HelpProvider = New System.Windows.Forms.HelpProvider()
        Me.mnuCloseAll = New System.Windows.Forms.MenuItem()
        CType(Me.uLstBar, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "")
        Me.ImageList1.Images.SetKeyName(1, "")
        Me.ImageList1.Images.SetKeyName(2, "")
        Me.ImageList1.Images.SetKeyName(3, "")
        Me.ImageList1.Images.SetKeyName(4, "")
        Me.ImageList1.Images.SetKeyName(5, "")
        Me.ImageList1.Images.SetKeyName(6, "")
        Me.ImageList1.Images.SetKeyName(7, "")
        Me.ImageList1.Images.SetKeyName(8, "")
        Me.ImageList1.Images.SetKeyName(9, "")
        Me.ImageList1.Images.SetKeyName(10, "")
        Me.ImageList1.Images.SetKeyName(11, "")
        Me.ImageList1.Images.SetKeyName(12, "")
        Me.ImageList1.Images.SetKeyName(13, "")
        Me.ImageList1.Images.SetKeyName(14, "")
        Me.ImageList1.Images.SetKeyName(15, "")
        Me.ImageList1.Images.SetKeyName(16, "")
        Me.ImageList1.Images.SetKeyName(17, "")
        Me.ImageList1.Images.SetKeyName(18, "")
        Me.ImageList1.Images.SetKeyName(19, "")
        Me.ImageList1.Images.SetKeyName(20, "")
        Me.ImageList1.Images.SetKeyName(21, "")
        Me.ImageList1.Images.SetKeyName(22, "")
        '
        'tbViewAllInvoices
        '
        Me.tbViewAllInvoices.ImageIndex = 0
        Me.tbViewAllInvoices.Name = "tbViewAllInvoices"
        Me.tbViewAllInvoices.Tag = "ViewInvoices"
        Me.tbViewAllInvoices.ToolTipText = "View All Invoices"
        '
        'tbAddNewInvoice
        '
        Me.tbAddNewInvoice.ImageIndex = 1
        Me.tbAddNewInvoice.Name = "tbAddNewInvoice"
        Me.tbAddNewInvoice.Tag = "NewInvoice"
        Me.tbAddNewInvoice.ToolTipText = "Add New Invoice"
        '
        'tbOpenInvoice
        '
        Me.tbOpenInvoice.ImageIndex = 2
        Me.tbOpenInvoice.Name = "tbOpenInvoice"
        Me.tbOpenInvoice.Tag = "OpenInvoice"
        Me.tbOpenInvoice.ToolTipText = "Open Invoice"
        '
        'tbrMain
        '
        Me.tbrMain.Appearance = System.Windows.Forms.ToolBarAppearance.Flat
        Me.tbrMain.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.tbViewAllInvoices, Me.tbInvSum, Me.tbAddNewInvoice, Me.tbOpenInvoice, Me.tbCusInv, Me.tbPrintInvoice, Me.tbInvAdvSearch, Me.tbSep1, Me.tbViewAllCustomers, Me.tbAddCustomer, Me.tbEditCustomer, Me.tbSep2, Me.tbViewAllMaterials, Me.tbAddMaterials, Me.tbMatAdvSrch, Me.tbsep3, Me.tbMiscReport, Me.tbMiscAdd, Me.tbEditMiscCats, Me.tbMiscExpAdvSrch, Me.Sep4, Me.tbInet, Me.tbCalc, Me.tbNotePad, Me.tbSep5, Me.tbExit})
        Me.tbrMain.ButtonSize = New System.Drawing.Size(40, 40)
        Me.tbrMain.DropDownArrows = True
        Me.tbrMain.ImageList = Me.ImageList1
        Me.tbrMain.Location = New System.Drawing.Point(0, 0)
        Me.tbrMain.Name = "tbrMain"
        Me.tbrMain.ShowToolTips = True
        Me.tbrMain.Size = New System.Drawing.Size(750, 28)
        Me.tbrMain.TabIndex = 4
        '
        'tbInvSum
        '
        Me.tbInvSum.ImageIndex = 10
        Me.tbInvSum.Name = "tbInvSum"
        Me.tbInvSum.Tag = "InvoiceSummary"
        Me.tbInvSum.ToolTipText = "Monthly Summary Report"
        '
        'tbCusInv
        '
        Me.tbCusInv.ImageIndex = 19
        Me.tbCusInv.Name = "tbCusInv"
        Me.tbCusInv.Tag = "CusInv"
        Me.tbCusInv.ToolTipText = "View Customer Invoices Report"
        '
        'tbPrintInvoice
        '
        Me.tbPrintInvoice.ImageIndex = 3
        Me.tbPrintInvoice.Name = "tbPrintInvoice"
        Me.tbPrintInvoice.Tag = "PrintInvoice"
        Me.tbPrintInvoice.ToolTipText = "Print Invoice"
        '
        'tbInvAdvSearch
        '
        Me.tbInvAdvSearch.ImageIndex = 20
        Me.tbInvAdvSearch.Name = "tbInvAdvSearch"
        Me.tbInvAdvSearch.Tag = "InvAdvSearch"
        Me.tbInvAdvSearch.ToolTipText = "Advanced Invoice Search"
        '
        'tbSep1
        '
        Me.tbSep1.Name = "tbSep1"
        Me.tbSep1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbViewAllCustomers
        '
        Me.tbViewAllCustomers.ImageIndex = 4
        Me.tbViewAllCustomers.Name = "tbViewAllCustomers"
        Me.tbViewAllCustomers.Tag = "ViewCustomers"
        Me.tbViewAllCustomers.ToolTipText = "View All Customers"
        '
        'tbAddCustomer
        '
        Me.tbAddCustomer.ImageIndex = 5
        Me.tbAddCustomer.Name = "tbAddCustomer"
        Me.tbAddCustomer.Tag = "AddCustomer"
        Me.tbAddCustomer.ToolTipText = "Add Customer"
        '
        'tbEditCustomer
        '
        Me.tbEditCustomer.ImageIndex = 6
        Me.tbEditCustomer.Name = "tbEditCustomer"
        Me.tbEditCustomer.Tag = "EditCustomer"
        Me.tbEditCustomer.ToolTipText = "Edit Customer"
        '
        'tbSep2
        '
        Me.tbSep2.Name = "tbSep2"
        Me.tbSep2.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbViewAllMaterials
        '
        Me.tbViewAllMaterials.ImageIndex = 7
        Me.tbViewAllMaterials.Name = "tbViewAllMaterials"
        Me.tbViewAllMaterials.Tag = "ViewMaterials"
        Me.tbViewAllMaterials.ToolTipText = "View Materials Report"
        '
        'tbAddMaterials
        '
        Me.tbAddMaterials.ImageIndex = 8
        Me.tbAddMaterials.Name = "tbAddMaterials"
        Me.tbAddMaterials.Tag = "AddMaterials"
        Me.tbAddMaterials.ToolTipText = "Add Materials"
        '
        'tbMatAdvSrch
        '
        Me.tbMatAdvSrch.ImageIndex = 11
        Me.tbMatAdvSrch.Name = "tbMatAdvSrch"
        Me.tbMatAdvSrch.Tag = "MatAdvSrch"
        Me.tbMatAdvSrch.ToolTipText = "Materials Advanced Search"
        '
        'tbsep3
        '
        Me.tbsep3.Name = "tbsep3"
        Me.tbsep3.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbMiscReport
        '
        Me.tbMiscReport.ImageIndex = 14
        Me.tbMiscReport.Name = "tbMiscReport"
        Me.tbMiscReport.Tag = "MiscRpt"
        Me.tbMiscReport.ToolTipText = "Misc Report"
        '
        'tbMiscAdd
        '
        Me.tbMiscAdd.ImageIndex = 15
        Me.tbMiscAdd.Name = "tbMiscAdd"
        Me.tbMiscAdd.Tag = "MiscAdd"
        Me.tbMiscAdd.ToolTipText = "Add Misc Expense"
        '
        'tbEditMiscCats
        '
        Me.tbEditMiscCats.ImageIndex = 18
        Me.tbEditMiscCats.Name = "tbEditMiscCats"
        Me.tbEditMiscCats.Tag = "EditMiscCats"
        Me.tbEditMiscCats.ToolTipText = "Edit Misc Categories"
        '
        'tbMiscExpAdvSrch
        '
        Me.tbMiscExpAdvSrch.ImageIndex = 21
        Me.tbMiscExpAdvSrch.Name = "tbMiscExpAdvSrch"
        Me.tbMiscExpAdvSrch.Tag = "MiscExpAdvSrch"
        Me.tbMiscExpAdvSrch.ToolTipText = "Misc Expense Advanced Search"
        '
        'Sep4
        '
        Me.Sep4.Name = "Sep4"
        Me.Sep4.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbInet
        '
        Me.tbInet.ImageIndex = 13
        Me.tbInet.Name = "tbInet"
        Me.tbInet.Tag = "IE"
        Me.tbInet.ToolTipText = "Launch Internet Explorer"
        '
        'tbCalc
        '
        Me.tbCalc.ImageIndex = 16
        Me.tbCalc.Name = "tbCalc"
        Me.tbCalc.Tag = "Calc"
        Me.tbCalc.ToolTipText = "Calculator"
        '
        'tbNotePad
        '
        Me.tbNotePad.ImageIndex = 17
        Me.tbNotePad.Name = "tbNotePad"
        Me.tbNotePad.Tag = "NotePad"
        Me.tbNotePad.ToolTipText = "Notepad"
        '
        'tbSep5
        '
        Me.tbSep5.Name = "tbSep5"
        Me.tbSep5.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbExit
        '
        Me.tbExit.ImageIndex = 9
        Me.tbExit.Name = "tbExit"
        Me.tbExit.Tag = "Exit"
        Me.tbExit.ToolTipText = "Exit"
        '
        'mnuMain
        '
        Me.mnuMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.mnuTools, Me.mnuWindow, Me.mnuHelp})
        '
        'mnuFile
        '
        Me.mnuFile.Index = 0
        Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuOpen, Me.mnuExit})
        Me.mnuFile.Text = "File"
        '
        'mnuOpen
        '
        Me.mnuOpen.Index = 0
        Me.mnuOpen.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MnuInvoices, Me.mnuCustomer, Me.mnuMaterial, Me.mnuMiscExp})
        Me.mnuOpen.Text = "&Open"
        '
        'MnuInvoices
        '
        Me.MnuInvoices.Index = 0
        Me.MnuInvoices.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuViewRegualrInvRpt, Me.mnuViewInvMoSummRpt, Me.mnuAddNewInv, Me.mnuOPenOneInvoice, Me.mnuPrintOneInvoice, Me.mnuViewCustomersInvoices, Me.mnuAdvInvSearch})
        Me.MnuInvoices.Text = "Invoices"
        '
        'mnuViewRegualrInvRpt
        '
        Me.mnuViewRegualrInvRpt.Index = 0
        Me.mnuViewRegualrInvRpt.Text = "&View Regular Report"
        '
        'mnuViewInvMoSummRpt
        '
        Me.mnuViewInvMoSummRpt.Index = 1
        Me.mnuViewInvMoSummRpt.Text = "View &Monthly Summary Report"
        '
        'mnuAddNewInv
        '
        Me.mnuAddNewInv.Index = 2
        Me.mnuAddNewInv.Text = "&Add New Invoice"
        '
        'mnuOPenOneInvoice
        '
        Me.mnuOPenOneInvoice.Index = 3
        Me.mnuOPenOneInvoice.Text = "&Open One Invoice"
        '
        'mnuPrintOneInvoice
        '
        Me.mnuPrintOneInvoice.Index = 4
        Me.mnuPrintOneInvoice.Text = "&Print One Invoice"
        '
        'mnuViewCustomersInvoices
        '
        Me.mnuViewCustomersInvoices.Index = 5
        Me.mnuViewCustomersInvoices.Text = "View Customers With Invoices"
        '
        'mnuAdvInvSearch
        '
        Me.mnuAdvInvSearch.Index = 6
        Me.mnuAdvInvSearch.Text = "Ad&vanced Seacrh"
        '
        'mnuCustomer
        '
        Me.mnuCustomer.Index = 1
        Me.mnuCustomer.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuViewRegularCustRpt, Me.mnuAddNewCustomers, Me.mnuFindCustByName, Me.mnuPrintCustomers})
        Me.mnuCustomer.Text = "&Customers"
        '
        'mnuViewRegularCustRpt
        '
        Me.mnuViewRegularCustRpt.Index = 0
        Me.mnuViewRegularCustRpt.Text = "&View Regular Report"
        '
        'mnuAddNewCustomers
        '
        Me.mnuAddNewCustomers.Index = 1
        Me.mnuAddNewCustomers.Text = "&Add New Customer"
        '
        'mnuFindCustByName
        '
        Me.mnuFindCustByName.Index = 2
        Me.mnuFindCustByName.Text = "&Find Customer By Name"
        '
        'mnuPrintCustomers
        '
        Me.mnuPrintCustomers.Index = 3
        Me.mnuPrintCustomers.Text = "&Print Customers"
        '
        'mnuMaterial
        '
        Me.mnuMaterial.Index = 2
        Me.mnuMaterial.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuViewMaterialRpt, Me.mnuAddNewMaterialItem, Me.mnuMatAdvSrch})
        Me.mnuMaterial.Text = "&Materials"
        '
        'mnuViewMaterialRpt
        '
        Me.mnuViewMaterialRpt.Index = 0
        Me.mnuViewMaterialRpt.Text = "&View Regular Report"
        '
        'mnuAddNewMaterialItem
        '
        Me.mnuAddNewMaterialItem.Index = 1
        Me.mnuAddNewMaterialItem.Text = "&Add New Item"
        '
        'mnuMatAdvSrch
        '
        Me.mnuMatAdvSrch.Index = 2
        Me.mnuMatAdvSrch.Text = "Advanced &Search"
        '
        'mnuMiscExp
        '
        Me.mnuMiscExp.Index = 3
        Me.mnuMiscExp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuViewMiscRpt, Me.mnuAddMiscExp, Me.mnuEditCategories, Me.mnuMiscExpAdvSrch})
        Me.mnuMiscExp.Text = "Misc Expenses"
        '
        'mnuViewMiscRpt
        '
        Me.mnuViewMiscRpt.Index = 0
        Me.mnuViewMiscRpt.Text = "&View Report"
        '
        'mnuAddMiscExp
        '
        Me.mnuAddMiscExp.Index = 1
        Me.mnuAddMiscExp.Text = "&Add"
        '
        'mnuEditCategories
        '
        Me.mnuEditCategories.Index = 2
        Me.mnuEditCategories.Text = "&Edit/Add Categories"
        '
        'mnuMiscExpAdvSrch
        '
        Me.mnuMiscExpAdvSrch.Index = 3
        Me.mnuMiscExpAdvSrch.Text = "&Advanced Search"
        '
        'mnuExit
        '
        Me.mnuExit.Index = 1
        Me.mnuExit.Text = "E&xit"
        '
        'mnuTools
        '
        Me.mnuTools.Index = 1
        Me.mnuTools.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuCalculator, Me.mnuNotepad, Me.mnuMaint, Me.mnuUndelCust, Me.mnuWindowPrice})
        Me.mnuTools.Text = "&Tools"
        '
        'mnuCalculator
        '
        Me.mnuCalculator.Index = 0
        Me.mnuCalculator.Text = "&Calculator"
        '
        'mnuNotepad
        '
        Me.mnuNotepad.Index = 1
        Me.mnuNotepad.Text = "&Notepad"
        '
        'mnuMaint
        '
        Me.mnuMaint.Index = 2
        Me.mnuMaint.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuCompact})
        Me.mnuMaint.Text = "Maintenance"
        '
        'mnuCompact
        '
        Me.mnuCompact.Index = 0
        Me.mnuCompact.Text = "Compact Database"
        '
        'mnuUndelCust
        '
        Me.mnuUndelCust.Index = 3
        Me.mnuUndelCust.Text = "Undelete Customers"
        '
        'mnuWindowPrice
        '
        Me.mnuWindowPrice.Index = 4
        Me.mnuWindowPrice.Text = "Window Price Form"
        '
        'mnuWindow
        '
        Me.mnuWindow.Index = 2
        Me.mnuWindow.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuTileH, Me.mnuTileV, Me.mnuCascade, Me.mnuArrange, Me.mnuCloseAll})
        Me.mnuWindow.Text = "&Window"
        '
        'mnuTileH
        '
        Me.mnuTileH.Index = 0
        Me.mnuTileH.Text = "Tile &Horizontally"
        '
        'mnuTileV
        '
        Me.mnuTileV.Index = 1
        Me.mnuTileV.Text = "Tile &Vertically"
        '
        'mnuCascade
        '
        Me.mnuCascade.Index = 2
        Me.mnuCascade.Text = "&Cascade"
        '
        'mnuArrange
        '
        Me.mnuArrange.Index = 3
        Me.mnuArrange.Text = "&Arrange"
        '
        'mnuHelp
        '
        Me.mnuHelp.Index = 3
        Me.mnuHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuABout, Me.mnuProgHelp})
        Me.mnuHelp.Text = "&Help"
        '
        'mnuABout
        '
        Me.mnuABout.Index = 0
        Me.mnuABout.Text = "&About Delort Invoices"
        '
        'mnuProgHelp
        '
        Me.mnuProgHelp.Index = 1
        Me.mnuProgHelp.Text = "Program Help"
        '
        'LargeIcons
        '
        Me.LargeIcons.ImageStream = CType(resources.GetObject("LargeIcons.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.LargeIcons.TransparentColor = System.Drawing.Color.Transparent
        Me.LargeIcons.Images.SetKeyName(0, "")
        Me.LargeIcons.Images.SetKeyName(1, "")
        Me.LargeIcons.Images.SetKeyName(2, "")
        Me.LargeIcons.Images.SetKeyName(3, "")
        Me.LargeIcons.Images.SetKeyName(4, "")
        Me.LargeIcons.Images.SetKeyName(5, "")
        Me.LargeIcons.Images.SetKeyName(6, "")
        Me.LargeIcons.Images.SetKeyName(7, "")
        Me.LargeIcons.Images.SetKeyName(8, "")
        Me.LargeIcons.Images.SetKeyName(9, "")
        Me.LargeIcons.Images.SetKeyName(10, "")
        Me.LargeIcons.Images.SetKeyName(11, "")
        Me.LargeIcons.Images.SetKeyName(12, "")
        Me.LargeIcons.Images.SetKeyName(13, "")
        Me.LargeIcons.Images.SetKeyName(14, "")
        Me.LargeIcons.Images.SetKeyName(15, "")
        Me.LargeIcons.Images.SetKeyName(16, "")
        Me.LargeIcons.Images.SetKeyName(17, "")
        Me.LargeIcons.Images.SetKeyName(18, "")
        Me.LargeIcons.Images.SetKeyName(19, "")
        Me.LargeIcons.Images.SetKeyName(20, "")
        Me.LargeIcons.Images.SetKeyName(21, "")
        '
        'uLstBar
        '
        Me.uLstBar.BorderStyle = Infragistics.Win.UIElementBorderStyle.Raised
        Me.uLstBar.Dock = System.Windows.Forms.DockStyle.Left
        Me.uLstBar.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Item1.Key = "ViewInvoices"
        Item1.LargeImageIndex = 0
        Item1.SmallImageIndex = 0
        Item1.Text = "Report"
        Appearance1.FontData.Name = "Verdana"
        Appearance1.Image = 4
        Appearance1.ImageBackgroundAlpha = Infragistics.Win.Alpha.UseAlphaLevel
        Appearance1.TextVAlignAsString = "Top"
        Item2.Appearance = Appearance1
        Item2.Key = "AddNewInvoice"
        Item2.LargeImageIndex = 1
        Item2.SmallImageIndex = 1
        Item2.Text = "Add"
        Item3.Key = "OpenInvoice"
        Item3.LargeImageIndex = 2
        Item3.SmallImageIndex = 2
        Item3.Text = "Open"
        Item4.Key = "PrintInvoice"
        Item4.LargeImageIndex = 3
        Item4.SmallImageIndex = 3
        Item4.Text = "Print"
        Item5.Key = "InvoiceSummary"
        Item5.LargeImageIndex = 10
        Item5.Text = "Summary"
        Item6.Key = "CustomerInvRpt"
        Item6.LargeImageIndex = 16
        Item6.Text = "Customer Report"
        Item7.Key = "AdvSearch"
        Item7.LargeImageIndex = 17
        Item7.Text = "Advanced Search"
        Group1.Items.Add(Item1)
        Group1.Items.Add(Item2)
        Group1.Items.Add(Item3)
        Group1.Items.Add(Item4)
        Group1.Items.Add(Item5)
        Group1.Items.Add(Item6)
        Group1.Items.Add(Item7)
        Group1.Text = "Invoices"
        Item8.Key = "ViewMaterials"
        Item8.LargeImageIndex = 9
        Item8.SmallImageIndex = 9
        Item8.Text = "Report"
        Item9.Key = "AddMaterials"
        Item9.LargeImageIndex = 8
        Item9.SmallImageIndex = 8
        Item9.Text = "Add"
        Item10.Key = "MatAdvSrch"
        Item10.LargeImageIndex = 18
        Item10.Text = "Advanced Search"
        Group2.Items.Add(Item8)
        Group2.Items.Add(Item9)
        Group2.Items.Add(Item10)
        Group2.Text = "Materials"
        Item11.Key = "ViewCustomers"
        Item11.LargeImageIndex = 6
        Item11.SmallImageIndex = 4
        Item11.Text = "All"
        Item12.Key = "Windows"
        Item12.LargeImageIndex = 20
        Item12.Text = "Windows"
        Item13.Key = "AddCustomers"
        Item13.LargeImageIndex = 7
        Item13.SmallImageIndex = 5
        Item13.Text = "Add"
        Item14.Key = "FindCustomer"
        Item14.LargeImageIndex = 5
        Item14.SmallImageIndex = 6
        Item14.Text = "Find"
        Item15.Key = "PrintCustomers"
        Item15.LargeImageIndex = 11
        Item15.Text = "Print"
        Group3.Items.Add(Item11)
        Group3.Items.Add(Item12)
        Group3.Items.Add(Item13)
        Group3.Items.Add(Item14)
        Group3.Items.Add(Item15)
        Group3.Text = "Customers"
        Item16.Key = "MiscRpt"
        Item16.LargeImageIndex = 14
        Item16.Tag = "MiscRpt"
        Item16.Text = "Report"
        Item17.Key = "MiscAdd"
        Item17.LargeImageIndex = 13
        Item17.Tag = "MiscAdd"
        Item17.Text = "Add "
        Item18.Key = "MiscEditCats"
        Item18.LargeImageIndex = 15
        Item18.Tag = "MiscEditCats"
        Item18.Text = "Edit Categories"
        Item19.Key = "MiscExpAdvSrch"
        Item19.LargeImageIndex = 19
        Item19.Tag = "MiscExpAdvSrch"
        Item19.Text = "Advanced Search"
        Group4.Items.Add(Item16)
        Group4.Items.Add(Item17)
        Group4.Items.Add(Item18)
        Group4.Items.Add(Item19)
        Group4.Text = "Misc Expenses"
        Appearance2.Image = 22
        Group5.Appearance = Appearance2
        Appearance3.Image = 22
        Item20.Appearance = Appearance3
        Item20.Key = "ViewChecks"
        Item20.LargeImageIndex = 21
        Item20.Text = "View Checks"
        Group5.Items.Add(Item20)
        Group5.Key = "Check Ledger"
        Group5.Text = "Financials"
        Me.uLstBar.Groups.Add(Group1)
        Me.uLstBar.Groups.Add(Group2)
        Me.uLstBar.Groups.Add(Group3)
        Me.uLstBar.Groups.Add(Group4)
        Me.uLstBar.Groups.Add(Group5)
        Appearance4.TextVAlignAsString = "Top"
        Me.uLstBar.ItemAppearance = Appearance4
        Me.uLstBar.LargeImageList = Me.LargeIcons
        Me.uLstBar.Location = New System.Drawing.Point(0, 28)
        Me.uLstBar.Name = "uLstBar"
        Me.uLstBar.ShowContextMenus = False
        Me.uLstBar.Size = New System.Drawing.Size(104, 518)
        Me.uLstBar.SmallImageList = Me.ImageList1
        '
        'HelpProvider
        '
        Me.HelpProvider.HelpNamespace = "InvoicesHelp.chm"
        '
        'mnuCloseAll
        '
        Me.mnuCloseAll.Index = 4
        Me.mnuCloseAll.Text = "C&lose All"
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(750, 546)
        Me.Controls.Add(Me.uLstBar)
        Me.Controls.Add(Me.tbrMain)
        Me.HelpProvider.SetHelpKeyword(Me, "MainForm")
        Me.HelpProvider.SetHelpNavigator(Me, System.Windows.Forms.HelpNavigator.KeywordIndex)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.Menu = Me.mnuMain
        Me.Name = "frmMain"
        Me.HelpProvider.SetShowHelp(Me, True)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Delort Invoices"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.uLstBar, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub tbrMain_ButtonClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles tbrMain.ButtonClick

        Try
            Select Case e.Button.Tag.ToString

                Case "ViewInvoices"
                    Me.ManageForms(New frmViewEditInvoices)
                Case "NewInvoice"
                    Me.ManageForms(New frmEditInvoice(Mode.AddNew))
                Case "OpenInvoice"
                    Me.ManageForms(New frmOpenOneInvoice)
                Case "CusInv"
                    Me.ManageForms(New frmCustomerInvoices)
                Case "PrintInvoice"
                    Me.ManageForms(New frmPrintOneInvoice)
                Case "InvAdvSearch"
                    ManageForms(New frmInvAdvancedSearch)
                Case "ViewCustomers"
                    Me.ManageForms(New frmCustomerReport)
                Case "AddCustomer"
                    Utilities.ViewAddEditCustomer(Me, Nothing, Nothing)
                Case "EditCustomer"
                    Me.ManageForms(New frmFindOneCustomer)
                Case "ViewMaterials"
                    Me.ManageForms(New frmViewMaterials)
                Case "AddMaterials"
                    Utilities.ViewAddMaterials(Me)
                Case "MatAdvSrch"
                    Me.ManageForms(New frmMatAdvSearch)
                Case "MiscRpt"
                    Utilities.ViewMiscExpReport(Me)
                Case "MiscAdd"
                    Utilities.ViewMiscAdd(Me)
                Case "EditMiscCats"
                    Utilities.ViewEditMiscCats(Me)
                Case "MiscExpAdvSrch"
                    Me.ManageForms(New frmMiscExpAdvSrch)
                Case "Calc"
                    Shell("Calc.exe", AppWinStyle.MaximizedFocus)
                Case "NotePad"
                    Shell("notepad.exe", AppWinStyle.MaximizedFocus)
                Case "Exit"
                    Me.Close()
                Case "InvoiceSummary"
                    Utilities.ShowInvSummaryForm(Me)
                Case "IE"
                    Shell("C:\Program Files\Internet Explorer\iexplore.exe", AppWinStyle.MaximizedFocus)
            End Select

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub


    Private Sub mnuAddInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try

            Me.ManageForms(New frmEditInvoice(Mode.AddNew))

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub mnuViewInvoices_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try
            Cursor = Cursors.WaitCursor

            Me.ManageForms(New frmViewEditInvoices)


        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            Cursor = Cursors.Default
        End Try


    End Sub

    Private Sub mnuAddCust_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try

            Utilities.ViewAddEditCustomer(Me, Nothing, Nothing)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try


    End Sub

    Private Sub mnuAddMaterials_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try

            Utilities.ViewAddMaterials(Me)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub mnuSearchByNum_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try

            Me.ManageForms(New frmOpenOneInvoice)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub mnuSearchCust_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try

            Utilities.ViewAddEditCustomer(Me, Nothing, Nothing)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub uLstBar_ItemSelected(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinListBar.ItemEventArgs) Handles uLstBar.ItemSelected

        Try
            Cursor = Cursors.WaitCursor

            Select Case e.Item.Key
                Case "AddNewInvoice"
                    Me.ManageForms(New frmEditInvoice(Mode.AddNew))
                Case "OpenInvoice"
                    Me.ManageForms(New frmOpenOneInvoice)
                Case "ViewInvoices"
                    Me.ManageForms(New frmViewEditInvoices)
                Case "ViewMaterials"
                    Me.ManageForms(New frmViewMaterials)
                Case "AddMaterials"
                    Utilities.ViewAddMaterials(Me)
                Case "MatAdvSrch"
                    Me.ManageForms(New frmMatAdvSearch)
                Case "ViewCustomers"
                    Me.ManageForms(New frmCustomerReport)
                Case "Windows"
                    Me.ManageForms(New frmWindowsCustomersReport)
                Case "AddCustomers"
                    Utilities.ViewAddEditCustomer(Me, Nothing, Nothing)
                Case "FindCustomer"
                    Me.ManageForms(New frmFindOneCustomer)
                Case "PrintInvoice"
                    Me.ManageForms(New frmPrintOneInvoice)
                Case "InvoiceSummary"
                    Utilities.ShowInvSummaryForm(Me)
                Case "PrintCustomers"
                    Me.ManageForms(New frmPrintCustomers)
                Case "MiscAdd"
                    Utilities.ViewMiscAdd(Me)
                Case "MiscRpt"
                    Utilities.ViewMiscExpReport(Me)
                Case "MiscEditCats"
                    Utilities.ViewEditMiscCats(Me)
                Case "CustomerInvRpt"
                    Me.ManageForms(New frmCustomerInvoices)
                Case "MiscExpAdvSrch"
                    Me.ManageForms(New frmMiscExpAdvSrch)
                Case "AdvSearch"
                    ManageForms(New frmInvAdvancedSearch)
                Case "ViewChecks"
                    ManageForms(New frmViewChecks)
            End Select

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            Cursor = Cursors.Default

        End Try

    End Sub

    Private Sub ManageForms(ByVal frm As Form)

        For Each mdiFrm As Form In Me.MdiChildren
            If mdiFrm.Name = frm.Name Then
                MessageBox.Show(mdiFrm.Text & " form is open.", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Sub
            End If
        Next

        With frm
            .MdiParent = Me
            .StartPosition = FormStartPosition.CenterScreen
            .Show()
            .WindowState = FormWindowState.Normal
        End With

    End Sub

    Private Sub frmMain_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        Try
            If Me.MdiChildren.Length > 0 Then
                MessageBox.Show("There are " & Me.MdiChildren.Length.ToString & " Form(s) Open. Please close all forms before exiting", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                e.Cancel = True
            End If


        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub mnuTileH_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuTileH.Click

        Try
            Me.LayoutMdi(MdiLayout.TileHorizontal)

        Catch ex As Exception

            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub mnuTileV_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuTileV.Click

        Try
            Me.LayoutMdi(MdiLayout.TileVertical)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub mnuCascade_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCascade.Click

        Try
            Me.LayoutMdi(MdiLayout.Cascade)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub mnuArrange_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuArrange.Click

        Try
            Me.LayoutMdi(MdiLayout.ArrangeIcons)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub mnuCalculator_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCalculator.Click
        Shell("Calc.exe", AppWinStyle.MaximizedFocus)

    End Sub

    Private Sub mnuAddNewInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try
            Me.ManageForms(New frmEditInvoice(Mode.AddNew))

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub


    Private Sub mnuAddNewCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try
            Utilities.ViewAddEditCustomer(Me, Nothing, Nothing)


        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub mnuAddNewMaterials_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try
            Utilities.ViewAddMaterials(Me)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub mnuEditOneCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try
            Me.ManageForms(New frmFindOneCustomer)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try


    End Sub

    Private Sub mnuExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuExit.Click

        Try
            Me.Close()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub mnuOpenOneInv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try
            Me.ManageForms(New frmOpenOneInvoice)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub mnuPrintOneInv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try
            Me.ManageForms(New frmPrintOneInvoice)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub mnuViewAllCustomers_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try
            Me.ManageForms(New frmCustomerReport)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub mnuViewAllInvoices_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try
            Me.ManageForms(New frmViewEditInvoices)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub mnuViewAllMaterials_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try
            Me.ManageForms(New frmViewMaterials)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub mnuNotepad_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuNotepad.Click

        Shell("Notepad.exe", AppWinStyle.NormalFocus)

    End Sub

    Private Sub mnuABout_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuABout.Click

        Try
            Utilities.ShowAbout()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = ProgramName
    End Sub

    Private Sub mnuCompact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCompact.Click

        Try

            If MessageBox.Show("Are you sure you want to compact the database?" & vbCrLf & "Note: It is recommended that you do a backup before compacting", ProgramName, MessageBoxButtons.YesNo) = DialogResult.No Then
                Exit Sub
            End If

            Cursor = Cursors.WaitCursor
            Maintenance.CompactDB()

            MessageBox.Show("Database Compacted", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Information)


        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            Cursor = Cursors.Default
        End Try

    End Sub

    Private Sub frmMain_MdiChildActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.MdiChildActivate

        Me.mnuCompact.Enabled = False

    End Sub

    Private Sub mnuMoSummary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Utilities.ShowInvSummaryForm(Me)

    End Sub

    Private Sub mnuAddNewCustomers_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuAddNewCustomers.Click

        Utilities.ViewAddEditCustomer(Me, Nothing, Nothing)

    End Sub

    Private Sub mnuOPenOneInvoice_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuOPenOneInvoice.Click

        Me.ManageForms(New frmOpenOneInvoice)

    End Sub

    Private Sub mnuAddNewInv_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuAddNewInv.Click

        Me.ManageForms(New frmEditInvoice(Mode.AddNew))

    End Sub

    Private Sub mnuAddNewMaterialItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuAddNewMaterialItem.Click

        Utilities.ViewAddMaterials(Me)

    End Sub

    Private Sub mnuFindCustByName_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFindCustByName.Click

        Me.ManageForms(New frmFindOneCustomer)

    End Sub

    Private Sub mnuPrintCustomers_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuPrintCustomers.Click
        Me.ManageForms(New frmPrintCustomers)
    End Sub

    Private Sub mnuViewInvMoSummRpt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuViewInvMoSummRpt.Click

        Utilities.ShowInvSummaryForm(Me)

    End Sub

    Private Sub mnuViewMaterialRpt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuViewMaterialRpt.Click

        Me.ManageForms(New frmViewMaterials)

    End Sub

    Private Sub mnuViewRegualrInvRpt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuViewRegualrInvRpt.Click

        Me.ManageForms(New frmViewEditInvoices)

    End Sub

    Private Sub mnuViewRegularCustRpt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuViewRegularCustRpt.Click

        Me.ManageForms(New frmCustomerReport)

    End Sub

    Private Sub mnuPrintOneInvoice_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuPrintOneInvoice.Click

        Me.ManageForms(New frmPrintOneInvoice)

    End Sub

    Private Sub mnuViewMiscRpt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuViewMiscRpt.Click

        Utilities.ViewMiscExpReport(Me)

    End Sub

    Private Sub mnuAddMiscExp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuAddMiscExp.Click

        Utilities.ViewMiscAdd(Me)

    End Sub

    Private Sub mnuEditCategories_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuEditCategories.Click
        Utilities.ViewEditMiscCats(Me)
    End Sub

    Private Sub mnuViewCustomersInvoices_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuViewCustomersInvoices.Click

        Me.ManageForms(New frmCustomerInvoices)

    End Sub

    Private Sub mnuAdvInvSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuAdvInvSearch.Click
        ManageForms(New frmInvAdvancedSearch)

    End Sub

    Private Sub mnuMatAdvSrch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuMatAdvSrch.Click

        Me.ManageForms(New frmMatAdvSearch)

    End Sub

    Private Sub mnuMiscExpAdvSrch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuMiscExpAdvSrch.Click

        Me.ManageForms(New frmMiscExpAdvSrch)

    End Sub

    Private Sub mnuUndelCust_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuUndelCust.Click
        Me.ManageForms(New frmDeletedCustomers)

    End Sub

    Private Sub mnuWindowPrice_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuWindowPrice.Click
        ManageForms(New frmWindowPrice)

    End Sub

    Private Sub mnuProgHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuProgHelp.Click

        Try
            If System.IO.File.Exists(Me.HelpProvider.HelpNamespace) Then
                System.Diagnostics.Process.Start(Me.HelpProvider.HelpNamespace)
            Else
                MessageBox.Show("Help file is missing.", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If

        Catch ex As Exception
            MessageBox.Show("Help file is missing or corrupted.", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub mnuCloseAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCloseAll.Click
        ' Close all child forms of the parent.
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub
End Class
