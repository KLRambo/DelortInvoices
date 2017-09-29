
#Region " Options "

Option Explicit On
Option Strict On

#End Region

Imports DelortBusObjects
Imports DelortBusObjects.Invoice
Imports System.Configuration.ConfigurationSettings
Imports System.Drawing.Color

Public Enum Mode
    AddNew = 0
    Edit = 1
End Enum


Public Class frmEditInvoice

    Inherits frmBaseWindow

    Public CallingForm As frmViewEditInvoices
    Public FormMode As Mode
    Private WithEvents _Invoice As New DelortBusObjects.Invoice(g_Settings)
    Private _CustomerID As Integer
    Private Loading As Boolean = True

    Private WithEvents _LineItem0 As New LineItem
    Private WithEvents _LineItem1 As LineItem
    Private WithEvents _LineItem2 As LineItem
    Private WithEvents _LineItem3 As LineItem
    Private WithEvents _LineItem4 As LineItem

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal Mode As Mode)
        MyBase.New()

        If Mode = Mode.AddNew Then
            Me.Loading = False
            Me.FormMode = Mode.AddNew
        Else
            Me.FormMode = Mode.Edit
            Me.Loading = True
        End If


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
    Friend WithEvents grpLineItems As System.Windows.Forms.GroupBox
    Friend WithEvents lblLineTotal As System.Windows.Forms.Label
    Friend WithEvents grpFooter As System.Windows.Forms.GroupBox
    Friend WithEvents BtnNewLine As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents lblLineTotal1 As System.Windows.Forms.Label
    Friend WithEvents txtMaterialCost1 As System.Windows.Forms.TextBox
    Friend WithEvents txtMaterialDesc1 As System.Windows.Forms.TextBox
    Friend WithEvents dtpLine0 As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblLineTotal0 As System.Windows.Forms.Label
    Friend WithEvents txtMaterialCost0 As System.Windows.Forms.TextBox
    Friend WithEvents txtMaterialDesc0 As System.Windows.Forms.TextBox
    Friend WithEvents txtLaborDesc1 As System.Windows.Forms.TextBox
    Friend WithEvents txtLaborCost0 As System.Windows.Forms.TextBox
    Friend WithEvents txtLaborDesc0 As System.Windows.Forms.TextBox
    Friend WithEvents txtLaborCost1 As System.Windows.Forms.TextBox
    Friend WithEvents txtMaterialDesc2 As System.Windows.Forms.TextBox
    Friend WithEvents dtpLine4 As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtMaterialDesc4 As System.Windows.Forms.TextBox
    Friend WithEvents txtLaborDesc4 As System.Windows.Forms.TextBox
    Friend WithEvents dtpLine3 As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpLine2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpLine1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblDate As System.Windows.Forms.Label
    Friend WithEvents lblMaterialCost As System.Windows.Forms.Label
    Friend WithEvents lblMaterialDesc As System.Windows.Forms.Label
    Friend WithEvents lblLaborCost As System.Windows.Forms.Label
    Friend WithEvents lblLaborDesc As System.Windows.Forms.Label
    Friend WithEvents txtLaborCost3 As System.Windows.Forms.TextBox
    Friend WithEvents txtLaborDesc3 As System.Windows.Forms.TextBox
    Friend WithEvents txtLaborCost2 As System.Windows.Forms.TextBox
    Friend WithEvents txtLaborDesc2 As System.Windows.Forms.TextBox
    Friend WithEvents txtMaterialDesc3 As System.Windows.Forms.TextBox
    Friend WithEvents txtLaborCost4 As System.Windows.Forms.TextBox
    Friend WithEvents txtMaterialCost2 As System.Windows.Forms.TextBox
    Friend WithEvents txtMaterialCost3 As System.Windows.Forms.TextBox
    Friend WithEvents txtMaterialCost4 As System.Windows.Forms.TextBox
    Friend WithEvents lblLineTotal4 As System.Windows.Forms.Label
    Friend WithEvents lblLineTotal2 As System.Windows.Forms.Label
    Friend WithEvents lblLineTotal3 As System.Windows.Forms.Label
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents chkL0 As System.Windows.Forms.CheckBox
    Friend WithEvents lblDelete As System.Windows.Forms.Label
    Friend WithEvents chkL1 As System.Windows.Forms.CheckBox
    Friend WithEvents chkL2 As System.Windows.Forms.CheckBox
    Friend WithEvents chkL3 As System.Windows.Forms.CheckBox
    Friend WithEvents chkL4 As System.Windows.Forms.CheckBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents btnNewInvoice As System.Windows.Forms.Button
    Friend WithEvents grpHeader As System.Windows.Forms.GroupBox
    Friend WithEvents dtpInvoiceDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblInvoiceDate As System.Windows.Forms.Label
    Friend WithEvents lblInvoiceNo As System.Windows.Forms.Label
    Friend WithEvents lblInvoiceNum As System.Windows.Forms.Label
    Friend WithEvents errInvoice As System.Windows.Forms.ErrorProvider
    Friend WithEvents grpFooter2 As System.Windows.Forms.GroupBox
    Friend WithEvents lblLabor As System.Windows.Forms.Label
    Friend WithEvents lblLaborCostTotal As System.Windows.Forms.Label
    Friend WithEvents lblMaterial As System.Windows.Forms.Label
    Friend WithEvents lblMaterialCostTotal As System.Windows.Forms.Label
    Friend WithEvents lblTotal As System.Windows.Forms.Label
    Friend WithEvents lblInvoiceTotal As System.Windows.Forms.Label
    Friend WithEvents btnAddNotes As System.Windows.Forms.Button
    Friend WithEvents lblCustomer As System.Windows.Forms.Label
    Friend WithEvents uCombo As Infragistics.Win.UltraWinEditors.UltraComboEditor
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEditInvoice))
        Me.grpLineItems = New System.Windows.Forms.GroupBox()
        Me.chkL4 = New System.Windows.Forms.CheckBox()
        Me.chkL3 = New System.Windows.Forms.CheckBox()
        Me.chkL2 = New System.Windows.Forms.CheckBox()
        Me.chkL1 = New System.Windows.Forms.CheckBox()
        Me.lblDelete = New System.Windows.Forms.Label()
        Me.chkL0 = New System.Windows.Forms.CheckBox()
        Me.lblLineTotal4 = New System.Windows.Forms.Label()
        Me.lblLineTotal2 = New System.Windows.Forms.Label()
        Me.txtMaterialCost2 = New System.Windows.Forms.TextBox()
        Me.txtMaterialDesc2 = New System.Windows.Forms.TextBox()
        Me.lblLineTotal1 = New System.Windows.Forms.Label()
        Me.txtMaterialCost1 = New System.Windows.Forms.TextBox()
        Me.txtMaterialDesc1 = New System.Windows.Forms.TextBox()
        Me.dtpLine4 = New System.Windows.Forms.DateTimePicker()
        Me.txtMaterialDesc4 = New System.Windows.Forms.TextBox()
        Me.txtLaborDesc4 = New System.Windows.Forms.TextBox()
        Me.dtpLine3 = New System.Windows.Forms.DateTimePicker()
        Me.dtpLine2 = New System.Windows.Forms.DateTimePicker()
        Me.dtpLine1 = New System.Windows.Forms.DateTimePicker()
        Me.lblDate = New System.Windows.Forms.Label()
        Me.dtpLine0 = New System.Windows.Forms.DateTimePicker()
        Me.lblLineTotal = New System.Windows.Forms.Label()
        Me.lblLineTotal0 = New System.Windows.Forms.Label()
        Me.lblMaterialCost = New System.Windows.Forms.Label()
        Me.lblMaterialDesc = New System.Windows.Forms.Label()
        Me.txtMaterialCost0 = New System.Windows.Forms.TextBox()
        Me.txtMaterialDesc0 = New System.Windows.Forms.TextBox()
        Me.lblLaborCost = New System.Windows.Forms.Label()
        Me.lblLaborDesc = New System.Windows.Forms.Label()
        Me.txtLaborCost3 = New System.Windows.Forms.TextBox()
        Me.txtLaborDesc3 = New System.Windows.Forms.TextBox()
        Me.txtLaborCost2 = New System.Windows.Forms.TextBox()
        Me.txtLaborDesc2 = New System.Windows.Forms.TextBox()
        Me.txtLaborCost1 = New System.Windows.Forms.TextBox()
        Me.txtLaborDesc1 = New System.Windows.Forms.TextBox()
        Me.txtLaborCost0 = New System.Windows.Forms.TextBox()
        Me.txtLaborDesc0 = New System.Windows.Forms.TextBox()
        Me.txtMaterialCost3 = New System.Windows.Forms.TextBox()
        Me.txtMaterialDesc3 = New System.Windows.Forms.TextBox()
        Me.lblLineTotal3 = New System.Windows.Forms.Label()
        Me.txtMaterialCost4 = New System.Windows.Forms.TextBox()
        Me.txtLaborCost4 = New System.Windows.Forms.TextBox()
        Me.grpFooter = New System.Windows.Forms.GroupBox()
        Me.btnAddNotes = New System.Windows.Forms.Button()
        Me.btnNewInvoice = New System.Windows.Forms.Button()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.BtnNewLine = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.lblLaborCostTotal = New System.Windows.Forms.Label()
        Me.lblMaterialCostTotal = New System.Windows.Forms.Label()
        Me.lblInvoiceTotal = New System.Windows.Forms.Label()
        Me.grpHeader = New System.Windows.Forms.GroupBox()
        Me.uCombo = New Infragistics.Win.UltraWinEditors.UltraComboEditor()
        Me.lblCustomer = New System.Windows.Forms.Label()
        Me.dtpInvoiceDate = New System.Windows.Forms.DateTimePicker()
        Me.lblInvoiceDate = New System.Windows.Forms.Label()
        Me.lblInvoiceNo = New System.Windows.Forms.Label()
        Me.lblInvoiceNum = New System.Windows.Forms.Label()
        Me.errInvoice = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.grpFooter2 = New System.Windows.Forms.GroupBox()
        Me.lblLabor = New System.Windows.Forms.Label()
        Me.lblMaterial = New System.Windows.Forms.Label()
        Me.lblTotal = New System.Windows.Forms.Label()
        Me.grpLineItems.SuspendLayout()
        Me.grpFooter.SuspendLayout()
        Me.grpHeader.SuspendLayout()
        CType(Me.uCombo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.errInvoice, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpFooter2.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpLineItems
        '
        Me.grpLineItems.BackColor = System.Drawing.Color.Transparent
        Me.grpLineItems.Controls.Add(Me.chkL4)
        Me.grpLineItems.Controls.Add(Me.chkL3)
        Me.grpLineItems.Controls.Add(Me.chkL2)
        Me.grpLineItems.Controls.Add(Me.chkL1)
        Me.grpLineItems.Controls.Add(Me.lblDelete)
        Me.grpLineItems.Controls.Add(Me.chkL0)
        Me.grpLineItems.Controls.Add(Me.lblLineTotal4)
        Me.grpLineItems.Controls.Add(Me.lblLineTotal2)
        Me.grpLineItems.Controls.Add(Me.txtMaterialCost2)
        Me.grpLineItems.Controls.Add(Me.txtMaterialDesc2)
        Me.grpLineItems.Controls.Add(Me.lblLineTotal1)
        Me.grpLineItems.Controls.Add(Me.txtMaterialCost1)
        Me.grpLineItems.Controls.Add(Me.txtMaterialDesc1)
        Me.grpLineItems.Controls.Add(Me.dtpLine4)
        Me.grpLineItems.Controls.Add(Me.txtMaterialDesc4)
        Me.grpLineItems.Controls.Add(Me.txtLaborDesc4)
        Me.grpLineItems.Controls.Add(Me.dtpLine3)
        Me.grpLineItems.Controls.Add(Me.dtpLine2)
        Me.grpLineItems.Controls.Add(Me.dtpLine1)
        Me.grpLineItems.Controls.Add(Me.lblDate)
        Me.grpLineItems.Controls.Add(Me.dtpLine0)
        Me.grpLineItems.Controls.Add(Me.lblLineTotal)
        Me.grpLineItems.Controls.Add(Me.lblLineTotal0)
        Me.grpLineItems.Controls.Add(Me.lblMaterialCost)
        Me.grpLineItems.Controls.Add(Me.lblMaterialDesc)
        Me.grpLineItems.Controls.Add(Me.txtMaterialCost0)
        Me.grpLineItems.Controls.Add(Me.txtMaterialDesc0)
        Me.grpLineItems.Controls.Add(Me.lblLaborCost)
        Me.grpLineItems.Controls.Add(Me.lblLaborDesc)
        Me.grpLineItems.Controls.Add(Me.txtLaborCost3)
        Me.grpLineItems.Controls.Add(Me.txtLaborDesc3)
        Me.grpLineItems.Controls.Add(Me.txtLaborCost2)
        Me.grpLineItems.Controls.Add(Me.txtLaborDesc2)
        Me.grpLineItems.Controls.Add(Me.txtLaborCost1)
        Me.grpLineItems.Controls.Add(Me.txtLaborDesc1)
        Me.grpLineItems.Controls.Add(Me.txtLaborCost0)
        Me.grpLineItems.Controls.Add(Me.txtLaborDesc0)
        Me.grpLineItems.Controls.Add(Me.txtMaterialCost3)
        Me.grpLineItems.Controls.Add(Me.txtMaterialDesc3)
        Me.grpLineItems.Controls.Add(Me.lblLineTotal3)
        Me.grpLineItems.Controls.Add(Me.txtMaterialCost4)
        Me.grpLineItems.Controls.Add(Me.txtLaborCost4)
        Me.grpLineItems.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpLineItems.Location = New System.Drawing.Point(10, 74)
        Me.grpLineItems.Name = "grpLineItems"
        Me.grpLineItems.Size = New System.Drawing.Size(869, 430)
        Me.grpLineItems.TabIndex = 2
        Me.grpLineItems.TabStop = False
        Me.grpLineItems.Text = "Line Items"
        '
        'chkL4
        '
        Me.chkL4.Enabled = False
        Me.chkL4.Location = New System.Drawing.Point(22, 353)
        Me.chkL4.Name = "chkL4"
        Me.chkL4.Size = New System.Drawing.Size(16, 21)
        Me.chkL4.TabIndex = 44
        Me.chkL4.Tag = "4"
        '
        'chkL3
        '
        Me.chkL3.Enabled = False
        Me.chkL3.Location = New System.Drawing.Point(22, 276)
        Me.chkL3.Name = "chkL3"
        Me.chkL3.Size = New System.Drawing.Size(16, 21)
        Me.chkL3.TabIndex = 43
        Me.chkL3.Tag = "3"
        '
        'chkL2
        '
        Me.chkL2.Enabled = False
        Me.chkL2.Location = New System.Drawing.Point(22, 199)
        Me.chkL2.Name = "chkL2"
        Me.chkL2.Size = New System.Drawing.Size(16, 21)
        Me.chkL2.TabIndex = 42
        Me.chkL2.Tag = "2"
        '
        'chkL1
        '
        Me.chkL1.Enabled = False
        Me.chkL1.Location = New System.Drawing.Point(22, 122)
        Me.chkL1.Name = "chkL1"
        Me.chkL1.Size = New System.Drawing.Size(16, 21)
        Me.chkL1.TabIndex = 41
        Me.chkL1.Tag = "1"
        '
        'lblDelete
        '
        Me.lblDelete.Location = New System.Drawing.Point(8, 24)
        Me.lblDelete.Name = "lblDelete"
        Me.lblDelete.Size = New System.Drawing.Size(48, 16)
        Me.lblDelete.TabIndex = 40
        Me.lblDelete.Text = "Delete:"
        '
        'chkL0
        '
        Me.chkL0.Location = New System.Drawing.Point(22, 46)
        Me.chkL0.Name = "chkL0"
        Me.chkL0.Size = New System.Drawing.Size(16, 21)
        Me.chkL0.TabIndex = 39
        Me.chkL0.Tag = "0"
        '
        'lblLineTotal4
        '
        Me.lblLineTotal4.BackColor = System.Drawing.Color.AliceBlue
        Me.lblLineTotal4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLineTotal4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLineTotal4.Location = New System.Drawing.Point(758, 353)
        Me.lblLineTotal4.Name = "lblLineTotal4"
        Me.lblLineTotal4.Size = New System.Drawing.Size(92, 20)
        Me.lblLineTotal4.TabIndex = 38
        Me.lblLineTotal4.Tag = "t"
        Me.lblLineTotal4.Text = "$0.00"
        Me.lblLineTotal4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblLineTotal2
        '
        Me.lblLineTotal2.BackColor = System.Drawing.Color.AliceBlue
        Me.lblLineTotal2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLineTotal2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLineTotal2.Location = New System.Drawing.Point(758, 199)
        Me.lblLineTotal2.Name = "lblLineTotal2"
        Me.lblLineTotal2.Size = New System.Drawing.Size(92, 20)
        Me.lblLineTotal2.TabIndex = 37
        Me.lblLineTotal2.Tag = "t"
        Me.lblLineTotal2.Text = "$0.00"
        Me.lblLineTotal2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtMaterialCost2
        '
        Me.txtMaterialCost2.Enabled = False
        Me.txtMaterialCost2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaterialCost2.Location = New System.Drawing.Point(663, 199)
        Me.txtMaterialCost2.Name = "txtMaterialCost2"
        Me.txtMaterialCost2.Size = New System.Drawing.Size(65, 21)
        Me.txtMaterialCost2.TabIndex = 16
        Me.txtMaterialCost2.Tag = "2"
        '
        'txtMaterialDesc2
        '
        Me.txtMaterialDesc2.Enabled = False
        Me.txtMaterialDesc2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaterialDesc2.Location = New System.Drawing.Point(455, 199)
        Me.txtMaterialDesc2.MaxLength = 256
        Me.txtMaterialDesc2.Multiline = True
        Me.txtMaterialDesc2.Name = "txtMaterialDesc2"
        Me.txtMaterialDesc2.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtMaterialDesc2.Size = New System.Drawing.Size(198, 58)
        Me.txtMaterialDesc2.TabIndex = 15
        Me.txtMaterialDesc2.Tag = "2"
        '
        'lblLineTotal1
        '
        Me.lblLineTotal1.BackColor = System.Drawing.Color.AliceBlue
        Me.lblLineTotal1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLineTotal1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLineTotal1.Location = New System.Drawing.Point(758, 122)
        Me.lblLineTotal1.Name = "lblLineTotal1"
        Me.lblLineTotal1.Size = New System.Drawing.Size(92, 20)
        Me.lblLineTotal1.TabIndex = 34
        Me.lblLineTotal1.Tag = "t"
        Me.lblLineTotal1.Text = "$0.00"
        Me.lblLineTotal1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtMaterialCost1
        '
        Me.txtMaterialCost1.Enabled = False
        Me.txtMaterialCost1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaterialCost1.Location = New System.Drawing.Point(663, 122)
        Me.txtMaterialCost1.Name = "txtMaterialCost1"
        Me.txtMaterialCost1.Size = New System.Drawing.Size(65, 21)
        Me.txtMaterialCost1.TabIndex = 11
        Me.txtMaterialCost1.Tag = "1"
        '
        'txtMaterialDesc1
        '
        Me.txtMaterialDesc1.Enabled = False
        Me.txtMaterialDesc1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaterialDesc1.Location = New System.Drawing.Point(455, 122)
        Me.txtMaterialDesc1.MaxLength = 256
        Me.txtMaterialDesc1.Multiline = True
        Me.txtMaterialDesc1.Name = "txtMaterialDesc1"
        Me.txtMaterialDesc1.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtMaterialDesc1.Size = New System.Drawing.Size(198, 58)
        Me.txtMaterialDesc1.TabIndex = 10
        Me.txtMaterialDesc1.Tag = "1"
        '
        'dtpLine4
        '
        Me.dtpLine4.Enabled = False
        Me.dtpLine4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpLine4.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpLine4.Location = New System.Drawing.Point(52, 353)
        Me.dtpLine4.Name = "dtpLine4"
        Me.dtpLine4.Size = New System.Drawing.Size(96, 21)
        Me.dtpLine4.TabIndex = 22
        Me.dtpLine4.Tag = "4"
        '
        'txtMaterialDesc4
        '
        Me.txtMaterialDesc4.Enabled = False
        Me.txtMaterialDesc4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaterialDesc4.Location = New System.Drawing.Point(455, 353)
        Me.txtMaterialDesc4.MaxLength = 256
        Me.txtMaterialDesc4.Multiline = True
        Me.txtMaterialDesc4.Name = "txtMaterialDesc4"
        Me.txtMaterialDesc4.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtMaterialDesc4.Size = New System.Drawing.Size(198, 58)
        Me.txtMaterialDesc4.TabIndex = 25
        Me.txtMaterialDesc4.Tag = "4"
        '
        'txtLaborDesc4
        '
        Me.txtLaborDesc4.Enabled = False
        Me.txtLaborDesc4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLaborDesc4.Location = New System.Drawing.Point(166, 353)
        Me.txtLaborDesc4.MaxLength = 256
        Me.txtLaborDesc4.Multiline = True
        Me.txtLaborDesc4.Name = "txtLaborDesc4"
        Me.txtLaborDesc4.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtLaborDesc4.Size = New System.Drawing.Size(204, 58)
        Me.txtLaborDesc4.TabIndex = 23
        Me.txtLaborDesc4.Tag = "4"
        '
        'dtpLine3
        '
        Me.dtpLine3.Enabled = False
        Me.dtpLine3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpLine3.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpLine3.Location = New System.Drawing.Point(52, 276)
        Me.dtpLine3.Name = "dtpLine3"
        Me.dtpLine3.Size = New System.Drawing.Size(96, 21)
        Me.dtpLine3.TabIndex = 17
        Me.dtpLine3.Tag = "3"
        '
        'dtpLine2
        '
        Me.dtpLine2.Enabled = False
        Me.dtpLine2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpLine2.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpLine2.Location = New System.Drawing.Point(52, 199)
        Me.dtpLine2.Name = "dtpLine2"
        Me.dtpLine2.Size = New System.Drawing.Size(96, 21)
        Me.dtpLine2.TabIndex = 12
        Me.dtpLine2.Tag = "2"
        '
        'dtpLine1
        '
        Me.dtpLine1.Enabled = False
        Me.dtpLine1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpLine1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpLine1.Location = New System.Drawing.Point(52, 122)
        Me.dtpLine1.Name = "dtpLine1"
        Me.dtpLine1.Size = New System.Drawing.Size(96, 21)
        Me.dtpLine1.TabIndex = 7
        Me.dtpLine1.Tag = "1"
        '
        'lblDate
        '
        Me.lblDate.Location = New System.Drawing.Point(75, 24)
        Me.lblDate.Name = "lblDate"
        Me.lblDate.Size = New System.Drawing.Size(39, 16)
        Me.lblDate.TabIndex = 22
        Me.lblDate.Text = "Date:"
        '
        'dtpLine0
        '
        Me.dtpLine0.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpLine0.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpLine0.Location = New System.Drawing.Point(52, 46)
        Me.dtpLine0.Name = "dtpLine0"
        Me.dtpLine0.Size = New System.Drawing.Size(96, 21)
        Me.dtpLine0.TabIndex = 2
        Me.dtpLine0.Tag = "0"
        '
        'lblLineTotal
        '
        Me.lblLineTotal.Location = New System.Drawing.Point(786, 24)
        Me.lblLineTotal.Name = "lblLineTotal"
        Me.lblLineTotal.Size = New System.Drawing.Size(72, 16)
        Me.lblLineTotal.TabIndex = 20
        Me.lblLineTotal.Text = "Line Total:"
        '
        'lblLineTotal0
        '
        Me.lblLineTotal0.BackColor = System.Drawing.Color.AliceBlue
        Me.lblLineTotal0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLineTotal0.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLineTotal0.Location = New System.Drawing.Point(758, 46)
        Me.lblLineTotal0.Name = "lblLineTotal0"
        Me.lblLineTotal0.Size = New System.Drawing.Size(92, 20)
        Me.lblLineTotal0.TabIndex = 19
        Me.lblLineTotal0.Tag = "t"
        Me.lblLineTotal0.Text = "$0.00"
        Me.lblLineTotal0.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblMaterialCost
        '
        Me.lblMaterialCost.Location = New System.Drawing.Point(663, 24)
        Me.lblMaterialCost.Name = "lblMaterialCost"
        Me.lblMaterialCost.Size = New System.Drawing.Size(89, 16)
        Me.lblMaterialCost.TabIndex = 18
        Me.lblMaterialCost.Text = "Material Cost:"
        '
        'lblMaterialDesc
        '
        Me.lblMaterialDesc.Location = New System.Drawing.Point(482, 24)
        Me.lblMaterialDesc.Name = "lblMaterialDesc"
        Me.lblMaterialDesc.Size = New System.Drawing.Size(144, 16)
        Me.lblMaterialDesc.TabIndex = 17
        Me.lblMaterialDesc.Text = "Description of Materials:"
        '
        'txtMaterialCost0
        '
        Me.txtMaterialCost0.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaterialCost0.Location = New System.Drawing.Point(663, 46)
        Me.txtMaterialCost0.Name = "txtMaterialCost0"
        Me.txtMaterialCost0.Size = New System.Drawing.Size(65, 21)
        Me.txtMaterialCost0.TabIndex = 6
        Me.txtMaterialCost0.Tag = "0"
        '
        'txtMaterialDesc0
        '
        Me.txtMaterialDesc0.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaterialDesc0.Location = New System.Drawing.Point(455, 46)
        Me.txtMaterialDesc0.MaxLength = 256
        Me.txtMaterialDesc0.Multiline = True
        Me.txtMaterialDesc0.Name = "txtMaterialDesc0"
        Me.txtMaterialDesc0.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtMaterialDesc0.Size = New System.Drawing.Size(198, 58)
        Me.txtMaterialDesc0.TabIndex = 5
        Me.txtMaterialDesc0.Tag = "0"
        '
        'lblLaborCost
        '
        Me.lblLaborCost.Location = New System.Drawing.Point(380, 24)
        Me.lblLaborCost.Name = "lblLaborCost"
        Me.lblLaborCost.Size = New System.Drawing.Size(82, 16)
        Me.lblLaborCost.TabIndex = 13
        Me.lblLaborCost.Text = "Labor Cost:"
        '
        'lblLaborDesc
        '
        Me.lblLaborDesc.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLaborDesc.Location = New System.Drawing.Point(197, 24)
        Me.lblLaborDesc.Name = "lblLaborDesc"
        Me.lblLaborDesc.Size = New System.Drawing.Size(135, 16)
        Me.lblLaborDesc.TabIndex = 12
        Me.lblLaborDesc.Text = "Description of Labor:"
        '
        'txtLaborCost3
        '
        Me.txtLaborCost3.Enabled = False
        Me.txtLaborCost3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLaborCost3.Location = New System.Drawing.Point(380, 276)
        Me.txtLaborCost3.Name = "txtLaborCost3"
        Me.txtLaborCost3.Size = New System.Drawing.Size(65, 21)
        Me.txtLaborCost3.TabIndex = 19
        Me.txtLaborCost3.Tag = "3"
        '
        'txtLaborDesc3
        '
        Me.txtLaborDesc3.Enabled = False
        Me.txtLaborDesc3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLaborDesc3.Location = New System.Drawing.Point(166, 276)
        Me.txtLaborDesc3.MaxLength = 256
        Me.txtLaborDesc3.Multiline = True
        Me.txtLaborDesc3.Name = "txtLaborDesc3"
        Me.txtLaborDesc3.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtLaborDesc3.Size = New System.Drawing.Size(204, 58)
        Me.txtLaborDesc3.TabIndex = 18
        Me.txtLaborDesc3.Tag = "3"
        '
        'txtLaborCost2
        '
        Me.txtLaborCost2.Enabled = False
        Me.txtLaborCost2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLaborCost2.Location = New System.Drawing.Point(380, 199)
        Me.txtLaborCost2.Name = "txtLaborCost2"
        Me.txtLaborCost2.Size = New System.Drawing.Size(65, 21)
        Me.txtLaborCost2.TabIndex = 14
        Me.txtLaborCost2.Tag = "2"
        '
        'txtLaborDesc2
        '
        Me.txtLaborDesc2.Enabled = False
        Me.txtLaborDesc2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLaborDesc2.Location = New System.Drawing.Point(166, 199)
        Me.txtLaborDesc2.MaxLength = 256
        Me.txtLaborDesc2.Multiline = True
        Me.txtLaborDesc2.Name = "txtLaborDesc2"
        Me.txtLaborDesc2.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtLaborDesc2.Size = New System.Drawing.Size(204, 58)
        Me.txtLaborDesc2.TabIndex = 13
        Me.txtLaborDesc2.Tag = "2"
        '
        'txtLaborCost1
        '
        Me.txtLaborCost1.Enabled = False
        Me.txtLaborCost1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLaborCost1.Location = New System.Drawing.Point(380, 122)
        Me.txtLaborCost1.Name = "txtLaborCost1"
        Me.txtLaborCost1.Size = New System.Drawing.Size(65, 21)
        Me.txtLaborCost1.TabIndex = 9
        Me.txtLaborCost1.Tag = "1"
        '
        'txtLaborDesc1
        '
        Me.txtLaborDesc1.Enabled = False
        Me.txtLaborDesc1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLaborDesc1.Location = New System.Drawing.Point(166, 122)
        Me.txtLaborDesc1.MaxLength = 256
        Me.txtLaborDesc1.Multiline = True
        Me.txtLaborDesc1.Name = "txtLaborDesc1"
        Me.txtLaborDesc1.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtLaborDesc1.Size = New System.Drawing.Size(204, 58)
        Me.txtLaborDesc1.TabIndex = 8
        Me.txtLaborDesc1.Tag = "1"
        '
        'txtLaborCost0
        '
        Me.txtLaborCost0.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLaborCost0.Location = New System.Drawing.Point(380, 45)
        Me.txtLaborCost0.Name = "txtLaborCost0"
        Me.txtLaborCost0.Size = New System.Drawing.Size(65, 21)
        Me.txtLaborCost0.TabIndex = 4
        Me.txtLaborCost0.Tag = "0"
        '
        'txtLaborDesc0
        '
        Me.txtLaborDesc0.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLaborDesc0.Location = New System.Drawing.Point(166, 45)
        Me.txtLaborDesc0.MaxLength = 256
        Me.txtLaborDesc0.Multiline = True
        Me.txtLaborDesc0.Name = "txtLaborDesc0"
        Me.txtLaborDesc0.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtLaborDesc0.Size = New System.Drawing.Size(204, 58)
        Me.txtLaborDesc0.TabIndex = 3
        Me.txtLaborDesc0.Tag = "0"
        '
        'txtMaterialCost3
        '
        Me.txtMaterialCost3.Enabled = False
        Me.txtMaterialCost3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaterialCost3.Location = New System.Drawing.Point(663, 276)
        Me.txtMaterialCost3.Name = "txtMaterialCost3"
        Me.txtMaterialCost3.Size = New System.Drawing.Size(65, 21)
        Me.txtMaterialCost3.TabIndex = 21
        Me.txtMaterialCost3.Tag = "3"
        '
        'txtMaterialDesc3
        '
        Me.txtMaterialDesc3.Enabled = False
        Me.txtMaterialDesc3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaterialDesc3.Location = New System.Drawing.Point(455, 276)
        Me.txtMaterialDesc3.MaxLength = 256
        Me.txtMaterialDesc3.Multiline = True
        Me.txtMaterialDesc3.Name = "txtMaterialDesc3"
        Me.txtMaterialDesc3.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtMaterialDesc3.Size = New System.Drawing.Size(198, 58)
        Me.txtMaterialDesc3.TabIndex = 20
        Me.txtMaterialDesc3.Tag = "3"
        '
        'lblLineTotal3
        '
        Me.lblLineTotal3.BackColor = System.Drawing.Color.AliceBlue
        Me.lblLineTotal3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLineTotal3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLineTotal3.Location = New System.Drawing.Point(758, 276)
        Me.lblLineTotal3.Name = "lblLineTotal3"
        Me.lblLineTotal3.Size = New System.Drawing.Size(92, 20)
        Me.lblLineTotal3.TabIndex = 37
        Me.lblLineTotal3.Tag = "t"
        Me.lblLineTotal3.Text = "$0.00"
        Me.lblLineTotal3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtMaterialCost4
        '
        Me.txtMaterialCost4.Enabled = False
        Me.txtMaterialCost4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaterialCost4.Location = New System.Drawing.Point(663, 353)
        Me.txtMaterialCost4.Name = "txtMaterialCost4"
        Me.txtMaterialCost4.Size = New System.Drawing.Size(65, 21)
        Me.txtMaterialCost4.TabIndex = 26
        Me.txtMaterialCost4.Tag = "4"
        '
        'txtLaborCost4
        '
        Me.txtLaborCost4.Enabled = False
        Me.txtLaborCost4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLaborCost4.Location = New System.Drawing.Point(380, 353)
        Me.txtLaborCost4.Name = "txtLaborCost4"
        Me.txtLaborCost4.Size = New System.Drawing.Size(65, 21)
        Me.txtLaborCost4.TabIndex = 24
        Me.txtLaborCost4.Tag = "4"
        '
        'grpFooter
        '
        Me.grpFooter.BackColor = System.Drawing.Color.Transparent
        Me.grpFooter.Controls.Add(Me.btnAddNotes)
        Me.grpFooter.Controls.Add(Me.btnNewInvoice)
        Me.grpFooter.Controls.Add(Me.btnPrint)
        Me.grpFooter.Controls.Add(Me.btnUpdate)
        Me.grpFooter.Controls.Add(Me.btnCancel)
        Me.grpFooter.Controls.Add(Me.BtnNewLine)
        Me.grpFooter.Location = New System.Drawing.Point(11, 508)
        Me.grpFooter.Name = "grpFooter"
        Me.grpFooter.Size = New System.Drawing.Size(541, 56)
        Me.grpFooter.TabIndex = 21
        Me.grpFooter.TabStop = False
        Me.grpFooter.Text = "Action Buttons"
        '
        'btnAddNotes
        '
        Me.btnAddNotes.BackColor = System.Drawing.Color.AliceBlue
        Me.btnAddNotes.Enabled = False
        Me.btnAddNotes.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddNotes.Image = CType(resources.GetObject("btnAddNotes.Image"), System.Drawing.Image)
        Me.btnAddNotes.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnAddNotes.Location = New System.Drawing.Point(445, 21)
        Me.btnAddNotes.Name = "btnAddNotes"
        Me.btnAddNotes.Size = New System.Drawing.Size(88, 24)
        Me.btnAddNotes.TabIndex = 47
        Me.btnAddNotes.Text = "&Notes"
        Me.btnAddNotes.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.btnAddNotes, "Notes Available After Invoice has been saved")
        Me.btnAddNotes.UseVisualStyleBackColor = False
        '
        'btnNewInvoice
        '
        Me.btnNewInvoice.BackColor = System.Drawing.Color.AliceBlue
        Me.btnNewInvoice.Enabled = False
        Me.btnNewInvoice.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNewInvoice.Image = CType(resources.GetObject("btnNewInvoice.Image"), System.Drawing.Image)
        Me.btnNewInvoice.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnNewInvoice.Location = New System.Drawing.Point(357, 21)
        Me.btnNewInvoice.Name = "btnNewInvoice"
        Me.btnNewInvoice.Size = New System.Drawing.Size(88, 24)
        Me.btnNewInvoice.TabIndex = 46
        Me.btnNewInvoice.Text = "N&ew Inv."
        Me.btnNewInvoice.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.btnNewInvoice, "Add New Invoice")
        Me.btnNewInvoice.UseVisualStyleBackColor = False
        '
        'btnPrint
        '
        Me.btnPrint.BackColor = System.Drawing.Color.AliceBlue
        Me.btnPrint.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrint.Image = CType(resources.GetObject("btnPrint.Image"), System.Drawing.Image)
        Me.btnPrint.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnPrint.Location = New System.Drawing.Point(272, 21)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(85, 24)
        Me.btnPrint.TabIndex = 30
        Me.btnPrint.Text = "&Print"
        Me.ToolTip1.SetToolTip(Me.btnPrint, "Prints Invoice")
        Me.btnPrint.UseVisualStyleBackColor = False
        '
        'btnUpdate
        '
        Me.btnUpdate.BackColor = System.Drawing.Color.AliceBlue
        Me.btnUpdate.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpdate.Image = CType(resources.GetObject("btnUpdate.Image"), System.Drawing.Image)
        Me.btnUpdate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnUpdate.Location = New System.Drawing.Point(185, 21)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(87, 24)
        Me.btnUpdate.TabIndex = 29
        Me.btnUpdate.Text = "  &Save"
        Me.ToolTip1.SetToolTip(Me.btnUpdate, "Saves Changes to the Database")
        Me.btnUpdate.UseVisualStyleBackColor = False
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.Color.AliceBlue
        Me.btnCancel.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.Image = CType(resources.GetObject("btnCancel.Image"), System.Drawing.Image)
        Me.btnCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCancel.Location = New System.Drawing.Point(95, 21)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(90, 24)
        Me.btnCancel.TabIndex = 28
        Me.btnCancel.Text = "E&xit"
        Me.ToolTip1.SetToolTip(Me.btnCancel, "Cancels Last Changes Before Update")
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'BtnNewLine
        '
        Me.BtnNewLine.BackColor = System.Drawing.Color.AliceBlue
        Me.BtnNewLine.Enabled = False
        Me.BtnNewLine.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnNewLine.Image = CType(resources.GetObject("BtnNewLine.Image"), System.Drawing.Image)
        Me.BtnNewLine.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnNewLine.Location = New System.Drawing.Point(5, 21)
        Me.BtnNewLine.Name = "BtnNewLine"
        Me.BtnNewLine.Size = New System.Drawing.Size(90, 24)
        Me.BtnNewLine.TabIndex = 27
        Me.BtnNewLine.Text = "&New Item"
        Me.BtnNewLine.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.BtnNewLine, "Add New Item")
        Me.BtnNewLine.UseVisualStyleBackColor = False
        '
        'lblLaborCostTotal
        '
        Me.lblLaborCostTotal.BackColor = System.Drawing.Color.AliceBlue
        Me.lblLaborCostTotal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLaborCostTotal.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLaborCostTotal.Location = New System.Drawing.Point(10, 24)
        Me.lblLaborCostTotal.Name = "lblLaborCostTotal"
        Me.lblLaborCostTotal.Size = New System.Drawing.Size(85, 20)
        Me.lblLaborCostTotal.TabIndex = 49
        Me.lblLaborCostTotal.Text = "$0.00"
        Me.lblLaborCostTotal.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.lblLaborCostTotal, "Total Labor For This Invoice")
        '
        'lblMaterialCostTotal
        '
        Me.lblMaterialCostTotal.BackColor = System.Drawing.Color.AliceBlue
        Me.lblMaterialCostTotal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblMaterialCostTotal.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMaterialCostTotal.Location = New System.Drawing.Point(119, 24)
        Me.lblMaterialCostTotal.Name = "lblMaterialCostTotal"
        Me.lblMaterialCostTotal.Size = New System.Drawing.Size(85, 20)
        Me.lblMaterialCostTotal.TabIndex = 47
        Me.lblMaterialCostTotal.Text = "$0.00"
        Me.lblMaterialCostTotal.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.lblMaterialCostTotal, "Total Materials This Invoice")
        '
        'lblInvoiceTotal
        '
        Me.lblInvoiceTotal.BackColor = System.Drawing.Color.AliceBlue
        Me.lblInvoiceTotal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblInvoiceTotal.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInvoiceTotal.ForeColor = System.Drawing.Color.Red
        Me.lblInvoiceTotal.Location = New System.Drawing.Point(228, 24)
        Me.lblInvoiceTotal.Name = "lblInvoiceTotal"
        Me.lblInvoiceTotal.Size = New System.Drawing.Size(85, 20)
        Me.lblInvoiceTotal.TabIndex = 45
        Me.lblInvoiceTotal.Text = "$0.00"
        Me.lblInvoiceTotal.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.lblInvoiceTotal, "Grand Total This Invoice")
        '
        'grpHeader
        '
        Me.grpHeader.BackColor = System.Drawing.Color.Transparent
        Me.grpHeader.Controls.Add(Me.uCombo)
        Me.grpHeader.Controls.Add(Me.lblCustomer)
        Me.grpHeader.Controls.Add(Me.dtpInvoiceDate)
        Me.grpHeader.Controls.Add(Me.lblInvoiceDate)
        Me.grpHeader.Controls.Add(Me.lblInvoiceNo)
        Me.grpHeader.Controls.Add(Me.lblInvoiceNum)
        Me.grpHeader.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpHeader.Location = New System.Drawing.Point(10, 8)
        Me.grpHeader.Name = "grpHeader"
        Me.grpHeader.Size = New System.Drawing.Size(869, 64)
        Me.grpHeader.TabIndex = 28
        Me.grpHeader.TabStop = False
        '
        'uCombo
        '
        Me.uCombo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.uCombo.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList
        Me.uCombo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.uCombo.Location = New System.Drawing.Point(384, 26)
        Me.uCombo.Name = "uCombo"
        Me.uCombo.Size = New System.Drawing.Size(219, 24)
        Me.uCombo.TabIndex = 35
        '
        'lblCustomer
        '
        Me.lblCustomer.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCustomer.Location = New System.Drawing.Point(306, 26)
        Me.lblCustomer.Name = "lblCustomer"
        Me.lblCustomer.Size = New System.Drawing.Size(72, 16)
        Me.lblCustomer.TabIndex = 34
        Me.lblCustomer.Text = "Customer:"
        '
        'dtpInvoiceDate
        '
        Me.dtpInvoiceDate.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpInvoiceDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpInvoiceDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpInvoiceDate.Location = New System.Drawing.Point(104, 26)
        Me.dtpInvoiceDate.Name = "dtpInvoiceDate"
        Me.dtpInvoiceDate.Size = New System.Drawing.Size(104, 22)
        Me.dtpInvoiceDate.TabIndex = 28
        '
        'lblInvoiceDate
        '
        Me.lblInvoiceDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInvoiceDate.Location = New System.Drawing.Point(8, 26)
        Me.lblInvoiceDate.Name = "lblInvoiceDate"
        Me.lblInvoiceDate.Size = New System.Drawing.Size(96, 16)
        Me.lblInvoiceDate.TabIndex = 33
        Me.lblInvoiceDate.Text = "Invoice Date:"
        '
        'lblInvoiceNo
        '
        Me.lblInvoiceNo.BackColor = System.Drawing.Color.AliceBlue
        Me.lblInvoiceNo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblInvoiceNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInvoiceNo.Location = New System.Drawing.Point(776, 26)
        Me.lblInvoiceNo.Name = "lblInvoiceNo"
        Me.lblInvoiceNo.Size = New System.Drawing.Size(72, 22)
        Me.lblInvoiceNo.TabIndex = 32
        Me.lblInvoiceNo.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblInvoiceNum
        '
        Me.lblInvoiceNum.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInvoiceNum.Location = New System.Drawing.Point(684, 26)
        Me.lblInvoiceNum.Name = "lblInvoiceNum"
        Me.lblInvoiceNum.Size = New System.Drawing.Size(80, 16)
        Me.lblInvoiceNum.TabIndex = 31
        Me.lblInvoiceNum.Text = "Invoice No:"
        '
        'errInvoice
        '
        Me.errInvoice.ContainerControl = Me
        '
        'grpFooter2
        '
        Me.grpFooter2.BackColor = System.Drawing.Color.Transparent
        Me.grpFooter2.Controls.Add(Me.lblLabor)
        Me.grpFooter2.Controls.Add(Me.lblLaborCostTotal)
        Me.grpFooter2.Controls.Add(Me.lblMaterial)
        Me.grpFooter2.Controls.Add(Me.lblMaterialCostTotal)
        Me.grpFooter2.Controls.Add(Me.lblTotal)
        Me.grpFooter2.Controls.Add(Me.lblInvoiceTotal)
        Me.grpFooter2.Location = New System.Drawing.Point(555, 508)
        Me.grpFooter2.Name = "grpFooter2"
        Me.grpFooter2.Size = New System.Drawing.Size(324, 56)
        Me.grpFooter2.TabIndex = 29
        Me.grpFooter2.TabStop = False
        '
        'lblLabor
        '
        Me.lblLabor.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLabor.Location = New System.Drawing.Point(28, 8)
        Me.lblLabor.Name = "lblLabor"
        Me.lblLabor.Size = New System.Drawing.Size(48, 16)
        Me.lblLabor.TabIndex = 50
        Me.lblLabor.Text = "Labor:"
        '
        'lblMaterial
        '
        Me.lblMaterial.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMaterial.Location = New System.Drawing.Point(131, 8)
        Me.lblMaterial.Name = "lblMaterial"
        Me.lblMaterial.Size = New System.Drawing.Size(61, 16)
        Me.lblMaterial.TabIndex = 48
        Me.lblMaterial.Text = "Materials:"
        '
        'lblTotal
        '
        Me.lblTotal.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotal.Location = New System.Drawing.Point(250, 8)
        Me.lblTotal.Name = "lblTotal"
        Me.lblTotal.Size = New System.Drawing.Size(41, 16)
        Me.lblTotal.TabIndex = 46
        Me.lblTotal.Text = "Total:"
        '
        'frmEditInvoice
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.BackColor = System.Drawing.Color.LightSteelBlue
        Me.ClientSize = New System.Drawing.Size(882, 576)
        Me.Controls.Add(Me.grpFooter2)
        Me.Controls.Add(Me.grpHeader)
        Me.Controls.Add(Me.grpFooter)
        Me.Controls.Add(Me.grpLineItems)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmEditInvoice"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Edit/View Invoice"
        Me.grpLineItems.ResumeLayout(False)
        Me.grpLineItems.PerformLayout()
        Me.grpFooter.ResumeLayout(False)
        Me.grpHeader.ResumeLayout(False)
        Me.grpHeader.PerformLayout()
        CType(Me.uCombo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.errInvoice, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpFooter2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cboCustomers_SelectionChangeCommitted(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try
            If Me.Loading = True Then Exit Sub

            If Me.FormMode = Mode.AddNew Then
                If Me.grpLineItems.Enabled = False Then
                    Me.grpLineItems.Enabled = True
                    Me.EnableLine(0, True)
                End If
            ElseIf Me.FormMode = Mode.Edit Then
                Me.Check_All_Lines_For_Validity()
            End If

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub LaborDesc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLaborDesc0.TextChanged, _
                                        txtLaborDesc1.TextChanged, _
                                        txtLaborDesc2.TextChanged, _
                                        txtLaborDesc3.TextChanged, _
                                        txtLaborDesc4.TextChanged


        If Me.Loading = True Then Exit Sub

        Try
            'This should never happen, but could if program crashed
            If Me._Invoice.LineItems.Count = 0 Then
                Me._LineItem0 = New LineItem
                Me._Invoice.LineItems.Add(Me._LineItem0)
            End If

            Dim txtBox As TextBox = CType(sender, TextBox)
            Me._Invoice.LineItems(Convert.ToInt32(txtBox.Tag)).LaborDesc = txtBox.Text.Trim

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub MaterialDesc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMaterialDesc0.TextChanged, _
                                        txtMaterialDesc1.TextChanged, _
                                        txtMaterialDesc2.TextChanged, _
                                        txtMaterialDesc3.TextChanged, _
                                        txtMaterialDesc4.TextChanged
        Try

            If Me.Loading = True Then Exit Sub

            Dim txtBox As TextBox = CType(sender, TextBox)
            Me._Invoice.LineItems(Convert.ToInt32(txtBox.Tag)).MaterialDesc = txtBox.Text.Trim

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


    End Sub

    Private Sub txtLaborCost0_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLaborCost0.TextChanged
        If Me.Loading = True Then Exit Sub

        Me._Invoice.LineItems(0).LaborCost = Me.txtLaborCost0.Text

    End Sub

    Private Sub txtMaterialCost0_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMaterialCost0.TextChanged
        If Me.Loading = True Then Exit Sub

        Me._Invoice.LineItems(0).MaterialCost = Me.txtMaterialCost0.Text


    End Sub

    Private Sub txtMaterialCost0_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaterialCost0.KeyPress, _
                                            txtMaterialCost1.KeyPress, _
                                            txtMaterialCost2.KeyPress, _
                                            txtMaterialCost3.KeyPress, _
                                            txtMaterialCost4.KeyPress, _
                                            txtLaborCost0.KeyPress, _
                                            txtLaborCost1.KeyPress, _
                                            txtLaborCost2.KeyPress, _
                                            txtLaborCost3.KeyPress, _
                                            txtLaborCost4.KeyPress




        If Me.Loading = True Then Exit Sub

        Dim txtBox As TextBox = CType(sender, TextBox)

        e.Handled = Utilities.NumericOnly(e, txtBox)

    End Sub

    Private Sub BtnNewLine_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnNewLine.Click
        Try
            Me.btnUpdate.Enabled = False
            Me.BtnNewLine.Enabled = False


            With Me._Invoice
                Select Case .LineItems.Count

                    Case 1
                        If .LineItems(0).IsValid Then
                            Me._LineItem1 = New LineItem
                            .LineItems.Add(_LineItem1)
                            .LineItems(1).IsValid = False
                            Me.EnableLine(1, True)
                            Me._LineItem1.IsValid = False
                        End If

                    Case 2

                        If .LineItems(1).IsValid Then
                            Me._LineItem2 = New LineItem
                            .LineItems.Add(_LineItem2)
                            .LineItems(2).IsValid = False
                            Me.EnableLine(2, True)
                            Me._LineItem2.IsValid = False
                        End If

                    Case 3
                        If .LineItems(2).IsValid Then
                            Me._LineItem3 = New LineItem
                            .LineItems.Add(_LineItem3)
                            .LineItems(3).IsValid = False
                            Me.EnableLine(3, True)
                            Me._LineItem3.IsValid = False
                        End If


                    Case 4
                        If .LineItems(3).IsValid Then
                            Me._LineItem4 = New LineItem
                            .LineItems.Add(_LineItem4)
                            .LineItems(4).IsValid = False
                            Me.EnableLine(4, True)
                        End If

                End Select

            End With


        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
#Region " Event Sinks "

    Private Sub _LineItem1_LaborCostChanged() Handles _LineItem1.LaborCostChanged
        If Me.Loading = True Then Exit Sub

        With Me._Invoice
            Me.lblLineTotal1.Text = .LineItems(1).Total.ToString("c")
            Me.lblInvoiceTotal.Text = .InvoiceTotal.ToString("c")
            Me.lblLaborCostTotal.Text = .LaborTotal.ToString("c")
            Me.lblMaterialCostTotal.Text = .MaterialTotal.ToString("c")
        End With

    End Sub
    Private Sub _LineItem2_LaborCostChanged() Handles _LineItem2.LaborCostChanged

        If Me.Loading = True Then Exit Sub

        With Me._Invoice
            Me.lblLineTotal2.Text = .LineItems(2).Total.ToString("c")
            Me.lblInvoiceTotal.Text = .InvoiceTotal.ToString("c")
            Me.lblLaborCostTotal.Text = .LaborTotal.ToString("c")
            Me.lblLaborCostTotal.Text = .LaborTotal.ToString("c")

        End With

    End Sub


    Private Sub txtLaborCost1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLaborCost1.TextChanged
        If Me.Loading = True Then Exit Sub

        Me._Invoice.LineItems(1).LaborCost = Me.txtLaborCost1.Text

    End Sub

    Private Sub _LineItem0_LaborCostChanged() Handles _LineItem0.LaborCostChanged
        If Me.Loading = True Then Exit Sub

        With Me._Invoice
            Me.lblLineTotal0.Text = .LineItems(0).Total.ToString("c")
            Me.lblInvoiceTotal.Text = .InvoiceTotal.ToString("c")
            Me.lblLaborCostTotal.Text = .LaborTotal.ToString("c")
        End With

    End Sub

    Private Sub _LineItem0_MaterialCostChanged() Handles _LineItem0.MaterialCostChanged
        If Me.Loading = True Then Exit Sub

        With Me._Invoice
            Me.lblLineTotal0.Text = .LineItems(0).Total.ToString("c")
            Me.lblInvoiceTotal.Text = .InvoiceTotal.ToString("c")
            Me.lblMaterialCostTotal.Text = .MaterialTotal.ToString("c")
        End With

    End Sub

    Private Sub txtMaterialCost1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMaterialCost1.TextChanged
        If Me.Loading = True Then Exit Sub

        Me._Invoice.LineItems(1).MaterialCost = Me.txtMaterialCost1.Text

    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click

        Try
            Cursor = Cursors.WaitCursor
            Utilities.PrintInvoice(Me._Invoice)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            Cursor = Cursors.Default

        End Try

    End Sub
    Private Sub _LineItem0_Valid(ByVal IsValid As Boolean) Handles _LineItem0.Valid

        If Me.Loading = True Then Exit Sub

        If Me._Invoice.LineItems.Count < 5 Then
            Me.BtnNewLine.Enabled = IsValid
        Else
            Me.BtnNewLine.Enabled = False
        End If

        If Me._LineItem0.IsValid = False And Me.Loading = False Then
            Me.errInvoice.SetError(Me.lblLineTotal0, "Line Not Valid, check values")
        Else
            Me.errInvoice.SetError(Me.lblLineTotal0, "")

        End If

        Check_All_Lines_For_Validity()

    End Sub
    Private Sub _LineItem1_Valid(ByVal IsValid As Boolean) Handles _LineItem1.Valid

        If Me.Loading = True Then Exit Sub

        If Me._Invoice.LineItems.Count < 5 Then
            Me.BtnNewLine.Enabled = IsValid
        Else
            Me.BtnNewLine.Enabled = False
        End If

        If Me._LineItem1.IsValid = False Then
            Me.errInvoice.SetError(Me.lblLineTotal1, "Line Not Valid, check values")
        Else
            Me.errInvoice.SetError(Me.lblLineTotal1, "")

        End If

        Check_All_Lines_For_Validity()

    End Sub
    Private Sub _LineItem2_Valid(ByVal IsValid As Boolean) Handles _LineItem2.Valid

        If Me.Loading = True Then Exit Sub

        If Me._Invoice.LineItems.Count < 5 Then
            Me.BtnNewLine.Enabled = IsValid
        Else
            Me.BtnNewLine.Enabled = False
        End If

        If Me._LineItem2.IsValid = False Then
            Me.errInvoice.SetError(Me.lblLineTotal2, "Line Not Valid, check values")
        Else
            Me.errInvoice.SetError(Me.lblLineTotal2, "")

        End If

        Check_All_Lines_For_Validity()


    End Sub
    Private Sub _LineItem3_Valid(ByVal IsValid As Boolean) Handles _LineItem3.Valid

        If Me.Loading = True Then Exit Sub

        If Me._Invoice.LineItems.Count < 5 Then
            Me.BtnNewLine.Enabled = IsValid
        Else
            Me.BtnNewLine.Enabled = False
        End If

        If Me._LineItem3.IsValid = False Then
            Me.errInvoice.SetError(Me.lblLineTotal3, "Line Not Valid, check values")
        Else
            Me.errInvoice.SetError(Me.lblLineTotal3, "")

        End If

        Check_All_Lines_For_Validity()


    End Sub
    Private Sub _LineItem4_Valid(ByVal IsValid As Boolean) Handles _LineItem4.Valid

        If Me.Loading = True Then Exit Sub

        If Me._Invoice.LineItems.Count < 5 Then
            Me.BtnNewLine.Enabled = IsValid
        Else
            Me.BtnNewLine.Enabled = False
        End If

        If Me._LineItem4.IsValid = False Then
            Me.errInvoice.SetError(Me.lblLineTotal4, "Line Not Valid, check values")
        Else
            Me.errInvoice.SetError(Me.lblLineTotal4, "")

        End If

        Check_All_Lines_For_Validity()


    End Sub
    Private Sub _LineItem3_LaborCostChanged() Handles _LineItem3.LaborCostChanged
        If Me.Loading = True Then Exit Sub

        With Me._Invoice
            Me.lblLineTotal3.Text = .LineItems(3).Total.ToString("c")
            Me.lblInvoiceTotal.Text = .InvoiceTotal.ToString("c")
            Me.lblLaborCostTotal.Text = .LaborTotal.ToString("c")
            Me.lblMaterialCostTotal.Text = .MaterialTotal.ToString("c")

        End With
    End Sub

    Private Sub _LineItem4_LaborCostChanged() Handles _LineItem4.LaborCostChanged

        If Me.Loading = True Then Exit Sub

        With Me._Invoice
            Me.lblLineTotal4.Text = .LineItems(4).Total.ToString("c")
            Me.lblInvoiceTotal.Text = .InvoiceTotal.ToString("c")
            Me.lblLaborCostTotal.Text = .LaborTotal.ToString("c")
        End With

    End Sub

    Private Sub _LineItem2_MaterialCostChanged() Handles _LineItem2.MaterialCostChanged
        If Me.Loading = True Then Exit Sub

        With Me._Invoice
            Me.lblLineTotal2.Text = .LineItems(2).Total.ToString("c")
            Me.lblInvoiceTotal.Text = .InvoiceTotal.ToString("c")
            Me.lblMaterialCostTotal.Text = .MaterialTotal.ToString("c")
        End With

    End Sub

    Private Sub _LineItem3_MaterialCostChanged() Handles _LineItem3.MaterialCostChanged
        If Me.Loading = True Then Exit Sub

        With Me._Invoice
            Me.lblLineTotal3.Text = .LineItems(3).Total.ToString("c")
            Me.lblInvoiceTotal.Text = .InvoiceTotal.ToString("c")
            Me.lblMaterialCostTotal.Text = .MaterialTotal.ToString("c")
        End With

    End Sub

    Private Sub _LineItem4_MaterialCostChanged() Handles _LineItem4.MaterialCostChanged
        If Me.Loading = True Then Exit Sub

        With Me._Invoice
            Me.lblLineTotal4.Text = .LineItems(4).Total.ToString("c")
            Me.lblInvoiceTotal.Text = .InvoiceTotal.ToString("c")
            Me.lblMaterialCostTotal.Text = .MaterialTotal.ToString("c")
        End With

    End Sub

    Private Sub txtLaborCost2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLaborCost2.TextChanged
        If Me.Loading = True Then Exit Sub

        Me._Invoice.LineItems(2).LaborCost = Me.txtLaborCost2.Text

    End Sub

    Private Sub txtLaborCost3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLaborCost3.TextChanged
        If Me.Loading = True Then Exit Sub
        Me._Invoice.LineItems(3).LaborCost = Me.txtLaborCost3.Text

    End Sub

    Private Sub txtLaborCost4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLaborCost4.TextChanged
        If Me.Loading = True Then Exit Sub

        Me._Invoice.LineItems(4).LaborCost = Me.txtLaborCost4.Text

    End Sub

    Private Sub txtMaterialCost2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMaterialCost2.TextChanged
        If Me.Loading = True Then Exit Sub

        Me._Invoice.LineItems(2).MaterialCost = Me.txtMaterialCost2.Text

    End Sub

    Private Sub txtMaterialCost3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMaterialCost3.TextChanged
        If Me.Loading = True Then Exit Sub

        Me._Invoice.LineItems(3).MaterialCost = Me.txtMaterialCost3.Text

    End Sub

    Private Sub txtMaterialCost4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMaterialCost4.TextChanged
        If Me.Loading = True Then Exit Sub

        Me._Invoice.LineItems(4).MaterialCost = Me.txtMaterialCost4.Text

    End Sub
#End Region

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click

        Try

            Me._Invoice.InsertUpdateDelete(Convert.ToInt32(Me.lblInvoiceNo.Text))

            RefreshAfterUpdate()

            MessageBox.Show("Invoice Updated Successfully", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Information)

            Me.btnNewInvoice.Enabled = True
            Me.btnAddNotes.Enabled = True
            btnUpdate.Enabled = False

            Me._Invoice.IsDirty = False


        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message)

        End Try

    End Sub

    Private Sub dtpLine0_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpLine0.ValueChanged
        If Me.Loading = True Then Exit Sub

        Me._Invoice.LineItems(0).ItemDate = Me.dtpLine0.Value

    End Sub

    Private Sub dtpLine1_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpLine1.ValueChanged
        If Me.Loading = True Then Exit Sub

        Me._Invoice.LineItems(1).ItemDate = Me.dtpLine1.Value

    End Sub

    Private Sub dtpLine2_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpLine2.ValueChanged
        If Me.Loading = True Then Exit Sub

        Me._Invoice.LineItems(2).ItemDate = Me.dtpLine2.Value

    End Sub

    Private Sub dtpLine3_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpLine3.ValueChanged
        If Me.Loading = True Then Exit Sub

        Me._Invoice.LineItems(3).ItemDate = Me.dtpLine3.Value

    End Sub

    Private Sub dtpLine4_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpLine4.ValueChanged
        If Me.Loading = True Then Exit Sub

        Me._Invoice.LineItems(4).ItemDate = Me.dtpLine4.Value

    End Sub

    Private Sub dtpInvoiceDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpInvoiceDate.ValueChanged

        If Me.Loading = True Then Exit Sub

        Me._Invoice.InvoiceDate = Convert.ToDateTime(Me.dtpInvoiceDate.Value)

        Me.Check_All_Lines_For_Validity()


    End Sub
    Public Sub LoadInvoice(ByVal dsInv As DataSet, ByVal dsLineItems As DataSet)

        With dsInv.Tables(0).Rows(0)
            Me._Invoice.CustomerName = .Item("Name").ToString
            Me._Invoice.InvoiceDate = Convert.ToDateTime(.Item("InvoiceDate"))
            Me._Invoice.CustID = Convert.ToInt32(.Item("CustomerID"))
            Me._Invoice.InvoiceNo = Convert.ToInt32(.Item("InvoiceNo"))
            Me.lblInvoiceNo.Text = .Item("InvoiceNo").ToString
            Me.lblMaterialCostTotal.Text = .Item("MaterialTotal").ToString
            Me.lblLaborCostTotal.Text = .Item("LaborTotal").ToString
            Me.lblInvoiceTotal.Text = .Item("InvoiceTotal").ToString
            Me._CustomerID = Convert.ToInt32(.Item("CustomerID"))
        End With

        Utilities.LoadUltraCustCombo(Me.uCombo, Me._CustomerID)

        With dsLineItems.Tables(0)

            For rowCount As Integer = 0 To dsLineItems.Tables(0).Rows.Count - 1
                Select Case rowCount
                    Case 0
                        _LineItem0 = New LineItem
                        Me.PopulateLineItemObject(_LineItem0, .Rows(0))
                    Case 1
                        _LineItem1 = New LineItem
                        Me.PopulateLineItemObject(_LineItem1, .Rows(1))
                    Case 2
                        _LineItem2 = New LineItem
                        Me.PopulateLineItemObject(_LineItem2, .Rows(2))

                    Case 3
                        _LineItem3 = New LineItem
                        Me.PopulateLineItemObject(_LineItem3, .Rows(3))

                    Case 4
                        _LineItem4 = New LineItem
                        Me.PopulateLineItemObject(_LineItem4, .Rows(4))

                End Select
            Next

        End With

        LoadControls()

        Me.Loading = False

    End Sub
    Private Sub LoadControls()

        With Me._Invoice
            'Load Controls For Invoice
            Me.lblMaterialCostTotal.Text = .MaterialTotal.ToString("c")
            Me.lblLaborCostTotal.Text = .LaborTotal.ToString("c")
            Me.lblInvoiceTotal.Text = .InvoiceTotal.ToString("c")
            Me.dtpInvoiceDate.Value = .InvoiceDate

            'This will take care of the line items
            For LineCount As Integer = 0 To .LineItems.Count - 1
                Me.PopulateLineItemControls(LineCount, .LineItems(LineCount))
                Me.EnableLine(LineCount, True)
            Next

            'TODO - can this be put into an event?
            'Set the new line button 
            If .LineItems.Count = 5 Or .LineItems.Count = 0 Then
                Me.BtnNewLine.Enabled = False
            Else
                Me.BtnNewLine.Enabled = True
            End If

        End With


    End Sub
    Private Sub PopulateLineItemControls(ByVal LineNo As Integer, ByVal LI As LineItem)
        'This routine populates each line item on form from the invoice's line items

        For Each ctrl As Control In Me.grpLineItems.Controls

            If TypeOf ctrl Is DateTimePicker Then
                If ctrl.Tag.ToString = LineNo.ToString Then
                    DirectCast(ctrl, DateTimePicker).Value = LI.ItemDate
                End If
            End If

            If TypeOf (ctrl) Is TextBox Then
                If ctrl.Tag.ToString = LineNo.ToString Then
                    Dim rootName As String = ctrl.Name.Substring(0, ctrl.Name.Length - 1)
                    Select Case rootName
                        Case "txtLaborCost"
                            If LI.LaborCost.Length > 0 Then
                                ctrl.Text = LI.LaborCost
                            Else
                                ctrl.Text = vbNullString
                            End If
                        Case "txtLaborDesc"
                            ctrl.Text = LI.LaborDesc
                        Case "txtMaterialDesc"
                            ctrl.Text = LI.MaterialDesc
                        Case "txtMaterialCost"
                            If LI.MaterialCost.Length > 0 Then
                                ctrl.Text = LI.MaterialCost
                            Else
                                ctrl.Text = vbNullString
                            End If
                    End Select
                End If
            End If
        Next

        'TODO - Find a more intuitive way to do this
        Select Case LineNo
            Case 0
                'populate line total
                Me.lblLineTotal0.Text = LI.Total.ToString("c")
            Case 1
                'populate line total
                Me.lblLineTotal1.Text = LI.Total.ToString("c")
            Case 2
                'populate line total
                Me.lblLineTotal2.Text = LI.Total.ToString("c")
            Case 3
                'populate line total
                Me.lblLineTotal3.Text = LI.Total.ToString("c")
            Case 4
                'populate line total
                Me.lblLineTotal4.Text = LI.Total.ToString("c")
        End Select


    End Sub
    Private Sub PopulateLineItemObject(ByVal LineItem As LineItem, ByVal dr As DataRow)

        With LineItem
            .ItemNo = Convert.ToInt32(dr.Item("ItemNo"))

            If Not IsDBNull(dr.Item("ItemDate")) Then
                .ItemDate = Convert.ToDateTime(dr.Item("ItemDate"))
            Else
                .ItemDate = #1/1/1799#
            End If
            .LaborCost = dr.Item("LaborCost").ToString
            .LaborDesc = dr.Item("LaborDesc").ToString
            .MaterialCost = dr.Item("MaterialCost").ToString
            .MaterialDesc = dr.Item("MaterialDesc").ToString
            .IsValid = True
            .Mode = LineItem.Crud.Update
            Me._Invoice.LineItems.Add(LineItem)
        End With

    End Sub

    Private Sub _LineItem1_MaterialCostChanged() Handles _LineItem1.MaterialCostChanged
        If Me.Loading = True Then Exit Sub

        With Me._Invoice
            Me.lblLineTotal1.Text = .LineItems(1).Total.ToString("c")
            Me.lblInvoiceTotal.Text = .InvoiceTotal.ToString("c")
            Me.lblMaterialCostTotal.Text = .MaterialTotal.ToString("c")
        End With

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()

    End Sub

    Private Sub RefreshAfterUpdate()

        Dim dsInv As DataSet
        Dim dsLineItems As DataSet
        Dim invNum As Integer = Convert.ToInt32(Me.lblInvoiceNo.Text)
        Me.Loading = True

        BlankOutControls(False)

        Me._Invoice = New DelortBusObjects.Invoice(g_Settings)

        _Invoice.CrudMode = DBMode.Update

        dsInv = _Invoice.GetInvoice(invNum)
        dsLineItems = _Invoice.GetLineItems(invNum)
        LoadInvoice(dsInv, dsLineItems)

        If Me._Invoice.LineItems.Count = 0 Then

            Me.BtnNewLine.Enabled = False
            Me.txtLaborCost0.Enabled = True
            Me.txtMaterialCost0.Enabled = True
            Me.txtLaborDesc0.Enabled = True
            Me.txtMaterialDesc0.Enabled = True
            Me.dtpLine0.Enabled = True
            Me._LineItem0 = New LineItem
            Me._LineItem0.IsValid = False
            Me._Invoice.LineItems.Add(_LineItem0)
        Else
            Me.BtnNewLine.Enabled = True
        End If

        Me.Loading = False

        If Not IsNothing(Me.CallingForm) Then
            CallingForm.RefreshReport()

        End If
    End Sub
    Private Sub BlankOutControls(ByVal bolNewInvoice As Boolean)

        Me.ClearLineItemControls()

        lblMaterialCostTotal.Text = "$0.00"
        lblLaborCostTotal.Text = "$0.00"
        lblInvoiceTotal.Text = "$0.00"

        If bolNewInvoice = True Then

            Me.uCombo.SelectedIndex = -1
        End If
    End Sub
    Private Sub DeleteCheckBoxesChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        Try
            Dim chkBox As CheckBox = DirectCast(sender, CheckBox)
            Dim LineNo As Integer = Convert.ToInt32(chkBox.Tag)


            If LineNo = 0 Then
                'This could happen if program crashed
                If Me._Invoice.LineItems.Count = 0 Then
                    Me._LineItem0 = New LineItem
                    Me._Invoice.LineItems.Add(Me._LineItem0)
                End If
            End If


            If Me.Loading = False Then

                If chkBox.Checked Then
                    Me._Invoice.LineItems(LineNo).Mode = LineItem.Crud.Delete
                    Me.ZeroLineCosts(LineNo)
                    Me.EnableLine(LineNo, False)
                    Me.Check_All_Lines_For_Validity()
                    Me._Invoice.LineItems(LineNo).IsValid = True
                Else
                    If _Invoice.LineItems(LineNo).ItemNo <> 0 Then
                        Me._Invoice.LineItems(LineNo).Mode = LineItem.Crud.Update
                    Else
                        Me._Invoice.LineItems(LineNo).Mode = LineItem.Crud.Insert
                    End If

                    Me.EnableLine(LineNo, True)
                    Me._Invoice.LineItems(LineNo).CheckValidity()

                    Select Case LineNo
                        Case 0
                            Me.LoadLineNumbersFromControls(Me.txtLaborCost0, Me.txtMaterialCost0)
                        Case 1
                            Me.LoadLineNumbersFromControls(Me.txtLaborCost1, Me.txtMaterialCost1)
                        Case 2
                            Me.LoadLineNumbersFromControls(Me.txtLaborCost2, Me.txtMaterialCost2)
                        Case 3
                            Me.LoadLineNumbersFromControls(Me.txtLaborCost3, Me.txtMaterialCost3)
                        Case 4
                            Me.LoadLineNumbersFromControls(Me.txtLaborCost4, Me.txtMaterialCost4)

                    End Select
                End If
            End If

            chkBox.Enabled = True

            Me._Invoice.LineItems(LineNo).MarkedForDeletion = chkBox.Checked


        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub
    Private Sub frmEditInvoice_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        Try
            If Me._Invoice.CustID = 0 Then Exit Sub

            Dim LegitLineItems As Integer

            For Each lineItem As LineItem In Me._Invoice.LineItems
                If lineItem.Mode = lineItem.Crud.Update Then
                    LegitLineItems += 1
                End If
            Next

            If LegitLineItems = 0 Then
                If MessageBox.Show("There are no Line items SAVED with this Invoice." & vbCrLf & "Select Cancel to add line items, OK to Delete the Invoice.", ProgramName, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                    Me._Invoice.Delete(Convert.ToInt32(Me.lblInvoiceNo.Text))
                    Me._Invoice = Nothing
                    Exit Sub
                Else
                    e.Cancel = True
                    Exit Sub
                End If

            End If

            'If Me.btnUpdate.Enabled = True Then
            If Me._Invoice.IsDirty Then
                If MessageBox.Show("Changes have been made and not saved." & vbCrLf & "You can select Cancel to go back and Update, or OK to discard changes.", ProgramName, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                    e.Cancel = False
                Else
                    e.Cancel = True
                End If

            End If

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
    Private Sub ClearLineItemControls()

        For Each ctrl As Control In Me.grpLineItems.Controls
            If TypeOf (ctrl) Is TextBox Then
                ctrl.Text = vbNullString
                ctrl.Enabled = False
            End If

            If TypeOf (ctrl) Is CheckBox Then
                ctrl.Enabled = False
                DirectCast(ctrl, CheckBox).Checked = False
            End If
            If TypeOf (ctrl) Is DateTimePicker Then
                DirectCast(ctrl, DateTimePicker).Value = Now
                ctrl.Enabled = False
            End If
            If TypeOf ctrl Is Label AndAlso Not IsNothing(ctrl.Tag) AndAlso ctrl.Tag.ToString = "t" Then
                ctrl.Text = "$0.00"
                Me.errInvoice.SetError(ctrl, "")
            End If

        Next

    End Sub

    Private Sub frmEditInvoice_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            If Me.FormMode = Mode.AddNew Then
                Me.Loading = True
                Me.lblInvoiceNo.Text = Me._Invoice.GetLatestInvoiceNo.ToString
                Me._Invoice.InvoiceDate = Now
                Me.btnUpdate.Enabled = False
                Me.grpLineItems.Enabled = False
                Me.btnUpdate.Enabled = False
                Me.btnPrint.Enabled = False
                Me._LineItem0.IsValid = False
                Me._Invoice.LineItems.Add(Me._LineItem0)
                Me.Text = "Add New Invoice"
                Utilities.LoadUltraCustCombo(Me.uCombo, 0)
            ElseIf Me.FormMode = Mode.Edit Then
                Me.btnAddNotes.Enabled = True
                Me._Invoice.CrudMode = DBMode.Update
                Me.Text = "Edit Invoice"
                Me.btnUpdate.Enabled = False
                Me.btnPrint.Enabled = True
                Me.btnNewInvoice.Enabled = True
            End If

            Me.Loading = False

            AddEventHandlers(Me.grpLineItems)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try



    End Sub

    Private Sub Check_All_Lines_For_Validity()

        If Loading = True Then Exit Sub

        Dim Valid As Boolean = True

        If Me._Invoice.LineItems.Count = 0 Then
            Me.btnUpdate.Enabled = False
            Me.btnPrint.Enabled = False
            Exit Sub
        End If

        'iterate thru all lineitems to see if any are invalid. If so, the
        ' invoice is in an invalid state
        For Each Item As LineItem In Me._Invoice.LineItems
            If Item.IsValid = False Then
                Valid = False
                Exit For
            End If
        Next

        Me.btnUpdate.Enabled = Valid
        Me.btnPrint.Enabled = Valid

        Me._Invoice.IsDirty = True


    End Sub

    Private Sub btnNewInvoice_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNewInvoice.Click

        Try
            Me._Invoice.IsDirty = True
            Me.AddNewInvoice()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try



    End Sub

    Private Sub LoadLineNumbersFromControls(ByVal LaborTextBox As TextBox, ByVal MaterialTextBox As TextBox)

        Dim LineNo As Integer = Convert.ToInt32(LaborTextBox.Tag)

        With LaborTextBox
            If .Text.Length > 0 Then
                Me._Invoice.LineItems(LineNo).LaborCost = .Text
            End If
        End With

        With MaterialTextBox
            If .Text.Length > 0 Then

                Me._Invoice.LineItems(LineNo).MaterialCost = .Text
            End If
        End With

    End Sub

    Private Sub ZeroLineCosts(ByVal LineNo As Integer)

        With Me._Invoice.LineItems(LineNo)
            .LaborCost = ""
            .MaterialCost = ""
        End With

    End Sub

    Private Sub AddNewInvoice()
        Me.Loading = True

        Me.BlankOutControls(True)
        Me._Invoice = New DelortBusObjects.Invoice(g_Settings)
        Me.lblInvoiceNo.Text = Me._Invoice.GetLatestInvoiceNo.ToString
        Me._Invoice.InvoiceDate = Now
        Me.dtpInvoiceDate.Value = Now
        Me.btnUpdate.Enabled = False
        Me.grpLineItems.Enabled = False
        Me.BtnNewLine.Enabled = False
        Me.btnUpdate.Enabled = False
        Me.btnAddNotes.Enabled = False
        Me._LineItem0.IsValid = False
        Me._LineItem0 = New LineItem
        Me._Invoice.LineItems.Add(Me._LineItem0)
        Me.FormMode = Mode.AddNew
        Me.Loading = False

    End Sub

    Private Sub btnAddNotes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddNotes.Click

        Try
            Dim frmNotes As New frmNotes(Convert.ToInt32(Me.lblInvoiceNo.Text))
            frmNotes.ShowDialog()

            frmNotes.Dispose()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub EnableLine(ByVal LineNo As Integer, ByVal Enabled As Boolean)

        For Each ctrl As Control In Me.grpLineItems.Controls

            If Not IsNothing(ctrl.Tag) Then
                If ctrl.Tag.ToString = LineNo.ToString Then
                    ctrl.Enabled = Enabled
                End If
            End If
        Next

    End Sub

    Private Sub LineTextBoxesLeave(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim Name As String = DirectCast(sender, TextBox).Name

        'If it's a dollar box, set dots to zero
        If Name.Substring(Name.Length - 5, 4).ToUpper = "COST" Then
            If DirectCast(sender, TextBox).Text.Trim = "." Then
                DirectCast(sender, TextBox).Text = "0"
            End If
        End If

        DirectCast(sender, TextBox).BackColor = Color.FromKnownColor(KnownColor.Window)

    End Sub

    Private Sub LineTextBoxesEnter(ByVal sender As Object, ByVal e As System.EventArgs)

        DirectCast(sender, TextBox).BackColor = Color.Wheat

    End Sub

    Private Sub AddEventHandlers(ByVal grpLine As Control)

        For Each ctrl As Control In grpLine.Controls
            If TypeOf (ctrl) Is TextBox Then
                AddHandler ctrl.Enter, AddressOf Me.LineTextBoxesEnter
                AddHandler ctrl.Leave, AddressOf Me.LineTextBoxesLeave
            End If
            If TypeOf (ctrl) Is CheckBox Then
                AddHandler DirectCast(ctrl, CheckBox).CheckedChanged, AddressOf Me.DeleteCheckBoxesChanged
            End If
        Next

    End Sub

    Private Sub _Invoice_Dirty(ByVal IsDirty As Boolean) Handles _Invoice.Dirty

        Me.btnPrint.Enabled = Not (IsDirty)
        Me.btnNewInvoice.Enabled = Not IsDirty

    End Sub


    Private Sub uCombo_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles uCombo.SelectionChangeCommitted

        Try
            If Me.Loading = True Then Exit Sub

            Me._Invoice.CustID = Convert.ToInt32(Me.uCombo.SelectedItem.DataValue)

            If Me.FormMode = Mode.AddNew Then
                If Me.grpLineItems.Enabled = False Then
                    Me.grpLineItems.Enabled = True
                    Me.EnableLine(0, True)
                End If
            ElseIf Me.FormMode = Mode.Edit Then
                Me.Check_All_Lines_For_Validity()
            End If

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

End Class
