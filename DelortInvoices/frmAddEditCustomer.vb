
#Region " Options "

Option Explicit On
Option Strict On

#End Region

#Region " Imports "

Imports DelortBusObjects
Imports System.Configuration
Imports System.Text.RegularExpressions

#End Region

Public Class frmAddEditCustomer
    Inherits frmBaseWindow

#Region " Variables "

    Private _FormMode As Mode
    Public Customer As Customer
    Public Loading As Boolean = True
    Public CallingForm As frmCustomerReport
    Public WithEvents _Customer As New Customer(g_Settings)

#End Region

#Region " Enums "

    Public Enum Mode
        AddNew = 0
        Edit = 1
    End Enum

#End Region

#Region " Properties "

    Public Property FormMode() As Mode
        Get
            Return _FormMode

        End Get

        Set(ByVal Value As Mode)
            _FormMode = Value
        End Set
    End Property

#End Region


#Region " Windows Form Designer generated code "

    Public Sub New(ByVal Mode As Mode)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Me.FormMode = Mode

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
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents cboState As System.Windows.Forms.ComboBox
    Friend WithEvents txtNotes As System.Windows.Forms.TextBox
    Friend WithEvents txtCell As System.Windows.Forms.TextBox
    Friend WithEvents txtPhone As System.Windows.Forms.TextBox
    Friend WithEvents txtZip As System.Windows.Forms.TextBox
    Friend WithEvents txtAddr2 As System.Windows.Forms.TextBox
    Friend WithEvents txtCity As System.Windows.Forms.TextBox
    Friend WithEvents grpMain As System.Windows.Forms.GroupBox
    Friend WithEvents lblNotes As System.Windows.Forms.Label
    Friend WithEvents lblCell As System.Windows.Forms.Label
    Friend WithEvents lblPhone As System.Windows.Forms.Label
    Friend WithEvents lblAddress1 As System.Windows.Forms.Label
    Friend WithEvents lblAddr2 As System.Windows.Forms.Label
    Friend WithEvents lblCity As System.Windows.Forms.Label
    Friend WithEvents lblState As System.Windows.Forms.Label
    Friend WithEvents lblZip As System.Windows.Forms.Label
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents grpFooter As System.Windows.Forms.GroupBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblFirstName As System.Windows.Forms.Label
    Friend WithEvents txtFirstName As System.Windows.Forms.TextBox
    Friend WithEvents txtLastName As System.Windows.Forms.TextBox
    Friend WithEvents txtAddr1 As System.Windows.Forms.TextBox
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents grpWindows As System.Windows.Forms.GroupBox
    Friend WithEvents chkWindows As System.Windows.Forms.CheckBox
    Friend WithEvents txtWindowPrice As System.Windows.Forms.TextBox
    Friend WithEvents chkSpring As System.Windows.Forms.CheckBox
    Friend WithEvents chkFall As System.Windows.Forms.CheckBox
    Friend WithEvents chkPerRequest As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAddEditCustomer))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.grpWindows = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.chkPerRequest = New System.Windows.Forms.CheckBox()
        Me.chkFall = New System.Windows.Forms.CheckBox()
        Me.chkSpring = New System.Windows.Forms.CheckBox()
        Me.txtWindowPrice = New System.Windows.Forms.TextBox()
        Me.chkWindows = New System.Windows.Forms.CheckBox()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.grpFooter = New System.Windows.Forms.GroupBox()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.grpMain = New System.Windows.Forms.GroupBox()
        Me.lblFirstName = New System.Windows.Forms.Label()
        Me.txtFirstName = New System.Windows.Forms.TextBox()
        Me.cboState = New System.Windows.Forms.ComboBox()
        Me.lblNotes = New System.Windows.Forms.Label()
        Me.txtNotes = New System.Windows.Forms.TextBox()
        Me.lblCell = New System.Windows.Forms.Label()
        Me.txtCell = New System.Windows.Forms.TextBox()
        Me.lblPhone = New System.Windows.Forms.Label()
        Me.txtPhone = New System.Windows.Forms.TextBox()
        Me.txtZip = New System.Windows.Forms.TextBox()
        Me.lblAddress1 = New System.Windows.Forms.Label()
        Me.lblAddr2 = New System.Windows.Forms.Label()
        Me.lblCity = New System.Windows.Forms.Label()
        Me.lblState = New System.Windows.Forms.Label()
        Me.lblZip = New System.Windows.Forms.Label()
        Me.lblName = New System.Windows.Forms.Label()
        Me.txtLastName = New System.Windows.Forms.TextBox()
        Me.txtAddr2 = New System.Windows.Forms.TextBox()
        Me.txtCity = New System.Windows.Forms.TextBox()
        Me.txtAddr1 = New System.Windows.Forms.TextBox()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.grpWindows.SuspendLayout()
        Me.grpFooter.SuspendLayout()
        Me.grpMain.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grpWindows
        '
        Me.grpWindows.BackColor = System.Drawing.Color.Transparent
        Me.grpWindows.Controls.Add(Me.Label1)
        Me.grpWindows.Controls.Add(Me.chkPerRequest)
        Me.grpWindows.Controls.Add(Me.chkFall)
        Me.grpWindows.Controls.Add(Me.chkSpring)
        Me.grpWindows.Controls.Add(Me.txtWindowPrice)
        Me.grpWindows.Controls.Add(Me.chkWindows)
        Me.grpWindows.Location = New System.Drawing.Point(4, 211)
        Me.grpWindows.Name = "grpWindows"
        Me.grpWindows.Size = New System.Drawing.Size(324, 86)
        Me.grpWindows.TabIndex = 9
        Me.grpWindows.TabStop = False
        Me.grpWindows.Text = "Windows Clients"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(8, 51)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 16)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Price:"
        '
        'chkPerRequest
        '
        Me.chkPerRequest.Enabled = False
        Me.chkPerRequest.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkPerRequest.Location = New System.Drawing.Point(201, 63)
        Me.chkPerRequest.Name = "chkPerRequest"
        Me.chkPerRequest.Size = New System.Drawing.Size(102, 16)
        Me.chkPerRequest.TabIndex = 4
        Me.chkPerRequest.Text = "Per Request"
        '
        'chkFall
        '
        Me.chkFall.Enabled = False
        Me.chkFall.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkFall.Location = New System.Drawing.Point(201, 39)
        Me.chkFall.Name = "chkFall"
        Me.chkFall.Size = New System.Drawing.Size(64, 16)
        Me.chkFall.TabIndex = 3
        Me.chkFall.Text = "Fall"
        '
        'chkSpring
        '
        Me.chkSpring.Enabled = False
        Me.chkSpring.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSpring.Location = New System.Drawing.Point(201, 15)
        Me.chkSpring.Name = "chkSpring"
        Me.chkSpring.Size = New System.Drawing.Size(78, 16)
        Me.chkSpring.TabIndex = 2
        Me.chkSpring.Text = "Spring"
        '
        'txtWindowPrice
        '
        Me.txtWindowPrice.Enabled = False
        Me.txtWindowPrice.Location = New System.Drawing.Point(78, 51)
        Me.txtWindowPrice.Name = "txtWindowPrice"
        Me.txtWindowPrice.Size = New System.Drawing.Size(58, 21)
        Me.txtWindowPrice.TabIndex = 1
        '
        'chkWindows
        '
        Me.chkWindows.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkWindows.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkWindows.Location = New System.Drawing.Point(8, 22)
        Me.chkWindows.Name = "chkWindows"
        Me.chkWindows.Size = New System.Drawing.Size(128, 16)
        Me.chkWindows.TabIndex = 0
        Me.chkWindows.Text = "Window Customer"
        '
        'lblStatus
        '
        Me.lblStatus.BackColor = System.Drawing.Color.Transparent
        Me.lblStatus.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblStatus.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus.Location = New System.Drawing.Point(6, 305)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(554, 24)
        Me.lblStatus.TabIndex = 8
        Me.lblStatus.Text = "Adding A New Customer"
        '
        'grpFooter
        '
        Me.grpFooter.BackColor = System.Drawing.Color.Transparent
        Me.grpFooter.Controls.Add(Me.btnDelete)
        Me.grpFooter.Controls.Add(Me.btnClear)
        Me.grpFooter.Controls.Add(Me.btnCancel)
        Me.grpFooter.Controls.Add(Me.btnUpdate)
        Me.grpFooter.Location = New System.Drawing.Point(335, 211)
        Me.grpFooter.Name = "grpFooter"
        Me.grpFooter.Size = New System.Drawing.Size(225, 86)
        Me.grpFooter.TabIndex = 7
        Me.grpFooter.TabStop = False
        '
        'btnDelete
        '
        Me.btnDelete.BackColor = System.Drawing.Color.AliceBlue
        Me.btnDelete.Enabled = False
        Me.btnDelete.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.Image = CType(resources.GetObject("btnDelete.Image"), System.Drawing.Image)
        Me.btnDelete.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDelete.Location = New System.Drawing.Point(12, 51)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(77, 24)
        Me.btnDelete.TabIndex = 13
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.btnDelete, "Delete This Customer")
        Me.btnDelete.UseVisualStyleBackColor = False
        '
        'btnClear
        '
        Me.btnClear.BackColor = System.Drawing.Color.AliceBlue
        Me.btnClear.Enabled = False
        Me.btnClear.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClear.Image = CType(resources.GetObject("btnClear.Image"), System.Drawing.Image)
        Me.btnClear.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnClear.Location = New System.Drawing.Point(131, 19)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(77, 24)
        Me.btnClear.TabIndex = 11
        Me.btnClear.Text = "C&lear"
        Me.btnClear.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.btnClear, "Clear All Text Boxes")
        Me.btnClear.UseVisualStyleBackColor = False
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.Color.AliceBlue
        Me.btnCancel.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.Image = CType(resources.GetObject("btnCancel.Image"), System.Drawing.Image)
        Me.btnCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCancel.Location = New System.Drawing.Point(12, 19)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(77, 24)
        Me.btnCancel.TabIndex = 12
        Me.btnCancel.Text = "E&xit"
        Me.ToolTip1.SetToolTip(Me.btnCancel, "Exit Form")
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'btnUpdate
        '
        Me.btnUpdate.BackColor = System.Drawing.Color.AliceBlue
        Me.btnUpdate.Enabled = False
        Me.btnUpdate.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpdate.Image = CType(resources.GetObject("btnUpdate.Image"), System.Drawing.Image)
        Me.btnUpdate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnUpdate.Location = New System.Drawing.Point(131, 51)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(77, 24)
        Me.btnUpdate.TabIndex = 10
        Me.btnUpdate.Text = "&Save"
        Me.btnUpdate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.btnUpdate, "Save Changes")
        Me.btnUpdate.UseVisualStyleBackColor = False
        '
        'grpMain
        '
        Me.grpMain.BackColor = System.Drawing.Color.Transparent
        Me.grpMain.Controls.Add(Me.lblFirstName)
        Me.grpMain.Controls.Add(Me.txtFirstName)
        Me.grpMain.Controls.Add(Me.cboState)
        Me.grpMain.Controls.Add(Me.lblNotes)
        Me.grpMain.Controls.Add(Me.txtNotes)
        Me.grpMain.Controls.Add(Me.lblCell)
        Me.grpMain.Controls.Add(Me.txtCell)
        Me.grpMain.Controls.Add(Me.lblPhone)
        Me.grpMain.Controls.Add(Me.txtPhone)
        Me.grpMain.Controls.Add(Me.txtZip)
        Me.grpMain.Controls.Add(Me.lblAddress1)
        Me.grpMain.Controls.Add(Me.lblAddr2)
        Me.grpMain.Controls.Add(Me.lblCity)
        Me.grpMain.Controls.Add(Me.lblState)
        Me.grpMain.Controls.Add(Me.lblZip)
        Me.grpMain.Controls.Add(Me.lblName)
        Me.grpMain.Controls.Add(Me.txtLastName)
        Me.grpMain.Controls.Add(Me.txtAddr2)
        Me.grpMain.Controls.Add(Me.txtCity)
        Me.grpMain.Controls.Add(Me.txtAddr1)
        Me.grpMain.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpMain.Location = New System.Drawing.Point(4, 8)
        Me.grpMain.Name = "grpMain"
        Me.grpMain.Size = New System.Drawing.Size(556, 200)
        Me.grpMain.TabIndex = 6
        Me.grpMain.TabStop = False
        Me.grpMain.Text = "* Required Field"
        '
        'lblFirstName
        '
        Me.lblFirstName.Location = New System.Drawing.Point(296, 23)
        Me.lblFirstName.Name = "lblFirstName"
        Me.lblFirstName.Size = New System.Drawing.Size(85, 16)
        Me.lblFirstName.TabIndex = 38
        Me.lblFirstName.Text = "*First Name:"
        '
        'txtFirstName
        '
        Me.txtFirstName.Location = New System.Drawing.Point(392, 23)
        Me.txtFirstName.Name = "txtFirstName"
        Me.txtFirstName.Size = New System.Drawing.Size(146, 21)
        Me.txtFirstName.TabIndex = 1
        '
        'cboState
        '
        Me.cboState.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboState.Location = New System.Drawing.Point(325, 97)
        Me.cboState.Name = "cboState"
        Me.cboState.Size = New System.Drawing.Size(67, 21)
        Me.cboState.TabIndex = 5
        '
        'lblNotes
        '
        Me.lblNotes.Location = New System.Drawing.Point(272, 134)
        Me.lblNotes.Name = "lblNotes"
        Me.lblNotes.Size = New System.Drawing.Size(48, 16)
        Me.lblNotes.TabIndex = 36
        Me.lblNotes.Text = "Notes:"
        '
        'txtNotes
        '
        Me.txtNotes.Location = New System.Drawing.Point(325, 134)
        Me.txtNotes.Multiline = True
        Me.txtNotes.Name = "txtNotes"
        Me.txtNotes.Size = New System.Drawing.Size(213, 58)
        Me.txtNotes.TabIndex = 9
        '
        'lblCell
        '
        Me.lblCell.Location = New System.Drawing.Point(19, 173)
        Me.lblCell.Name = "lblCell"
        Me.lblCell.Size = New System.Drawing.Size(80, 16)
        Me.lblCell.TabIndex = 34
        Me.lblCell.Text = "Cell Phone:"
        '
        'txtCell
        '
        Me.txtCell.Location = New System.Drawing.Point(114, 171)
        Me.txtCell.Name = "txtCell"
        Me.txtCell.Size = New System.Drawing.Size(138, 21)
        Me.txtCell.TabIndex = 8
        '
        'lblPhone
        '
        Me.lblPhone.Location = New System.Drawing.Point(19, 134)
        Me.lblPhone.Name = "lblPhone"
        Me.lblPhone.Size = New System.Drawing.Size(64, 16)
        Me.lblPhone.TabIndex = 32
        Me.lblPhone.Text = "Phone:"
        '
        'txtPhone
        '
        Me.txtPhone.Location = New System.Drawing.Point(114, 134)
        Me.txtPhone.Name = "txtPhone"
        Me.txtPhone.Size = New System.Drawing.Size(138, 21)
        Me.txtPhone.TabIndex = 7
        '
        'txtZip
        '
        Me.txtZip.Location = New System.Drawing.Point(482, 97)
        Me.txtZip.Name = "txtZip"
        Me.txtZip.Size = New System.Drawing.Size(56, 21)
        Me.txtZip.TabIndex = 6
        Me.txtZip.Text = "60544"
        '
        'lblAddress1
        '
        Me.lblAddress1.Location = New System.Drawing.Point(11, 60)
        Me.lblAddress1.Name = "lblAddress1"
        Me.lblAddress1.Size = New System.Drawing.Size(80, 16)
        Me.lblAddress1.TabIndex = 29
        Me.lblAddress1.Text = "*Address1:"
        '
        'lblAddr2
        '
        Me.lblAddr2.Location = New System.Drawing.Point(316, 60)
        Me.lblAddr2.Name = "lblAddr2"
        Me.lblAddr2.Size = New System.Drawing.Size(68, 16)
        Me.lblAddr2.TabIndex = 28
        Me.lblAddr2.Text = "Address2:"
        '
        'lblCity
        '
        Me.lblCity.Location = New System.Drawing.Point(11, 97)
        Me.lblCity.Name = "lblCity"
        Me.lblCity.Size = New System.Drawing.Size(56, 16)
        Me.lblCity.TabIndex = 27
        Me.lblCity.Text = "*City:"
        '
        'lblState
        '
        Me.lblState.Location = New System.Drawing.Point(272, 97)
        Me.lblState.Name = "lblState"
        Me.lblState.Size = New System.Drawing.Size(50, 16)
        Me.lblState.TabIndex = 26
        Me.lblState.Text = "State:"
        '
        'lblZip
        '
        Me.lblZip.Location = New System.Drawing.Point(416, 97)
        Me.lblZip.Name = "lblZip"
        Me.lblZip.Size = New System.Drawing.Size(40, 16)
        Me.lblZip.TabIndex = 25
        Me.lblZip.Text = "*Zip:"
        '
        'lblName
        '
        Me.lblName.Location = New System.Drawing.Point(11, 23)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(88, 16)
        Me.lblName.TabIndex = 24
        Me.lblName.Text = "*Last Name:"
        '
        'txtLastName
        '
        Me.txtLastName.Location = New System.Drawing.Point(114, 23)
        Me.txtLastName.Name = "txtLastName"
        Me.txtLastName.Size = New System.Drawing.Size(138, 21)
        Me.txtLastName.TabIndex = 0
        '
        'txtAddr2
        '
        Me.txtAddr2.Location = New System.Drawing.Point(392, 60)
        Me.txtAddr2.Name = "txtAddr2"
        Me.txtAddr2.Size = New System.Drawing.Size(146, 21)
        Me.txtAddr2.TabIndex = 3
        '
        'txtCity
        '
        Me.txtCity.Location = New System.Drawing.Point(114, 97)
        Me.txtCity.Name = "txtCity"
        Me.txtCity.Size = New System.Drawing.Size(138, 21)
        Me.txtCity.TabIndex = 4
        Me.txtCity.Text = "Plainfield"
        '
        'txtAddr1
        '
        Me.txtAddr1.Location = New System.Drawing.Point(114, 60)
        Me.txtAddr1.Name = "txtAddr1"
        Me.txtAddr1.Size = New System.Drawing.Size(186, 21)
        Me.txtAddr1.TabIndex = 2
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'frmAddEditCustomer
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(566, 332)
        Me.Controls.Add(Me.grpWindows)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.grpFooter)
        Me.Controls.Add(Me.grpMain)
        Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmAddEditCustomer"
        Me.Text = "Customer Data Entry Form"
        Me.grpWindows.ResumeLayout(False)
        Me.grpWindows.PerformLayout()
        Me.grpFooter.ResumeLayout(False)
        Me.grpMain.ResumeLayout(False)
        Me.grpMain.PerformLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub txtAddr1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAddr1.TextChanged

        Me._Customer.Addr1 = Me.txtAddr1.Text.Trim

        If Loading = True Then Exit Sub

        If Me.txtAddr1.Text.Trim.Length < 2 Then
            Me.ErrorProvider1.SetError(Me.txtAddr1, "Address Must Have at Least 2 Characters")
        Else
            Me.ErrorProvider1.SetError(Me.txtAddr1, "")
        End If
    End Sub

    Private Sub txtLastName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLastName.TextChanged

        Me._Customer.LastName = Me.txtLastName.Text.Trim
        If Loading = True Then Exit Sub

        If Me.txtLastName.Text.Trim.Length < 2 Then
            Me.ErrorProvider1.SetError(Me.txtLastName, "Last Name Must Have at Least 2 Characters")
        Else
            Me.ErrorProvider1.SetError(Me.txtLastName, "")
        End If

    End Sub

    Private Sub txtAddr2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAddr2.TextChanged

        Me._Customer.Addr2 = Me.txtAddr2.Text.Trim

    End Sub

    Private Sub txtCity_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCity.TextChanged

        Me._Customer.City = Me.txtCity.Text.Trim

        If Loading = True Then Exit Sub

        If Me.txtCity.Text.Trim.Length < 2 Then
            Me.ErrorProvider1.SetError(Me.txtCity, "City Must Have at Least 2 Characters")
        Else
            Me.ErrorProvider1.SetError(Me.txtCity, "")
        End If

    End Sub

    Private Sub txtzip_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtZip.TextChanged

        Me._Customer.Zip = Me.txtZip.Text.Trim
        If Loading = True Then Exit Sub

        If Me.txtZip.Text.Trim.Length > 0 Then
            If Not Regex.IsMatch(Me.txtZip.Text, ConfigurationManager.AppSettings("ZipCodeMask")) Then
                Me.ErrorProvider1.SetError(Me.txtZip, "Not a zipcode")
            Else
                Me.ErrorProvider1.SetError(Me.txtZip, "")
            End If
        Else
            Me.ErrorProvider1.SetError(Me.txtZip, "Required Field")
        End If


    End Sub

    Private Sub btnUpdateClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click

        Try
            If Me._Customer.DBMode = Customer.CrudMode.Insert Then
                Dim CustomerInDB As Boolean = Me._Customer.CheckForCustomer()
                If CustomerInDB Then
                    If MessageBox.Show("A customer by that name is already in your database." & vbCrLf & "Click OK to Add, or Cancel to stop.", ProgramName, MessageBoxButtons.OKCancel) = DialogResult.Cancel Then
                        Exit Sub
                    End If

                End If
            End If

            Me._Customer.InsertUpdateDelete()

            Me.UpdateCallingForm()

            With Me.lblStatus
                .ForeColor = System.Drawing.Color.Red
                .Text = Me._Customer.FirstName & " " & Me._Customer.LastName & " Updated."
            End With

            If Me._Customer.DBMode = Customer.CrudMode.Insert Then
                Me.ResetForm()
            End If

            If Me.FormMode = Mode.AddNew Then
                Me.txtCity.Text = "Plainfield"
                Me.txtZip.Text = "60544"
                Me._Customer.IsDirty = False
            End If

            Me.btnUpdate.Enabled = False


        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub _Customer_Valid(ByVal Valid As Boolean) Handles _Customer.Valid

        If Loading = True Then Exit Sub

        Me.btnUpdate.Enabled = Valid


    End Sub

    Private Sub txtPhone_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPhone.TextChanged


        Me._Customer.Phone = Me.txtPhone.Text.Trim

        If Me.txtPhone.Text.Trim.Length > 0 Then
            If Not Regex.IsMatch(Me.txtPhone.Text.Trim, ConfigurationManager.AppSettings("PhoneMask")) Then
                Me.ErrorProvider1.SetError(Me.txtPhone, "Not in phone format")
            Else
                Me.ErrorProvider1.SetError(Me.txtPhone, "")

            End If
        Else
            Me.ErrorProvider1.SetError(Me.txtPhone, "")

        End If

    End Sub

    Private Sub txtCell_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCell.TextChanged


        Me._Customer.CellPhone = Me.txtCell.Text.Trim

        If Me.txtCell.Text.Length > 0 Then
            If Not Regex.IsMatch(Me.txtCell.Text, ConfigurationManager.AppSettings("PhoneMask")) Then
                Me.ErrorProvider1.SetError(Me.txtCell, "Not in phone format")
            Else
                Me.ErrorProvider1.SetError(Me.txtCell, "")

            End If
        Else
            Me.ErrorProvider1.SetError(Me.txtCell, "")

        End If


    End Sub

    Private Sub txtNotes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNotes.TextChanged

        Me._Customer.Notes = Me.txtNotes.Text.Trim

    End Sub

    Private Sub frmNewCustomer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            LoadStatesCombo()

            If Me.FormMode = Mode.Edit Then
                LoadForm()
                Me.lblStatus.Text = "Editing"
            ElseIf Me.FormMode = Mode.AddNew Then
                Me.btnDelete.Enabled = False
                Me._Customer.State = Me.cboState.Text
                Me._Customer.WindowCustomer = "0"
            End If

            Me.Loading = False
            Me._Customer.IsDirty = False
            Me.btnUpdate.Enabled = False


        Catch ex As Exception

            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try


    End Sub
    Private Sub LoadStatesCombo()

        Dim States() As String = ConfigurationManager.AppSettings("States").Split(",".ToCharArray)

        With cboState
            For i As Integer = 0 To States.Length - 1
                .Items.Add(States(i).ToString)
            Next
            .SelectedIndex = 0
        End With

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click

        Try

            Me.Close()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click

        Try

            Me.ResetForm()
            Me._Customer.IsDirty = True

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
    Private Sub LoadForm()

        'Load the controls
        With Me.Customer
            Me.txtFirstName.Text = .FirstName
            Me.txtLastName.Text = .LastName
            Me.txtAddr1.Text = .Addr1
            Me.txtAddr2.Text = .Addr2
            Me.txtCity.Text = .City
            Me.cboState.Text = .State
            Me.txtZip.Text = .Zip
            Me.txtPhone.Text = .Phone
            Me.txtCell.Text = .CellPhone
            Me.txtNotes.Text = .Notes
            Me._Customer.CustID = .CustID

            If Me.Customer.WindowCustomer = "1" Then
                Me.chkWindows.Checked = True
                Me.txtWindowPrice.Text = .WindowPrice.ToString
                Select Case Me.Customer.WindowsCleanedIn
                    Case Customer.WindowsCleanedWhen.All
                        Me.chkFall.Checked = True
                        Me.chkSpring.Checked = True
                        Me.chkPerRequest.Checked = True
                    Case Customer.WindowsCleanedWhen.Fall
                        Me.chkFall.Checked = True
                    Case Customer.WindowsCleanedWhen.FallAndOnRequest
                        Me.chkFall.Checked = True
                        Me.chkPerRequest.Checked = True
                    Case Customer.WindowsCleanedWhen.OnRequest
                        Me.chkPerRequest.Checked = True
                    Case Customer.WindowsCleanedWhen.Spring
                        Me.chkSpring.Checked = True
                    Case Customer.WindowsCleanedWhen.SpringAndFall
                        Me.chkFall.Checked = True
                        Me.chkSpring.Checked = True
                    Case Customer.WindowsCleanedWhen.SpringAndOnRequest
                        Me.chkSpring.Checked = True
                        Me.chkPerRequest.Checked = True
                End Select
            End If
        End With

        Me._Customer.DBMode = Customer.CrudMode.Update


    End Sub

    Private Sub txtFirstName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFirstName.TextChanged
        Me._Customer.FirstName = Me.txtFirstName.Text.Trim

        If Loading = True Then Exit Sub

        If Me.txtFirstName.Text.Trim.Length < 1 Then
            Me.ErrorProvider1.SetError(Me.txtFirstName, "First Name Must Have at Least 1 Character")
        Else
            Me.ErrorProvider1.SetError(Me.txtFirstName, "")
        End If

    End Sub

    Private Sub _Customer_DeleteValid(ByVal IsValid As Boolean) Handles _Customer.DeleteValid

        Me.btnDelete.Enabled = IsValid

    End Sub
    Public Sub UpdateCallingForm()

        If Not Me.CallingForm Is Nothing Then
            Me.CallingForm.ReloadData()

        End If

    End Sub

    Private Sub ResetForm()

        Me.Loading = True

        With Me._Customer
            .Addr1 = ""
            .Addr2 = ""
            .CellPhone = ""
            .City = ""
            .FirstName = ""
            .LastName = ""
            .Phone = ""
            .Notes = ""
            .Zip = ""
            .WindowCustomer = "0"
            .WindowPrice = Nothing
            .WindowsCleanedIn = Customer.WindowsCleanedWhen.NA
        End With

        For Each ctrl As Control In Me.grpMain.Controls
            If TypeOf (ctrl) Is TextBox Then
                ctrl.Text = ""
                Me.ErrorProvider1.SetError(ctrl, "")
            End If

        Next

        'Clear Windows Group
        Me.chkWindows.Checked = False
        'Me.txtWindowPrice.Enabled = False



        Me.Loading = False
        Me._Customer.IsDirty = False


    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click

        Try

            If MessageBox.Show("Are you sure you want to delete " & _Customer.FullName & "?", ProgramName, MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) = DialogResult.Cancel Then
                Exit Sub
            End If

            Me._Customer.DBMode = Customer.CrudMode.Delete
            Me._Customer.InsertUpdateDelete()

            ResetForm()

            Me.FormMode = Mode.AddNew
            Me.UpdateCallingForm()
            Me._Customer.DBMode = Customer.CrudMode.Insert

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try


    End Sub

    Private Sub frmAddEditCustomer_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        Try
            If Me._Customer.IsDirty = True Then
                If MessageBox.Show("Changes have been made and not saved.You can select Cancel to go back and Update, or OK to discard changes.", ProgramName, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                    e.Cancel = False
                    If Not IsNothing(Me._Customer) Then
                        _Customer = Nothing
                    End If
                Else
                    e.Cancel = True
                End If

            End If


        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub _Customer_Dirty() Handles _Customer.Dirty

        If Me._Customer.IsDirty And Me._Customer.IsValid Then
            Me.btnUpdate.Enabled = True
        End If
        If Me._Customer.IsDirty Then Me.btnClear.Enabled = True

    End Sub

    Private Sub cboState_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboState.SelectionChangeCommitted

        Me._Customer.State = Me.cboState.Text

    End Sub

    Private Sub txtWindowPrice_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWindowPrice.TextChanged

        If IsNumeric(Me.txtWindowPrice.Text) Then
            Me._Customer.WindowPrice = Convert.ToDouble(Me.txtWindowPrice.Text)
        Else
            Me._Customer.WindowPrice = Nothing

        End If

    End Sub

    Private Sub txtWindowPrice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtWindowPrice.KeyPress
        e.Handled = Utilities.NumericOnly(e, Me.txtWindowPrice)

    End Sub

    Private Sub chkWindows_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkWindows.CheckedChanged

        If Me.chkWindows.Checked Then
            Me._Customer.WindowCustomer = "1"
        Else
            Me._Customer.WindowCustomer = "0"
            ClearWindowGroup()
        End If

    End Sub

    Private Sub _Customer_WindowsCustomer(ByVal IsWindowClient As Boolean) Handles _Customer.WindowsCustomer

        Me.txtWindowPrice.Enabled = IsWindowClient
        Me.chkFall.Enabled = IsWindowClient
        Me.chkPerRequest.Enabled = IsWindowClient
        Me.chkSpring.Enabled = IsWindowClient

    End Sub
    Private Sub ClearWindowGroup()

        Me.chkSpring.Checked = False
        Me.chkFall.Checked = False
        Me.chkPerRequest.Checked = False
        Me.txtWindowPrice.Text = ""

    End Sub

    Private Sub WhenWindowsCleaned(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSpring.CheckedChanged, chkFall.CheckedChanged, chkPerRequest.CheckedChanged

        With Me._Customer
            If Me.chkFall.Checked And Me.chkSpring.Checked And Me.chkPerRequest.Checked Then
                .WindowsCleanedIn = Customer.WindowsCleanedWhen.All
            ElseIf Me.chkFall.Checked And Me.chkSpring.Checked Then
                .WindowsCleanedIn = Customer.WindowsCleanedWhen.SpringAndFall
            ElseIf Me.chkFall.Checked And Me.chkPerRequest.Checked Then
                .WindowsCleanedIn = Customer.WindowsCleanedWhen.FallAndOnRequest
            ElseIf Me.chkSpring.Checked And Me.chkPerRequest.Checked Then
                .WindowsCleanedIn = Customer.WindowsCleanedWhen.SpringAndOnRequest
            ElseIf Me.chkFall.Checked Then
                .WindowsCleanedIn = Customer.WindowsCleanedWhen.Fall
            ElseIf Me.chkSpring.Checked Then
                .WindowsCleanedIn = Customer.WindowsCleanedWhen.Spring
            ElseIf Me.chkPerRequest.Checked Then
                .WindowsCleanedIn = Customer.WindowsCleanedWhen.OnRequest
            Else
                .WindowsCleanedIn = Customer.WindowsCleanedWhen.NA
            End If
        End With

    End Sub


End Class
