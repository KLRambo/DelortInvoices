Option Strict On

Imports DelortBusObjects
Imports System.Configuration.ConfigurationSettings

Public Class frmInvoiceSummaryRpt

    Inherits frmBaseWindow

    Private dsInv As DataSet
    Private Invoices As New DelortBusObjects.Invoice(g_settings)


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
    Friend WithEvents grpHeader As System.Windows.Forms.GroupBox
    Friend WithEvents ugSummary As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents cboFromMonth As System.Windows.Forms.ComboBox
    Friend WithEvents cboFromYear As System.Windows.Forms.ComboBox
    Friend WithEvents grpHeader2 As System.Windows.Forms.GroupBox
    Friend WithEvents cboToYear As System.Windows.Forms.ComboBox
    Friend WithEvents cboToMonth As System.Windows.Forms.ComboBox
    Friend WithEvents btnRun As System.Windows.Forms.Button
    Friend WithEvents grpFooter As System.Windows.Forms.GroupBox
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents grpHeader3 As System.Windows.Forms.GroupBox
    Friend WithEvents txtLabor As System.Windows.Forms.TextBox
    Friend WithEvents lblLabor As System.Windows.Forms.Label
    Friend WithEvents lblMaterials As System.Windows.Forms.Label
    Friend WithEvents txtMaterials As System.Windows.Forms.TextBox
    Friend WithEvents lblTotal As System.Windows.Forms.Label
    Friend WithEvents txtTotals As System.Windows.Forms.TextBox
    Friend WithEvents lblNoOfInv As System.Windows.Forms.Label
    Friend WithEvents txtNumOfInv As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmInvoiceSummaryRpt))
        Me.grpMain = New System.Windows.Forms.GroupBox()
        Me.ugSummary = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.grpHeader = New System.Windows.Forms.GroupBox()
        Me.cboFromYear = New System.Windows.Forms.ComboBox()
        Me.cboFromMonth = New System.Windows.Forms.ComboBox()
        Me.grpHeader2 = New System.Windows.Forms.GroupBox()
        Me.cboToYear = New System.Windows.Forms.ComboBox()
        Me.cboToMonth = New System.Windows.Forms.ComboBox()
        Me.grpHeader3 = New System.Windows.Forms.GroupBox()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.btnRun = New System.Windows.Forms.Button()
        Me.grpFooter = New System.Windows.Forms.GroupBox()
        Me.lblNoOfInv = New System.Windows.Forms.Label()
        Me.txtNumOfInv = New System.Windows.Forms.TextBox()
        Me.lblTotal = New System.Windows.Forms.Label()
        Me.txtTotals = New System.Windows.Forms.TextBox()
        Me.lblMaterials = New System.Windows.Forms.Label()
        Me.txtMaterials = New System.Windows.Forms.TextBox()
        Me.lblLabor = New System.Windows.Forms.Label()
        Me.txtLabor = New System.Windows.Forms.TextBox()
        Me.grpMain.SuspendLayout()
        CType(Me.ugSummary, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpHeader.SuspendLayout()
        Me.grpHeader2.SuspendLayout()
        Me.grpHeader3.SuspendLayout()
        Me.grpFooter.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpMain
        '
        Me.grpMain.BackColor = System.Drawing.Color.Transparent
        Me.grpMain.Controls.Add(Me.ugSummary)
        Me.grpMain.Location = New System.Drawing.Point(9, 58)
        Me.grpMain.Name = "grpMain"
        Me.grpMain.Size = New System.Drawing.Size(515, 414)
        Me.grpMain.TabIndex = 1
        Me.grpMain.TabStop = False
        '
        'ugSummary
        '
        Appearance1.TextHAlignAsString = "Left"
        Me.ugSummary.DisplayLayout.Appearance = Appearance1
        Me.ugSummary.DisplayLayout.MaxColScrollRegions = 1
        Me.ugSummary.DisplayLayout.MaxRowScrollRegions = 1
        Me.ugSummary.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugSummary.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.ugSummary.DisplayLayout.Override.AllowGroupBy = Infragistics.Win.DefaultableBoolean.[False]
        Me.ugSummary.DisplayLayout.Override.AllowGroupSwapping = Infragistics.Win.UltraWinGrid.AllowGroupSwapping.NotAllowed
        Me.ugSummary.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.ugSummary.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.ugSummary.DisplayLayout.Override.CellPadding = 0
        Me.ugSummary.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugSummary.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugSummary.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.ugSummary.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ugSummary.Location = New System.Drawing.Point(11, 15)
        Me.ugSummary.Name = "ugSummary"
        Me.ugSummary.Size = New System.Drawing.Size(493, 385)
        Me.ugSummary.TabIndex = 1
        '
        'grpHeader
        '
        Me.grpHeader.BackColor = System.Drawing.Color.Transparent
        Me.grpHeader.Controls.Add(Me.cboFromYear)
        Me.grpHeader.Controls.Add(Me.cboFromMonth)
        Me.grpHeader.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpHeader.Location = New System.Drawing.Point(9, 0)
        Me.grpHeader.Name = "grpHeader"
        Me.grpHeader.Size = New System.Drawing.Size(163, 56)
        Me.grpHeader.TabIndex = 2
        Me.grpHeader.TabStop = False
        Me.grpHeader.Text = "Select From Date:"
        '
        'cboFromYear
        '
        Me.cboFromYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFromYear.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboFromYear.Location = New System.Drawing.Point(93, 23)
        Me.cboFromYear.Name = "cboFromYear"
        Me.cboFromYear.Size = New System.Drawing.Size(58, 21)
        Me.cboFromYear.TabIndex = 1
        '
        'cboFromMonth
        '
        Me.cboFromMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFromMonth.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboFromMonth.Location = New System.Drawing.Point(14, 23)
        Me.cboFromMonth.Name = "cboFromMonth"
        Me.cboFromMonth.Size = New System.Drawing.Size(58, 21)
        Me.cboFromMonth.TabIndex = 0
        '
        'grpHeader2
        '
        Me.grpHeader2.BackColor = System.Drawing.Color.Transparent
        Me.grpHeader2.Controls.Add(Me.cboToYear)
        Me.grpHeader2.Controls.Add(Me.cboToMonth)
        Me.grpHeader2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpHeader2.Location = New System.Drawing.Point(183, 0)
        Me.grpHeader2.Name = "grpHeader2"
        Me.grpHeader2.Size = New System.Drawing.Size(164, 56)
        Me.grpHeader2.TabIndex = 3
        Me.grpHeader2.TabStop = False
        Me.grpHeader2.Text = "Select To Date:"
        '
        'cboToYear
        '
        Me.cboToYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboToYear.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboToYear.Location = New System.Drawing.Point(93, 23)
        Me.cboToYear.Name = "cboToYear"
        Me.cboToYear.Size = New System.Drawing.Size(56, 21)
        Me.cboToYear.TabIndex = 5
        '
        'cboToMonth
        '
        Me.cboToMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboToMonth.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboToMonth.Location = New System.Drawing.Point(14, 23)
        Me.cboToMonth.Name = "cboToMonth"
        Me.cboToMonth.Size = New System.Drawing.Size(58, 21)
        Me.cboToMonth.TabIndex = 4
        '
        'grpHeader3
        '
        Me.grpHeader3.BackColor = System.Drawing.Color.Transparent
        Me.grpHeader3.Controls.Add(Me.btnPrint)
        Me.grpHeader3.Controls.Add(Me.btnRun)
        Me.grpHeader3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpHeader3.Location = New System.Drawing.Point(358, 0)
        Me.grpHeader3.Name = "grpHeader3"
        Me.grpHeader3.Size = New System.Drawing.Size(166, 56)
        Me.grpHeader3.TabIndex = 4
        Me.grpHeader3.TabStop = False
        Me.grpHeader3.Text = "Action Buttons"
        '
        'btnPrint
        '
        Me.btnPrint.Enabled = False
        Me.btnPrint.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrint.Image = CType(resources.GetObject("btnPrint.Image"), System.Drawing.Image)
        Me.btnPrint.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnPrint.Location = New System.Drawing.Point(90, 23)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(64, 25)
        Me.btnPrint.TabIndex = 2
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnRun
        '
        Me.btnRun.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRun.Image = CType(resources.GetObject("btnRun.Image"), System.Drawing.Image)
        Me.btnRun.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnRun.Location = New System.Drawing.Point(8, 23)
        Me.btnRun.Name = "btnRun"
        Me.btnRun.Size = New System.Drawing.Size(64, 25)
        Me.btnRun.TabIndex = 1
        Me.btnRun.Text = "&Run"
        Me.btnRun.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'grpFooter
        '
        Me.grpFooter.BackColor = System.Drawing.Color.Transparent
        Me.grpFooter.Controls.Add(Me.lblNoOfInv)
        Me.grpFooter.Controls.Add(Me.txtNumOfInv)
        Me.grpFooter.Controls.Add(Me.lblTotal)
        Me.grpFooter.Controls.Add(Me.txtTotals)
        Me.grpFooter.Controls.Add(Me.lblMaterials)
        Me.grpFooter.Controls.Add(Me.txtMaterials)
        Me.grpFooter.Controls.Add(Me.lblLabor)
        Me.grpFooter.Controls.Add(Me.txtLabor)
        Me.grpFooter.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpFooter.Location = New System.Drawing.Point(9, 480)
        Me.grpFooter.Name = "grpFooter"
        Me.grpFooter.Size = New System.Drawing.Size(515, 56)
        Me.grpFooter.TabIndex = 5
        Me.grpFooter.TabStop = False
        Me.grpFooter.Text = "Totals:"
        '
        'lblNoOfInv
        '
        Me.lblNoOfInv.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNoOfInv.Location = New System.Drawing.Point(4, 20)
        Me.lblNoOfInv.Name = "lblNoOfInv"
        Me.lblNoOfInv.Size = New System.Drawing.Size(46, 14)
        Me.lblNoOfInv.TabIndex = 7
        Me.lblNoOfInv.Text = "Count:"
        '
        'txtNumOfInv
        '
        Me.txtNumOfInv.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtNumOfInv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtNumOfInv.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNumOfInv.Location = New System.Drawing.Point(51, 20)
        Me.txtNumOfInv.Name = "txtNumOfInv"
        Me.txtNumOfInv.ReadOnly = True
        Me.txtNumOfInv.Size = New System.Drawing.Size(39, 21)
        Me.txtNumOfInv.TabIndex = 6
        Me.txtNumOfInv.Text = "0"
        Me.txtNumOfInv.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblTotal
        '
        Me.lblTotal.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotal.Location = New System.Drawing.Point(383, 20)
        Me.lblTotal.Name = "lblTotal"
        Me.lblTotal.Size = New System.Drawing.Size(40, 14)
        Me.lblTotal.TabIndex = 5
        Me.lblTotal.Text = "Total:"
        '
        'txtTotals
        '
        Me.txtTotals.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtTotals.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtTotals.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTotals.Location = New System.Drawing.Point(424, 20)
        Me.txtTotals.Name = "txtTotals"
        Me.txtTotals.ReadOnly = True
        Me.txtTotals.Size = New System.Drawing.Size(85, 21)
        Me.txtTotals.TabIndex = 4
        Me.txtTotals.Text = "$0.00"
        Me.txtTotals.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblMaterials
        '
        Me.lblMaterials.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMaterials.Location = New System.Drawing.Point(230, 20)
        Me.lblMaterials.Name = "lblMaterials"
        Me.lblMaterials.Size = New System.Drawing.Size(64, 14)
        Me.lblMaterials.TabIndex = 3
        Me.lblMaterials.Text = "Materials:"
        '
        'txtMaterials
        '
        Me.txtMaterials.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtMaterials.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMaterials.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaterials.Location = New System.Drawing.Point(294, 20)
        Me.txtMaterials.Name = "txtMaterials"
        Me.txtMaterials.ReadOnly = True
        Me.txtMaterials.Size = New System.Drawing.Size(89, 21)
        Me.txtMaterials.TabIndex = 2
        Me.txtMaterials.Text = "$0.00"
        Me.txtMaterials.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblLabor
        '
        Me.lblLabor.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLabor.Location = New System.Drawing.Point(94, 20)
        Me.lblLabor.Name = "lblLabor"
        Me.lblLabor.Size = New System.Drawing.Size(44, 14)
        Me.lblLabor.TabIndex = 1
        Me.lblLabor.Text = "Labor:"
        '
        'txtLabor
        '
        Me.txtLabor.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtLabor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtLabor.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLabor.Location = New System.Drawing.Point(139, 20)
        Me.txtLabor.Name = "txtLabor"
        Me.txtLabor.ReadOnly = True
        Me.txtLabor.Size = New System.Drawing.Size(89, 21)
        Me.txtLabor.TabIndex = 0
        Me.txtLabor.Text = "$0.00"
        Me.txtLabor.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'frmInvoiceSummaryRpt
        '
        Me.AccessibleRole = System.Windows.Forms.AccessibleRole.None
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(530, 544)
        Me.Controls.Add(Me.grpFooter)
        Me.Controls.Add(Me.grpHeader3)
        Me.Controls.Add(Me.grpHeader2)
        Me.Controls.Add(Me.grpHeader)
        Me.Controls.Add(Me.grpMain)
        Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmInvoiceSummaryRpt"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Invoice Summary Report"
        Me.grpMain.ResumeLayout(False)
        CType(Me.ugSummary, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpHeader.ResumeLayout(False)
        Me.grpHeader2.ResumeLayout(False)
        Me.grpHeader3.ResumeLayout(False)
        Me.grpFooter.ResumeLayout(False)
        Me.grpFooter.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub frmInvoiceSummaryRpt_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Utilities.LoadMonthCombo(Me.cboFromMonth)
            Utilities.LoadMonthCombo(Me.cboToMonth)
            Utilities.LoadYearsCombo(Me.cboToYear)
            Utilities.LoadYearsCombo(Me.cboFromYear)

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click

        Try
            Cursor = Cursors.WaitCursor


            Dim sbFromDate As New System.Text.StringBuilder
            Dim sbToDate As New System.Text.StringBuilder
            Dim FromDate As Date
            Dim ToDate As Date

            With sbFromDate
                .Append(Me.cboFromMonth.SelectedIndex + 1)
                .Append("/")
                .Append("01")
                .Append("/")
                .Append(Me.cboFromYear.Text)
                FromDate = Convert.ToDateTime(.ToString)
            End With

            With sbToDate
                .Append(Me.cboToMonth.SelectedIndex + 1)
                .Append("/")
                .Append("01")
                .Append("/")
                .Append(Me.cboToYear.Text)
                ToDate = Convert.ToDateTime(.ToString)
            End With

            If FromDate > ToDate Then
                MessageBox.Show("'To Date' can't be less than 'From Date'", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If

            If FromDate > Now Or ToDate > Now Then
                MessageBox.Show("Cant report on future dates", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If

            Me.btnPrint.Enabled = True

            dsInv = Invoices.GetMonthlyReports(FromDate, ToDate)

            Me.ugSummary.DataSource = dsInv.Tables(0)

            LoadTotals()

        Catch ex As Exception
            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            Cursor = Cursors.Default

        End Try

    End Sub


    Private Sub ugSummary_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugSummary.InitializeLayout


        With e.Layout.Bands(0)

            .Columns(0).Header.Caption = "Month"
            .Columns(0).Width = 82
            .Columns(0).CellAppearance.ForeColor = SystemColors.HotTrack
            .Columns(1).Header.Caption = "No. of Invoices"
            .Columns(1).Width = 110
            .Columns(2).Header.Caption = "Labor"
            .Columns(2).Width = 85
            .Columns(2).Format = "$#,##0.00"
            .Columns(3).Header.Caption = "Materials"
            .Columns(3).Width = 85
            .Columns(3).Format = "$#,##0.00"
            .Columns(4).Header.Caption = "Total"
            .Columns(4).Width = 107
            .Columns(4).Format = "$#,##0.00"
            .Columns(5).Hidden = True
            .Columns(6).Hidden = True
            .Columns(7).Hidden = True
            .Columns(8).Hidden = True
        End With

        'e.Layout.Override.BorderStyleRowSelector = Infragistics.Win.UIElementBorderStyle.None
    End Sub
    Private Sub LoadTotals()

        Dim RowCount As Integer = Me.dsInv.Tables(0).Rows.Count - 1

        Dim LaborTotal As Double
        Dim MaterialTotal As Double
        Dim ReportTotal As Double

        LaborTotal = Convert.ToDouble(dsInv.Tables(0).Rows(RowCount).Item("LaborTotal").ToString)
        MaterialTotal = Convert.ToDouble(dsInv.Tables(0).Rows(RowCount).Item("MaterialTotal").ToString)
        ReportTotal = Convert.ToDouble(dsInv.Tables(0).Rows(RowCount).Item("MasterTotal").ToString)

        Me.txtNumOfInv.Text = dsInv.Tables(0).Rows(RowCount).Item("InvoiceCount").ToString
        Me.txtLabor.Text = LaborTotal.ToString("c")
        Me.txtMaterials.Text = MaterialTotal.ToString("c")
        Me.txtTotals.Text = ReportTotal.ToString("c")

    End Sub

    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click

        Try
            Cursor = Cursors.WaitCursor
            Utilities.PrintInvoiceSummary(Me.dsInv)

        Catch ex As Exception

            Errhandler.Logging.LogError(ex)
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally

            Cursor = Cursors.Default

        End Try

    End Sub

    Private Sub cboFromMonth_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboFromMonth.SelectionChangeCommitted, cboFromYear.SelectionChangeCommitted, cboToMonth.SelectionChangeCommitted, cboToYear.SelectionChangeCommitted

        Me.ugSummary.DataSource = Nothing
        Me.txtLabor.Text = "$0.00"
        Me.txtMaterials.Text = "$0.00"
        Me.txtTotals.Text = "$0.00"

        Me.dsInv = Nothing

        Me.btnPrint.Enabled = False

    End Sub

    Private Sub frmInvoiceSummaryRpt_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        If Not IsNothing(dsInv) Then
            dsInv.Dispose()
        End If

        Invoices = Nothing

    End Sub
End Class
