
#Region " Options "

Option Explicit On
Option Strict On

#End Region

#Region " Imports "

Imports System.IO
Imports System.Configuration.ConfigurationSettings
Imports DelortBusObjects
Imports Infragistics.Win.UltraWinEditors

#End Region

Public Class Utilities

    Public Enum ReportType
        Windows = 0
        NonWindows = 1
        All = 2
    End Enum

    Public Enum CrudMode
        UPDATE = 0
        INSERT = 1
        DELETE = 2
    End Enum

    Public Shared Sub LoadMonthCombo(ByRef cboBox As ComboBox)

        Dim Month As Integer = DatePart(DateInterval.Month, Now)

        With cboBox
            .Items.Add("Jan")
            .Items.Add("Feb")
            .Items.Add("Mar")
            .Items.Add("Apr")
            .Items.Add("May")
            .Items.Add("Jun")
            .Items.Add("Jul")
            .Items.Add("Aug")
            .Items.Add("Sep")
            .Items.Add("Oct")
            .Items.Add("Nov")
            .Items.Add("Dec")
            'set the selected index to current month
            .SelectedIndex = Month - 1

        End With

    End Sub

    Public Shared Sub LoadYearsCombo(ByRef cboBox As ComboBox)

        'display last 10 years
        Dim Year As Integer = DatePart(DateInterval.Year, Now)

        For i As Integer = Year - 9 To Year
            With cboBox
                .Items.Add(i.ToString)
            End With
        Next

        'Display this year
        cboBox.SelectedIndex = 9

    End Sub

    Public Shared Sub LoadCustCombo(ByRef cboBox As ComboBox, ByVal CustID As Integer)

        Dim ds As New DataSet
        Dim Customer As New Customer(g_Settings)
        Dim Index As Integer = 0

        ds = Customer.GetAllCustomersForCombo

        If ds.Tables(0).Rows.Count > 0 Then
            With cboBox
                .DataSource = ds.Tables(0)
                .DisplayMember = "Name"
                .ValueMember = "CustomerID"
                .SelectedIndex = -1
            End With
        End If

        'set select index

        If CustID = 0 Then
            cboBox.SelectedIndex = -1
        Else
            For Each drv As DataRowView In cboBox.Items

                If Convert.ToInt32(drv("CustomerID")) = CustID Then
                    cboBox.SelectedIndex = Index
                    Exit Sub
                End If

                Index += 1
            Next
        End If

    End Sub

    Public Shared Sub LoadUltraCustCombo(ByRef cboBox As UltraComboEditor, ByVal CustID As Integer)

        Dim ds As New DataSet
        Dim Customer As New Customer(g_Settings)
        Dim Index As Integer = 0

        If CustID = 0 Then
            ds = Customer.GetActiveCustomersForCombo
        Else
            ds = Customer.GetAllCustomersForCombo

        End If

        If ds.Tables(0).Rows.Count > 0 Then
            With cboBox
                .DataSource = ds.Tables(0)
                .DisplayMember = "Name"
                .ValueMember = "CustomerID"
                .SelectedIndex = -1
            End With
        End If

        'set select index

        If CustID = 0 Then
            cboBox.SelectedIndex = -1
        Else
            For Each Item As Infragistics.Win.ValueListItem In cboBox.Items
                If Convert.ToInt32(cboBox.Items(Index).DataValue) = CustID Then
                    cboBox.SelectedIndex = Index
                    Exit Sub
                End If

                Index += 1
            Next
        End If

    End Sub
    Public Shared Sub LoadMiscExpTypes(ByRef cboBox As ComboBox, ByRef Mode As String)

        Dim ds As New DataSet
        Dim Expenses As New DelortBusObjects.MiscExp((g_Settings))

        If Mode = "All" Then
            ds = Expenses.GetExpenseTypesAll
        Else
            ds = Expenses.GetExpenseTypesActive
        End If


        If ds.Tables(0).Rows.Count > 0 Then
            With cboBox
                .DataSource = ds.Tables(0)
                .DisplayMember = "Description"
                .ValueMember = "CategoryID"
                .SelectedIndex = 0
                .DropDownStyle = ComboBoxStyle.DropDownList
            End With
        End If


    End Sub
    Public Shared Function NumericOnly(ByVal e As System.Windows.Forms.KeyPressEventArgs, ByVal TextBox As TextBox) As Boolean

        Dim KeyAscii As Integer
        Dim PositiveOnly As Boolean = True
        Dim IntegerOnly As Boolean = False


        KeyAscii = AscW(e.KeyChar) ' NOTE: used as a flag when = 0 to mark bad values

        Select Case KeyAscii

            Case 48 To 57, 8, 13 ' Digits 0 - 9, Backspace, CR.


            Case 45 ' Minus sign.

                KeyAscii = 0

            Case 46 ' Period (decimal point).

                ' If we already have one period, throw it away

                If InStr(TextBox.Text, ".") <> 0 Then

                    KeyAscii = 0

                End If


            Case Else ' Provide no handling for other keys.

                KeyAscii = 0

        End Select

        ' If we want to throw the keystroke away, then set the event

        ' as already handled. Otherwise, let the keystroke be

        ' handled normally.

        If KeyAscii = 0 Then

            Return True

        Else

            Return False

        End If

    End Function

    Public Shared Function OpenInvoice(ByVal Invnum As Integer, ByVal MDIForm As Form) As Boolean

        For Each frm As Form In MDIForm.MdiChildren
            If TypeOf (frm) Is frmEditInvoice AndAlso CType(frm, frmEditInvoice).lblInvoiceNo.Text = CStr(Invnum) Then
                MessageBox.Show("That Invoice is already open.", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Hand)
                Return True
            End If
        Next

        Dim Invoice As New DelortBusObjects.Invoice(g_Settings)
        Dim dsInv As DataSet
        Dim dsLineItems As DataSet

        dsInv = Invoice.GetInvoice(Invnum)

        'If no records, jump out of here
        If dsInv.Tables(0).Rows.Count = 0 Then Return False

        dsLineItems = Invoice.GetLineItems(Invnum)

        Dim EditInv As New frmEditInvoice(Mode.Edit)

        With EditInv
            .MdiParent = MDIForm
            .LoadInvoice(dsInv, dsLineItems)
            .Show()
        End With

        'EditInv.Dispose()

        Return True

    End Function

    Public Shared Function OpenInvoice(ByVal Invnum As Integer, ByVal CallingForm As frmViewEditInvoices, ByVal MDIForm As Form) As Boolean

        For Each frm As Form In MDIForm.MdiChildren
            If TypeOf (frm) Is frmEditInvoice AndAlso CType(frm, frmEditInvoice).lblInvoiceNo.Text = CStr(Invnum) Then
                MessageBox.Show("That Invoice is already open.", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Hand)
                Return True
            End If
        Next

        Dim Invoice As New DelortBusObjects.Invoice(g_Settings)
        Dim dsInv As DataSet
        Dim dsLineItems As DataSet

        dsInv = Invoice.GetInvoice(Invnum)

        'If no records, jump out of here
        If dsInv.Tables(0).Rows.Count = 0 Then Return False

        dsLineItems = Invoice.GetLineItems(Invnum)

        Dim EditInv As New frmEditInvoice(Mode.Edit)

        With EditInv
            .MdiParent = MDIForm
            .LoadInvoice(dsInv, dsLineItems)
            .CallingForm = CallingForm
            .Show()
        End With

        Return True

    End Function

#Region " Manage Program Forms "

    Shared Sub ShowInvSummaryForm(ByVal MDIParent As Form)

        Dim frmInvSumm As New frmInvoiceSummaryRpt

        With frmInvSumm
            .MdiParent = MDIParent
            .Show()
        End With

    End Sub

    Shared Sub PrintInvoice(ByVal InvNum As Integer)

        Dim InvRpt As New Invoice
        Dim tempfrm As New ActRptPrvw
        Dim Cust As DelortBusObjects.Customer

        Dim Invoice As New DelortBusObjects.Invoice(g_Settings)

        Dim Inv As DelortBusObjects.Invoice
        Inv = Invoice.GetInvoiceObject(InvNum)

        If Inv Is Nothing Then
            Throw New Exception("Invoice Not Found")
        End If

        'This could happen if program crashed; duplicated it in test
        If Inv.LineItems.Count = 0 Then
            Throw New Exception("No Valid Line Items To Print")

        End If

        'This could happen if program has an error
        If Inv.LineItems.Count = 1 Then
            Inv.LineItems(0).CheckValidity()
            If Inv.LineItems(0).IsValid = False Then
                Throw New Exception("No Valid Line Items To Print")
            End If
        End If

        'Now create a customer object
        Dim Customer As New DelortBusObjects.Customer(g_Settings)

        Cust = Customer.GetCustomer(Inv.CustID)

        Try
            With InvRpt
                ._Invoice = Inv
                ._Customer = Cust
                .Run()
                tempfrm.ARVeiwer.Document = .Document

            End With

            tempfrm.ShowDialog()

        Finally

            InvRpt.Dispose()
            tempfrm.Dispose()
            Invoice = Nothing

        End Try

    End Sub

    Shared Sub PrintInvoice(ByVal Inv As DelortBusObjects.Invoice)

        'This could happen if program crashed; duplicated it in test
        If Inv.LineItems.Count = 0 Then
            Throw New Exception("No Valid Line Items To Print")

        End If

        If Inv.LineItems.Count = 1 Then
            If Inv.LineItems(0).IsValid = False Then
                Throw New Exception("No Valid Line Items To Print")
            End If
        End If

        'Create Customer Object Now

        Dim Customers As New Customer(g_Settings)
        Dim cust As Customer

        cust = Customers.GetCustomer(Inv.CustID)

        Dim InvRpt As New Invoice
        Dim tempfrm As New ActRptPrvw

        Try
            With InvRpt
                ._Invoice = Inv
                ._Customer = cust
                .Run()
                tempfrm.ARVeiwer.Document = .Document

            End With

            tempfrm.ShowDialog()

        Finally

            InvRpt.Dispose()
            tempfrm.Dispose()
            'Invoice = Nothing


        End Try

    End Sub

    Shared Sub PrintInvoiceSummary(ByVal ds As DataSet)

        Dim InvRpt As New InvoiceSummaryRpt
        Dim tempfrm As New ActRptPrvw

        Try
            With InvRpt
                .dsReport = ds
                .Run()
                tempfrm.ARVeiwer.Document = .Document
            End With

            tempfrm.ShowDialog()

        Finally

            InvRpt.Dispose()
            tempfrm.Dispose()


        End Try

    End Sub

    Shared Sub PrintMiscExpByCategory(ByVal ds As DataSet, ByVal StartDate As String, ByVal EndDate As String)

        'Dim row As DataRow
        Dim sb As New System.Text.StringBuilder
        Dim pad As Char = Convert.ToChar(" ")
        Dim GrandTotal As Double
        Dim Total As Double

        'Write File header
        With sb

            .Append(" ".PadRight(20, pad))
            .Append("****** Misc Expenses Summary Report *****")
            .Append(vbCrLf & vbCrLf)

            If Not IsNothing(StartDate) Then
                .Append(" ".PadRight(25, pad))
                .Append("Between Dates ")
                .Append(StartDate)
                .Append(" And ")
                .Append(EndDate)
                .Append(vbCrLf & vbCrLf & vbCrLf & vbCrLf)
            Else
                .Append(" ".PadRight(32, pad))
                .Append("Grand Total Report")
                .Append(vbCrLf & vbCrLf & vbCrLf & vbCrLf)
            End If

            .Append("Category".PadRight(30, pad))
            .Append("Total".PadRight(25, pad))
            .Append(vbCrLf)

            'next line
            .Append("--------".PadRight(30, pad))
            .Append("-------".PadRight(25, pad))
            .Append(vbCrLf)
            .Append(vbCrLf)

        End With

        For Each row As DataRow In ds.Tables(0).Rows
            With row
                If Convert.ToInt32(.Item("CategoryID")) > 0 Then
                    sb.Append(.Item("Description").ToString.PadRight(25, pad))
                    If Not IsDBNull(.Item("Total")) Then
                        Total = Convert.ToDouble(.Item("Total"))
                        GrandTotal += Total
                        sb.Append(Total.ToString("c").PadLeft(11, pad))
                    Else
                        sb.Append("$0.00".PadLeft(11, pad))
                    End If
                    sb.Append(vbCrLf & vbCrLf)
                End If
            End With
        Next

        With sb
            .Append(vbCrLf & vbCrLf & vbCrLf)
            .Append(" ".PadRight(14, pad))
            .Append("Grand Total: " & GrandTotal.ToString("c"))
        End With

        If File.Exists("MiscSumRpt.txt") Then
            File.Delete("MiscSumRpt.txt")
        End If

        Dim sw As StreamWriter = File.AppendText("MiscSumRpt.txt")

        With sw
            ' Update the underlying file.
            .Write(sb.ToString)
            .Flush()
            .Close()
        End With

        System.Diagnostics.Process.Start("notepad.exe", "MiscSumRpt.txt")

        ds.Dispose()


    End Sub

    Shared Sub PrintMiscExpByMonth(ByVal ds As DataSet, ByVal Year As String)

        'Dim row As DataRow
        Dim sb As New System.Text.StringBuilder
        Dim pad As Char = Convert.ToChar(" ")
        Dim GrandTotal As Double
        Dim Total As Double
        Dim FileName As String = "MiscMoSumRpt.txt"
        Dim Month As Integer

        'Write File header
        With sb

            .Append(" ".PadRight(20, pad))
            .Append("****** Misc Expenses Monthly Summary Report *****")
            .Append(vbCrLf & vbCrLf)

            .Append(" ".PadRight(35, pad))
            .Append("For the Year:  ")
            .Append(Year)
            .Append(vbCrLf & vbCrLf & vbCrLf & vbCrLf)

            .Append("Month".PadRight(30, pad))
            .Append("Total".PadRight(25, pad))
            .Append(vbCrLf)

            'next line
            .Append("-----".PadRight(30, pad))
            .Append("-----".PadRight(25, pad))
            .Append(vbCrLf)
            .Append(vbCrLf)

        End With

        For Each row As DataRow In ds.Tables(0).Rows
            With row
                Month = Convert.ToInt32(.Item("Mo"))
                sb.Append(MonthName(Month).PadRight(25, pad))
                If Not IsDBNull(.Item("MoSum")) Then
                    Total = Convert.ToDouble(.Item("MoSum"))
                    GrandTotal += Total
                    sb.Append(Total.ToString("c").PadLeft(11, pad))
                Else
                    sb.Append("$0.00".PadLeft(11, pad))
                End If
                sb.Append(vbCrLf & vbCrLf)
            End With
        Next

        With sb
            .Append(vbCrLf & vbCrLf & vbCrLf)
            .Append(" ".PadRight(12, pad))
            .Append("Grand Total: " & GrandTotal.ToString("c").PadLeft(11, pad))
        End With

        If File.Exists(FileName) Then
            File.Delete(FileName)
        End If

        Dim sw As StreamWriter = File.AppendText(FileName)

        With sw
            ' Update the underlying file.
            .Write(sb.ToString)
            .Flush()
            .Close()
        End With

        System.Diagnostics.Process.Start("notepad.exe", FileName)

        ds.Dispose()


    End Sub

    Shared Sub PrintCustomers(ByVal ds As DataSet, ByVal FileName As String, ByVal Spacing As String, ByVal ReportType As ReportType)


        'Dim row As DataRow
        Dim sb As New System.Text.StringBuilder
        Dim pad As Char = Convert.ToChar(" ")
        Dim Header As String

        'Write File header
        With sb
            .Append("Name".PadRight(25, pad))
            .Append("Address".PadRight(25, pad))
            .Append("Phone".PadRight(15, pad))
            If ReportType = ReportType.Windows Or ReportType = ReportType.All Then
                .Append("Price".PadRight(7, pad))
                .Append("Spring".PadRight(7, pad))
                .Append("Fall".PadRight(7, pad))
            End If
            .Append("Notes".PadRight(20))
            .Append(vbCrLf)
            'next line
            .Append("----".PadRight(25, pad))
            .Append("-------".PadRight(25, pad))
            .Append("-----".PadRight(15, pad))
            .Append("-----".PadRight(7, pad))
            If ReportType = ReportType.Windows Or ReportType = ReportType.All Then
                .Append("------".PadRight(7, pad))
                .Append("----".PadRight(7, pad))
                .Append("-----")
            End If
            .Append(vbCrLf)
            .Append(vbCrLf)
        End With

        Header = sb.ToString

        For Each row As DataRow In ds.Tables(0).Rows
            sb.Append(ReturnFullName(row.Item("LastName").ToString, row.Item("FirstName").ToString))
            sb.Append(row.Item("Addr1").ToString.PadRight(25, pad))
            sb.Append(row.Item("Phone").ToString.PadRight(15, pad))
            If row.Item("WindowsCustomer").ToString = "1" Then
                If row.Item("WindowPrice").ToString.Length > 0 Then
                    sb.Append("$")
                End If
                sb.Append(row.Item("WindowPrice").ToString.PadRight(7, pad))
                sb.Append(ReturnCleanedWhenData(row.Item("WindowsCleanedWhen").ToString))
            End If
            sb.Append(row.Item("Notes").ToString.PadRight(10, pad).Substring(0, 10))
            sb.Append(vbCrLf)
            If Spacing = "Double" Then
                sb.Append(vbCrLf)
            End If
        Next

        If File.Exists(FileName) Then
            File.Delete(FileName)
        End If

        Dim sw As StreamWriter = File.AppendText(FileName)

        With sw
            ' Update the underlying file.
            .Write(sb.ToString)
            .Flush()
            .Close()
        End With

        System.Diagnostics.Process.Start("notepad.exe", FileName)

        ds.Dispose()


    End Sub

    Shared Sub PrintCustomersActiveReports(ByVal ds As DataSet, ByVal FileName As String, ByVal Spacing As String, ByVal ReportType As ReportType)


        'Dim row As DataRow
        Dim sb As New System.Text.StringBuilder
        Dim pad As Char = Convert.ToChar(" ")
        Dim Header As String

        'Write File header
        With sb
            .Append("Name".PadRight(25, pad))
            .Append("Address".PadRight(25, pad))
            .Append("Phone".PadRight(15, pad))
            If ReportType = ReportType.Windows Or ReportType = ReportType.All Then
                .Append("Price".PadRight(7, pad))
                .Append("Spring".PadRight(7, pad))
                .Append("Fall".PadRight(7, pad))
            End If
            .Append(vbCrLf)
            'next line
            .Append("----".PadRight(25, pad))
            .Append("-------".PadRight(25, pad))
            .Append("-----".PadRight(15, pad))
            '.Append("-----".PadRight(7, pad))
            If ReportType = ReportType.Windows Or ReportType = ReportType.All Then
                .Append("------".PadRight(7, pad))
                .Append("----".PadRight(7, pad))
            End If
            .Append(vbCrLf)
            .Append(vbCrLf)
        End With

        Header = sb.ToString

        sb = New System.Text.StringBuilder

        For Each row As DataRow In ds.Tables(0).Rows
            sb.Append(ReturnFullName(row.Item("LastName").ToString, row.Item("FirstName").ToString))
            sb.Append(row.Item("Addr1").ToString.PadRight(25, pad))
            sb.Append(row.Item("Phone").ToString.PadRight(15, pad))
            If row.Item("WindowsCustomer").ToString = "1" Then
                If row.Item("WindowPrice").ToString.Length > 0 Then
                    sb.Append("$")
                End If
                sb.Append(row.Item("WindowPrice").ToString.PadRight(7, pad))
                sb.Append(ReturnCleanedWhenData(row.Item("WindowsCleanedWhen").ToString))
            End If
            sb.Append(vbCrLf)
            If Spacing = "Double" Then
                sb.Append(vbCrLf)
            End If
        Next

        Dim CustRpt As New PrintCustomers
        Dim tempfrm As New ActRptPrvw

        With CustRpt
            .Header = Header
            .CustomerData = sb.ToString
            .Run()
            tempfrm.ARVeiwer.Document = .Document
            tempfrm.Text = "Customer Report"
        End With

        tempfrm.ShowDialog()

        ds.Dispose()


    End Sub

    Shared Sub PrintMaterials(ByVal Materials As DataSet, ByVal Year As String)

        Dim tempfrm As New ActRptPrvw
        Dim MaterialsRpt As New MaterialRpt
        Dim GT As Double

        Try
            If Materials.Tables(0).Rows.Count = 0 Then
                MessageBox.Show("No Reports for that year", "Materials Report")
                Exit Sub

            End If

            For Each row As DataRow In Materials.Tables(0).Rows
                GT += Convert.ToDouble(row.Item(1))
            Next

            With MaterialsRpt
                .CurrentDate = Now.ToShortDateString
                .GrandTotal = GT.ToString
                .Header = "Material Report For " & Year
                Dim dvMat As DataView
                dvMat = New DataView(Materials.Tables(0))
                .DataSource = dvMat
                .Run()
                tempfrm.ARVeiwer.Document = .Document
            End With

            tempfrm.Text = "Materials Report"

            tempfrm.ShowDialog()

        Finally

            Materials.Dispose()

        End Try

    End Sub

    Shared Sub ViewEditMiscCats(ByRef MdiForm As Form)

        For Each frm As Form In MdiForm.MdiChildren
            If TypeOf (frm) Is frmCustomerReport Then
                MessageBox.Show("That form is already open", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Hand)
                Exit Sub
            End If
        Next

        Dim frmView As New frmCategoriesEdit
        Dim MiscExp As New MiscExp((g_Settings))

        Dim MiscCats As DataSet = MiscExp.GetExpenseCategories

        With frmView
            .dsCategories = MiscCats
            .ugCats.DataSource = MiscCats
            .MdiParent = MdiForm
            .Show()
        End With

    End Sub

    Shared Sub ViewAddMaterials(ByRef MdiForm As Form)

        For Each frm As Form In MdiForm.MdiChildren
            If TypeOf (frm) Is frmMaterialsDataEntry Then
                MessageBox.Show("That form is already open", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Hand)
                Exit Sub
            End If
        Next

        Dim frmMaterials As New frmMaterialsDataEntry(frmMaterialsDataEntry.Mode.Add)
        With frmMaterials
            .MdiParent = MdiForm
            .Show()
        End With


    End Sub

    Shared Sub ViewEditMaterials(ByRef MdiForm As Form, ByVal CallingForm As frmViewMaterials, ByVal ds As DataSet)

        For Each frm As Form In MdiForm.MdiChildren
            If TypeOf (frm) Is frmMaterialsDataEntry Then
                MessageBox.Show("That form is already open", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Hand)
                Exit Sub
            End If
        Next

        Dim frmMaterials As New frmMaterialsDataEntry(frmMaterialsDataEntry.Mode.Edit)
        With frmMaterials
            .MdiParent = MdiForm
            .CallingForm = CallingForm
            .dsMaterialItem = ds
            .Show()
        End With


    End Sub

    Shared Sub ViewMiscExpItem(ByRef MdiForm As Form, ByVal CallingForm As frmMiscExpRpt, ByVal ds As DataSet)

        For Each frm As Form In MdiForm.MdiChildren
            If TypeOf (frm) Is frmMiscExpAdd Then
                MessageBox.Show("That form is already open", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Hand)
                Exit Sub
            End If
        Next

        Dim frmMiscExp As New frmMiscExpAdd(frmMiscExpAdd.Mode.Edit)
        With frmMiscExp
            .MdiParent = MdiForm
            .CallingForm = CallingForm
            .dsExpItem = ds
            .Show()
        End With

    End Sub

    Shared Sub ViewAddEditCustomer(ByVal MdiForm As Form, ByVal CusObj As Customer, ByVal CallingForm As frmCustomerReport)

        Dim frmEditCust As frmAddEditCustomer
        If Not IsNothing(CusObj) Then
            'Make sure not to open 2 sessions for same customer
            For Each frm As Form In MdiForm.MdiChildren

                If TypeOf (frm) Is frmAddEditCustomer Then
                    If DirectCast(frm, frmAddEditCustomer).txtFirstName.Text = CusObj.FirstName And _
                        DirectCast(frm, frmAddEditCustomer).txtLastName.Text = CusObj.LastName And _
                        DirectCast(frm, frmAddEditCustomer).txtAddr1.Text = CusObj.Addr1 Then
                        MessageBox.Show("Form is already open for " & CusObj.FirstName & " " & CusObj.LastName, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Hand)

                        Exit Sub

                    End If

                End If
            Next

        End If

        If IsNothing(CusObj) Then
            frmEditCust = New frmAddEditCustomer(frmAddEditCustomer.Mode.AddNew)
        Else
            frmEditCust = New frmAddEditCustomer(frmAddEditCustomer.Mode.Edit)
        End If

        With frmEditCust
            .Customer = CusObj
            .CallingForm = CallingForm
            .MdiParent = MdiForm
            .Show()
        End With

    End Sub

    Shared Sub ViewAddEditWindowCustomer(ByVal MdiForm As Form, ByVal CusObj As Customer, ByVal CallingForm As frmWindowsCustomersReport)

        Dim frmEditCust As frmAddEditCustomer
        If Not IsNothing(CusObj) Then
            'Make sure not to open 2 sessions for same customer
            For Each frm As Form In MdiForm.MdiChildren

                If TypeOf (frm) Is frmAddEditCustomer Then
                    If DirectCast(frm, frmAddEditCustomer).txtFirstName.Text = CusObj.FirstName And _
                        DirectCast(frm, frmAddEditCustomer).txtLastName.Text = CusObj.LastName And _
                        DirectCast(frm, frmAddEditCustomer).txtAddr1.Text = CusObj.Addr1 Then
                        MessageBox.Show("Form is already open for " & CusObj.FirstName & " " & CusObj.LastName, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Hand)

                        Exit Sub

                    End If

                End If
            Next

        End If

        If IsNothing(CusObj) Then
            frmEditCust = New frmAddEditCustomer(frmAddEditCustomer.Mode.AddNew)
        Else
            frmEditCust = New frmAddEditCustomer(frmAddEditCustomer.Mode.Edit)
        End If

        With frmEditCust
            .Customer = CusObj
            '.CallingForm = CallingForm
            .MdiParent = MdiForm
            .Show()
        End With

    End Sub

    Shared Sub ViewMaterialReport(ByRef MdiForm As Form)

        Dim frmMatRpt As New frmViewMaterials

        For Each frm As Form In MdiForm.MdiChildren
            If TypeOf (frm) Is frmViewMaterials Then
                MessageBox.Show("Form is already open", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Hand)
                Exit Sub
            End If
        Next

        With frmMatRpt
            .MdiParent = MdiForm
            .Show()

        End With
    End Sub

    Shared Sub ViewMiscExpReport(ByRef MdiForm As Form)

        Dim frmExpRpt As New frmMiscExpRpt

        For Each frm As Form In MdiForm.MdiChildren
            If TypeOf (frm) Is frmMiscExpRpt Then
                MessageBox.Show("Form is already open", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Hand)
                Exit Sub
            End If
        Next

        With frmExpRpt
            .MdiParent = MdiForm
            .Show()

        End With
    End Sub

    Shared Sub ViewMiscAdd(ByRef MdiForm As Form)

        Dim frmAddMisc As New frmMiscExpAdd(frmMiscExpAdd.Mode.Add)

        For Each frm As Form In MdiForm.MdiChildren
            If TypeOf (frm) Is frmMiscExpAdd Then
                MessageBox.Show("Form is already open", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Hand)
                Exit Sub
            End If
        Next

        With frmAddMisc
            .MdiParent = MdiForm
            .Show()
        End With
    End Sub

    Shared Sub ShowAbout()

        Dim frmAbout As New frmAbout
        frmAbout.ShowDialog()

    End Sub

    Shared Sub ViewEditCategory(ByVal CatID As Integer, ByVal DBMode As CrudMode)

        Dim MiscExpCat As MiscCategory

        If CatID <> 0 Then
            MiscExpCat = New MiscCategory(g_Settings, CatID)
        Else
            MiscExpCat = New MiscCategory(g_Settings)
        End If

        Dim frmView As New frmEditMiscCat

        With frmView
            .MiscCat = MiscExpCat
            Select Case DBMode
                Case CrudMode.DELETE
                    .MiscCat.CrudMode = MiscCategory.DBMode.Delete
                Case CrudMode.INSERT
                    .MiscCat.CrudMode = MiscCategory.DBMode.Insert
                Case CrudMode.UPDATE
                    .MiscCat.CrudMode = MiscCategory.DBMode.Update
            End Select
            .ShowDialog()
        End With

    End Sub

#End Region

    Shared Function ReturnFullName(ByVal strLastName As String, ByVal strFirstName As String) As String

        Dim pad As Char = Convert.ToChar(" ")
        Dim FullName As String = Trim(strLastName) & ", " & Trim(strFirstName)
        Return FullName.PadRight(25, pad)

    End Function

    Shared Function ReturnCleanedWhenData(ByVal strWhen As String) As String

        Dim pad As Char = Convert.ToChar(" ")

        Select Case strWhen
            Case "0" 'Spring
                Return " X".PadRight(14, pad)
            Case "1" 'Fall
                Return " ".PadRight(7, pad) & "X".PadRight(7, pad)
            Case "3", "6" 'Spring and Fall, and All
                Return " X".PadRight(7, pad) & "X".PadRight(7, pad)
            Case Else
                Return ""
        End Select

    End Function


End Class
