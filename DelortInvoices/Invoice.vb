Imports System
Imports DataDynamics.ActiveReports
Imports DataDynamics.ActiveReports.Document
Imports System.Configuration.ConfigurationSettings

Public Class Invoice
    Inherits ActiveReport
    Public _Invoice As DelortBusObjects.Invoice
    Public _Customer As DelortBusObjects.Customer

	Public Sub New()
	MyBase.New()
        Me.InitializeReport()
        _Invoice = Nothing
        _Customer = Nothing
    End Sub
#Region "ActiveReports Designer generated code"
    Private WithEvents PageHeader As DataDynamics.ActiveReports.PageHeader = Nothing
    Private WithEvents Detail As DataDynamics.ActiveReports.Detail = Nothing
    Private WithEvents PageFooter As DataDynamics.ActiveReports.PageFooter = Nothing
	Private txtCustName As DataDynamics.ActiveReports.TextBox = Nothing
	Private LblInvoiceHeader As DataDynamics.ActiveReports.Label = Nothing
	Private lblInvNo As DataDynamics.ActiveReports.Label = Nothing
	Private txtInvNo As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtDate As DataDynamics.ActiveReports.TextBox = Nothing
	Private Shape1 As DataDynamics.ActiveReports.Shape = Nothing
	Private lblNameAddress As DataDynamics.ActiveReports.Label = Nothing
	Private Line1 As DataDynamics.ActiveReports.Line = Nothing
	Private txtDate0 As DataDynamics.ActiveReports.TextBox = Nothing
	Private lblLineItems As DataDynamics.ActiveReports.Label = Nothing
	Private txtLaborDesc0 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtLaborCost0 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtMaterialDesc0 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtMaterialCost0 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtLineTotal0 As DataDynamics.ActiveReports.TextBox = Nothing
	Private Line2 As DataDynamics.ActiveReports.Line = Nothing
	Private txtdate1 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtLaborDesc1 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtLaborCost1 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtMaterialDesc1 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtMaterialCost1 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtLineTotal1 As DataDynamics.ActiveReports.TextBox = Nothing
	Private lblLineDateHeader As DataDynamics.ActiveReports.Label = Nothing
	Private lblLaborDescHeader As DataDynamics.ActiveReports.Label = Nothing
	Private lblLaborCostHeader As DataDynamics.ActiveReports.Label = Nothing
	Private lblMatDescHeader As DataDynamics.ActiveReports.Label = Nothing
	Private lblMaterialCostHeader As DataDynamics.ActiveReports.Label = Nothing
	Private Label1 As DataDynamics.ActiveReports.Label = Nothing
	Private Label2 As DataDynamics.ActiveReports.Label = Nothing
	Private lblLineTotalHeader As DataDynamics.ActiveReports.Label = Nothing
	Private Line3 As DataDynamics.ActiveReports.Line = Nothing
	Private txtDate2 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtLaborDesc2 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtLaborCost2 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtMaterialDesc2 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtMaterialCost2 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtLineTotal2 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtDate3 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtLaborDesc3 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtLaborCost3 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtMaterialDesc3 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtMaterialCost3 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtLineTotal3 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtDate4 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtLaborDesc4 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtLaborCost4 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtMaterialDesc4 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtMaterialCost4 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtLineTotal4 As DataDynamics.ActiveReports.TextBox = Nothing
	Private Line5 As DataDynamics.ActiveReports.Line = Nothing
	Private Line4 As DataDynamics.ActiveReports.Line = Nothing
	Private MainBox As DataDynamics.ActiveReports.Shape = Nothing
	Private lblInvoiceTotal As DataDynamics.ActiveReports.TextBox = Nothing
	Private Label9 As DataDynamics.ActiveReports.Label = Nothing
	Private txtMaterialTotal As DataDynamics.ActiveReports.TextBox = Nothing
	Private Label10 As DataDynamics.ActiveReports.Label = Nothing
	Private txtLaborTotal As DataDynamics.ActiveReports.TextBox = Nothing
	Private Label11 As DataDynamics.ActiveReports.Label = Nothing
	Private Shape2 As DataDynamics.ActiveReports.Shape = Nothing
	Private Label12 As DataDynamics.ActiveReports.Label = Nothing
	Private Shape3 As DataDynamics.ActiveReports.Shape = Nothing
	Public Sub InitializeReport()
		Me.LoadLayout(Me.GetType, "DelortInvoices.Invoice.rpx")
		Me.PageHeader = CType(Me.Sections("PageHeader"),DataDynamics.ActiveReports.PageHeader)
		Me.Detail = CType(Me.Sections("Detail"),DataDynamics.ActiveReports.Detail)
		Me.PageFooter = CType(Me.Sections("PageFooter"),DataDynamics.ActiveReports.PageFooter)
		Me.txtCustName = CType(Me.PageHeader.Controls(0),DataDynamics.ActiveReports.TextBox)
		Me.LblInvoiceHeader = CType(Me.PageHeader.Controls(1),DataDynamics.ActiveReports.Label)
		Me.lblInvNo = CType(Me.PageHeader.Controls(2),DataDynamics.ActiveReports.Label)
		Me.txtInvNo = CType(Me.PageHeader.Controls(3),DataDynamics.ActiveReports.TextBox)
		Me.txtDate = CType(Me.PageHeader.Controls(4),DataDynamics.ActiveReports.TextBox)
		Me.Shape1 = CType(Me.PageHeader.Controls(5),DataDynamics.ActiveReports.Shape)
		Me.lblNameAddress = CType(Me.PageHeader.Controls(6),DataDynamics.ActiveReports.Label)
		Me.Line1 = CType(Me.Detail.Controls(0),DataDynamics.ActiveReports.Line)
		Me.txtDate0 = CType(Me.Detail.Controls(1),DataDynamics.ActiveReports.TextBox)
		Me.lblLineItems = CType(Me.Detail.Controls(2),DataDynamics.ActiveReports.Label)
		Me.txtLaborDesc0 = CType(Me.Detail.Controls(3),DataDynamics.ActiveReports.TextBox)
		Me.txtLaborCost0 = CType(Me.Detail.Controls(4),DataDynamics.ActiveReports.TextBox)
		Me.txtMaterialDesc0 = CType(Me.Detail.Controls(5),DataDynamics.ActiveReports.TextBox)
		Me.txtMaterialCost0 = CType(Me.Detail.Controls(6),DataDynamics.ActiveReports.TextBox)
		Me.txtLineTotal0 = CType(Me.Detail.Controls(7),DataDynamics.ActiveReports.TextBox)
		Me.Line2 = CType(Me.Detail.Controls(8),DataDynamics.ActiveReports.Line)
		Me.txtdate1 = CType(Me.Detail.Controls(9),DataDynamics.ActiveReports.TextBox)
		Me.txtLaborDesc1 = CType(Me.Detail.Controls(10),DataDynamics.ActiveReports.TextBox)
		Me.txtLaborCost1 = CType(Me.Detail.Controls(11),DataDynamics.ActiveReports.TextBox)
		Me.txtMaterialDesc1 = CType(Me.Detail.Controls(12),DataDynamics.ActiveReports.TextBox)
		Me.txtMaterialCost1 = CType(Me.Detail.Controls(13),DataDynamics.ActiveReports.TextBox)
		Me.txtLineTotal1 = CType(Me.Detail.Controls(14),DataDynamics.ActiveReports.TextBox)
		Me.lblLineDateHeader = CType(Me.Detail.Controls(15),DataDynamics.ActiveReports.Label)
		Me.lblLaborDescHeader = CType(Me.Detail.Controls(16),DataDynamics.ActiveReports.Label)
		Me.lblLaborCostHeader = CType(Me.Detail.Controls(17),DataDynamics.ActiveReports.Label)
		Me.lblMatDescHeader = CType(Me.Detail.Controls(18),DataDynamics.ActiveReports.Label)
		Me.lblMaterialCostHeader = CType(Me.Detail.Controls(19),DataDynamics.ActiveReports.Label)
		Me.Label1 = CType(Me.Detail.Controls(20),DataDynamics.ActiveReports.Label)
		Me.Label2 = CType(Me.Detail.Controls(21),DataDynamics.ActiveReports.Label)
		Me.lblLineTotalHeader = CType(Me.Detail.Controls(22),DataDynamics.ActiveReports.Label)
		Me.Line3 = CType(Me.Detail.Controls(23),DataDynamics.ActiveReports.Line)
		Me.txtDate2 = CType(Me.Detail.Controls(24),DataDynamics.ActiveReports.TextBox)
		Me.txtLaborDesc2 = CType(Me.Detail.Controls(25),DataDynamics.ActiveReports.TextBox)
		Me.txtLaborCost2 = CType(Me.Detail.Controls(26),DataDynamics.ActiveReports.TextBox)
		Me.txtMaterialDesc2 = CType(Me.Detail.Controls(27),DataDynamics.ActiveReports.TextBox)
		Me.txtMaterialCost2 = CType(Me.Detail.Controls(28),DataDynamics.ActiveReports.TextBox)
		Me.txtLineTotal2 = CType(Me.Detail.Controls(29),DataDynamics.ActiveReports.TextBox)
		Me.txtDate3 = CType(Me.Detail.Controls(30),DataDynamics.ActiveReports.TextBox)
		Me.txtLaborDesc3 = CType(Me.Detail.Controls(31),DataDynamics.ActiveReports.TextBox)
		Me.txtLaborCost3 = CType(Me.Detail.Controls(32),DataDynamics.ActiveReports.TextBox)
		Me.txtMaterialDesc3 = CType(Me.Detail.Controls(33),DataDynamics.ActiveReports.TextBox)
		Me.txtMaterialCost3 = CType(Me.Detail.Controls(34),DataDynamics.ActiveReports.TextBox)
		Me.txtLineTotal3 = CType(Me.Detail.Controls(35),DataDynamics.ActiveReports.TextBox)
		Me.txtDate4 = CType(Me.Detail.Controls(36),DataDynamics.ActiveReports.TextBox)
		Me.txtLaborDesc4 = CType(Me.Detail.Controls(37),DataDynamics.ActiveReports.TextBox)
		Me.txtLaborCost4 = CType(Me.Detail.Controls(38),DataDynamics.ActiveReports.TextBox)
		Me.txtMaterialDesc4 = CType(Me.Detail.Controls(39),DataDynamics.ActiveReports.TextBox)
		Me.txtMaterialCost4 = CType(Me.Detail.Controls(40),DataDynamics.ActiveReports.TextBox)
		Me.txtLineTotal4 = CType(Me.Detail.Controls(41),DataDynamics.ActiveReports.TextBox)
		Me.Line5 = CType(Me.Detail.Controls(42),DataDynamics.ActiveReports.Line)
		Me.Line4 = CType(Me.Detail.Controls(43),DataDynamics.ActiveReports.Line)
		Me.MainBox = CType(Me.Detail.Controls(44),DataDynamics.ActiveReports.Shape)
		Me.lblInvoiceTotal = CType(Me.PageFooter.Controls(0),DataDynamics.ActiveReports.TextBox)
		Me.Label9 = CType(Me.PageFooter.Controls(1),DataDynamics.ActiveReports.Label)
		Me.txtMaterialTotal = CType(Me.PageFooter.Controls(2),DataDynamics.ActiveReports.TextBox)
		Me.Label10 = CType(Me.PageFooter.Controls(3),DataDynamics.ActiveReports.Label)
		Me.txtLaborTotal = CType(Me.PageFooter.Controls(4),DataDynamics.ActiveReports.TextBox)
		Me.Label11 = CType(Me.PageFooter.Controls(5),DataDynamics.ActiveReports.Label)
		Me.Shape2 = CType(Me.PageFooter.Controls(6),DataDynamics.ActiveReports.Shape)
		Me.Label12 = CType(Me.PageFooter.Controls(7),DataDynamics.ActiveReports.Label)
		Me.Shape3 = CType(Me.PageFooter.Controls(8),DataDynamics.ActiveReports.Shape)
	End Sub

#End Region

    Private Sub Invoice_DataInitialize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.DataInitialize

        Fields.Add("CustomerName")
        Fields.Add("InvoiceNumber")
        Fields.Add("MaterialTotal")
        Fields.Add("LaborTotal")
        Fields.Add("InvoiceTotal")
        Fields.Add("InvoiceDate")

        'Line 0 
        Fields.Add("Date0")
        Fields.Add("LaborDesc0")
        Fields.Add("LaborCost0")
        Fields.Add("MaterialDesc0")
        Fields.Add("MaterialCost0")
        Fields.Add("LineTotal0")

        'line 1
        Fields.Add("Date1")
        Fields.Add("LaborDesc1")
        Fields.Add("LaborCost1")
        Fields.Add("MaterialDesc1")
        Fields.Add("MaterialCost1")
        Fields.Add("LineTotal1")

        'Line 2
        Fields.Add("Date2")
        Fields.Add("LaborDesc2")
        Fields.Add("LaborCost2")
        Fields.Add("MaterialDesc2")
        Fields.Add("MaterialCost2")
        Fields.Add("LineTotal2")

        'Line 3
        Fields.Add("Date3")
        Fields.Add("LaborDesc3")
        Fields.Add("LaborCost3")
        Fields.Add("MaterialDesc3")
        Fields.Add("MaterialCost3")
        Fields.Add("LineTotal3")

        'Line 4
        Fields.Add("Date4")
        Fields.Add("LaborDesc4")
        Fields.Add("LaborCost4")
        Fields.Add("MaterialDesc4")
        Fields.Add("MaterialCost4")
        Fields.Add("LineTotal4")

    End Sub

    Private Sub Invoice_ReportStart(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.ReportStart

        Me.PageSettings.Orientation = PageOrientation.Portrait
        Me.PageSettings.PaperKind = Printing.PaperKind.Letter
        Me.PageSettings.PaperWidth = 8.5F
        Me.PageSettings.PaperHeight = 11.0F
        Me.PageSettings.Margins.Left = 0.25F
        Me.PageSettings.Margins.Right = 0.25F
        Me.PageSettings.Margins.Bottom = 0.2F
        Me.PageSettings.Margins.Top = 0.2F

    End Sub
    Private Sub Invoice_FetchData(ByVal sender As Object, ByVal eArgs As DataDynamics.ActiveReports.ActiveReport.FetchEventArgs) Handles MyBase.FetchData

        Me.Fields("CustomerName").Value = Me.GetNameAndAddress
        Me.Fields("InvoiceNumber").Value = Me._Invoice.InvoiceNo.ToString
        Me.Fields("MaterialTotal").Value = Me._Invoice.MaterialTotal.ToString("c")
        Me.Fields("LaborTotal").Value = Me._Invoice.LaborTotal.ToString("c")
        Me.Fields("InvoiceTotal").Value = Me._Invoice.InvoiceTotal.ToString("c")
        Me.Fields("InvoiceDate").Value = Me._Invoice.InvoiceDate.ToShortDateString

        Dim LineNo As Integer

        For Each LineItem As DelortBusObjects.LineItem In Me._Invoice.LineItems
            With Me._Invoice
                Select Case LineNo
                    Case 0
                        Me.Fields("Date0").Value = .LineItems(0).ItemDate.ToShortDateString.Trim
                        Me.Fields("LaborDesc0").Value = .LineItems(0).LaborDesc.ToString
                        Me.Fields("LaborCost0").Value = .LineItems(0).LaborCost.ToString
                        Me.Fields("MaterialDesc0").Value = .LineItems(0).MaterialDesc.ToString
                        Me.Fields("MaterialCost0").Value = .LineItems(0).MaterialCost.ToString
                        Me.Fields("LineTotal0").Value = .LineItems(0).Total.ToString("c")
                    Case 1
                        Me.Fields("Date1").Value = .LineItems(1).ItemDate.ToShortDateString
                        Me.Fields("LaborDesc1").Value = .LineItems(1).LaborDesc.ToString
                        Me.Fields("LaborCost1").Value = .LineItems(1).LaborCost.ToString
                        Me.Fields("MaterialDesc1").Value = .LineItems(1).MaterialDesc.ToString
                        Me.Fields("MaterialCost1").Value = .LineItems(1).MaterialCost.ToString()
                        Me.Fields("LineTotal1").Value = .LineItems(1).Total.ToString("c")

                        'Make controls visible
                        Me.Line2.Visible = True

                    Case 2
                        Me.Fields("Date2").Value = .LineItems(2).ItemDate.ToShortDateString
                        Me.Fields("LaborDesc2").Value = .LineItems(2).LaborDesc.ToString
                        Me.Fields("LaborCost2").Value = .LineItems(2).LaborCost.ToString
                        Me.Fields("MaterialDesc2").Value = .LineItems(2).MaterialDesc.ToString
                        Me.Fields("MaterialCost2").Value = .LineItems(2).MaterialCost.ToString()
                        Me.Fields("LineTotal2").Value = .LineItems(2).Total.ToString("c")

                        'Make Line visible
                        Me.Line3.Visible = True


                    Case 3
                        Me.Fields("Date3").Value = .LineItems(3).ItemDate.ToShortDateString
                        Me.Fields("LaborDesc3").Value = .LineItems(3).LaborDesc.ToString
                        Me.Fields("LaborCost3").Value = .LineItems(3).LaborCost.ToString
                        Me.Fields("MaterialDesc3").Value = .LineItems(3).MaterialDesc.ToString
                        Me.Fields("MaterialCost3").Value = .LineItems(3).MaterialCost.ToString
                        Me.Fields("LineTotal3").Value = .LineItems(3).Total.ToString("c")

                        'Make controls visible
                        Me.Line4.Visible = True

                    Case 4
                        Me.Fields("Date4").Value = .LineItems(4).ItemDate.ToShortDateString
                        Me.Fields("LaborDesc4").Value = .LineItems(4).LaborDesc.ToString
                        Me.Fields("LaborCost4").Value = .LineItems(4).LaborCost.ToString
                        Me.Fields("MaterialDesc4").Value = .LineItems(4).MaterialDesc.ToString
                        Me.Fields("MaterialCost4").Value = .LineItems(4).MaterialCost.ToString
                        Me.Fields("LineTotal4").Value = .LineItems(4).Total.ToString("c")

                        'Make controls visible
                        Me.Line5.Visible = True
                End Select

            End With

            LineNo += 1

        Next

        With Me.lblNameAddress
            .Text = "John Delort & Sons" & vbCrLf
            .Text &= "Homeowner Services" & vbCrLf
            .Text &= "1716 Lynwood" & vbCrLf
            .Text &= "Crest Hill, IL 60403" & vbCrLf
            .Text &= "815-744-1280"
        End With


    End Sub
    Private Function FlipName(ByVal strName As String) As String

        Dim FullName() As String = strName.Split(CChar(","))

        Return FullName(1) & " " & FullName(0)

    End Function

    Private Function GetNameAndAddress() As String

        Dim sb As New System.Text.StringBuilder

        With sb
            .Append(Me._Customer.FirstName)
            .Append(" ")
            .Append(Me._Customer.LastName)
            .Append(vbCrLf)
            .Append(Me._Customer.Addr1)
            .Append(vbCrLf)
            If Me._Customer.Addr2.Length > 0 Then
                .Append(Me._Customer.Addr2)
                .Append(vbCrLf)
            End If
            .Append(Me._Customer.City)
            .Append(",")
            .Append(" ")
            .Append(Me._Customer.State)
            .Append("  ")
            .Append(Me._Customer.Zip)
            Return .ToString

        End With

    End Function

End Class
