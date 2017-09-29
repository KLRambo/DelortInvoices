Imports System
Imports DataDynamics.ActiveReports
Imports DataDynamics.ActiveReports.Document

Public Class MaterialRpt

    Inherits ActiveReport

    Public Materials As New DataSet
    Public GrandTotal As String
    Public CurrentDate As String
    Public Header As String


    Public Sub New()

        MyBase.New()
        InitializeReport()

    End Sub

#Region "ActiveReports Designer generated code"
    Private WithEvents ReportHeader As DataDynamics.ActiveReports.ReportHeader = Nothing
    Private WithEvents PageHeader As DataDynamics.ActiveReports.PageHeader = Nothing
    Private WithEvents GroupHeader1 As DataDynamics.ActiveReports.GroupHeader = Nothing
    Private WithEvents Detail As DataDynamics.ActiveReports.Detail = Nothing
    Private WithEvents GroupFooter1 As DataDynamics.ActiveReports.GroupFooter = Nothing
    Private WithEvents PageFooter As DataDynamics.ActiveReports.PageFooter = Nothing
    Private WithEvents ReportFooter As DataDynamics.ActiveReports.ReportFooter = Nothing
	Private Label1 As DataDynamics.ActiveReports.Label = Nothing
	Private TextBox1 As DataDynamics.ActiveReports.TextBox = Nothing
	Private Label2 As DataDynamics.ActiveReports.Label = Nothing
	Private Label3 As DataDynamics.ActiveReports.Label = Nothing
	Private Label5 As DataDynamics.ActiveReports.Label = Nothing
	Private Label6 As DataDynamics.ActiveReports.Label = Nothing
	Private TextBox2 As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtdte As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtDate As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtDesc As DataDynamics.ActiveReports.TextBox = Nothing
	Private Line1 As DataDynamics.ActiveReports.Line = Nothing
	Private Label8 As DataDynamics.ActiveReports.Label = Nothing
	Private txtGrandTotal As DataDynamics.ActiveReports.TextBox = Nothing
	Public Sub InitializeReport()
		Me.LoadLayout(Me.GetType, "DelortInvoices.MaterialRpt.rpx")
		Me.ReportHeader = CType(Me.Sections("ReportHeader"),DataDynamics.ActiveReports.ReportHeader)
		Me.PageHeader = CType(Me.Sections("PageHeader"),DataDynamics.ActiveReports.PageHeader)
		Me.GroupHeader1 = CType(Me.Sections("GroupHeader1"),DataDynamics.ActiveReports.GroupHeader)
		Me.Detail = CType(Me.Sections("Detail"),DataDynamics.ActiveReports.Detail)
		Me.GroupFooter1 = CType(Me.Sections("GroupFooter1"),DataDynamics.ActiveReports.GroupFooter)
		Me.PageFooter = CType(Me.Sections("PageFooter"),DataDynamics.ActiveReports.PageFooter)
		Me.ReportFooter = CType(Me.Sections("ReportFooter"),DataDynamics.ActiveReports.ReportFooter)
		Me.Label1 = CType(Me.PageHeader.Controls(0),DataDynamics.ActiveReports.Label)
		Me.TextBox1 = CType(Me.PageHeader.Controls(1),DataDynamics.ActiveReports.TextBox)
		Me.Label2 = CType(Me.PageHeader.Controls(2),DataDynamics.ActiveReports.Label)
		Me.Label3 = CType(Me.PageHeader.Controls(3),DataDynamics.ActiveReports.Label)
		Me.Label5 = CType(Me.PageHeader.Controls(4),DataDynamics.ActiveReports.Label)
		Me.Label6 = CType(Me.PageHeader.Controls(5),DataDynamics.ActiveReports.Label)
		Me.TextBox2 = CType(Me.PageHeader.Controls(6),DataDynamics.ActiveReports.TextBox)
		Me.txtdte = CType(Me.PageHeader.Controls(7),DataDynamics.ActiveReports.TextBox)
		Me.txtDate = CType(Me.Detail.Controls(0),DataDynamics.ActiveReports.TextBox)
		Me.txtDesc = CType(Me.Detail.Controls(1),DataDynamics.ActiveReports.TextBox)
		Me.Line1 = CType(Me.Detail.Controls(2),DataDynamics.ActiveReports.Line)
		Me.Label8 = CType(Me.ReportFooter.Controls(0),DataDynamics.ActiveReports.Label)
		Me.txtGrandTotal = CType(Me.ReportFooter.Controls(1),DataDynamics.ActiveReports.TextBox)
	End Sub

#End Region

    Private Sub MaterialRpt_DataInitialize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.DataInitialize

        Me.Fields.Add("GrandTotal")
        Me.Fields.Add("CurrentDate")
        Me.Fields.Add("Header")

    End Sub

    Private Sub MaterialRpt_FetchData(ByVal sender As Object, ByVal eArgs As DataDynamics.ActiveReports.ActiveReport.FetchEventArgs) Handles MyBase.FetchData


        Me.Fields("GrandTotal").Value = Me.GrandTotal
        Me.Fields("CurrentDate").Value = Me.CurrentDate
        Me.Fields("Header").Value = Me.Header

    End Sub

    

    Private Sub MaterialRpt_ReportStart(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.ReportStart
        Me.PageSettings.Orientation = PageOrientation.Portrait
        '  Me.PageSettings.PaperKind = Printing.PaperKind.Letter
        Me.PageSettings.PaperWidth = 8.5F
        Me.PageSettings.PaperHeight = 9.75F
        Me.PageSettings.Margins.Left = 0.25F
        Me.PageSettings.Margins.Right = 0.25F
        Me.PageSettings.Margins.Bottom = 0.25F
        Me.PageSettings.Margins.Top = 0.25F
    End Sub
End Class
