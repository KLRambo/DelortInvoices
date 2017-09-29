Imports System
Imports DataDynamics.ActiveReports
Imports DataDynamics.ActiveReports.Document

Public Class PrintCustomers
    Inherits ActiveReport

    Public CustomerData As String
    Public Header As String

    Public Sub New()
        MyBase.New()
        InitializeReport()
    End Sub
#Region "ActiveReports Designer generated code"
    Private WithEvents PageHeader As DataDynamics.ActiveReports.PageHeader = Nothing
    Private WithEvents Detail As DataDynamics.ActiveReports.Detail = Nothing
    Private WithEvents PageFooter As DataDynamics.ActiveReports.PageFooter = Nothing
	Private lblHeader As DataDynamics.ActiveReports.Label = Nothing
	Private txtCustomers As DataDynamics.ActiveReports.TextBox = Nothing
	Public Sub InitializeReport()
		Me.LoadLayout(Me.GetType, "DelortInvoices.PrintCustomers.rpx")
		Me.PageHeader = CType(Me.Sections("PageHeader"),DataDynamics.ActiveReports.PageHeader)
		Me.Detail = CType(Me.Sections("Detail"),DataDynamics.ActiveReports.Detail)
		Me.PageFooter = CType(Me.Sections("PageFooter"),DataDynamics.ActiveReports.PageFooter)
		Me.lblHeader = CType(Me.PageHeader.Controls(0),DataDynamics.ActiveReports.Label)
		Me.txtCustomers = CType(Me.Detail.Controls(0),DataDynamics.ActiveReports.TextBox)
	End Sub

#End Region

    Private Sub Customers_DataInitialize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.DataInitialize

        Me.Fields.Add("CustomerData")
        Me.Fields.Add("Header")


    End Sub

    Private Sub Customers_FetchData(ByVal sender As Object, ByVal eArgs As DataDynamics.ActiveReports.ActiveReport.FetchEventArgs) Handles MyBase.FetchData
        Me.Fields("CustomerData").Value = Me.CustomerData
        Me.Fields("Header").Value = Me.Header

    End Sub
End Class
