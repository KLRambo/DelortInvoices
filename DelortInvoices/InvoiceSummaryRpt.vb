Imports System
Imports DataDynamics.ActiveReports
Imports DataDynamics.ActiveReports.Document

Public Class InvoiceSummaryRpt
Inherits ActiveReport
    Public dsReport As DataSet

    Public Sub New()
        MyBase.New()
        InitializeReport()
    End Sub
#Region "ActiveReports Designer generated code"
    Private WithEvents PageHeader As DataDynamics.ActiveReports.PageHeader = Nothing
    Private WithEvents Detail As DataDynamics.ActiveReports.Detail = Nothing
    Private WithEvents PageFooter As DataDynamics.ActiveReports.PageFooter = Nothing
	Private lblPageHeader As DataDynamics.ActiveReports.Label = Nothing
	Private txtDetails As DataDynamics.ActiveReports.TextBox = Nothing
	Private lblNumOfInv As DataDynamics.ActiveReports.Label = Nothing
	Private txtNoOfInv As DataDynamics.ActiveReports.TextBox = Nothing
	Private lblLaborSTotal As DataDynamics.ActiveReports.Label = Nothing
	Private txtLaborTotal As DataDynamics.ActiveReports.TextBox = Nothing
	Private txtMaterials As DataDynamics.ActiveReports.TextBox = Nothing
	Private lblTotalMat As DataDynamics.ActiveReports.Label = Nothing
	Private txtGrandTotal As DataDynamics.ActiveReports.TextBox = Nothing
	Private lblTotal As DataDynamics.ActiveReports.Label = Nothing
	Private Shape1 As DataDynamics.ActiveReports.Shape = Nothing
	Public Sub InitializeReport()
		Me.LoadLayout(Me.GetType, "DelortInvoices.InvoiceSummaryRpt.rpx")
		Me.PageHeader = CType(Me.Sections("PageHeader"),DataDynamics.ActiveReports.PageHeader)
		Me.Detail = CType(Me.Sections("Detail"),DataDynamics.ActiveReports.Detail)
		Me.PageFooter = CType(Me.Sections("PageFooter"),DataDynamics.ActiveReports.PageFooter)
		Me.lblPageHeader = CType(Me.PageHeader.Controls(0),DataDynamics.ActiveReports.Label)
		Me.txtDetails = CType(Me.Detail.Controls(0),DataDynamics.ActiveReports.TextBox)
		Me.lblNumOfInv = CType(Me.PageFooter.Controls(0),DataDynamics.ActiveReports.Label)
		Me.txtNoOfInv = CType(Me.PageFooter.Controls(1),DataDynamics.ActiveReports.TextBox)
		Me.lblLaborSTotal = CType(Me.PageFooter.Controls(2),DataDynamics.ActiveReports.Label)
		Me.txtLaborTotal = CType(Me.PageFooter.Controls(3),DataDynamics.ActiveReports.TextBox)
		Me.txtMaterials = CType(Me.PageFooter.Controls(4),DataDynamics.ActiveReports.TextBox)
		Me.lblTotalMat = CType(Me.PageFooter.Controls(5),DataDynamics.ActiveReports.Label)
		Me.txtGrandTotal = CType(Me.PageFooter.Controls(6),DataDynamics.ActiveReports.TextBox)
		Me.lblTotal = CType(Me.PageFooter.Controls(7),DataDynamics.ActiveReports.Label)
		Me.Shape1 = CType(Me.PageFooter.Controls(8),DataDynamics.ActiveReports.Shape)
	End Sub

#End Region
   

    Private Sub InvoiceSummaryRpt_DataInitialize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.DataInitialize

        Me.Fields.Add("DetailInfo")
        Me.Fields.Add("NumOfInv")
        Me.Fields.Add("TotalLabor")
        Me.Fields.Add("TotalMat")
        Me.Fields.Add("GrandTotal")


    End Sub

    Private Sub InvoiceSummaryRpt_FetchData(ByVal sender As Object, ByVal eArgs As DataDynamics.ActiveReports.ActiveReport.FetchEventArgs) Handles MyBase.FetchData

        'Dim row As DataRow
        Dim sb As New System.Text.StringBuilder
        Dim Materials As Double
        Dim strMaterials As String
        Dim Labor As Double
        Dim StrLabor As String
        Dim MoTotal As Double
        Dim strMoTotal As String
        Dim SpaceCount As Integer
        Try
            With sb
                .Append("Month")
                .Append(vbTab & vbTab & vbTab)
                .Append("Count")
                .Append(vbTab & vbTab)
                .Append("Labor")
                .Append(vbTab & vbTab & vbTab)
                .Append("Materials")
                .Append(vbTab & vbTab)
                .Append("Total")
                .Append(vbCrLf)
                .Append("--------------------------------------")
                .Append("-------------------------------------")
                .Append("------------------------------------")
                .Append(vbCrLf)
            End With


            For Each row As DataRow In Me.dsReport.Tables(0).Rows

                Labor = Convert.ToDouble(row.Item("Labor"))
                StrLabor = Labor.ToString("c")
                Materials = Convert.ToDouble(row.Item(3))
                strMaterials = Materials.ToString("c")
                MoTotal = Convert.ToDouble(row.Item(4))
                strMoTotal = MoTotal.ToString("c")

                With sb
                    SpaceCount = (9 - row.Item(0).ToString.Length)
                    .Append(row.Item(0).ToString & Space(SpaceCount))
                    .Append(vbTab & vbTab)
                    SpaceCount = (3 - row.Item(1).ToString.Length)
                    .Append(row.Item(1).ToString & Space(SpaceCount))
                    .Append(vbTab & vbTab)
                    SpaceCount = (11 - StrLabor.Length)
                    .Append(StrLabor & Space(SpaceCount))
                    .Append(vbTab & vbTab)
                    SpaceCount = (11 - strMaterials.Length)
                    .Append(strMaterials & Space(SpaceCount))
                    .Append(vbTab & vbTab)
                    SpaceCount = (11 - strMoTotal.Length)
                    .Append(strMoTotal & Space(SpaceCount))
                    .Append(vbCrLf & vbCrLf)
                End With
            Next

            Me.Fields("DetailInfo").Value = sb.ToString

            With Me.dsReport.Tables(0)
                Dim LastRow As Integer = .Rows.Count - 1
                Me.Fields("NumOfInv").Value = .Rows(LastRow).Item("InvoiceCount")
                Me.Fields("TotalLabor").Value = .Rows(LastRow).Item("LaborTotal")
                Me.Fields("TotalMat").Value = .Rows(LastRow).Item("MaterialTotal")
                Me.Fields("GrandTotal").Value = .Rows(LastRow).Item("MasterTotal")
            End With

        Catch ex As ArgumentException
            Throw New Exception("Report has too large of a number!")
        Catch ex As Exception
            Throw ex
        End Try

    End Sub



End Class
