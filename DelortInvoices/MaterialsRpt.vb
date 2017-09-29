Option Explicit On
Option Strict On


Imports System.Collections.Specialized
Imports AccessUtils
Imports Microsoft.Reporting.WinForms


Public Class MaterialsRpt

    Public FromDate As String
    Public ToDate As String


    Private Sub MaterialsRpt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim ds As New DataSet
        Dim params As New NameValueCollection
        Dim sbMaterials As New System.Text.StringBuilder

        With params
            .Add("Date1", Me.FromDate)
            .Add("Date2", Me.ToDate)
        End With

        With sbMaterials
            .Append("SELECT MaterialNo,DatePurchased,MaterialDesc,MaterialCost,Notes from Materials")
            .Append(" WHERE DatePurchased BETWEEN @Date1 AND @Date2 ")
            .Append(" ORDER BY DatePurchased")
        End With

        ds = Utils.GetDataSet(My.Settings.DBConn, sbMaterials.ToString, params)

        If ds.Tables(0).Rows.Count > 0 Then

            Me.MaterialLinesBindingSource.DataMember = "MaterialLines"
            Me.MaterialLinesBindingSource.DataSource = ds.Tables(0)

            Me.ReportViewer1.RefreshReport()
        Else
            MessageBox.Show("No items for selected dates", ProgramName)
            Me.Close()

        End If

    End Sub

    ''' <summary>
    ''' This code not being used, just for future reference
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetReportParameters()

        'Set Processing Mode
        ReportViewer1.ProcessingMode = ProcessingMode.Local

        ' Set report server and report path
        ReportViewer1.LocalReport.ReportPath = "C:\Kevin\DelortInvoices\DelortInvoices\MaterialsRpt.rdlc"


        Dim pInfo As ReportParameterInfoCollection
        Dim paramList As New Generic.List(Of ReportParameter)

        paramList.Add(New ReportParameter("EmpID", "288", False))
        paramList.Add(New ReportParameter("ReportMonth", "12", False))
        paramList.Add(New ReportParameter("ReportYear", "2003", False))

        ReportViewer1.LocalReport.SetParameters(paramList)


        pInfo = ReportViewer1.LocalReport.GetParameters()

        ' Process and render the report
        ReportViewer1.RefreshReport()

    End Sub

End Class