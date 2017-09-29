<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MaterialsRpt
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim ReportDataSource1 As Microsoft.Reporting.WinForms.ReportDataSource = New Microsoft.Reporting.WinForms.ReportDataSource()
        Me.MaterialLinesBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.MaterialRptDS = New DelortInvoices.MaterialRptDS()
        Me.ReportViewer1 = New Microsoft.Reporting.WinForms.ReportViewer()
        Me.MaterialsBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.InvoicesDataSet = New DelortInvoices.InvoicesDataSet()
        Me.MaterialsTableAdapter = New DelortInvoices.InvoicesDataSetTableAdapters.MaterialsTableAdapter()
        CType(Me.MaterialLinesBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MaterialRptDS, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MaterialsBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.InvoicesDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MaterialLinesBindingSource
        '
        Me.MaterialLinesBindingSource.DataMember = "MaterialLines"
        Me.MaterialLinesBindingSource.DataSource = Me.MaterialRptDS
        '
        'MaterialRptDS
        '
        Me.MaterialRptDS.DataSetName = "MaterialRptDS"
        Me.MaterialRptDS.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'ReportViewer1
        '
        Me.ReportViewer1.Dock = System.Windows.Forms.DockStyle.Fill
        ReportDataSource1.Name = "dsMaterials"
        ReportDataSource1.Value = Me.MaterialLinesBindingSource
        Me.ReportViewer1.LocalReport.DataSources.Add(ReportDataSource1)
        Me.ReportViewer1.LocalReport.ReportEmbeddedResource = "DelortInvoices.MaterialsRpt.rdlc"
        Me.ReportViewer1.Location = New System.Drawing.Point(0, 0)
        Me.ReportViewer1.Name = "ReportViewer1"
        Me.ReportViewer1.Size = New System.Drawing.Size(813, 512)
        Me.ReportViewer1.TabIndex = 0
        '
        'MaterialsBindingSource
        '
        Me.MaterialsBindingSource.DataMember = "Materials"
        Me.MaterialsBindingSource.DataSource = Me.InvoicesDataSet
        '
        'InvoicesDataSet
        '
        Me.InvoicesDataSet.DataSetName = "InvoicesDataSet"
        Me.InvoicesDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'MaterialsTableAdapter
        '
        Me.MaterialsTableAdapter.ClearBeforeFill = True
        '
        'MaterialsRpt
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(813, 512)
        Me.Controls.Add(Me.ReportViewer1)
        Me.Name = "MaterialsRpt"
        Me.Text = "Materials Report"
        CType(Me.MaterialLinesBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MaterialRptDS, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MaterialsBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.InvoicesDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents InvoicesDataSet As DelortInvoices.InvoicesDataSet
    Friend WithEvents MaterialsBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents MaterialsTableAdapter As DelortInvoices.InvoicesDataSetTableAdapters.MaterialsTableAdapter
    Friend WithEvents ReportViewer1 As Microsoft.Reporting.WinForms.ReportViewer
    Friend WithEvents MaterialLinesBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents MaterialRptDS As DelortInvoices.MaterialRptDS
End Class
