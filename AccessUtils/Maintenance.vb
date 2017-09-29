Imports System
Imports System.IO
Public Class Maintenance

    Shared Sub CompactDB()

        Dim jro As JRO.JetEngine
        jro = New JRO.JetEngine

        'Now make a backup of the database
        File.Copy("Invoices.mdb", "InvoicesTemp.mdb", True)
        File.Delete("Invoices.mdb")

        jro.CompactDatabase("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=InvoicesTemp.mdb", _
        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Invoices.mdb;Jet OLEDB:Engine Type=5")

        File.Delete("InvoicesTemp.mdb")

    End Sub
End Class
