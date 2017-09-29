
#Region " Options "

Option Explicit On
Option Strict On

#End Region

#Region " Imports "

Imports System.Collections.Specialized
Imports AccessUtils

#End Region

Public NotInheritable Class Invoice

#Region " Public Constructors "
    Public Sub New(ByVal Settings As NameValueCollection)

        Me._InvoiceNo = InvoiceNo
        Me.Appsettings = Settings

    End Sub
    Public Sub New()

    End Sub
#End Region

#Region " Enums "

    Public Enum DBMode
        Insert = 0
        Update = 1
        Delete = 2
    End Enum

#End Region

#Region " Private Variables "

    Private _InvoiceNo As Integer
    Private _CustID As Integer
    Private _InvoiceDate As Date
    Private _DueDate As Date
    Private _IsDirty As Boolean = False
    Private _Mode As DBMode = DBMode.Insert
    Private _CustomerName As String
    Private _LineItems As New LineItems
    Private Appsettings As New NameValueCollection
    Public Event Dirty(ByVal IsDirty As Boolean)


#End Region

#Region " Public Properties "

    Public Property CrudMode() As DBMode
        Get
            Return Me._Mode
        End Get
        Set(ByVal Value As DBMode)
            Me._Mode = Value

        End Set
    End Property
    Public Property CustomerName() As String
        Get
            Return _CustomerName
        End Get
        Set(ByVal Value As String)
            _CustomerName = Value

        End Set
    End Property

    Public Property InvoiceNo() As Integer
        Get
            Return Me._InvoiceNo
        End Get
        Set(ByVal Value As Integer)
            Me._InvoiceNo = Value
        End Set
    End Property
    Public Property CustID() As Integer
        Get
            Return Me._CustID
        End Get
        Set(ByVal Value As Integer)
            Me._CustID = Value

        End Set
    End Property

    Public Property InvoiceDate() As Date

        Get
            Return Me._InvoiceDate

        End Get

        Set(ByVal Value As Date)
            Me._InvoiceDate = Value

        End Set

    End Property

    Public Property DueDate() As Date

        Get
            Return Me._DueDate
        End Get

        Set(ByVal Value As Date)
            Me._DueDate = Value

        End Set
    End Property

    Public ReadOnly Property InvoiceTotal() As Double

        Get
            Dim Total As Double

            For Each Line As LineItem In Me.LineItems
                Total += Me.BlankToZero(Line.LaborCost) + Me.BlankToZero(Line.MaterialCost)
            Next

            Return Total

        End Get

    End Property

    Public ReadOnly Property LaborTotal() As Double
        Get
            Dim Total As Double

            For Each Li As LineItem In Me.LineItems
                Total += Me.BlankToZero(Li.LaborCost)
            Next
            Return Total

        End Get
    End Property

    Public ReadOnly Property MaterialTotal() As Double

        Get
            Dim Total As Double

            For Each Li As LineItem In Me.LineItems
                Total += Me.BlankToZero(Li.MaterialCost)
            Next
            Return Total
        End Get

    End Property
    Public Property LineItems() As LineItems

        Get
            Return Me._LineItems

        End Get

        Set(ByVal Value As LineItems)

            Me._LineItems = Value

        End Set

    End Property
    Public Property IsDirty() As Boolean

        Get
            Return Me._IsDirty
        End Get

        Set(ByVal Value As Boolean)
            Me._IsDirty = Value
            RaiseEvent Dirty(Value)

        End Set

    End Property

#End Region

#Region " Private Routines "

    Private Function BlankToZero(ByVal data As Object) As Double

        If IsNumeric(data.ToString) Then
            Return Convert.ToDouble(data.ToString)
        Else
            Return 0
        End If

    End Function

    Private Function GetLastDayInMonth(ByVal dDate As Date) As Date

        dDate = DateAdd(DateInterval.Month, 1, dDate)
        dDate = Convert.ToDateTime(Month(dDate).ToString() & "/" & "1/" & Year(dDate).ToString())
        dDate = DateAdd(DateInterval.Day, -1, dDate)

        Return dDate

    End Function

#End Region

#Region " Public Methods "

    Public Sub Delete(ByVal InvNo As Integer)

        Dim params As New NameValueCollection
        Dim strSQL As String

        strSQL = "DELETE FROM Invoices WHERE InvoiceNo=@InvoiceNo"

        With params
            .Add("InvoiceNo", InvNo.ToString)
        End With

        Utils.ExecuteNonQuery(Appsettings("Connstr"), strSQL, params)


    End Sub

    Public Sub InsertUpdateDelete(ByVal invNo As Integer)

        Dim LineItemUpdates As Integer = 0
        Dim LineItemDeletes As Integer = 0
        Dim LineItemInserts As Integer = 0

        Dim LineItemUpdateSQl As String = String.Empty
        Dim LineItemDeleteSQL As String = String.Empty
        Dim LineItemInsertSql As String = String.Empty

        Dim InvoiceParams As New NameValueCollection
        Dim sbInvoiceSQL As New System.Text.StringBuilder
        Dim sbLineItemInsert As New System.Text.StringBuilder
        Dim sbLineItemUpdate As New System.Text.StringBuilder
        Dim sbLineItemDelete As New System.Text.StringBuilder

        Dim LineItemCount As Integer = 0

        'Set the invoice up first, will be either UPDATE or INSERT
        If Me.CrudMode = DBMode.Update Then
            With InvoiceParams
                .Add("CustomerID", Me.CustID.ToString)
                .Add("InvoiceDate", Me.InvoiceDate.ToShortDateString)
                .Add("LaborTotal", Me.LaborTotal.ToString)
                .Add("MaterialTotal", Me.MaterialTotal.ToString)
                .Add("InvoiceTotal", Me.InvoiceTotal.ToString)
                .Add("InvoiceNo", invNo.ToString)
            End With
            With sbInvoiceSQL
                .Append("UPDATE Invoices")
                .Append(" SET CustomerID=@CustomerID,InvoiceDate=@InvoiceDate,LaborTotal=@LaborTotal,MaterialTotal=@MaterialTotal,InvoiceTotal=@InvoiceTotal")
                .Append(" WHERE InvoiceNo=@InvoiceNo")
            End With
        ElseIf Me.CrudMode = DBMode.Insert Then
            With InvoiceParams
                .Add("InvoiceNo", invNo.ToString)
                .Add("CustomerID", Me.CustID.ToString)
                .Add("InvoiceDate", Me.InvoiceDate.ToShortDateString)
                .Add("LaborTotal", Me.LaborTotal.ToString)
                .Add("MaterialTotal", Me.MaterialTotal.ToString)
                .Add("InvoiceTotal", Me.InvoiceTotal.ToString)
            End With
            With sbInvoiceSQL
                .Append("INSERT INTO Invoices (")
                .Append("InvoiceNo,CustomerID,InvoiceDate,LaborTotal,MaterialTotal,InvoiceTotal)")
                .Append(" Values (")
                .Append("@InvoiceNo,@CustomerID,@InvoiceDate,@LaborTotal,@MaterialTotal,@InvoiceTotal)")
            End With
        End If

        'Now check the LineItems to see what there
        For LineItemCount = 0 To Me.LineItems.Count - 1
            If Me.LineItems(LineItemCount).Mode = LineItem.Crud.Update Then
                LineItemUpdates += 1
            ElseIf Me.LineItems(LineItemCount).Mode = LineItem.Crud.Delete Then
                LineItemDeletes += 1
            ElseIf Me.LineItems(LineItemCount).Mode = LineItem.Crud.Insert Then
                LineItemInserts += 1
            End If
        Next

        LineItemCount = 0

        Dim LineItemParamsDelete(LineItemDeletes - 1) As NameValueCollection
        Dim LineItemParamsInsert(LineItemInserts - 1) As NameValueCollection
        Dim LineItemParamsUpdate(LineItemUpdates - 1) As NameValueCollection


        'Set up the update Params
        If LineItemUpdates > 0 Then
            With sbLineItemUpdate
                .Append("UPDATE LineItems")
                .Append(" SET ItemDate=@ItemDate,LaborDesc=@LaborDesc,LaborCost=@LaborCost,MaterialDesc=@MaterialDesc,MaterialCost=@MaterialCost,Total=@Total")
                .Append(" WHERE ItemNo=@ItemNo")
                LineItemUpdateSQl = .ToString
            End With

            For Each LineItem As LineItem In Me.LineItems
                If LineItem.Mode = LineItem.Crud.Update Then
                    LineItemParamsUpdate(LineItemCount) = New NameValueCollection
                    LineItemParamsUpdate(LineItemCount).Add("ItemDate", LineItem.ItemDate.ToShortDateString)
                    LineItemParamsUpdate(LineItemCount).Add("LaborDesc", LineItem.LaborDesc)
                    LineItemParamsUpdate(LineItemCount).Add("LaborCost", LineItem.LaborCost)
                    LineItemParamsUpdate(LineItemCount).Add("MaterialDesc", LineItem.MaterialDesc)
                    LineItemParamsUpdate(LineItemCount).Add("MaterialCost", LineItem.MaterialCost)
                    LineItemParamsUpdate(LineItemCount).Add("Total", LineItem.Total.ToString)
                    LineItemParamsUpdate(LineItemCount).Add("ItemNo", LineItem.ItemNo.ToString)
                    LineItemCount += 1
                End If
            Next
        End If

        LineItemCount = 0

        'Set up insert
        If LineItemInserts > 0 Then
            With sbLineItemInsert
                .Append("INSERT INTO LineItems (")
                .Append("InvoiceNo,ItemDate,LaborDesc,LaborCost,MaterialDesc,MaterialCost,Total )")
                .Append(" VALUES (")
                .Append("@InvoiceNo,@ItemDate,@LaborDesc,@LaborCost,@MaterialDesc,@MaterialCost,@Total)")
                LineItemInsertSql = .ToString
            End With

            For Each LineItem As LineItem In Me.LineItems
                If LineItem.Mode = LineItem.Crud.Insert Then
                    LineItemParamsInsert(LineItemCount) = New NameValueCollection
                    LineItemParamsInsert(LineItemCount).Add("InvoiceNo", invNo.ToString)
                    LineItemParamsInsert(LineItemCount).Add("ItemDate", LineItem.ItemDate.ToString)
                    LineItemParamsInsert(LineItemCount).Add("LaborDesc", LineItem.LaborDesc)
                    LineItemParamsInsert(LineItemCount).Add("LaborCost", LineItem.LaborCost)
                    LineItemParamsInsert(LineItemCount).Add("MaterialDesc", LineItem.MaterialDesc)
                    LineItemParamsInsert(LineItemCount).Add("MaterialCost", LineItem.MaterialCost)
                    LineItemParamsInsert(LineItemCount).Add("Total", LineItem.Total.ToString)
                    LineItemCount += 1
                End If
            Next

        End If

        LineItemCount = 0

        'Set up Deletes
        If LineItemDeletes > 0 Then

            With sbLineItemDelete
                .Append("DELETE FROM LineItems")
                .Append(" WHERE ItemNo=@ItemNo")
                LineItemDeleteSQL = .ToString
            End With

            For Each LineItem As LineItem In Me.LineItems
                If LineItem.Mode = LineItem.Crud.Delete Then
                    LineItemParamsDelete(LineItemCount) = New NameValueCollection
                    LineItemParamsDelete(LineItemCount).Add("ItemNo", LineItem.ItemNo.ToString)
                    LineItemCount += 1
                End If
            Next
        End If

        Try
            Utils.InsertUpdateDelete(Appsettings("Connstr"), sbInvoiceSQL.ToString, LineItemInsertSql, LineItemUpdateSQl, LineItemDeleteSQL, InvoiceParams, LineItemParamsInsert, LineItemParamsUpdate, LineItemParamsDelete)

        Catch ex As Exception
            If Me.CrudMode = DBMode.Insert Then
                Dim DeleteSQL As String = "Delete From Invoices Where InvoiceNo=" & invNo
                Utils.ExecuteNonQuery(Appsettings("Connstr"), DeleteSQL, Nothing)
            End If

            Throw ex

        End Try

    End Sub

    Public Function GetMonthlyReports(ByVal Date1 As Date, ByVal Date2 As Date) As DataSet

        Dim MonthTotalCount As Integer = Convert.ToInt32(DateDiff(DateInterval.Month, Date1, Date2))
        Dim params As New NameValueCollection
        Dim sbInv As New System.Text.StringBuilder
        Dim ds As New DataSet
        Dim dsFinal As New DataSet
        Dim Count As Integer
        Dim MonthStart As Date
        Dim MonthEnd As Date
        Dim Row As DataRow = Nothing
        Dim col As DataColumn

        Dim LaborTotal As Double
        Dim MaterialTotal As Double
        Dim MasterDollarTotal As Double
        Dim TotalInvCount As Integer

        dsFinal.Tables.Add("Summary")

        'TODO - put this into a Typed dataset
        With dsFinal.Tables("Summary")
            col = .Columns.Add("Month", Type.GetType("System.String"))
            col = .Columns.Add("Number", Type.GetType("System.Int32"))
            col = .Columns.Add("Labor", Type.GetType("System.Double"))
            col = .Columns.Add("Material", Type.GetType("System.Double"))
            col = .Columns.Add("Total", Type.GetType("System.Double"))
            col = .Columns.Add("MaterialTotal", Type.GetType("System.Double"))
            col = .Columns.Add("LaborTotal", Type.GetType("System.Double"))
            col = .Columns.Add("MasterTotal", Type.GetType("System.Double"))
            col = .Columns.Add("InvoiceCount", Type.GetType("System.Double"))
        End With

        For Count = 0 To MonthTotalCount
            If Count = 0 Then
                MonthStart = Date1
            Else
                MonthStart = DateAdd(DateInterval.Day, 1, MonthEnd)
            End If

            MonthEnd = Me.GetLastDayInMonth(MonthStart)

            With params
                .Clear()
                .Add("Date1", MonthStart.ToShortDateString)
                .Add("Date2", MonthEnd.ToShortDateString)
            End With

            sbInv = New System.Text.StringBuilder

            With sbInv
                .Append("Select Count(LaborTotal)as NumOfInv, Sum(LaborTotal) as LaborSum,Sum(MaterialTotal) as MaterialSum,Sum(InvoiceTotal) as InvoiceSum from Invoices where InvoiceDate Between @Date1 and @Date2")
            End With


            ds = Utils.GetDataSet(Appsettings("Connstr"), sbInv.ToString, params)

            ds.Tables(0).TableName = "Invoices"


            With dsFinal.Tables("Summary")
                Row = .NewRow
                'set values
                Row.Item("Month") = MonthStart.Month.ToString & "/" & MonthStart.Year.ToString
                Row.Item("Number") = Convert.ToInt32(ds.Tables(0).Rows(0).Item("NumOfInv"))

                If Not IsDBNull(ds.Tables(0).Rows(0).Item("LaborSum")) Then
                    Row.Item("Labor") = Convert.ToDouble(ds.Tables(0).Rows(0).Item("LaborSum"))
                Else
                    Row.Item("Labor") = 0
                End If

                If Not IsDBNull(ds.Tables(0).Rows(0).Item("MaterialSum")) Then
                    Row.Item("Material") = Convert.ToDouble(ds.Tables(0).Rows(0).Item("MaterialSum"))
                Else
                    Row.Item("Material") = 0
                End If
                If Not IsDBNull(ds.Tables(0).Rows(0).Item("InvoiceSum")) Then
                    Row.Item("Total") = Convert.ToDouble(ds.Tables(0).Rows(0).Item("InvoiceSum"))
                Else
                    Row.Item("Total") = 0
                End If

                .Rows.Add(Row)

            End With


            'Tally Report Totals
            With ds.Tables(0).Rows(0)
                If Not IsDBNull(.Item("LaborSum")) Then
                    LaborTotal += Convert.ToDouble(.Item("LaborSum"))
                End If
                If Not IsDBNull(.Item("MaterialSum")) Then
                    MaterialTotal += Convert.ToDouble(ds.Tables(0).Rows(0).Item("MaterialSum"))
                End If
                If Not IsDBNull(.Item("InvoiceSum")) Then
                    MasterDollarTotal += Convert.ToDouble(ds.Tables(0).Rows(0).Item("InvoiceSum"))

                End If
                If Not IsDBNull(.Item("NumOfInv")) Then
                    TotalInvCount += Convert.ToInt32(.Item("NumOfInv"))
                End If

            End With

        Next

        Row.Item("MaterialTotal") = MaterialTotal
        Row.Item("LaborTotal") = LaborTotal
        Row.Item("MasterTotal") = MasterDollarTotal
        Row.Item("InvoiceCount") = TotalInvCount

        Return dsFinal


    End Function

    Public Function GetCustomerInvoices(ByVal intCustID As Integer) As DataSet

        Dim params As New NameValueCollection
        Dim sbInv As New System.Text.StringBuilder
        Dim ds As New DataSet

        With params
            .Add("CustomerID", intCustID.ToString)
        End With

        With sbInv
            .Append("SELECT InvoiceNo,InvoiceDate,LaborTotal,MaterialTotal,InvoiceTotal FROM Invoices")
            .Append(" WHERE CustomerID = @CustomerID")
        End With

        ds = Utils.GetDataSet(Appsettings("Connstr"), sbInv.ToString, params)
        ds.Tables(0).TableName = "Invoices"

        Return ds

    End Function

    Public Function GetCustomerInvoices(ByVal intCustID As Integer, ByVal Date1 As String, ByVal Date2 As String) As DataSet

        Dim params As New NameValueCollection
        Dim sbInv As New System.Text.StringBuilder
        Dim ds As New DataSet

        With params
            .Add("CustomerID", intCustID.ToString)
            .Add("Date1", Date1)
            .Add("Date2", Date2)
        End With

        With sbInv
            .Append("SELECT InvoiceNo,InvoiceDate,LaborTotal,MaterialTotal,InvoiceTotal FROM Invoices")
            .Append(" WHERE CustomerID = @CustomerID ")
            .Append(" AND InvoiceDate BETWEEN @Date1 AND @Date2")
        End With


        ds = Utils.GetDataSet(Appsettings("Connstr"), sbInv.ToString, params)

        ds.Tables(0).TableName = "Invoices"

        Return ds

    End Function

    Public Function GetAllInvoices() As DataSet

        Dim sbInv As New System.Text.StringBuilder
        Dim ds As New DataSet

        With sbInv
            .Append("SELECT Customers.CustomerID,Customers.LastName,Customers.FirstName, Invoices.InvoiceNo, Invoices.InvoiceDate,Invoices.LaborTotal,Invoices.MaterialTotal,Invoices.InvoiceTotal ")
            .Append("FROM Customers ")
            .Append(" INNER JOIN Invoices")
            .Append(" ON Customers.CustomerID = Invoices.CustomerID")
            .Append(" ORDER BY InvoiceNo")
        End With

        ds = Utils.GetDataSet(Appsettings("Connstr"), sbInv.ToString, Nothing)
        ds.Tables(0).TableName = "Invoices"

        Return ds


    End Function

    Public Function GetAllInvoices(ByVal Date1 As String, ByVal Date2 As String) As DataSet

        Dim params As New NameValueCollection
        Dim sbInv As New System.Text.StringBuilder
        Dim ds As New DataSet

        With params
            .Add("Date1", Date1)
            .Add("Date2", Date2)
        End With

        With sbInv
            .Append("SELECT Customers.CustomerID,Customers.LastName,Customers.FirstName, Invoices.InvoiceNo, Invoices.InvoiceDate,Invoices.LaborTotal,Invoices.MaterialTotal,Invoices.InvoiceTotal ")
            .Append("FROM Customers")
            .Append(" INNER JOIN Invoices")
            .Append(" ON Customers.CustomerID = Invoices.CustomerID")
            .Append(" WHERE InvoiceDate BETWEEN @Date1 AND @Date2 ")
            .Append(" ORDER BY InvoiceNo")
        End With

        ds = Utils.GetDataSet(Appsettings("Connstr"), sbInv.ToString, params)
        ds.Tables(0).TableName = "Invoices"

        Return ds


    End Function

    Public Function GetInvoice(ByVal InvNo As Integer) As DataSet

        Dim params As New NameValueCollection
        Dim sbInv As New System.Text.StringBuilder

        With params
            .Add("InvoiceNo", InvNo.ToString)
        End With

        With sbInv
            .Append("SELECT Customers.CustomerID,Customers.LastName + ',' + Customers.FirstName as Name, Invoices.InvoiceNo, Invoices.InvoiceDate,Invoices.LaborTotal,Invoices.MaterialTotal, Invoices.InvoiceTotal")
            .Append(" FROM Invoices")
            .Append(" INNER JOIN Customers")
            .Append(" ON Invoices.CustomerID = Customers.CustomerID")
            .Append(" WHERE InvoiceNo=@InvoiceNo ")
        End With

        Return Utils.GetDataSet(Appsettings("Connstr"), sbInv.ToString, params)

    End Function

    Public Function GetInvoiceObject(ByVal InvNo As Integer) As Invoice

        Dim params As New NameValueCollection
        Dim sbInv As New System.Text.StringBuilder
        Dim dsInv As DataSet
        Dim dsLineItems As DataSet

        With params
            .Add("InvoiceNo", InvNo.ToString)
        End With

        With sbInv
            .Append("SELECT Customers.CustomerID,Customers.LastName+ ',' + Customers.FirstName as Name, Invoices.InvoiceNo, Invoices.InvoiceDate,Invoices.LaborTotal,Invoices.MaterialTotal,Invoices.InvoiceTotal ")
            .Append(" FROM Invoices")
            .Append(" INNER JOIN Customers")
            .Append(" ON Invoices.CustomerID = Customers.CustomerID")
            .Append(" WHERE InvoiceNo=@InvoiceNo ")
        End With

        dsInv = Utils.GetDataSet(Appsettings("Connstr"), sbInv.ToString, params)

        'see if there is such an invoice before we look for line items
        If dsInv.Tables(0).Rows.Count = 0 Then
            Return Nothing
        End If

        dsLineItems = Me.GetLineItems(InvNo)

        With dsInv.Tables(0).Rows(0)
            Me._CustomerName = .Item("Name").ToString
            Me._CustID = Convert.ToInt32(.Item("CustomerID"))
            Me.InvoiceDate = Convert.ToDateTime(.Item("InvoiceDate"))
            Me._InvoiceNo = InvNo
        End With

        For LineCount As Integer = 0 To dsLineItems.Tables(0).Rows.Count - 1
            Dim LI As New LineItem(Me.InvoiceNo)
            With dsLineItems.Tables(0)
                LI.ItemDate = Convert.ToDateTime(.Rows(LineCount).Item("ItemDate").ToString)
                LI.LaborCost = .Rows(LineCount).Item("LaborCost").ToString
                LI.LaborDesc = .Rows(LineCount).Item("LaborDesc").ToString
                LI.MaterialCost = .Rows(LineCount).Item("MaterialCost").ToString
                LI.MaterialDesc = .Rows(LineCount).Item("MaterialDesc").ToString
            End With
            Me.LineItems.Add(LI)
        Next

        Return Me

    End Function

    Public Function GetLineItems(ByVal invNo As Integer) As DataSet

        Dim params As New NameValueCollection
        Dim sbInv As New System.Text.StringBuilder

        With params
            .Add("InvoiceNo", invNo.ToString)
        End With

        With sbInv
            .Append("SELECT ItemNo,ItemDate,LaborDesc,LaborCost,MaterialDesc,MaterialCost,Total")
            .Append(" FROM LineItems")
            .Append(" WHERE InvoiceNo=@InvoiceNo ")
        End With

        Return Utils.GetDataSet(Appsettings("Connstr"), sbInv.ToString, params)

    End Function

    Public Function GetLatestInvoiceNo() As Integer

        Try
            Dim InvoiceNo As Integer
            InvoiceNo = Convert.ToInt32(Utils.ExecuteScalar(Appsettings("Connstr"), "SELECT MAX(InvoiceNo) FROM Invoices", Nothing))

            Return InvoiceNo + 1

        Catch ex As Exception
            Return 1
        End Try

    End Function

#End Region

End Class
