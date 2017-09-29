
#Region " Options "

Option Explicit On
Option Strict On

#End Region

#Region " Imports "

Imports AccessUtils
Imports System.Collections.Specialized

#End Region

Public Class Materials

#Region " Enums "

    Public Enum DBMode
        Insert = 0
        Update = 1
        Delete = 2
    End Enum

#End Region

#Region " Private Variables "

    Private Appsettings As New NameValueCollection
    Private _MaterialNo As Integer = 0
    Private _MaterialDesc As String = ""
    Private _MaterialCost As Double = 0
    Private _Notes As String = ""
    Private _Date As Date = Now
    Private _IsDirty As Boolean
    Private _isValid As Boolean
    Private WithEvents _BrokenRules As BrokenRules
    Public Event Valid(ByVal IsValid As Boolean)
    Private _CrudMode As DBMode = DBMode.Insert
    Public Event DisableView()
    Public Event Dirty()


#End Region

#Region " Properties "

    Public Property IsDirty() As Boolean

        Get
            Return Me._IsDirty

        End Get

        Set(ByVal Value As Boolean)

            Me._IsDirty = Value

        End Set

    End Property

    Public Property IsValid() As Boolean

        Get
            Return _isValid
        End Get

        Set(ByVal Value As Boolean)
            _isValid = Value

        End Set

    End Property

    Public Property PurchaseDate() As Date

        Get
            Return Me._Date

        End Get

        Set(ByVal Value As Date)

            Me.IsDirty = True

            Me._Date = Value

            Me.IsDirty = True
            RaiseEvent Dirty()

        End Set

    End Property

    Public Property MaterialNo() As Integer

        Get
            Return Me._MaterialNo
        End Get

        Set(ByVal Value As Integer)

            Me._MaterialNo = Value

        End Set

    End Property

    Public Property Notes() As String

        Get
            Return Me._Notes
        End Get

        Set(ByVal Value As String)

            Me.IsDirty = True
            RaiseEvent Dirty()
            Me._Notes = Value

        End Set

    End Property

    Public Property MaterialDesc() As String

        Get
            Return Me._MaterialDesc

        End Get

        Set(ByVal Value As String)

            Me._MaterialDesc = Value
            RaiseEvent Dirty()

            Me.IsDirty = True

            If Value.Length > 0 Then
                Me._BrokenRules.BrokenRule("Desc", False)
            Else
                Me._BrokenRules.BrokenRule("Desc", True)
            End If

        End Set

    End Property

    Public Property MaterialCost() As Double
        Get
            Return Me._MaterialCost
        End Get

        Set(ByVal Value As Double)

            Me._MaterialCost = Value
            Me.IsDirty = True
            RaiseEvent Dirty()


            If Value = 0 Then

                Me._BrokenRules.BrokenRule("Cost", True)
            Else
                Me._BrokenRules.BrokenRule("Cost", False)
            End If

        End Set
    End Property

    Public Property CrudMode() As DBMode

        Get
            Return Me._CrudMode
        End Get

        Set(ByVal Value As DBMode)
            Me._CrudMode = Value
            If Value = DBMode.Update Then
                RaiseEvent DisableView()
            End If
        End Set
    End Property

#End Region

#Region " Public Methods "

    Public Sub InsertUpdateDelete()


        Dim Params As New NameValueCollection
        Dim sbSQL As New System.Text.StringBuilder


        If Me.CrudMode = DBMode.Update Then
            With Params
                .Add("DatePurchased", Me.PurchaseDate.ToShortDateString)
                .Add("MaterialDesc", Me.MaterialDesc)
                .Add("MaterialCost", Me.MaterialCost.ToString)
                .Add("Notes", Me.Notes)
                .Add("MaterialNo", Me.MaterialNo.ToString)
            End With

            With sbSQL
                .Append("UPDATE Materials")
                .Append(" SET DatePurchased=@DatePurchased,")
                .Append("MaterialDesc=@MaterialDesc,MaterialCost=@MaterialCost,")
                .Append("Notes=@Notes")
                .Append(" WHERE MaterialNo=@MaterialNo")
            End With

        ElseIf Me.CrudMode = DBMode.Insert Then
            With Params
                .Add("DatePurchased", Me.PurchaseDate.ToShortDateString)
                .Add("MaterialDesc", Me.MaterialDesc)
                .Add("MaterialCost", Me.MaterialCost.ToString)
                .Add("Notes", Me.Notes)
            End With
            With sbSQL
                .Append("INSERT INTO Materials (")
                .Append("DatePurchased,MaterialDesc,MaterialCost,Notes)")
                .Append(" Values (")
                .Append("@DatePurchased,@MaterialDesc,@MaterialCost,@Notes)")
            End With
        ElseIf Me.CrudMode = DBMode.Delete Then
            With Params
                .Add("MaterialNo", Me.MaterialNo.ToString)
            End With
            With sbSQL
                .Append("DELETE FROM Materials")
                .Append(" WHERE MaterialNo=@MaterialNo")
            End With

        End If

        Utils.ExecuteNonQuery(Appsettings("Connstr"), sbSQL.ToString, Params)

        Me.IsDirty = False


    End Sub

    Public Function GetMoSummary(ByVal Date1 As Date, ByVal date2 As Date) As DataSet

        Dim Params As New NameValueCollection
        Dim sbSQL As New System.Text.StringBuilder
        Dim ds As DataSet
        Dim Row As DataRow
        Dim col As DataColumn

        Dim dsFinal As New DataSet

        dsFinal.Tables.Add("Summary")

        With dsFinal.Tables("Summary")
            col = .Columns.Add("MonthName", Type.GetType("System.String"))
            col = .Columns.Add("Total", Type.GetType("System.Double"))
        End With

        With Params
            .Add("Date1", Date1.ToShortDateString)
            .Add("Date2", date2.ToShortDateString)
        End With

        With sbSQL
            .Append("SELECT DatePart(""m"",datepurchased) as MonthName,")
            .Append("Sum(Materials.MaterialCost) AS SumOfMaterialCost  FROM(Materials)")
            .Append(" WHERE DatePurchased BETWEEN @Date1 AND @Date2")
            .Append(" GROUP BY DatePart(""m"",datepurchased);")
        End With

        ds = Utils.GetDataSet(Appsettings("Connstr"), sbSQL.ToString, Params)


        Dim i As Integer


        With ds.Tables(0)


            For Each d_row As DataRow In ds.Tables(0).Rows

                Row = dsFinal.Tables("Summary").NewRow

                Select Case .Rows(i).Item(0).ToString

                    Case "1"
                        Row.Item("MonthName") = "January"
                    Case "2"
                        Row.Item("MonthName") = "February"

                    Case "3"
                        Row.Item("MonthName") = "March"

                    Case "4"
                        Row.Item("MonthName") = "April"

                    Case "5"
                        Row.Item("MonthName") = "May"

                    Case "6"
                        Row.Item("MonthName") = "June"

                    Case "7"
                        Row.Item("MonthName") = "July"

                    Case "8"
                        Row.Item("MonthName") = "August"

                    Case "9"
                        Row.Item("MonthName") = "September"

                    Case "10"
                        Row.Item("MonthName") = "October"

                    Case "11"
                        Row.Item("MonthName") = "November"

                    Case "12"
                        Row.Item("MonthName") = "December"

                End Select

                Row.Item("Total") = Convert.ToDouble(.Rows(i).Item("SumOfMaterialCost"))
                dsFinal.Tables("Summary").Rows.Add(Row)

                i += 1
            Next

        End With

        Return dsFinal


    End Function

    Public Function GetMaterials(ByVal Date1 As Date, ByVal Date2 As Date) As DataSet


        Dim Params As New NameValueCollection
        Dim sbSQL As New System.Text.StringBuilder
        Dim ds As DataSet


        With Params
            .Add("Date1", Date1.ToShortDateString)
            .Add("Date2", Date2.ToShortDateString)
        End With

        With sbSQL
            .Append("SELECT MaterialNo,DatePurchased,MaterialDesc,MaterialCost,Notes FROM Materials")
            .Append(" WHERE DatePurchased BETWEEN @Date1 AND @Date2 ORDER BY DatePurchased")
        End With

        ds = Utils.GetDataSet(Appsettings("Connstr"), sbSQL.ToString, Params)

        ds.Tables(0).TableName = "Materials"

        Return ds


    End Function
    Public Function GetMaterials() As DataSet


        Dim sbSQL As New System.Text.StringBuilder
        Dim ds As DataSet

        With sbSQL
            .Append("SELECT MaterialNo,DatePurchased,MaterialDesc,MaterialCost,Notes FROM Materials ORDER BY DatePurchased")
        End With

        ds = Utils.GetDataSet(Appsettings("Connstr"), sbSQL.ToString, Nothing)

        ds.Tables(0).TableName = "Materials"

        Return ds


    End Function
    Public Function GetMaterialItem(ByVal MaterialNum As Integer) As DataSet


        Dim Params As New NameValueCollection
        Dim sbSQL As New System.Text.StringBuilder

        With Params
            .Add("MaterialNo", MaterialNum.ToString)
        End With

        With sbSQL
            .Append("SELECT MaterialNo,DatePurchased,MaterialDesc,MaterialCost,Notes FROM Materials")
            .Append(" WHERE MaterialNo=@MaterialNo")
        End With

        Return Utils.GetDataSet(Appsettings("Connstr"), sbSQL.ToString, Params)

    End Function

#End Region

#Region " Public Constructor "

    Public Sub New(ByVal Settings As NameValueCollection)

        Appsettings = Settings
        _BrokenRules = New BrokenRules

        With Me._BrokenRules
            .BrokenRule("Desc", True)
            .BrokenRule("Cost", True)

        End With
    End Sub

#End Region

#Region " Event Sinks "

    Private Sub _BrokenRules_ValidOrNot(ByVal isValid As Boolean) Handles _BrokenRules.ValidOrNot

        RaiseEvent Valid(isValid)
        Me.IsValid = isValid

    End Sub

#End Region

End Class


