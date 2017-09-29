
#Region " Options "

Option Explicit On
Option Strict On

#End Region

#Region " Imports "

Imports AccessUtils
Imports System.Collections.Specialized
Imports System.Data.OleDb
#End Region

Public Class MiscExp

#Region " Constructor(s) "

    Public Sub New(ByVal Settings As NameValueCollection)

        Appsettings = Settings
        _BrokenRules = New BrokenRules

        With Me._BrokenRules
            .BrokenRule("Desc", True)
            .BrokenRule("Cost", True)
            .BrokenRule("Type", True)
        End With
    End Sub

#End Region

#Region " Private Variables "

    Private Appsettings As New NameValueCollection
    Private _IsDirty As Boolean
    Private _isValid As Boolean
    Private WithEvents _BrokenRules As BrokenRules
    Private _CrudMode As DBMode = DBMode.Insert
    Private _ItemNo As Integer
    Private _Date As Date = Now
    Private _Description As String
    Private _Cost As Double
    Private _Notes As String = ""
    Private _Category As Integer = 0


#End Region

#Region " Enums "

    Public Enum DBMode
        Insert = 0
        Update = 1
        Delete = 2
    End Enum

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

    Public Property IsValid() As Boolean

        Get
            Return _isValid
        End Get

        Set(ByVal Value As Boolean)
            _isValid = Value

        End Set

    End Property
    Public Property ItemNo() As Integer

        Get
            Return Me._ItemNo
        End Get

        Set(ByVal Value As Integer)

            Me._ItemNo = Value

        End Set

    End Property

    Public Property Notes() As String

        Get
            Return Me._Notes
        End Get

        Set(ByVal Value As String)

            Me._Notes = Value
            Me.IsDirty = True
            RaiseEvent Dirty()


        End Set

    End Property
    Public Property ExpDate() As Date
        Get
            Return _Date
        End Get
        Set(ByVal Value As Date)
            _Date = Value
            Me.IsDirty = True
            RaiseEvent Dirty()

        End Set
    End Property

    Public Property Description() As String

        Get
            Return Me._Description
        End Get

        Set(ByVal Value As String)

            Me._Description = Value
            RaiseEvent Dirty()

            Me.IsDirty = True

            If Value.Length > 0 Then
                Me._BrokenRules.BrokenRule("Desc", False)
            Else
                Me._BrokenRules.BrokenRule("Desc", True)
            End If
        End Set

    End Property
    Public Property Cost() As Double

        Get
            Return Me._Cost
        End Get

        Set(ByVal Value As Double)
            Me._Cost = Value

            RaiseEvent Dirty()
            Me.IsDirty = True

            If Value = 0 Then
                Me._BrokenRules.BrokenRule("Cost", True)
            Else
                Me._BrokenRules.BrokenRule("Cost", False)
            End If

        End Set
    End Property
    Public Property Category() As Integer

        Get
            Return Me._Category
        End Get

        Set(ByVal Value As Integer)

            Me._Category = Value
            RaiseEvent Dirty()

            Me.IsDirty = True

            If Value = 0 Then
                Me._BrokenRules.BrokenRule("Type", True)
            Else
                Me._BrokenRules.BrokenRule("Type", False)
            End If
        End Set

    End Property

#End Region

#Region " Events "
    Public Event Valid(ByVal IsValid As Boolean)
    Public Event Dirty()
    Public Event DisableView()

#End Region

#Region " Public Routines "

    Public Sub InsertUpdateDelete()

        Dim Params As New NameValueCollection
        Dim sbSQL As New System.Text.StringBuilder


        If Me.CrudMode = DBMode.Update Or Me.CrudMode = DBMode.Insert Then
            With Params
                .Add("Date", Me.ExpDate.ToShortDateString)
                .Add("Description", Me.Description)
                .Add("Cost", Me.Cost.ToString)
                .Add("Notes", Me.Notes)
                .Add("CategoryID", Me.Category.ToString)
                If Me.CrudMode = DBMode.Update Then
                    .Add("ItemNo", Me.ItemNo.ToString)
                End If
            End With
        End If

        If Me.CrudMode = DBMode.Update Then
            With sbSQL
                .Append("UPDATE MiscExps")
                .Append(" SET DateOfExp=@Date,")
                .Append("Description=@Description,Cost=@Cost,")
                .Append("Notes=@Notes,CategoryID=@CategoryID")
                .Append(" WHERE ItemNo=@ItemNo")
            End With

        ElseIf Me.CrudMode = DBMode.Insert Then
            With sbSQL
                .Append("INSERT INTO MiscExps (")
                .Append("DateOfExp,Description,Cost,Notes,CategoryID)")
                .Append(" Values (")
                .Append("@Date,@Description,@Cost,@Notes,@CategoryID)")
            End With
        ElseIf Me.CrudMode = DBMode.Delete Then
            With Params
                .Add("ItemNo", Me.ItemNo.ToString)
            End With
            With sbSQL
                .Append("DELETE FROM MiscExps")
                .Append(" WHERE ItemNo=@ItemNo")
            End With

        End If


        Utils.ExecuteNonQuery(Appsettings("Connstr"), sbSQL.ToString, Params)

        Me.IsDirty = False


    End Sub
    Public Function GetMiscExps(ByVal Date1 As String, ByVal Date2 As String) As DataSet


        Dim Params As New NameValueCollection
        Dim sbSQL As New System.Text.StringBuilder
        Dim ds As DataSet

        With Params
            .Add("Date1", Date1)
            .Add("Date2", Date2)
        End With

        With sbSQL
            .Append("Select MiscExps.ItemNo,MiscExps.DateOfExp, ")
            .Append("MiscExpCategories.Description AS Category, ")
            .Append("MiscExps.Description, MiscExps.Cost,MiscExps.Notes ")
            .Append("FROM MiscExps ")
            .Append("INNER JOIN MiscExpCategories ON ")
            .Append("MiscExps.CategoryID = MiscExpCategories.CategoryID")
            .Append(" WHERE DateOfExp BETWEEN @Date1 AND @Date2 ORDER BY DateOfExp")
        End With


        ds = Utils.GetDataSet(Appsettings("Connstr"), sbSQL.ToString, Params)

        ds.Tables(0).TableName = "MiscExp"

        Return ds


    End Function
    Public Function GetMiscExps() As DataSet


        Dim sbSQL As New System.Text.StringBuilder
        Dim ds As DataSet

        With sbSQL
            .Append("Select MiscExps.ItemNo,MiscExps.DateOfExp, ")
            .Append("MiscExpCategories.Description AS Category, ")
            .Append("MiscExps.Description, MiscExps.Cost, MiscExps.Notes ")
            .Append("FROM MiscExps ")
            .Append("INNER JOIN MiscExpCategories ON ")
            .Append("MiscExps.CategoryID = MiscExpCategories.CategoryID")
            .Append(" ORDER BY DateOfExp")
        End With

        ds = Utils.GetDataSet(Appsettings("Connstr"), sbSQL.ToString, Nothing)

        ds.Tables(0).TableName = "MiscExp"

        Return ds


    End Function
    Public Function GetMiscExpItem(ByVal ExpNum As Integer) As DataSet

        Dim Params As New NameValueCollection
        Dim sbSQL As New System.Text.StringBuilder

        With Params
            .Add("ItemNo", ExpNum.ToString)
        End With

        With sbSQL
            .Append("SELECT * FROM MiscExps")
            .Append(" WHERE ItemNo=@ItemNo")
        End With

        Return Utils.GetDataSet(Appsettings("Connstr"), sbSQL.ToString, Params)

    End Function
    Public Function GetExpenseTypesAll() As DataSet

        Dim strSQl As String = "Select CategoryID, Description from MiscExpCategories "

        Return Utils.GetDataSet(Appsettings("Connstr"), strSQl, Nothing)


    End Function

    Public Function GetExpenseTypesActive() As DataSet

        Dim strSQl As String = "Select CategoryID, Description from MiscExpCategories WHERE Active=1"

        Return Utils.GetDataSet(Appsettings("Connstr"), strSQl, Nothing)

    End Function

    Public Function GetExpenseCategories() As DataSet

        Dim strSQl As String = "Select CategoryID, Description, Active from MiscExpCategories WHERE CategoryID <> 0"
        Return Utils.GetDataSet(Appsettings("Connstr"), strSQl, Nothing)

    End Function
    Public Function GetGroupSummaryReport() As DataSet

        Dim sbSQL As New System.Text.StringBuilder

        With sbSQL
            .Append("SELECT SUM(MiscExps.Cost) AS Total, MiscExpCategories.CategoryID, MiscExpCategories.Description FROM MiscExps ")
            .Append("RIGHT JOIN MiscExpCategories ON MiscExpCategories.CategoryID = MiscExps.CategoryID ")
            .Append("GROUP BY MiscExps.CategoryID, MiscExpCategories.Description,MiscExpCategories.CategoryID ORDER BY MiscExpCategories.CategoryID ;")
        End With

        Return Utils.GetDataSet(Appsettings("Connstr"), sbSQL.ToString, Nothing)

    End Function
    Public Function GetGroupSummaryReport(ByVal FromDate As String, ByVal ToDate As String) As DataSet

        Dim sbSQL As New System.Text.StringBuilder
        Dim Params As New NameValueCollection

        With Params
            .Add("FromDate", FromDate)
            .Add("ToDate", ToDate)

        End With

        With sbSQL
            .Append("Select SUM(MiscExps.Cost) AS Total, MiscExpCategories.CategoryID, MiscExpCategories.Description FROM MiscExps ")
            .Append("RIGHT JOIN MiscExpCategories ON MiscExpCategories.CategoryID = MiscExps.CategoryID ")
            .Append("WHERE DateOfExp BETWEEN @FromDate AND @ToDate ")
            .Append("GROUP BY  MiscExps.CategoryID, MiscExpCategories.Description,MiscExpCategories.CategoryID ORDER BY MiscExpCategories.CategoryID ;")
        End With

        Return Utils.GetDataSet(Appsettings("Connstr"), sbSQL.ToString, Params)

    End Function
    Public Function GetMoSummaryReport(ByVal Year As String) As DataSet

        Dim sbSQL As New System.Text.StringBuilder
        Dim Params As New NameValueCollection
        Dim FromDate As String = "01/01/" & Year
        Dim ToDate As String = "12/31/" & Year

        With Params
            .Add("FromDate", FromDate)
            .Add("ToDate", ToDate)
        End With

        With sbSQL
            .Append("SELECT Month(MiscExps.DateOfExp) AS Mo, Sum(MiscExps.Cost) AS MoSum ")
            .Append("FROM MiscExps ")
            .Append("WHERE MiscExps.DateOfExp Between @FromDate And @ToDate ")
            .Append("GROUP BY Month(MiscExps.DateOfExp) ")
            .Append("ORDER BY Month(MiscExps.DateOfExp);")

        End With

        Return Utils.GetDataSet(Appsettings("Connstr"), sbSQL.ToString, Params)

    End Function

#End Region

#Region " Event Sinks "

    Private Sub _BrokenRules_ValidOrNot(ByVal isValid As Boolean) Handles _BrokenRules.ValidOrNot

        RaiseEvent Valid(isValid)
        Me._isValid = isValid

    End Sub

#End Region

End Class
