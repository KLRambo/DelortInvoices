
#Region " Options "

Option Explicit On 
Option Strict On

#End Region

#Region " Imports "

Imports AccessUtils
Imports System.Collections.Specialized

#End Region

Public NotInheritable Class Check

#Region " Enums "

    Public Enum DBMode
        Insert = 0
        Update = 1
        Delete = 2
    End Enum

#End Region

#Region " Private variables "

    Private _Type As String
    Private _Number As String
    Private _Amount As Double
    Private _TransDate As Date
    Private _CrudMode As DBMode
    Private _LedgerID As Integer
    Private _Balance As Double
    Private _Description As String = ""
    Private Appsettings As New NameValueCollection
    Private WithEvents _BrokenRules As New BrokenRules

#End Region

#Region " Events "

    Public Event Valid(ByVal isValid As Boolean)

#End Region

#Region " Properties "
    Public isDirty As Boolean = False

    Public Property CrudMode() As DBMode

        Get
            Return Me._CrudMode
        End Get

        Set(ByVal Value As DBMode)
            Me._CrudMode = Value
        End Set
    End Property

    Public Property LedgerID() As Integer
        Get
            Return Me._LedgerID
        End Get

        Set(ByVal Value As Integer)
            Me._LedgerID = Value
            isDirty = True
        End Set

    End Property

    Public Property Number() As String
        Get
            Return Me._Number
        End Get

        Set(ByVal Value As String)

            If Value.Length > 0 Then
                Me._BrokenRules.BrokenRule("Number", False)

            Else
                Me._BrokenRules.BrokenRule("Number", True)
            End If

            Me._Number = Value
            isDirty = True

        End Set
    End Property

    Public Property Amount() As Double
        Get
            Return _Amount

        End Get

        Set(ByVal Value As Double)

            If Value <> 0 Then
                Me._BrokenRules.BrokenRule("Amount", False)

            Else
                Me._BrokenRules.BrokenRule("Amount", True)
            End If

            Me._Amount = Value
            isDirty = True

        End Set
    End Property

    Public Property Type() As String

        Get
            Return Me._Type
        End Get

        Set(ByVal Value As String)

            If Value.Length > 0 Then
                Me._BrokenRules.BrokenRule("Type", False)

            Else
                Me._BrokenRules.BrokenRule("Type", True)
            End If

            Me._Type = Value

        End Set

    End Property

    Public Property Balance() As Double
        Get
            Return Me._Balance
        End Get
        Set(ByVal Value As Double)
            Me._Balance = Value
        End Set
    End Property

    Public Property TransDate() As Date
        Get
            Return Me._TransDate
        End Get

        Set(ByVal Value As Date)

            If IsDate(Value) Then
                Me._BrokenRules.BrokenRule("TransDate", False)

            Else
                Me._BrokenRules.BrokenRule("TransDate", True)
            End If

            Me._TransDate = Value
            isDirty = True

        End Set

    End Property

    Public Property Description() As String
        Get

            Return Me._Description

        End Get

        Set(ByVal Value As String)

            If Value.Trim.Length > 0 Then
                Me._BrokenRules.BrokenRule("Description", False)

            Else
                Me._BrokenRules.BrokenRule("Description", True)
            End If

            Me._Description = Value
            isDirty = True

        End Set

    End Property


#End Region

#Region " Constructor "

    Public Sub New(ByVal Settings As NameValueCollection)

        Appsettings = Settings

        With _BrokenRules
            .BrokenRule("Number", True)
            .BrokenRule("TransDate", True)
            .BrokenRule("Description", True)
            .BrokenRule("Type", True)
            .BrokenRule("Amount", True)
        End With

    End Sub

#End Region

#Region " Public Methods "

    Public Function GetChecks(ByVal FromDate As String, ByVal ToDate As String) As DataSet

        Dim sb As System.Text.StringBuilder
        Dim Params As NameValueCollection
        Dim ds As DataSet

        'Set params for first table
        Params = New NameValueCollection
        With Params
            .Add("FromDate", FromDate)
            .Add("ToDate", ToDate)
        End With

        'Add paramters for first table to collection
        Params = New NameValueCollection

        'Set params for second table
        Params = New NameValueCollection
        With Params
            .Add("FromDate", FromDate)
            .Add("ToDate", ToDate)
        End With

        sb = New System.Text.StringBuilder
        With sb
            .Append("SELECT LedgerID, CodeNumber, TransDate, Desc, Type,Amount,Balance  ")
            .Append("FROM CheckLedger WHERE TransDate BETWEEN @FromDate AND @ToDate ORDER BY TransDate")
        End With

        'Populate the dataset
        ds = Utils.GetDataSet(Appsettings("Connstr"), sb.ToString, Params)
        isDirty = False

        Return ds

    End Function

    Public Sub InsertUpdateDelete()

        If Me.CrudMode <> DBMode.Delete AndAlso Me._BrokenRules.Count <> 0 Then
            Throw New Exception("Not all required fields have been entered. Record not saved.")
        End If

        Dim Params As New NameValueCollection
        Dim sbSQL As New System.Text.StringBuilder

        If Me.CrudMode = DBMode.Update Then
            With Params
                .Add("Number", Me.Number)
                .Add("Date", Me.TransDate.ToShortDateString)
                .Add("Description", Me.Description)
                .Add("Type", Me.Type.ToString)
                .Add("Amount", Me.Amount.ToString)
                .Add("Balance", Me.Balance.ToString)
                .Add("LedgerID", CStr(Me.LedgerID))
            End With

            With sbSQL
                .Append("UPDATE CheckLedger")
                .Append(" SET CodeNumber = @Number,TransDate=@TransDate,")
                .Append("[Desc]=@Description, Type=@Type, ")
                .Append("Amount=@Amount,Balance=@Balance")
                .Append(" WHERE LedgerID=@LedgerID")
            End With

        ElseIf Me.CrudMode = DBMode.Insert Then
            With Params
                .Add("Number", Me.Number)
                .Add("Date", Me.TransDate.ToShortDateString)
                .Add("Description", Me.Description)
                .Add("Type", Me.Type.ToString)
                .Add("Amount", Me.Amount.ToString)
                .Add("Balance", Me.Balance.ToString)
                .Add("LedgerID", CStr(Me.LedgerID))
            End With
            With sbSQL
                .Append("INSERT INTO CheckLedger (")
                .Append("CodeNumber,TransDate,[Desc],Type,Amount,Balance)")
                .Append(" Values (")
                .Append("@Number,@TransDate,@Description,@Type,@Amount,@Balance)")
            End With
        ElseIf Me.CrudMode = DBMode.Delete Then
            With Params
                .Add("LedgerID", Me.LedgerID.ToString)
            End With
            With sbSQL
                .Append("DELETE FROM CheckLedger")
                .Append(" WHERE LedgerID=@LedgerID")
            End With

        End If

        Me.isDirty = False

        Utils.ExecuteNonQuery(Appsettings("Connstr"), sbSQL.ToString, Params)

    End Sub
    Public Function GetLatestLedgerNo() As Integer

        Try
            Dim LedgerID As Integer
            LedgerID = Convert.ToInt32(Utils.ExecuteScalar(Appsettings("Connstr"), "SELECT MAX(LedgerID) FROM CheckLedger", Nothing))

            Return LedgerID

        Catch ex As Exception
            Return 1
        End Try

    End Function


#End Region

#Region " Custom event handlers "

    Private Sub _BrokenRules_ValidOrNot(ByVal isValid As Boolean) Handles _BrokenRules.ValidOrNot
        RaiseEvent Valid(isValid)

    End Sub

#End Region

End Class
