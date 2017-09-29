
#Region " Options "

Option Explicit On
Option Strict On

#End Region

#Region " Imports "

Imports System.Collections.Specialized
Imports AccessUtils
Imports System.Text.RegularExpressions

#End Region

Public NotInheritable Class Customer

#Region " Public Constructors "
    Public Sub New(ByVal Settings As NameValueCollection)

        Me.Appsettings = Settings
        Me._BrokenRules = New BrokenRules

        With _BrokenRules
            .BrokenRule("FirstName", True)
            .BrokenRule("LastName", True)
            .BrokenRule("Addr1", True)
            .BrokenRule("City", True)
            .BrokenRule("Zip", True)
        End With

    End Sub
    Public Sub New()

    End Sub
#End Region

#Region " Private Variables"

    Private _CusNo As Integer
    Private _FirstName As String = ""
    Private _LastName As String = ""
    Private _Addr1 As String = ""
    Private _Addr2 As String = ""
    Private _City As String = "PlainField"
    Private _State As String = "IL"
    Private _Zip As String = "60544"
    Private _Phone As String = ""
    Private _CellPhone As String = ""
    Private _Notes As String = ""
    Private _WindowCustomer As String = "0"
    Private _WindowsCleanedWhen As WindowsCleanedWhen = WindowsCleanedWhen.NA
    Private _WindowPrice As Double = 0
    Private WithEvents _BrokenRules As BrokenRules
    Private Appsettings As New NameValueCollection
    Private _CrudMode As CrudMode
    Private _IsDirty As Boolean
    Private _IsValid As Boolean

    Event Valid(ByVal Valid As Boolean)
    Event DeleteValid(ByVal IsValid As Boolean)
    Event Dirty()
    Event WindowsCustomer(ByVal IsWindowClient As Boolean)


#End Region

#Region " Enums "

    Public Enum WindowsCleanedWhen
        Spring = 0
        Fall = 1
        OnRequest = 2
        SpringAndFall = 3
        SpringAndOnRequest = 4
        FallAndOnRequest = 5
        All = 6
        NA = 7
    End Enum

    Public Enum CrudMode
        Insert = 0
        Update = 1
        Delete = 2
    End Enum

#End Region

#Region " Public Properties "

    Public Property IsValid() As Boolean
        Get
            Return Me._IsValid

        End Get
        Set(ByVal Value As Boolean)
            Me._IsValid = Value

        End Set
    End Property
    Public Property IsDirty() As Boolean
        Get
            Return Me._IsDirty
        End Get
        Set(ByVal Value As Boolean)
            Me._IsDirty = Value

        End Set
    End Property
    Public Property DBMode() As CrudMode
        Get
            Return _CrudMode

        End Get

        Set(ByVal Value As CrudMode)

            _CrudMode = Value

            If Value = CrudMode.Update Then
                RaiseEvent DeleteValid(True)
            Else
                RaiseEvent DeleteValid(False)
            End If

        End Set
    End Property

    Public Property CustID() As Integer
        Get
            Return Me._CusNo
        End Get

        Set(ByVal Value As Integer)
            Me._CusNo = Value
        End Set
    End Property

    Public Property FirstName() As String

        Get
            Return Me._FirstName
        End Get

        Set(ByVal Value As String)

            Me._FirstName = StrConv(Value, VbStrConv.ProperCase)

            Me._IsDirty = True
            RaiseEvent Dirty()

            If Value.Length > 0 Then
                Me._BrokenRules.BrokenRule("FirstName", False)

            Else
                Me._BrokenRules.BrokenRule("FirstName", True)
            End If

        End Set

    End Property
    Public Property LastName() As String
        Get
            Return Me._LastName
        End Get

        Set(ByVal Value As String)

            Me._LastName = StrConv(Value, VbStrConv.ProperCase)
            Me._IsDirty = True
            RaiseEvent Dirty()


            If Value.Length > 1 Then
                Me._BrokenRules.BrokenRule("LastName", False)
            Else
                Me._BrokenRules.BrokenRule("LastName", True)
            End If

        End Set
    End Property

    Public Property Addr1() As String

        Get
            Return Me._Addr1
        End Get

        Set(ByVal Value As String)

            Me._Addr1 = Value
            Me._IsDirty = True

            RaiseEvent Dirty()

            If Value.Length > 1 Then
                Me._BrokenRules.BrokenRule("Addr1", False)

            Else
                Me._BrokenRules.BrokenRule("Addr1", True)
            End If

        End Set

    End Property

    Public Property Addr2() As String

        Get
            Return Me._Addr2
        End Get

        Set(ByVal Value As String)
            Me._Addr2 = Value
            Me._IsDirty = True
            RaiseEvent Dirty()

        End Set

    End Property

    Public Property City() As String

        Get
            Return Me._City
        End Get

        Set(ByVal Value As String)
            Me._City = Value
            Me._IsDirty = True
            RaiseEvent Dirty()

            If Value.Length > 1 Then
                Me._BrokenRules.BrokenRule("City", False)

            Else
                Me._BrokenRules.BrokenRule("City", True)
            End If
        End Set

    End Property

    Public Property State() As String

        Get
            Return Me._State
        End Get

        Set(ByVal Value As String)

            Me._State = Value
            Me._IsDirty = True
            RaiseEvent Dirty()


        End Set

    End Property

    Public Property Zip() As String

        Get
            Return Me._Zip
        End Get

        Set(ByVal Value As String)

            Me._Zip = Value
            Me._IsDirty = True
            RaiseEvent Dirty()


            If Value.Length > 0 Then
                If Not Regex.IsMatch(Value, Appsettings("ZipCodeMask")) Then
                    Me._BrokenRules.BrokenRule("Zip", True)
                Else
                    Me._BrokenRules.BrokenRule("Zip", False)
                End If
            Else
                Me._BrokenRules.BrokenRule("Zip", True)
            End If

        End Set

    End Property

    Public Property Phone() As String

        Get
            Return Me._Phone

        End Get

        Set(ByVal Value As String)

            Me._Phone = Value
            Me._IsDirty = True

            RaiseEvent Dirty()


            If Value.Length > 0 Then
                If Not Regex.IsMatch(Value, Appsettings("PhoneMask")) Then
                    Me._BrokenRules.BrokenRule("Phone", True)
                Else
                    Me._BrokenRules.BrokenRule("Phone", False)
                End If
            Else
                Me._BrokenRules.BrokenRule("Phone", False)
            End If
        End Set

    End Property

    Public Property CellPhone() As String

        Get
            Return Me._CellPhone

        End Get

        Set(ByVal Value As String)

            Me._CellPhone = Value
            Me._IsDirty = True
            RaiseEvent Dirty()

            If Value.Length > 0 Then

                If Not Regex.IsMatch(Value, Appsettings("PhoneMask")) Then
                    Me._BrokenRules.BrokenRule("CellPhone", True)
                Else
                    Me._BrokenRules.BrokenRule("CellPhone", False)

                End If
            Else
                Me._BrokenRules.BrokenRule("CellPhone", False)
            End If
        End Set

    End Property

    Public Property Notes() As String

        Get
            Return Me._Notes
        End Get

        Set(ByVal Value As String)

            Me._Notes = Value
            Me._IsDirty = True
            RaiseEvent Dirty()


        End Set

    End Property
    Public ReadOnly Property FullName() As String
        Get
            Return Me._FirstName & " " & Me._LastName
        End Get
    End Property
    Public Property WindowCustomer() As String

        Get
            Return Me._WindowCustomer
        End Get

        Set(ByVal Value As String)

            Me._WindowCustomer = Value

            Me._IsDirty = True

            RaiseEvent Dirty()

            If Value = "0" Then
                RaiseEvent WindowsCustomer(False)

                Me._WindowPrice = 0
                Me.WindowsCleanedIn = WindowsCleanedWhen.NA
                Me._BrokenRules.BrokenRule("WindowsCleanedWhen", False)
                Me._BrokenRules.BrokenRule("WindowsPrice", False)

            Else
                RaiseEvent WindowsCustomer(True)
                Me._BrokenRules.BrokenRule("WindowsCleanedWhen", True)
                Me._BrokenRules.BrokenRule("WindowsPrice", True)
            End If

        End Set

    End Property

    Public Property WindowsCleanedIn() As WindowsCleanedWhen
        Get
            Return Me._WindowsCleanedWhen
        End Get
        Set(ByVal Value As WindowsCleanedWhen)
            Me._WindowsCleanedWhen = Value

            Me._IsDirty = True
            RaiseEvent Dirty()

            If Me.WindowCustomer = "1" Then
                If Value = WindowsCleanedWhen.NA Then
                    Me._BrokenRules.BrokenRule("WindowsCleanedWhen", True)
                Else
                    Me._BrokenRules.BrokenRule("WindowsCleanedWhen", False)
                End If
            End If

        End Set
    End Property
    Public Property WindowPrice() As Double

        Get
            Return Me._WindowPrice
        End Get

        Set(ByVal Value As Double)

            Me._WindowPrice = Value

            Me._IsDirty = True
            RaiseEvent Dirty()

            If Me.WindowCustomer = "1" Then
                If Value = 0 Then
                    Me._BrokenRules.BrokenRule("WindowsPrice", True)

                Else
                    Me._BrokenRules.BrokenRule("WindowsPrice", False)
                End If
            End If
        End Set

    End Property
#End Region

#Region " Public methods "

    Public Sub InsertUpdateDelete()

        Dim params As New NameValueCollection
        Dim sb As New System.Text.StringBuilder

        If Me.DBMode = CrudMode.Insert Or Me.DBMode = CrudMode.Update Then
            With params
                .Add("LastName", Me.LastName)
                .Add("FirstName", Me.FirstName)
                .Add("Addr1", Me.Addr1)
                .Add("Addr2", Me.Addr2)
                .Add("City", Me.City)
                .Add("State", Me.State)
                .Add("Zip", Me.Zip)
                .Add("Phone", Me.Phone)
                .Add("Cell", Me.CellPhone)
                .Add("Notes", Me.Notes)
                .Add("WindowsCustomer", Me.WindowCustomer)
                .Add("WindowsCleanedWhen", Convert.ToInt32(WindowsCleanedIn).ToString)
                .Add("WindowPrice", Me.WindowPrice.ToString)
                If Me.DBMode = CrudMode.Update Then
                    .Add("CustomerID", Me.CustID.ToString)
                End If
            End With
        End If

        If DBMode = CrudMode.Insert Then
            With sb
                .Append("INSERT INTO Customers(")
                .Append("LastName,FirstName,Addr1,Addr2,City,State,Zip,Phone,CellPhone,Notes,WindowsCustomer,WindowsCleanedWhen,WindowPrice)")
                .Append(" Values(")
                .Append("@LastName,@FirstName,@Addr1,@Addr2,@City,@State,@Zip,@Phone,@Cell,@Notes,@WindowsCustomer,@WindowsCleanedWhen,@WindowPrice)")
            End With
        End If

        If DBMode = CrudMode.Update Then
            With sb
                .Append("UPDATE Customers")
                .Append(" SET LastName=@LastName,FirstName=@FirstName,Addr1=@Addr1,Addr2=@Addr2,")
                .Append("City=@City,State=@State,Zip=@Zip,Phone=@Phone,CellPhone=@CellPhone,Notes=@Notes,")
                .Append("WindowsCustomer=@WindowsCustomer,WindowsCleanedWhen=@WindowsCleanedWhen,WindowPrice=@WindowPrice")
                .Append(" WHERE CustomerID=@CustomerID")
            End With

        End If

        'Customers are never deleted, just marked as inactive
        If Me.DBMode = CrudMode.Delete Then
            With params
                .Add("CustomerID", Me.CustID.ToString)
            End With
            With sb
                .Append("UPDATE Customers SET IsActive='0' WHERE")
                .Append(" CustomerID=@CustomerID")
            End With
        End If

        Utils.ExecuteNonQuery(Appsettings("Connstr"), sb.ToString, params)
        Me.IsDirty = False

    End Sub
    Public Sub UndeleteCustomer(ByVal CustomerId As Integer)

        Dim params As New NameValueCollection
        Dim sb As New System.Text.StringBuilder

        params.Add("CustomerID", CustomerId.ToString)

        With sb
            .Append("UPDATE Customers SET IsActive='1' WHERE")
            .Append(" CustomerID=@CustomerID")
        End With

        Utils.ExecuteNonQuery(Appsettings("Connstr"), sb.ToString, params)

    End Sub

    Public Sub UpdateWindowPricing(ByVal SQL As String)

        Utils.ExecuteNonQuery(Appsettings("Connstr"), SQL, Nothing)

    End Sub

    Public Function CheckForCustomer() As Boolean

        Dim params As New NameValueCollection
        Dim sb As New System.Text.StringBuilder

        With params
            .Add("LastName", Me.LastName)
            .Add("FirstName", Me.FirstName)
        End With

        With sb
            .Append("SELECT Count(*) FROM Customers")
            .Append(" WHERE LastName=@LastName AND FirstName=@FirstName")
        End With

        If Utils.ExecuteScalar(Appsettings("Connstr"), sb.ToString, params) > 0 Then
            Return True
        Else
            Return False
        End If


    End Function

    Public Function GetAllCustomers() As DataSet

        Dim sb As New System.Text.StringBuilder
        Dim ds As DataSet


        With sb
            .Append("SELECT CustomerID, LastName,FirstName,Addr1,Addr2,City,State,Zip,Phone,CellPhone,Notes,WindowsCustomer,WindowPrice,WindowsCleanedWhen")
            .Append(" FROM Customers WHERE IsActive='1' ORDER BY LastName,FirstName")
        End With

        ds = Utils.GetDataSet(Appsettings("Connstr"), sb.ToString, Nothing)

        ds.Tables(0).TableName = "Customers"

        Return ds

    End Function

    Public Function GetDeletedCustomers() As DataSet

        Dim sb As New System.Text.StringBuilder
        Dim ds As DataSet


        With sb
            .Append("SELECT CustomerID, LastName,FirstName,Addr1,Addr2,City,State,Zip,Phone,CellPhone,Notes,WindowsCustomer,WindowPrice,WindowsCleanedWhen")
            .Append(" FROM Customers WHERE IsActive='0' ORDER BY LastName,FirstName")
        End With

        ds = Utils.GetDataSet(Appsettings("Connstr"), sb.ToString, Nothing)

        ds.Tables(0).TableName = "Customers"

        Return ds

    End Function
    Public Function GetWindowCustomers() As DataSet

        Dim sb As New System.Text.StringBuilder
        Dim ds As DataSet


        With sb
            .Append("SELECT CustomerID, LastName,FirstName,Addr1,Addr2,City,State,Zip,Phone,CellPhone,Notes,WindowsCustomer,WindowPrice,WindowsCleanedWhen")
            .Append(" FROM Customers WHERE WindowsCustomer='1'AND IsActive='1' ORDER BY LastName,FirstName")
        End With

        ds = Utils.GetDataSet(Appsettings("Connstr"), sb.ToString, Nothing)

        ds.Tables(0).TableName = "Customers"

        Return ds

    End Function

    Public Function GetNonWindowCustomers() As DataSet

        Dim sb As New System.Text.StringBuilder
        Dim ds As DataSet


        With sb
            .Append("SELECT CustomerID, LastName,FirstName,Addr1,Addr2,City,State,Zip,Phone,CellPhone,Notes,WindowsCustomer,WindowPrice,WindowsCleanedWhen")
            .Append(" FROM Customers WHERE (WindowsCustomer IS NULL or WindowsCustomer ='0') AND IsActive='1' ORDER BY LastName,FirstName")
        End With

        ds = Utils.GetDataSet(Appsettings("Connstr"), sb.ToString, Nothing)

        ds.Tables(0).TableName = "Customers"

        Return ds

    End Function
    Public Function GetCustomer() As Customer

        Dim sb As New System.Text.StringBuilder
        Dim params As New NameValueCollection
        Dim dr As OleDb.OleDbDataReader

        With params
            .Add("LastName", Me.LastName)
            .Add("FirstName", Me.FirstName)
        End With

        With sb
            .Append("SELECT CustomerID, LastName,FirstName,Addr1,Addr2,City,State,Zip,Phone,CellPhone,Notes,")
            .Append("WindowsCustomer,WindowsCleanedWhen,WindowPrice")
            .Append(" FROM Customers WHERE LastName = @LastName AND FirstName = @FirstName AND IsActive='1'")
        End With


        dr = Utils.GetDataReader(Appsettings("Connstr"), sb.ToString, params)

        If Not dr.HasRows Then
            Return New Customer
        End If

        If dr.Depth > 0 Then
            Throw New Exception("More than one cutomer found")
        End If

        LoadPropsFromDataReader(dr)


        dr.Close()

        Return Me


    End Function
    Public Function GetAllCustomersForCombo() As DataSet

        Dim strSQl As String = "Select CustomerID, LastName + ',' + FirstName as Name from Customers ORDER BY LastName,FirstName"

        Return Utils.GetDataSet(Appsettings("Connstr"), strSQl, Nothing)


    End Function

    Public Function GetActiveCustomersForCombo() As DataSet

        Dim strSQl As String = "Select CustomerID, LastName + ',' + FirstName as Name from Customers WHERE IsActive='1' ORDER BY LastName,FirstName"

        Return Utils.GetDataSet(Appsettings("Connstr"), strSQl, Nothing)


    End Function

    Public Function GetCustomer(ByVal CustID As Integer) As Customer

        Dim sb As New System.Text.StringBuilder
        Dim params As New NameValueCollection
        Dim dr As OleDb.OleDbDataReader


        With params
            .Add("CustomerID", CustID.ToString)
        End With

        With sb
            .Append("SELECT CustomerID, LastName,FirstName,Addr1,Addr2,City,State,Zip,Phone,CellPhone,Notes,")
            .Append("WindowsCustomer,WindowsCleanedWhen,WindowPrice")
            .Append(" FROM Customers WHERE CustomerID=@CustomerID ")
        End With

        dr = Utils.GetDataReader(Appsettings("Connstr"), sb.ToString, params)

        If dr.Depth > 0 Then
            Throw New Exception("More than one cutomer found")
        End If

        LoadPropsFromDataReader(dr)

        dr.Close()

        Return Me


    End Function

#End Region

#Region " Private Routines "

    Private Sub LoadPropsFromDataReader(ByVal dr As OleDb.OleDbDataReader)

        With dr
            While .Read
                Me.CustID = Convert.ToInt32(.Item(0))
                Me.LastName = .GetString(1)
                Me.FirstName = .GetString(2)
                If .IsDBNull(3) Then
                    Me.Addr1 = ""
                Else
                    Me.Addr1 = .GetString(3).ToString
                End If

                If .IsDBNull(4) Then
                    Me.Addr2 = ""
                Else
                    Me.Addr2 = .GetString(4)
                End If

                If .IsDBNull(5) Then
                    Me.City = ""
                Else
                    Me.City = .GetString(5)
                End If

                If .IsDBNull(6) Then
                    Me.State = ""
                Else
                    Me.State = .GetString(6)
                End If

                If .IsDBNull(7) Then
                    Me.Zip = ""
                Else
                    Me.Zip = .GetString(7)
                End If

                If .IsDBNull(8) Then
                    Me.Phone = ""
                Else
                    Me.Phone = .GetString(8)

                End If

                If .IsDBNull(9) Then
                    Me.CellPhone = ""
                Else
                    Me.CellPhone = .GetString(9)

                End If

                If .IsDBNull(10) Then
                    Me.Notes = ""
                Else
                    Me.Notes = .GetString(10)

                End If

                If .IsDBNull(11) Then
                    Me.WindowCustomer = "0"
                Else
                    Me.WindowCustomer = .GetString(11)
                End If

                If .IsDBNull(12) Then
                    Me.WindowsCleanedIn = WindowsCleanedWhen.NA
                Else
                    Me.WindowsCleanedIn = CType(.Item(12), WindowsCleanedWhen)
                End If

                If .IsDBNull(13) Then
                    Me.WindowPrice = 0
                Else
                    Me.WindowPrice = Convert.ToDouble(.Item(13))
                End If
            End While

        End With


    End Sub

    Private Sub _BrokenRules_ValidOrNot(ByVal isValid As Boolean) Handles _BrokenRules.ValidOrNot

        RaiseEvent Valid(isValid)
        Me.IsValid = isValid

    End Sub

#End Region

End Class
