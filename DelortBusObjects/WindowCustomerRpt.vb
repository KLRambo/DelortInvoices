
#Region " Options "

Option Explicit On
Option Strict On

#End Region

#Region " Imports "

Imports AccessUtils
Imports System.Collections.Specialized
Imports TDS

#End Region

Public Class WindowCustomerRpt

#Region " Public Constructors "

    Public Sub New(ByVal Settings As NameValueCollection)

        Me.AppSettings = Settings

    End Sub

#End Region

#Region " Private variables "

    Private _Filter As FilterBy
    Private _SearchBy As SearchBy
    Private _CustomerNo As Integer
    Private _LastName As String = ""
    Private _FirstName As String = ""
    Private _DataView As New DataView
    Private _MasterDataSource As DataSet
    Private _LastNameStartsWith As Boolean = True
    Private _FirstNameStartsWith As Boolean = True
    Private AppSettings As New NameValueCollection

#End Region

#Region " Enums "

    Enum FilterBy
        All = 0
        FirstName = 1
        LastName = 2
        FirstAndLastName = 3
    End Enum

    Enum SearchBy
        All = 0
        Name = 1
    End Enum

#End Region

#Region " Events "

    Public Event ValidRecords(ByVal Recs As Boolean)
    Public Event SearchingByAll(ByVal Value As Boolean)

#End Region

#Region " Public Properties "

    Public Property MasterDataSource() As DataSet
        Get
            Return Me._MasterDataSource
        End Get

        Set(ByVal Value As DataSet)

            Me._MasterDataSource = Value

            If Value.Tables.Count = 1 Then
                Value.Tables(0).TableName = "Customers"
                Me.DataView.Table = Value.Tables(0)
            End If

        End Set

    End Property
    Public ReadOnly Property Filter() As FilterBy
        Get
            Return Me._Filter
        End Get

    End Property

    Public Property CustomerNo() As Integer

        Get
            Return Me._CustomerNo
        End Get

        Set(ByVal Value As Integer)
            Me._CustomerNo = Value
        End Set

    End Property
    Public Property FirstName() As String
        Get
            Return Me._FirstName

        End Get
        Set(ByVal Value As String)
            Me._FirstName = Value
            Me.SetFilter()
        End Set

    End Property
    Public Property LastName() As String
        Get
            Return Me._LastName
        End Get
        Set(ByVal Value As String)
            Me._LastName = Value
            Me.SetFilter()
        End Set
    End Property
    Public Property DataView() As DataView

        Get
            Return Me._DataView
        End Get

        Set(ByVal Value As DataView)
            Me._DataView = Value

        End Set

    End Property
    Public Property LastNameStartsWith() As Boolean

        Get
            Return Me._LastNameStartsWith
        End Get

        Set(ByVal Value As Boolean)
            Me._LastNameStartsWith = Value
            SetFilter()
        End Set

    End Property
    Public Property FirstNameStartsWith() As Boolean

        Get
            Return Me._FirstNameStartsWith
        End Get

        Set(ByVal Value As Boolean)
            Me._FirstNameStartsWith = Value
            SetFilter()
        End Set

    End Property
    Public Property SearchingBy() As SearchBy

        Get
            Return Me._SearchBy
        End Get

        Set(ByVal Value As SearchBy)
            Me._SearchBy = Value
            If Value = SearchBy.All Then
                Me.LastName = ""
                Me.FirstName = ""
                RaiseEvent SearchingByAll(True)
            Else
                RaiseEvent SearchingByAll(False)
            End If

            SetFilter()

        End Set

    End Property
#End Region

#Region " Public methods "

    Public Sub SetMasterDataSource()

        Dim sb As New System.Text.StringBuilder

        With sb
            .Append("Select *")
            .Append(" FROM Customers WHERE IsActive='1'AND WindowsCustomer = '1' ORDER BY LastName,FirstName")
        End With

        Me.MasterDataSource = Utils.GetDataSet(AppSettings("Connstr"), sb.ToString, Nothing)

        SetFilter()

    End Sub

    Public Sub SetFilter()

        If Me.MasterDataSource Is Nothing Then Exit Sub

        Dim LastNameFilter As String
        Dim FirstNameFilter As String

        Me.DataView.Table = Me.MasterDataSource.Tables(0)

        Select Case Me.SearchingBy
            Case SearchBy.All
                Me.DataView.Table = Me.MasterDataSource.Tables(0)
                Me.DataView.RowFilter = "LastName Like '%'"
            Case SearchBy.Name
                If Me.LastNameStartsWith = True Then
                    LastNameFilter = Me._LastName & "%"
                Else
                    LastNameFilter = "%" & Me._LastName & "%"
                End If

                If Me.FirstNameStartsWith = True Then
                    FirstNameFilter = Me._FirstName & "%"
                Else
                    FirstNameFilter = "%" & Me._FirstName & "%"
                End If

                Me.DataView.RowFilter = "FirstName Like '" & FirstNameFilter & "' AND LastName Like '" & LastNameFilter & "'"

        End Select

        RaiseEvent ValidRecords(Me.DataView.Count > 0)

    End Sub


#End Region

End Class

