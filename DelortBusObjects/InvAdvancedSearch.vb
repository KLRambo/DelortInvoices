
#Region " Options "

Option Explicit On
Option Strict On

#End Region

#Region " Imports "

Imports System.Collections.Specialized
Imports AccessUtils

#End Region

Public Class InvAdvancedSearch
    Inherits BaseBusiness

#Region " Public Constructors "

    Public Sub New(ByVal Settings As NameValueCollection)

        MyBase.New(Settings)
        Me._BrokenRules = New BrokenRules

        With _BrokenRules
            .BrokenRule("SearchText", True)
            .BrokenRule("DatesValid", True)
        End With

    End Sub

#End Region

#Region " Private Variables "

    Private _ToDate As Date
    Private _FromDate As Date
    Private _dsData As DataSet
    Private _ReportText As String
    Private _ReportType As SearchType
    Private _SearchText As String = ""
    Protected WithEvents _BrokenRules As BrokenRules

#End Region

#Region " Constants "

    Private Const NoRecords As String = "No Records, Click Run and/or Change Criteria"
    Private Const RecsFound As String = " Record(s) Matched Search Criteria"

#End Region

#Region " Public Events "

    Public Event ReadyToSearch(ByVal Valid As Boolean)
    Public Event HasData(ByVal Data As Boolean)
    Public Event ReportTextChanged()
    Public Event SearchDataRequired(ByVal bolRequired As Boolean)

#End Region

#Region " Enum(s) "

    Public Enum SearchType

        LaborDesc = 0
        MaterialDesc = 1
        InvoiceNotes = 2

    End Enum

#End Region

#Region " Public Properties "

    Public Property SearchText() As String

        Get
            Return Me._SearchText
        End Get

        Set(ByVal Value As String)

            Me._SearchText = "%" & Value & "%"

            PropertyChanged()

            If Value.Trim.Length > 0 Then
                Me._BrokenRules.BrokenRule("SearchText", False)
            Else
                Me._BrokenRules.BrokenRule("SearchText", True)

            End If

        End Set

    End Property

    Public Property ReportType() As SearchType

        Get
            Return Me._ReportType
        End Get

        Set(ByVal Value As SearchType)

            Me._ReportType = Value
            PropertyChanged()
        End Set

    End Property


    Public Property FromDate() As Date
        Get

            Return Me._FromDate

        End Get

        Set(ByVal Value As Date)

            Me._FromDate = Value

            If Value > Me._ToDate Then
                Me._BrokenRules.BrokenRule("DatesValid", True)
            Else
                Me._BrokenRules.BrokenRule("DatesValid", False)

            End If
            PropertyChanged()

        End Set

    End Property

    Public Property ToDate() As Date

        Get

            Return Me._ToDate

        End Get

        Set(ByVal Value As Date)

            Me._ToDate = Value
            If Value < Me._FromDate Then

                Me._BrokenRules.BrokenRule("DatesValid", True)
            Else
                Me._BrokenRules.BrokenRule("DatesValid", False)

            End If
            PropertyChanged()

        End Set
    End Property

    Public Property DataSource() As DataSet

        Get

            Return Me._dsData

        End Get

        Set(ByVal Value As DataSet)

            If IsNothing(Value) OrElse Value.Tables(0).Rows.Count = 0 Then
                Me.ReportText = NoRecords
                RaiseEvent HasData(False)
            Else
                Me.ReportText = Value.Tables(0).Rows.Count & RecsFound
                RaiseEvent HasData(True)
            End If

            Me._dsData = Value


        End Set

    End Property

    Public Property ReportText() As String

        Get
            Return _ReportText

        End Get

        Set(ByVal Value As String)
            _ReportText = Value

            RaiseEvent ReportTextChanged()

        End Set

    End Property

#End Region

#Region " Public Methods "

    Public Overridable Function GetData() As DataSet

        Dim params As New NameValueCollection
        Dim sbInv As New System.Text.StringBuilder
        Dim ds As New DataSet

        With params
            .Add("SearchText", Me.SearchText)
            .Add("FromDate", Me.FromDate.ToShortDateString)
            .Add("ToDate", Me.ToDate.ToShortDateString)

        End With

        Select Case Me.ReportType
            Case SearchType.InvoiceNotes
                With sbInv
                    .Append("SELECT DISTINCTROW Customers.LastName, Customers.FirstName, Invoices.InvoiceNo, Invoices.InvoiceDate ")
                    .Append("FROM (Customers INNER JOIN Invoices ON Customers.CustomerID = Invoices.CustomerID) INNER JOIN InvoiceNotes ON Invoices.InvoiceNo = InvoiceNotes.InvoiceNo ")
                    .Append(" WHERE (((InvoiceNotes.Notes) Like @SearchText)")
                    .Append(" AND Invoices.InvoiceDate BETWEEN @FromDate AND @ToDate); ")
                End With
            Case SearchType.LaborDesc, SearchType.MaterialDesc
                With sbInv
                    .Append("SELECT DISTINCTROW Customers.LastName, Customers.FirstName, Invoices.InvoiceNo, Invoices.InvoiceDate ")
                    .Append("FROM (Customers INNER JOIN Invoices ON Customers.CustomerID = Invoices.CustomerID) INNER JOIN LineItems ON Invoices.InvoiceNo = LineItems.InvoiceNo ")
                    If Me.ReportType = SearchType.LaborDesc Then
                        .Append("WHERE (((LineItems.LaborDesc) Like  @SearchText)")
                        .Append(" AND Invoices.InvoiceDate BETWEEN @FromDate AND @ToDate );")
                    Else
                        .Append("WHERE (((LineItems.MaterialDesc) Like  @SearchText)")
                        .Append(" AND Invoices.InvoiceDate BETWEEN @FromDate AND @ToDate );")
                    End If

                End With
        End Select

        ds = Utils.GetDataSet(Appsettings("Connstr"), sbInv.ToString, params)

        Me.DataSource = ds

        Return ds


    End Function

#End Region

#Region " Private Routines "


    Private Sub PropertyChanged()

        Me.DataSource = Nothing

        If Me._SearchText.Trim.Length > 2 Then
            RaiseEvent SearchDataRequired(False)
        Else
            RaiseEvent SearchDataRequired(True)

        End If

    End Sub


#End Region

#Region " Even Sinks "

    Private Sub _BrokenRules_ValidOrNot(ByVal isValid As Boolean) Handles _BrokenRules.ValidOrNot
        RaiseEvent ReadyToSearch(isValid)

    End Sub

#End Region

End Class
