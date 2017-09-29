
#Region " Options "

Option Explicit On
Option Strict On

#End Region

#Region " Imports "

Imports System.Collections.Specialized
Imports AccessUtils

#End Region

Public Class MatAdvSearch

    Inherits BaseBusiness

#Region " Public Constructors "

    Public Sub New(ByVal Settings As NameValueCollection)

        MyBase.New(Settings)

        With BrokenRules
            .BrokenRule("SearchText", True)
            .BrokenRule("DatesValid", True)
            .BrokenRule("PriceRangeValid", False)
        End With

    End Sub

#End Region

#Region " Private Variables "

    Protected WithEvents BrokenRules As New BrokenRules
    Private _ReportType As SearchType
    Private _SearchText As String = ""
    Private _dsData As New DataSet
    Private _FromDate As Date
    Private _ToDate As Date
    Private _PriceFrom As Double
    Private _PriceTo As Double
    Private _Text As String

#End Region

#Region " Constants "

    Private Const NoRecords As String = "No Records, Click Run and/or Change Criteria"
    Private Const RecsFound As String = " Record(s) Matched Search Criteria"

#End Region

#Region " Enums "

    Public Enum SearchType

        Desc = 0
        Notes = 1
        PriceRange = 2

    End Enum

#End Region

#Region " Public Events "

    Public Event SearchDataRequired(ByVal bolRequired As Boolean)
    Public Event IsPriceRange(ByVal bolPriceRange As Boolean)
    Public Event CostRangeValid(ByVal bolValid As Boolean)
    Public Event DateRangeValid(ByVal isValid As Boolean)
    Public Event ReadyToSearch(ByVal Valid As Boolean)
    Public Event HasData(ByVal Data As Boolean)
    Public Event DataSourceChanged()
    Public Event TextChanged()

#End Region

#Region " Public Properties "

    Public Property SearchText() As String

        Get
            Return Me._SearchText

        End Get

        Set(ByVal Value As String)

            Me._SearchText = "%" & Value & "%"
            Me.PropertyChanged()


            If Me.ReportType <> SearchType.PriceRange Then
                If Value.Trim.Length > 0 Then
                    Me.BrokenRules.BrokenRule("SearchText", False)
                Else
                    Me.BrokenRules.BrokenRule("SearchText", True)
                End If
            End If

        End Set

    End Property

    Public Property ReportType() As SearchType

        Get

            Return Me._ReportType

        End Get

        Set(ByVal Value As SearchType)

            Me._ReportType = Value
            Me.PropertyChanged()

            RaiseEvent IsPriceRange(Value = SearchType.PriceRange)

            Select Case Value
                Case SearchType.PriceRange
                    'Price range doesn't require search data
                    Me.BrokenRules.BrokenRule("SearchText", False)
                Case SearchType.Desc, SearchType.Notes
                    'Need to account for "%%"
                    If Me.SearchText.Trim.Length > 2 Then
                        Me.BrokenRules.BrokenRule("SearchText", False)
                    Else
                        Me.BrokenRules.BrokenRule("SearchText", True)
                    End If

            End Select

        End Set

    End Property


    Public Property FromDate() As Date
        Get

            Return Me._FromDate

        End Get

        Set(ByVal Value As Date)

            Me._FromDate = Value
            Me.PropertyChanged()

            If Value > Me._ToDate Then
                Me.BrokenRules.BrokenRule("DatesValid", True)
                RaiseEvent DateRangeValid(False)
            Else
                Me.BrokenRules.BrokenRule("DatesValid", False)
                RaiseEvent DateRangeValid(True)

            End If

        End Set

    End Property

    Public Property ToDate() As Date

        Get

            Return Me._ToDate

        End Get

        Set(ByVal Value As Date)

            Me._ToDate = Value

            Me.PropertyChanged()

            If Value < Me._FromDate Then
                Me.BrokenRules.BrokenRule("DatesValid", True)
                RaiseEvent DateRangeValid(False)

            Else
                Me.BrokenRules.BrokenRule("DatesValid", False)
                RaiseEvent DateRangeValid(True)
            End If

        End Set

    End Property

    Public Property PriceFrom() As Double

        Get
            Return Me._PriceFrom

        End Get

        Set(ByVal Value As Double)

            Me.PropertyChanged()

            Me._PriceFrom = Value

            If Me.ReportType = SearchType.PriceRange Then

                If Value <= Me.PriceTo Then
                    Me.BrokenRules.BrokenRule("PriceRangeValid", False)
                    RaiseEvent CostRangeValid(True)
                Else
                    Me.BrokenRules.BrokenRule("PriceRangeValid", True)
                    RaiseEvent CostRangeValid(False)

                End If

            End If

        End Set

    End Property
    Public Property PriceTo() As Double

        Get

            Return Me._PriceTo

        End Get

        Set(ByVal Value As Double)

            Me.PropertyChanged()

            Me._PriceTo = Value

            If Me.ReportType = SearchType.PriceRange Then

                If Value >= Me.PriceFrom Then
                    Me.BrokenRules.BrokenRule("PriceRangeValid", False)
                    RaiseEvent CostRangeValid(True)

                Else
                    Me.BrokenRules.BrokenRule("PriceRangeValid", True)
                    RaiseEvent CostRangeValid(False)

                End If

            End If


        End Set

    End Property

    Public Property ReportText() As String

        Get

            Return Me._Text

        End Get

        Set(ByVal Value As String)

            Me._Text = Value
            RaiseEvent TextChanged()

        End Set

    End Property

    Public Property DataSource() As DataSet

        Get

            Return Me._dsData

        End Get

        Set(ByVal Value As DataSet)

            RaiseEvent DataSourceChanged()

            If IsNothing(Value) OrElse Value.Tables(0).Rows.Count = 0 Then
                Me.ReportText = MatAdvSearch.NoRecords
                RaiseEvent HasData(False)
            Else
                Me.ReportText = Value.Tables(0).Rows.Count & MatAdvSearch.RecsFound
                RaiseEvent HasData(True)
            End If

            Me._dsData = Value


        End Set

    End Property

#End Region

#Region " Public Routines "

    Public Overridable Function GetData() As DataSet

        Dim params As New NameValueCollection
        Dim sb As New System.Text.StringBuilder
        Dim ds As New DataSet

        Select Case Me.ReportType
            Case SearchType.Desc, SearchType.Notes
                With params
                    .Add("Date1", Me.FromDate.ToShortDateString)
                    .Add("Date2", Me.ToDate.ToShortDateString)
                    .Add("Text", Me.SearchText)
                End With
            Case SearchType.PriceRange
                With params
                    .Add("FromPrice", Me.PriceFrom.ToString)
                    .Add("ToPrice", Me.PriceTo.ToString)
                    .Add("Date1", Me.FromDate.ToShortDateString)
                    .Add("Date2", Me.ToDate.ToShortDateString)
                End With
        End Select

        Select Case Me.ReportType
            Case SearchType.Desc
                With sb
                    .Append("SELECT MaterialNo,DatePurchased,MaterialDesc,MaterialCost,Notes FROM Materials")
                    .Append(" WHERE DatePurchased BETWEEN @Date1 AND @Date2 AND MaterialDesc LIKE  @Text ORDER BY DatePurchased")
                End With
            Case SearchType.Notes
                With sb
                    .Append("SELECT MaterialNo,DatePurchased,MaterialDesc,MaterialCost,Notes FROM Materials")
                    .Append(" WHERE DatePurchased BETWEEN @Date1 AND @Date2 AND Notes LIKE @Text ORDER BY DatePurchased")
                End With
            Case SearchType.PriceRange
                With sb
                    .Append("SELECT MaterialNo,DatePurchased,MaterialDesc,MaterialCost,Notes FROM Materials")
                    .Append(" WHERE MaterialCost BETWEEN @FromPrice AND @ToPrice AND DatePurchased BETWEEN @Date1 AND @Date2  ORDER BY DatePurchased")
                End With

        End Select


        ds = Utils.GetDataSet(Appsettings("Connstr"), sb.ToString, params)

        Me.DataSource = ds

        Return ds

    End Function


#End Region

#Region " Private Routines "


    Private Sub PropertyChanged()

        Me.DataSource = Nothing

        RaiseEvent DataSourceChanged()

        If Me.ReportType <> SearchType.PriceRange Then
            If Me._SearchText.Trim.Length > 2 Then
                RaiseEvent SearchDataRequired(False)
            Else
                RaiseEvent SearchDataRequired(True)

            End If
        Else
            RaiseEvent SearchDataRequired(False)
        End If

    End Sub
#End Region

#Region " Event Sinks "

    Private Sub BrokenRules_ValidOrNot(ByVal isValid As Boolean) Handles BrokenRules.ValidOrNot

        RaiseEvent ReadyToSearch(isValid)

    End Sub
#End Region

End Class
