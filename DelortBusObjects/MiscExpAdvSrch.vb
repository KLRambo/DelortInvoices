
#Region " Options "

Option Explicit On
Option Strict On

#End Region

#Region " Imports "

Imports System.Collections.Specialized
Imports AccessUtils

#End Region

Public Class MiscExpAdvSrch
    Inherits MatAdvSearch

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

#Region " Over-Ridden Function(s) "

    Public Overrides Function GetData() As DataSet

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
                    .Append("SELECT M.ItemNo,M.DateOfExp,MC.Description,M.Description,M.Cost,M.Notes,M.CategoryID FROM MiscExps AS M ")
                    .Append("INNER JOIN MiscExpCategories AS MC ON ")
                    .Append("M.CategoryID = MC.CategoryID")
                    .Append(" WHERE M.DateOfExp BETWEEN @Date1 AND @Date2 AND M.Description LIKE @Text ORDER BY DateOfExp")
                End With
            Case SearchType.Notes
                With sb
                    .Append("SELECT M.ItemNo,M.DateOfExp,MC.Description,M.Description,M.Cost,M.Notes,M.CategoryID FROM MiscExps AS M ")
                    .Append("INNER JOIN MiscExpCategories AS MC ON ")
                    .Append("M.CategoryID = MC.CategoryID")
                    .Append(" WHERE M.DateOfExp BETWEEN @Date1 AND @Date2 AND M.Notes LIKE @Text ORDER BY DateOfExp")
                End With
            Case SearchType.PriceRange
                With sb
                    .Append("SELECT M.ItemNo,M.DateOfExp,MC.Description,M.Description,M.Cost,M.Notes,M.CategoryID FROM MiscExps AS M ")
                    .Append("INNER JOIN MiscExpCategories AS MC ON ")
                    .Append("M.CategoryID = MC.CategoryID")
                    .Append(" WHERE Cost BETWEEN @FromPrice AND @ToPrice AND DateOfExp BETWEEN @Date1 AND @Date2  ORDER BY DateOfExp")
                End With

        End Select


        ds = Utils.GetDataSet(Appsettings("Connstr"), sb.ToString, params)

        Me.DataSource = ds

        Return ds

    End Function

#End Region

End Class
