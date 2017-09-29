
#Region " Options "

Option Explicit On
Option Strict On

#End Region

#Region " Imports "

Imports AccessUtils
Imports System.Collections.Specialized

#End Region

Public Class MiscCategory

#Region " Constructor(s) "

    Public Sub New(ByVal Settings As NameValueCollection, ByVal CatId As Integer)

        Appsettings = Settings
        Load(CatId)

    End Sub

    Public Sub New(ByVal Settings As NameValueCollection)
        Appsettings = Settings
    End Sub

#End Region

#Region " Private variables "

    Private _ID As Integer
    Private _Description As String
    Private _Active As Integer
    Private _CrudMode As DBMode = DBMode.Insert
    Private Appsettings As New NameValueCollection


#End Region

#Region " Enums "

    Public Enum DBMode
        Insert = 0
        Update = 1
        Delete = 2
    End Enum

#End Region

#Region " Public Properties "

    Public Property ID() As Integer

        Get
            Return Me._ID

        End Get

        Set(ByVal Value As Integer)
            Me._ID = Value
        End Set

    End Property

    Public Property Description() As String

        Get
            Return Me._Description
        End Get

        Set(ByVal Value As String)
            Me._Description = Value.Trim

        End Set

    End Property

    Public Property Active() As Integer

        Get
            Return Me._Active
        End Get

        Set(ByVal Value As Integer)
            Me._Active = Value

        End Set

    End Property

    Public Property CrudMode() As DBMode

        Get
            Return Me._CrudMode
        End Get

        Set(ByVal Value As DBMode)
            Me._CrudMode = Value
        End Set

    End Property

#End Region

#Region " Public methods "

    Public Sub InsertUpdateDelete()

        Dim Params As New NameValueCollection
        Dim sbSQL As New System.Text.StringBuilder

        Select Case Me.CrudMode
            Case DBMode.Update
                CheckForExisitingOnUpdate()
                With Params
                    .Add("Description", Me.Description)
                    .Add("Active", Me.Active.ToString)
                    .Add("CategoryID", Me.ID.ToString)
                End With
            Case DBMode.Insert
                CheckForExisitingOnInsert()
                With Params
                    .Add("CategoryID", Me.GetCatID.ToString)
                    .Add("Description", Me.Description)
                    .Add("Active", Me.Active.ToString)
                End With
            Case DBMode.Delete
                Params.Add("CategoryID", Me.ID.ToString)
        End Select

        Select Case Me.CrudMode
            Case DBMode.Update
                With sbSQL
                    .Append("UPDATE MiscExpCategories ")
                    .Append("SET Description=@Description, ")
                    .Append("Active=@Active ")
                    .Append("WHERE CategoryID=@CategoryID")
                End With

            Case DBMode.Insert
                With sbSQL
                    .Append("INSERT INTO MiscExpCategories (")
                    .Append("CategoryID,Description,Active)")
                    .Append(" Values (")
                    .Append("@CategoryID,@Description,@Active)")
                End With
            Case DBMode.Delete
                With sbSQL
                    .Append("DELETE FROM MiscExpCategories ")
                    .Append("WHERE CategoryID=@CategoryID")
                End With
        End Select


        Utils.ExecuteNonQuery(Appsettings("Connstr"), sbSQL.ToString, Params)

    End Sub

    Private Function Load(ByVal CategoryId As Integer) As MiscCategory

        Dim sb As New System.Text.StringBuilder
        Dim params As New NameValueCollection
        Dim dr As OleDb.OleDbDataReader


        With params
            .Add("CategoryID", CategoryId.ToString)
        End With

        With sb
            .Append("SELECT CategoryID,Description,Active ")
            .Append(" FROM MiscExpCategories WHERE CategoryID=@CategoryID ")
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

    Private Function GetCatID() As Integer

        Try
            Dim CatID As Integer
            CatID = Convert.ToInt32(Utils.ExecuteScalar(Appsettings("Connstr"), "SELECT MAX(CategoryID) FROM MiscExpCategories", Nothing))

            Return CatID + 1

        Catch ex As Exception
            Return 1
        End Try

    End Function

    Private Sub LoadPropsFromDataReader(ByVal dr As OleDb.OleDbDataReader)

        With dr

            While .Read
                Me.ID = Convert.ToInt32(.Item(0))
                Me.Description = .GetString(1)
                Me.Active = Convert.ToInt32(.Item(2))

            End While

        End With
    End Sub

    Private Sub CheckForExisitingOnInsert()

        Dim strsql As String

        strsql = "Select count(*) from MiscExpCategories  where Description='" & Me.Description & "'"
        Dim Result As Integer = Utils.ExecuteScalar(Appsettings("Connstr"), strsql, Nothing)

        If Result > 0 Then
            Throw New Exception("That category already exists!")
        End If
    End Sub
    Private Sub CheckForExisitingOnUpdate()

        Dim strsql As String

        strsql = "Select count(*) from MiscExpCategories  where Description='" & Me.Description & "' AND CategoryID <> " & Me.ID
        Dim Result As Integer = Utils.ExecuteScalar(Appsettings("Connstr"), strsql, Nothing)

        If Result > 0 Then
            Throw New Exception("That category already exists!")
        End If

    End Sub

#End Region

End Class
