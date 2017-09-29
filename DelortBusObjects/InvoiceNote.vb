
#Region " Options "

Option Explicit On
Option Strict On

#End Region

#Region " Imports "

Imports AccessUtils
Imports System.Collections.Specialized

#End Region

Public NotInheritable Class InvoiceNote

#Region " Public Constructor(s)"
    Public Sub New(ByVal Settings As NameValueCollection, ByVal InvNum As Integer)
        _InvoiceNo = InvNum
        Appsettings = Settings

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

    Private Appsettings As New NameValueCollection
    Private _NoteNo As Integer = 0
    Private _InvoiceNo As Integer = 0
    Private _Note As String = String.Empty
    Public Event Valid(ByVal IsValid As Boolean)
    Private _CrudMode As DBMode = DBMode.Insert
    Private _IsDirty As Boolean = False


#End Region

#Region " Properties "

    Public Property NoteNo() As Integer

        Get
            Return Me._NoteNo
        End Get

        Set(ByVal Value As Integer)

            Me._NoteNo = Value

        End Set

    End Property

    Public ReadOnly Property InvoiceNo() As Integer

        Get
            Return Me._InvoiceNo
        End Get


    End Property

    Public Property Note() As String

        Get
            Return Me._Note
        End Get

        Set(ByVal Value As String)
            Me._Note = Value
            Me.IsDirty = True

            If Value.Length > 0 Then
                RaiseEvent Valid(True)
            Else
                RaiseEvent Valid(False)
            End If

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
    Public Property IsDirty() As Boolean
        Get
            Return Me._IsDirty
        End Get
        Set(ByVal Value As Boolean)

            Me._IsDirty = Value
            If Value = False Then
                RaiseEvent Valid(False)

            End If

        End Set
    End Property
#End Region

#Region " Public Routines "

    Public Sub InsertUpdateDelete()


        Dim Params As New NameValueCollection
        Dim sbSQL As New System.Text.StringBuilder


        If Me.CrudMode = DBMode.Update Then
            With Params
                .Add("Notes", Me.Note)
                .Add("NoteNo", Me.NoteNo.ToString)
            End With

            With sbSQL
                .Append("UPDATE InvoiceNotes")
                .Append(" SET Notes=@Notes")
                .Append(" WHERE NoteNo=@NoteNo")
            End With

        ElseIf Me.CrudMode = DBMode.Insert Then
            With Params
                .Add("InvoiceNo", Me.InvoiceNo.ToString)
                .Add("Notes", Me.Note)
            End With
            With sbSQL
                .Append("INSERT INTO InvoiceNotes (")
                .Append("InvoiceNo,Notes)")
                .Append(" Values (")
                .Append("@InvoiceNo,@Notes)")
            End With
        End If

        Utils.ExecuteNonQuery(Appsettings("Connstr"), sbSQL.ToString, Params)
        Me.IsDirty = False



    End Sub
    Public Sub DeleteNoteByNoteNo(ByVal NoteNo As Integer)

        Dim Params As New NameValueCollection
        Dim sbSQL As New System.Text.StringBuilder

        With Params
            .Add("NoteNo", NoteNo.ToString)
        End With

        With sbSQL
            .Append("DELETE FROM InvoiceNotes ")
            .Append(" WHERE NoteNo=@NoteNo")
        End With

        Utils.ExecuteNonQuery(Appsettings("Connstr"), sbSQL.ToString, Params)



    End Sub
    Public Function GetAllInvoiceNotes(ByVal invNum As Integer) As DataSet


        Dim Params As New NameValueCollection
        Dim sbSQL As New System.Text.StringBuilder
        Dim ds As DataSet


        With Params
            .Add("InvoiceNo", invNum.ToString)
        End With

        With sbSQL
            .Append("SELECT NoteNo,Notes FROM InvoiceNotes ")
            .Append(" WHERE InvoiceNo=@InvoiceNo")
        End With

        ds = Utils.GetDataSet(Appsettings("Connstr"), sbSQL.ToString, Params)

        ds.Tables(0).TableName = "Notes"

        Return ds


    End Function
    Public Function GetNote(ByVal NoteNo As Integer) As String

        Dim Params As New NameValueCollection
        Dim sbSQL As New System.Text.StringBuilder
        Dim ds As DataSet


        With Params
            .Add("InvoiceNo", NoteNo.ToString)
        End With

        With sbSQL
            .Append("SELECT Notes FROM InvoiceNotes ")
            .Append(" WHERE NoteNo=@NoteNo")
        End With

        ds = Utils.GetDataSet(Appsettings("Connstr"), sbSQL.ToString, Params)


        Return ds.Tables(0).Rows(0).Item(0).ToString

    End Function

#End Region

End Class
