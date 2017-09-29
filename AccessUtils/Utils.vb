
#Region " Options "

Option Explicit On
Option Strict On

#End Region

#Region " Imports "

Imports System.Collections.Specialized
Imports System.Data.OleDb

#End Region

Public Class Utils

#Region " Shared methods "

    Shared Sub ExecuteNonQuery(ByVal Connstr As String, _
                               ByVal SQL As String, _
                               ByVal Params As NameValueCollection)

        Dim cnnOpen As New OleDbConnection(Connstr)
        Dim cmd As New OleDbCommand

        Try

            cnnOpen.Open()

            cmd = New OleDbCommand(SQL, cnnOpen)

            With cmd

                .CommandType = CommandType.Text
                .CommandTimeout = 500

                If Not IsNothing(Params) Then
                    For Each param As String In Params.AllKeys
                        .Parameters.AddWithValue(String.Format("@{0}", param), Params(param))
                    Next
                End If

                .ExecuteNonQuery()

            End With

        Catch ex As Exception
            Throw ex

        Finally
            cnnOpen.Close()
            cnnOpen.Dispose()

        End Try

    End Sub

    Shared Function GetDataSet(ByVal Connstr As String, _
                               ByVal SQL As String, _
                               ByVal Params As NameValueCollection) As DataSet

        Dim cnnOpen As New OleDbConnection(Connstr)
        Dim cmd As New OleDbCommand
        Dim ds As New DataSet
        Dim da As New OleDbDataAdapter

        Try
            cnnOpen.Open()
            cmd = New OleDbCommand(SQL, cnnOpen)

            With cmd
                .CommandType = CommandType.Text
                .CommandTimeout = 500
                .Parameters.Clear()
                If Not Params Is Nothing Then
                    For Each param As String In Params.AllKeys
                        .Parameters.AddWithValue(String.Format("@{0}", param), Params(param))
                    Next
                End If

            End With

            With da
                .SelectCommand = cmd
                .Fill(ds)
            End With

            Return ds

        Finally
            'clean up
            cnnOpen.Close()
            cnnOpen.Dispose()
            cmd.Dispose()
            ds.Dispose()
            da.Dispose()
        End Try


    End Function

    Shared Function GetDatasetWithMultipleTables(ByVal Connstr As String, _
                                                 ByVal SQL() As String, _
                                                 ByVal ParamsAll() As NameValueCollection, _
                                                 ByVal ds As DataSet) As DataSet


        Dim cnnOpen As New OleDbConnection(Connstr)
        Dim cmd As New OleDbCommand
        Dim da As New OleDbDataAdapter
        Dim Count As Integer
        Dim params As NameValueCollection

        Try
            cnnOpen.Open()

            For Each Table As DataTable In ds.Tables
                cmd = New OleDbCommand(SQL(Count), cnnOpen)

                With cmd
                    .CommandType = CommandType.Text
                    .CommandTimeout = 500
                    .Parameters.Clear()
                    If Not ParamsAll Is Nothing Then
                        params = ParamsAll(Count)
                        If Not params Is Nothing Then
                            For Each param As String In params.AllKeys
                                .Parameters.AddWithValue(String.Format("@{0}", param), params(param))
                            Next
                        End If
                    End If
                End With

                With da
                    .SelectCommand = cmd
                    .Fill(ds.Tables(Count))
                End With

                Count += 1

            Next

            Return ds

        Finally
            'clean up
            cnnOpen.Close()
            cnnOpen.Dispose()
            cmd.Dispose()
            ds.Dispose()
            da.Dispose()

        End Try

    End Function

    Shared Function GetDataReader(ByVal Connstr As String, _
                                  ByVal SQL As String, _
                                  ByVal Params As NameValueCollection) As OleDbDataReader


        Dim cnnOpen As New OleDbConnection(Connstr)
        Dim cmd As New OleDbCommand
        Dim dr As OleDb.OleDbDataReader = Nothing

        cnnOpen.Open()
        cmd = New OleDbCommand(SQL, cnnOpen)

        With cmd
            .CommandType = CommandType.Text
            .CommandTimeout = 500
            .Parameters.Clear()
            If Not Params Is Nothing Then
                For Each param As String In Params.AllKeys
                    .Parameters.AddWithValue(String.Format("@{0}", param), Params(param))
                Next
                dr = .ExecuteReader
            End If
        End With

        Return dr

    End Function

    Shared Sub ExecuteTransNonQuery(ByVal Connstr As String, _
                                    ByVal SQL1 As String, _
                                    ByVal SQL2 As String, _
                                    ByVal Params1 As NameValueCollection, _
                                    ByVal Params2() As NameValueCollection)


        Dim cnnOpen As New OleDbConnection(Connstr)
        Dim cmd As New OleDbCommand
        Dim Trans As OleDbTransaction = Nothing
        Dim Count As Integer

        Try
            cnnOpen.Open()

            Trans = cnnOpen.BeginTransaction(IsolationLevel.ReadUncommitted)

            cmd = New OleDbCommand(SQL1, cnnOpen)

            cmd.Transaction = Trans

            With cmd
                .CommandType = CommandType.Text
                .CommandTimeout = 500

                For Each param As String In Params1.AllKeys
                    .Parameters.AddWithValue(String.Format("@{0}", param), Params1(param))
                Next
                .ExecuteNonQuery()
            End With

            'line items here

            For Count = 0 To Params2.Length - 1
                With cmd
                    .Parameters.Clear()
                    .CommandText = SQL2

                    For Each param As String In Params2(Count).AllKeys
                        .Parameters.AddWithValue(String.Format("@{0}", param), Params2(Count)(param))
                    Next

                    .ExecuteNonQuery()
                End With
            Next Count

            'After everything is done, commit
            Trans.Commit()

        Catch ex As Exception
            Trans.Rollback()

            Throw ex

        Finally

            cnnOpen.Close()
            cnnOpen.Dispose()

        End Try

    End Sub

    Shared Sub InsertUpdateDelete(ByVal Connstr As String, _
                                  ByVal InvoiceSQl As String, _
                                  ByVal InsertSQL As String, _
                                  ByVal UpdateSQL As String, _
                                  ByVal DeleteSQL As String, _
                                  ByVal InvoiceParams As NameValueCollection, _
                                  ByVal InsertParams() As NameValueCollection, _
                                  ByVal UpdateParams() As NameValueCollection, _
                                  ByVal DeleteParams() As NameValueCollection)

        Dim cnnOpen As New OleDbConnection(Connstr)
        Dim cmd As New OleDbCommand
        Dim Count As Integer

        Try

            cnnOpen.Open()

            cmd = New OleDbCommand(InvoiceSQl, cnnOpen)

            With cmd
                .CommandType = CommandType.Text
                .CommandTimeout = 500

                For Each param As String In InvoiceParams.AllKeys
                    .Parameters.AddWithValue(String.Format("@{0}", param), InvoiceParams(param))
                Next
                .ExecuteNonQuery()
            End With

            Count = 0

            '//////////////////////////////////////////////////////////
            'Do the insert here, if any
            If InsertSQL.Length > 0 Then

                For Count = 0 To InsertParams.Length - 1
                    With cmd
                        .Parameters.Clear()
                        .CommandText = InsertSQL
                        For Each param As String In InsertParams(Count).AllKeys
                            .Parameters.AddWithValue(String.Format("@{0}", param), InsertParams(Count)(param))
                        Next

                        .ExecuteNonQuery()
                    End With
                Next Count
            End If
            '//////////////////////////////////////////////////////////

            Count = 0
            '//////////////////////////////////////////////////////////
            'Do the updates here

            If UpdateSQL.Length > 0 Then

                For Count = 0 To UpdateParams.Length - 1
                    With cmd
                        .Parameters.Clear()
                        .CommandText = UpdateSQL

                        For Each param As String In UpdateParams(Count).AllKeys
                            .Parameters.AddWithValue(String.Format("@{0}", param), UpdateParams(Count)(param))
                        Next

                        .ExecuteNonQuery()
                    End With
                Next Count
            End If

            Count = 0

            '//////////////////////////////////////////////////////////
            'Do the Deletes here
            If DeleteSQL.Length > 0 Then
                For Count = 0 To DeleteParams.Length - 1
                    With cmd
                        .Parameters.Clear()
                        .CommandText = DeleteSQL

                        For Each param As String In DeleteParams(Count).AllKeys
                            .Parameters.AddWithValue(String.Format("@{0}", param), DeleteParams(Count)(param))
                        Next
                        .ExecuteNonQuery()
                    End With
                Next Count
            End If

        Catch ex As Exception

            Throw ex

        Finally

            cnnOpen.Close()
            cnnOpen.Dispose()

        End Try

    End Sub

    Shared Function ExecuteScalar(ByVal Connstr As String, _
                                  ByVal SQL As String, _
                                  ByVal Params As NameValueCollection) As Integer

        Dim cnnOpen As New OleDbConnection(Connstr)
        Dim cmd As New OleDbCommand
        Dim Count As Integer

        Try
            cnnOpen.Open()
            cmd = New OleDbCommand(SQL, cnnOpen)

            With cmd
                .CommandType = CommandType.Text
                .CommandTimeout = 500
                If Not Params Is Nothing Then
                    For Each param As String In Params.AllKeys
                        .Parameters.AddWithValue(String.Format("@{0}", param), Params(param))
                    Next
                End If
            End With

            Count = Convert.ToInt32(cmd.ExecuteScalar())

            Return Count

        Finally
            cnnOpen.Close()
            cnnOpen.Dispose()

        End Try

    End Function

#End Region

End Class
