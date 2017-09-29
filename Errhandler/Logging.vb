
#Region " Options "

Option Explicit On
Option Strict On

#End Region

#Region " Imports "

Imports System
Imports System.IO

#End Region

Public Class Logging

#Region " Shared methods "

    Shared Sub LogError(ByVal ex As Exception)

        Dim sw As StreamWriter = File.AppendText("DelortErrors.log")

        With sw
            .Write(ControlChars.CrLf & "Log Entry : ")
            .WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString())
            .WriteLine("{0}", ex.Message & " --- " & ex.StackTrace)
            .WriteLine("-------------------------------")
            ' Update the underlying file.
            .Flush()
            .Close()
        End With

    End Sub

#End Region

End Class
