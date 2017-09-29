
#Region " Options "

Option Explicit On
Option Strict On

#End Region

#Region " Imports "

Imports System.Collections.Specialized
Imports AccessUtils

#End Region

Public Class BaseBusiness

    Protected Appsettings As New NameValueCollection

#Region " Public Constructors "

    Public Sub New(ByVal Settings As NameValueCollection)

        Me.Appsettings = Settings

    End Sub

#End Region

End Class
