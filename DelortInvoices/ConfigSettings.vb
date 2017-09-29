Option Strict On
Option Explicit On 

Imports System.Collections.Specialized
Imports System.Reflection.Assembly

Public Class ConfigSettings

    Public Shared Function GetAppSettings(ByVal ConfigFileLoc As String) As NameValueCollection

        Dim settings As New NameValueCollection
        Dim reader As System.Xml.XmlTextReader = Nothing
        Dim doc As New System.Xml.XmlDocument

        If Not System.IO.File.Exists(ConfigFileLoc) Then
            Return settings
        End If

        Try
            reader = New System.Xml.XmlTextReader(ConfigFileLoc)

            doc.Load(reader)

            Dim configurationNode As System.Xml.XmlNode = _
            doc.SelectSingleNode("configuration")

            If configurationNode Is Nothing Then
                Return settings
            End If

            Dim appSettingsNode As System.Xml.XmlNode = _
            configurationNode.SelectSingleNode("appSettings")

            If appSettingsNode Is Nothing Then
                Return settings
            End If

            Dim nodes As System.Xml.XmlNodeList = _
            appSettingsNode.SelectNodes("add")

            For Each node As System.Xml.XmlNode In nodes
                settings.Add(node.Attributes("key").Value, node.Attributes("value").Value)
            Next

        Catch ex As System.Xml.XmlException
            'do nothing, for now

        Finally
            reader.Close()
        End Try

        Return settings

    End Function


End Class

