Imports System.Collections.Specialized
Imports System.Configuration.ConfigurationSettings
Imports System.Configuration

Module mMain

    Private frmSplash As New Splash
    Private Mainform As New frmMain
    Public ProgramName As String = "Delort and Sons"
    Public g_Settings As New NameValueCollection


    Public Sub Main()

        g_Settings = ConfigSettings.GetAppSettings(ConfigurationManager.AppSettings("ConfigLoc"))


        Dim appProc() As Process
        Dim strModName, strProcName As String

        Try

            strModName = Process.GetCurrentProcess.MainModule.ModuleName()
            strProcName = System.IO.Path.GetFileNameWithoutExtension(strModName)
            appProc = Process.GetProcessesByName(strProcName)

            If appProc.Length > 1 Then
                MessageBox.Show("There is an instance of this application running.", ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Hand)
                Exit Sub
            End If

            frmSplash.ShowDialog()

            If Not IsNothing(frmSplash) Then
                frmSplash.Dispose()
                frmSplash = Nothing
            End If

            Mainform.ShowDialog()

        Catch ex As Exception
            MessageBox.Show("Unexpected error in application", "Application Error", _
            MessageBoxButtons.OK, MessageBoxIcon.Stop)

        Finally

            Application.Exit()

            If Not IsNothing(Mainform) Then
                Mainform.Dispose()
                Mainform = Nothing
            End If

            If Not IsNothing(frmSplash) Then
                frmSplash.Dispose()
            End If

        End Try

    End Sub


End Module
