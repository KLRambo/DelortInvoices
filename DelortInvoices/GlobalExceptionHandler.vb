Public Class GlobalExceptionHandler

    Public Sub OnThreadException(ByVal sender As Object, _
                ByVal t As System.Threading.ThreadExceptionEventArgs)


        Dim result As DialogResult

        Try
            result = MessageBox.Show("Unexpected error in application", "Application Error", _
            MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Stop)

        Catch ex As Exception
            Try
                MessageBox.Show("Fatal Error", "Fatal Error", MessageBoxButtons.OK, MessageBoxIcon.Stop)

            Finally
                Application.Exit()

            End Try

            If result = DialogResult.Abort Then
                Application.Exit()
            End If
        End Try
    End Sub

End Class
