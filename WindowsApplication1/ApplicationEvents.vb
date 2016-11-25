Namespace My

    ' The following events are available for MyApplication:
    ' 
    ' Startup: Raised when the application starts, before the startup form is created.
    ' Shutdown: Raised after all application forms are closed.  This event is not raised if the application terminates abnormally.
    ' UnhandledException: Raised if the application encounters an unhandled exception.
    ' StartupNextInstance: Raised when launching a single-instance application and the application is already active. 
    ' NetworkAvailabilityChanged: Raised when the network connection is connected or disconnected.
    Partial Friend Class MyApplication

        Private Sub MyApplication_Startup(sender As Object, e As ApplicationServices.StartupEventArgs) Handles Me.Startup
            Try
                Dim prcProcess As Process() = Process.GetProcessesByName("ffmpeg")
                For Each Item As Process In prcProcess
                    Item.Kill()
                    Item.Close()
                    Item = Nothing
                Next
            Catch ex As Exception
            End Try
        End Sub

    End Class

End Namespace

