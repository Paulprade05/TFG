Imports Microsoft.VisualBasic.ApplicationServices

Namespace My
    ' The following events are available for MyApplication:
    ' Startup: Raised when the application starts, before the startup form is created.
    ' Shutdown: Raised after all application forms are closed.  This event is not raised if the application terminates abnormally.
    ' UnhandledException: Raised if the application encounters an unhandled exception.
    ' StartupNextInstance: Raised when launching a single-instance application and the application is already active. 
    ' NetworkAvailabilityChanged: Raised when the network connection is connected or disconnected.

    Partial Friend Class MyApplication

        ' Captura de excepciones no manejadas: las registramos en optima_log.txt
        ' y mostramos un mensaje al usuario en lugar de que la app se cierre sin más.
        Private Sub MyApplication_UnhandledException(sender As Object, e As UnhandledExceptionEventArgs) Handles Me.UnhandledException
            Try
                LogErrores.Registrar("UnhandledException", e.Exception)
            Catch
                ' Si el propio log falla, no podemos hacer mucho más.
            End Try

            Try
                MessageBox.Show(
                    "Se ha producido un error inesperado y se ha registrado en el archivo de log." & vbCrLf & vbCrLf &
                    "Detalle: " & e.Exception.Message,
                    "Error inesperado",
                    MessageBoxButtons.OK, MessageBoxIcon.Error)
            Catch
            End Try

            ' Indicamos a la aplicación que continúe (no se cierre por este error).
            e.ExitApplication = False
        End Sub

        Private Sub MyApplication_Startup(sender As Object, e As StartupEventArgs) Handles Me.Startup
            Try
                LogErrores.Info("Application", "Inicio de aplicación OPTIMA")
            Catch
            End Try
        End Sub

    End Class
End Namespace
