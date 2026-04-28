Imports System.IO

' =====================================================================
'  LogErrores
'  Utilidad simple de logging a fichero para registrar errores que
'  hasta ahora solo se mostraban en un MessageBox y se perdían.
'  ---------------------------------------------------------------------
'  Escribe en la carpeta de la aplicación, fichero "optima_log.txt".
'  Formato:  [yyyy-MM-dd HH:mm:ss] [USR=Paul] [Modulo] Mensaje
'  Si el fichero crece más de 1 MB lo rota a "optima_log.old.txt".
'  ---------------------------------------------------------------------
'  Uso típico:
'    Catch ex As Exception
'        LogErrores.Registrar("FrmFacturas.Guardar", ex)
'        MessageBox.Show("Ha ocurrido un error...")
'    End Try
' =====================================================================
Public Class LogErrores

    Private Shared ReadOnly _ruta As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "optima_log.txt")
    Private Shared ReadOnly _rutaOld As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "optima_log.old.txt")
    Private Shared ReadOnly _lock As New Object()
    Private Const MAX_BYTES As Long = 1024 * 1024 ' 1 MB

    ''' <summary>Registra una excepción incluyendo stack trace.</summary>
    Public Shared Sub Registrar(modulo As String, ex As Exception)
        Try
            EscribirLinea($"[ERROR] [{modulo}] {ex.GetType().Name}: {ex.Message}")
            If ex.StackTrace IsNot Nothing Then
                EscribirLinea($"        StackTrace: {ex.StackTrace.Replace(vbCrLf, " | ")}")
            End If
        Catch
            ' Si el log mismo falla, no propagamos: es preferible perder el log
            ' a que se rompa el flujo del usuario.
        End Try
    End Sub

    ''' <summary>Registra un evento informativo.</summary>
    Public Shared Sub Info(modulo As String, mensaje As String)
        Try
            EscribirLinea($"[INFO]  [{modulo}] {mensaje}")
        Catch
        End Try
    End Sub

    ''' <summary>Registra un aviso (algo no crítico pero anormal).</summary>
    Public Shared Sub Aviso(modulo As String, mensaje As String)
        Try
            EscribirLinea($"[AVISO] [{modulo}] {mensaje}")
        Catch
        End Try
    End Sub

    Private Shared Sub EscribirLinea(texto As String)
        SyncLock _lock
            ' Rotación simple si el fichero ha crecido demasiado
            Try
                If File.Exists(_ruta) AndAlso New FileInfo(_ruta).Length > MAX_BYTES Then
                    If File.Exists(_rutaOld) Then File.Delete(_rutaOld)
                    File.Move(_ruta, _rutaOld)
                End If
            Catch
                ' No bloqueante
            End Try

            Dim usr As String = If(ComunSesionActual.Usuario, "?")
            Dim linea As String = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [USR={usr}] {texto}"
            File.AppendAllText(_ruta, linea & Environment.NewLine)
        End SyncLock
    End Sub

End Class
