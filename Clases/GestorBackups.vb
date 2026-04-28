Imports System.Data.SQLite
Imports System.IO

' =====================================================================
'  GestorBackups
'  Centraliza la lógica de copias de seguridad de la BD.
'  ---------------------------------------------------------------------
'  Antes el código de backup estaba duplicado en PagHome (auto) y en
'  FrmConfiguracion (manual), con bugs:
'    - PagHome usaba una ruta de origen incorrecta ("BD\Optima.db") que
'      no existía, así que las copias automáticas NUNCA se hacían.
'    - Ambos hacían File.Copy en caliente, lo que es peligroso si la BD
'      está en medio de una escritura.
'  ---------------------------------------------------------------------
'  Ahora todo pasa por aquí y se usa el comando "VACUUM INTO" de SQLite,
'  que es la forma correcta y atómica de hacer backup de una BD viva
'  sin riesgo de corrupción ni de bloquear escrituras.
' =====================================================================
Public Class GestorBackups

    ''' <summary>Ruta absoluta del fichero Optima.db actual.</summary>
    Public Shared ReadOnly Property RutaBDActual() As String
        Get
            Return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Optima.db")
        End Get
    End Property

    ''' <summary>Ruta de la "papelera" donde se mueven las BDs sustituidas al restaurar.</summary>
    Public Shared ReadOnly Property RutaPapelera() As String
        Get
            Return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PapeleraDB")
        End Get
    End Property

    ' =================================================================
    '  Hace una copia AUTOMÁTICA. Devuelve la ruta del fichero creado o
    '  Nothing si no se pudo. No muestra mensajes al usuario.
    ' =================================================================
    Public Shared Function HacerCopiaAutomatica(carpetaDestino As String) As String
        Try
            Dim carpeta As String = carpetaDestino
            If String.IsNullOrWhiteSpace(carpeta) Then
                carpeta = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CopiasSeguridad")
            End If
            If Not Directory.Exists(carpeta) Then Directory.CreateDirectory(carpeta)

            Dim nombre As String = "OptimaDB_AUTO_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".sqlite"
            Dim destino As String = Path.Combine(carpeta, nombre)

            HacerVacuumInto(destino)
            LimpiarCopiasAntiguas(carpeta, "OptimaDB_AUTO_*.sqlite", 5)

            LogErrores.Info("GestorBackups.Auto", "Copia automática creada: " & destino)
            Return destino
        Catch ex As Exception
            ' En segundo plano no avisamos al usuario, pero sí dejamos rastro.
            LogErrores.Registrar("GestorBackups.Auto", ex)
            Return Nothing
        End Try
    End Function

    ' =================================================================
    '  Hace una copia MANUAL. Lanza excepciones si falla, para que el
    '  formulario las muestre al usuario.
    ' =================================================================
    Public Shared Function HacerCopiaManual(carpetaDestino As String) As String
        If String.IsNullOrWhiteSpace(carpetaDestino) OrElse Not Directory.Exists(carpetaDestino) Then
            Throw New ArgumentException("La carpeta de destino no es válida.")
        End If

        Dim nombre As String = "OptimaDB_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".sqlite"
        Dim destino As String = Path.Combine(carpetaDestino, nombre)

        HacerVacuumInto(destino)

        LogErrores.Info("GestorBackups.Manual", "Copia manual creada: " & destino)
        Return destino
    End Function

    ' =================================================================
    '  Restaura una BD desde un fichero. La BD actual se mueve a la
    '  papelera para que se pueda recuperar si la restauración fue
    '  un error. Lanza excepciones si falla.
    ' =================================================================
    Public Shared Sub Restaurar(rutaFicheroOrigen As String)
        If Not File.Exists(rutaFicheroOrigen) Then
            Throw New FileNotFoundException("No se encuentra el archivo de copia.", rutaFicheroOrigen)
        End If

        ' 1. Crear la papelera si no existe
        If Not Directory.Exists(RutaPapelera) Then Directory.CreateDirectory(RutaPapelera)

        ' 2. Liberar la conexión activa para que Windows suelte el fichero
        Try
            Dim c = ConexionBD.GetConnection()
            If c IsNot Nothing AndAlso c.State = ConnectionState.Open Then c.Close()
        Catch
        End Try
        SQLiteConnection.ClearAllPools()
        GC.Collect()
        GC.WaitForPendingFinalizers()
        System.Threading.Thread.Sleep(500)

        ' 3. Mover la BD actual a papelera
        If File.Exists(RutaBDActual) Then
            Dim destinoPapelera As String = Path.Combine(RutaPapelera, "BD_Eliminada_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".sqlite")
            File.Move(RutaBDActual, destinoPapelera)
        End If

        ' 4. Copiar la nueva
        File.Copy(rutaFicheroOrigen, RutaBDActual)

        LogErrores.Info("GestorBackups.Restaurar", $"BD restaurada desde {rutaFicheroOrigen}")
    End Sub

    ' =================================================================
    '  Tamaño total en MB de la papelera de BDs eliminadas.
    ' =================================================================
    Public Shared Function TamanoPapeleraMB() As Decimal
        If Not Directory.Exists(RutaPapelera) Then Return 0
        Dim bytes As Long = 0
        For Each f As FileInfo In New DirectoryInfo(RutaPapelera).GetFiles()
            bytes += f.Length
        Next
        Return Math.Round(CDec(bytes) / (1024D * 1024D), 2)
    End Function

    ''' <summary>Vacía la papelera. Devuelve el número de archivos borrados.</summary>
    Public Shared Function VaciarPapelera() As Integer
        If Not Directory.Exists(RutaPapelera) Then Return 0
        Dim n As Integer = 0
        For Each f As FileInfo In New DirectoryInfo(RutaPapelera).GetFiles()
            Try
                f.Delete()
                n += 1
            Catch
                ' Si está bloqueado, lo dejamos; el usuario verá un total más bajo
            End Try
        Next
        Return n
    End Function

    ' =================================================================
    '  IMPLEMENTACIÓN INTERNA — VACUUM INTO
    '  Esto es la forma "correcta" de hacer backup de SQLite mientras
    '  la BD está siendo usada: SQLite gestiona internamente los locks
    '  y nos da una copia consistente y compactada.
    ' =================================================================
    Private Shared Sub HacerVacuumInto(destino As String)
        ' Si ya existe (raro, mismo segundo), borramos para evitar error.
        If File.Exists(destino) Then File.Delete(destino)

        Dim c = ConexionBD.GetConnection()
        If c.State <> ConnectionState.Open Then c.Open()

        ' VACUUM INTO no acepta parámetros en algunos drivers, así que
        ' hay que escapar las comillas simples del path por seguridad.
        Dim destinoEscapado As String = destino.Replace("'", "''")
        Using cmd As New SQLiteCommand($"VACUUM INTO '{destinoEscapado}'", c)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    ' =================================================================
    '  Mantiene como mucho N copias automáticas. Las más antiguas se
    '  borran. Solo afecta a las copias que coinciden con el patrón
    '  pasado (no toca las copias manuales).
    ' =================================================================
    Private Shared Sub LimpiarCopiasAntiguas(carpeta As String, patron As String, maximo As Integer)
        Try
            Dim dir As New DirectoryInfo(carpeta)
            Dim archivos = dir.GetFiles(patron) _
                              .OrderBy(Function(f) f.CreationTime) _
                              .ToList()
            If archivos.Count > maximo Then
                For i As Integer = 0 To archivos.Count - maximo - 1
                    Try : archivos(i).Delete() : Catch : End Try
                Next
            End If
        Catch ex As Exception
            LogErrores.Registrar("GestorBackups.LimpiarCopiasAntiguas", ex)
        End Try
    End Sub

End Class
