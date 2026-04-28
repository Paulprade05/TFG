Imports System.Data.SQLite ' Necesitas el paquete NuGet System.Data.SQLite
Imports System.IO

Public Class ConexionBD
    ' 1. Variable estática (Shared en VB) para la instancia única
    Private Shared _conexion As SQLiteConnection

    ' 2. Ruta de la base de datos (Unificada)
    Private Shared ReadOnly _dbPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Optima.db")
    Private Shared ReadOnly _connectionString As String = $"Data Source={_dbPath}"

    ' 3. Constructor privado (Private Sub New) para evitar instanciar desde fuera
    Private Sub New()
    End Sub

    ' 4. Método COMPARTIDO (Shared) para obtener la conexión desde cualquier parte
    Public Shared Function GetConnection() As SQLiteConnection
        If _conexion Is Nothing Then
            CrearBaseDatosSiNoExiste()
            Try
                _conexion = New SQLiteConnection(_connectionString)
                _conexion.Open()
                ' === ACTIVAR FOREIGN KEYS ===
                ' Por defecto SQLite NO aplica las FK declaradas. Hay que activarlas explícitamente
                ' en cada conexión nueva. Sin esto, los ON DELETE CASCADE no se ejecutan y se pueden
                ' insertar registros con referencias rotas.
                ActivarForeignKeys()
            Catch ex As Exception
                MessageBox.Show("Error crítico de conexión: " & ex.Message)
                Return Nothing ' Salimos aquí para evitar el error de instancia
            End Try
        End If

        ' Solo verificamos el estado si la conexión realmente existe
        If _conexion IsNot Nothing AndAlso _conexion.State = ConnectionState.Closed Then
            _conexion.Open()
            ActivarForeignKeys() ' Re-activar FK al reabrir
        End If

        Return _conexion
    End Function

    Private Shared Sub ActivarForeignKeys()
        Try
            Using cmd As New SQLiteCommand("PRAGMA foreign_keys = ON;", _conexion)
                cmd.ExecuteNonQuery()
            End Using
        Catch
            ' Si por algún motivo no se pueden activar, seguimos. La app funciona, solo que sin
            ' integridad referencial automática (que es como estaba antes de este parche).
        End Try
    End Sub

    Private Shared Sub CrearBaseDatosSiNoExiste()
        If Not File.Exists(_dbPath) Then
            SQLiteConnection.CreateFile(_dbPath)
        End If
    End Sub
End Class
