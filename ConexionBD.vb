Imports System.Data.SQLite ' Necesitas el paquete NuGet System.Data.SQLite
Imports System.IO

Public Class ConexionBD
    ' 1. Variable estática (Shared en VB) para la instancia única
    Private Shared _conexion As SQLiteConnection

    ' 2. Ruta de la base de datos
    Private Shared _dbPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "BD\Optima.db")
    Private Shared _connectionString As String = "Data Source=" & _dbPath & ";Version=3;"

    ' 3. Constructor privado (Private Sub New) para evitar instanciar desde fuera
    Private Sub New()
    End Sub

    ' 4. Método COMPARTIDO (Shared) para obtener la conexión desde cualquier parte
    Public Shared Function GetConnection() As SQLiteConnection
        ' Si la conexión no existe, la configuramos
        If _conexion Is Nothing Then
            CrearBaseDatosSiNoExiste()
            Try
                _conexion = New SQLiteConnection(_connectionString)
                _conexion.Open()
            Catch ex As Exception
                MessageBox.Show("Error al conectar a la base de datos " & _dbPath)
            End Try

        End If

        ' Si por alguna razón se cerró, la reabrimos
        If _conexion.State <> System.Data.ConnectionState.Open Then
            _conexion.Open()
        End If

        Return _conexion
    End Function

    Private Shared Sub CrearBaseDatosSiNoExiste()
        If Not File.Exists(_dbPath) Then
            SQLiteConnection.CreateFile(_dbPath)
        End If
    End Sub
End Class