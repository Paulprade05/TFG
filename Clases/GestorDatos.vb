Imports System.Data.SQLite

Public Class GestorDatos

    ' ===================================================================
    ' MÉTODO 1: Para cargar UNA TABLA ENTERA (Ej: "Clientes", "Articulos")
    ' Uso: Dim dt As DataTable = GestorDatos.LeerTabla("Clientes")
    ' ===================================================================
    Public Shared Function LeerTabla(nombreTabla As String) As DataTable
        Dim dt As New DataTable()

        Try
            Dim conexion = ConexionBD.GetConnection()

            ' Abrimos conexión si no está abierta
            If conexion.State <> ConnectionState.Open Then conexion.Open()

            ' NOTA: Los nombres de tabla no se pueden pasar como parámetros (@tabla),
            ' así que concatenamos el string. Asegúrate de escribir bien el nombre en el código.
            Dim sql As String = $"SELECT * FROM {nombreTabla}"

            Using cmd As New SQLiteCommand(sql, conexion)
                Using da As New SQLiteDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show($"Error al leer la tabla '{nombreTabla}': {ex.Message}", "Error BD", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return dt
    End Function

    ' ===================================================================
    ' MÉTODO 2: Para consultas PERSONALIZADAS (Con WHERE, ORDER BY, JOIN...)
    ' Uso: Dim dt = GestorDatos.EjecutarSelect("SELECT * FROM Pedidos WHERE Estado = 'Pendiente'")
    ' ===================================================================
    Public Shared Function EjecutarSelect(sql As String) As DataTable
        Dim dt As New DataTable()

        Try
            Dim conexion = ConexionBD.GetConnection()
            If conexion.State <> ConnectionState.Open Then conexion.Open()

            Using cmd As New SQLiteCommand(sql, conexion)
                Using da As New SQLiteDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show($"Error en la consulta: {ex.Message}", "Error SQL", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return dt
    End Function

End Class