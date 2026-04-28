Imports System.Data.SQLite

' =====================================================================
'  NumeradorDocumentos
'  Helper centralizado para generar el próximo número correlativo de
'  un documento (PRE-, PED-, ALB-, FAC-, etc.).
'  ---------------------------------------------------------------------
'  Antes cada formulario implementaba su propia función con esta consulta:
'    SELECT NumeroX FROM TablaX WHERE NumeroX LIKE 'PRE-%' ORDER BY NumeroX DESC LIMIT 1
'  Eso es ORDEN LEXICOGRÁFICO, así que al pasar de PRE-999 a PRE-1000 el
'  string "PRE-999" sigue siendo "mayor" que "PRE-1000" alfabéticamente y
'  la numeración se rompe.
'
'  Esta utilidad ordena por la parte numérica (después del último guión)
'  y soporta cualquier cantidad de dígitos. Empieza desde 001 (3 dígitos)
'  pero se amplía solo si hace falta.
' =====================================================================
Public Class NumeradorDocumentos

    ''' <summary>
    ''' Devuelve el siguiente número correlativo de un documento.
    ''' Ejemplo: SiguienteNumero("ALB-", "Albaranes", "NumeroAlbaran") -> "ALB-005"
    ''' </summary>
    ''' <param name="prefijo">Prefijo incluido el guión, p.ej. "ALB-"</param>
    ''' <param name="tabla">Tabla en la que buscar el último número</param>
    ''' <param name="columna">Columna con el número del documento</param>
    Public Shared Function SiguienteNumero(prefijo As String, tabla As String, columna As String) As String
        Dim siguiente As Integer = 1

        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()

            ' Ordenamos por la parte numérica posterior al último guión (CAST a INTEGER)
            ' para que después de XXX-999 venga XXX-1000 y no se rompa el orden lexicográfico.
            Dim sql As String = $"SELECT MAX(CAST(SUBSTR({columna}, LENGTH(@prefijo) + 1) AS INTEGER)) " &
                                $"FROM {tabla} WHERE {columna} LIKE @pat"

            Using cmd As New SQLiteCommand(sql, c)
                cmd.Parameters.AddWithValue("@prefijo", prefijo)
                cmd.Parameters.AddWithValue("@pat", prefijo & "%")
                Dim res = cmd.ExecuteScalar()
                If res IsNot Nothing AndAlso Not IsDBNull(res) Then
                    siguiente = Convert.ToInt32(res) + 1
                End If
            End Using
        Catch
            ' Si por lo que sea falla, generamos uno único pero "feo" basado en ticks,
            ' para no bloquear al usuario. Mejor que crashear.
            Return prefijo & DateTime.Now.Ticks.ToString().Substring(12)
        End Try

        ' Padding mínimo a 3 dígitos. Si el número es mayor de 999, el formato D3 lo respeta
        ' y no recorta nada (1000 -> "1000"). Solo añade ceros por la izquierda hasta llegar a 3.
        Return prefijo & siguiente.ToString("D3")
    End Function

End Class
