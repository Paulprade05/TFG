Imports System.Data.SQLite

Public Class FrmPedidos

    ' CAMBIO: Usamos String para el código del pedido (Ej: "PED-26-001")
    Private _numeroPedidoActual As String = ""
    Private _dtLineas As DataTable
    Private _idsParaBorrar As New List(Of Integer)
    ' Asegúrate de que llevan el "New"
    Private WithEvents cboFormaPago As New ComboBox()
    Private WithEvents cboRuta As New ComboBox()
    Private lblFormaPago As New Label() With {.Text = "Forma de Pago", .AutoSize = True, .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold), .ForeColor = Color.WhiteSmoke}
    Private lblRuta As New Label() With {.Text = "Ruta Asignada", .AutoSize = True, .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold), .ForeColor = Color.WhiteSmoke}

    ' Método para dar formato profesional a los selectores de fecha


    Private Sub CalculoTotalesVisuales(ByRef fila As DataGridViewRow)
        If fila.IsNewRow Then Return
        Try
            Dim cant As Decimal : Decimal.TryParse(fila.Cells("Cantidad").Value?.ToString(), cant)
            Dim prec As Decimal = 0 : Decimal.TryParse(fila.Cells("PrecioUnitario").Value?.ToString(), prec)
            Dim desc As Decimal = 0 : Decimal.TryParse(fila.Cells("Descuento").Value?.ToString(), desc)

            Dim total As Decimal = (cant * prec) - ((cant * prec) * (desc / 100))
            fila.Cells("Total").Value = total
        Catch ex As Exception
            fila.Cells("Total").Value = 0
        End Try
    End Sub

    ' Helper actualizado para recibir Total
    Private Sub EjecutarComandoLinea(sql As String, conn As SQLiteConnection, trans As SQLiteTransaction, idPadre As Integer, orden As Integer, idArt As Object, row As DataRow, idLinea As Object)
        Using cmd As New SQLiteCommand(sql, conn)
            cmd.Transaction = trans
            If sql.Contains("INSERT") Then cmd.Parameters.AddWithValue("@idPadre", idPadre)
            cmd.Parameters.AddWithValue("@orden", orden)
            cmd.Parameters.AddWithValue("@idArt", idArt)
            cmd.Parameters.AddWithValue("@desc", row("Descripcion"))
            cmd.Parameters.AddWithValue("@cant", row("Cantidad"))
            cmd.Parameters.AddWithValue("@prec", row("PrecioUnitario"))
            cmd.Parameters.AddWithValue("@dcto", row("Descuento"))
            ' AÑADIDO: Parámetro Total
            cmd.Parameters.AddWithValue("@tot", row("Total"))

            If idLinea IsNot Nothing Then cmd.Parameters.AddWithValue("@id", idLinea)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    ' Helper para reducir código repetido en Guardar




    Private Function ObtenerUltimoID() As Integer
        Dim id As Integer = 0
        Try
            Dim conexion = ConexionBD.GetConnection()
            ' CAMBIO 7: Tabla Pedidos
            Using cmd As New SQLiteCommand("SELECT MAX(ID_Pedido) FROM Pedidos", conexion)
                Dim res = cmd.ExecuteScalar()
                If res IsNot Nothing AndAlso Not IsDBNull(res) Then id = Convert.ToInt32(res)
            End Using
        Catch
        End Try
        Return id
    End Function

    Private Function GenerarProximoNumero() As String
        ' CAMBIO 8: Prefijo PED
        Dim prefijo As String = $"PED-"
        Dim nuevoNumero As String = $"{prefijo}001"

        Try
            Dim conexion = ConexionBD.GetConnection()
            ' CAMBIO 9: Buscar en Pedidos
            Dim sql As String = "SELECT NumeroPedido FROM Pedidos WHERE NumeroPedido LIKE @patron ORDER BY ID_Pedido DESC LIMIT 1"
            Using cmd As New SQLiteCommand(sql, conexion)
                cmd.Parameters.AddWithValue("@patron", prefijo & "%")
                Dim res = cmd.ExecuteScalar()
                If res IsNot Nothing AndAlso Not IsDBNull(res) Then
                    Dim partes As String() = res.ToString().Split("-"c)
                    If partes.Length = 3 Then
                        nuevoNumero = $"{prefijo}{(Convert.ToInt32(partes(2)) + 1).ToString("D3")}"
                    End If
                End If
            End Using
        Catch
        End Try
        Return nuevoNumero
    End Function

    ' --- EVENTOS Y LÓGICA COMPARTIDA ---
    ' (Copiar igual que en Presupuestos: LimpiarFormulario, DataGridView_CellEndEdit, CalculoTotales, etc.)
    ' Solo recuerda cambiar _idPresupuestoActual por _numeroPedidoActual en esos métodos.






    ' Este evento salta cuando el foco sale de la caja de texto ID Cliente
    ' O cuando lo llamamos manualmente desde el código (como en la importación)

    Private Sub ButtonNuevo_Click(sender As Object, e As EventArgs) Handles ButtonNuevoPresup.Click
        ' Simplemente llamamos al método de limpieza que ya se encarga de todo
        LimpiarFormulario()
    End Sub
    Private Sub ButtonBorrar_Click(sender As Object, e As EventArgs) Handles ButtonBorrar.Click
        ' 1. Seguridad: Si no hay pedido cargado (ID=0), no hacemos nada
        If _numeroPedidoActual = "" Then
            MessageBox.Show("No puedes borrar un pedido que aún no existe (es nuevo).")
            Return
        End If

        ' 2. Confirmación del usuario
        If MessageBox.Show("¿Estás seguro de ELIMINAR este pedido y todas sus líneas?", "Confirmar Borrado", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No Then
            Return
        End If

        Dim conexion = ConexionBD.GetConnection()
        Dim transaccion As SQLiteTransaction = Nothing

        Try
            If conexion.State <> ConnectionState.Open Then conexion.Open()
            transaccion = conexion.BeginTransaction()

            ' ====================================================================
            ' PASO A: Borrar las LÍNEAS (Tabla Hija)
            ' ====================================================================
            Dim sqlLineas As String = "DELETE FROM LineasPedido WHERE ID_Pedido = @id"
            Using cmd As New SQLiteCommand(sqlLineas, conexion)
                cmd.Transaction = transaccion
                cmd.Parameters.AddWithValue("@id", _numeroPedidoActual)
                cmd.ExecuteNonQuery()
            End Using

            ' ====================================================================
            ' PASO B: Borrar la CABECERA (Tabla Padre)
            ' ====================================================================
            Dim sqlCabecera As String = "DELETE FROM Pedidos WHERE ID_Pedido = @id"
            Using cmd As New SQLiteCommand(sqlCabecera, conexion)
                cmd.Transaction = transaccion
                cmd.Parameters.AddWithValue("@id", _numeroPedidoActual)
                cmd.ExecuteNonQuery()
            End Using

            ' Confirmar cambios
            transaccion.Commit()

            MessageBox.Show("Pedido eliminado correctamente.", "Eliminado", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' 3. Reseteamos el formulario
            LimpiarFormulario()

        Catch ex As Exception
            ' Si falla, deshacemos todo
            If transaccion IsNot Nothing Then transaccion.Rollback()
            MessageBox.Show("Error crítico al borrar: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles ButtonGuardar.Click
        ' 1. VALIDACIÓN BÁSICA
        If String.IsNullOrWhiteSpace(TextBoxIdCliente.Text) Then
            MessageBox.Show("Error: Debes seleccionar un cliente válido.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim esNuevo As Boolean = String.IsNullOrEmpty(_numeroPedidoActual)
        If esNuevo Then
            TextBoxPedido.Text = GenerarProximoNumeroPedido()
            _numeroPedidoActual = TextBoxPedido.Text
        End If

        Dim conexion = ConexionBD.GetConnection()
        Dim transaccion As SQLiteTransaction = Nothing

        Try
            If conexion.State <> ConnectionState.Open Then conexion.Open()
            transaccion = conexion.BeginTransaction()

            ' =========================================================
            ' 2. CÁLCULO DE TOTALES (De las líneas)
            ' =========================================================
            Dim sumaBase As Decimal = 0
            If _dtLineas IsNot Nothing Then
                For Each row As DataRow In _dtLineas.Rows
                    If row.RowState <> DataRowState.Deleted Then
                        Dim c As Decimal = 0 : Decimal.TryParse(row("Cantidad").ToString(), c)
                        Dim p As Decimal = 0 : Decimal.TryParse(row("PrecioUnitario").ToString(), p)
                        Dim d As Decimal = 0 : Decimal.TryParse(row("Descuento").ToString(), d)
                        Dim tot As Decimal = (c * p) * (1 - (d / 100))
                        If _dtLineas.Columns.Contains("Total") Then row("Total") = tot
                        sumaBase += tot
                    End If
                Next
            End If
            Dim sumaTotal As Decimal = sumaBase * 1.21D ' Asumiendo 21% de IVA general

            ' =========================================================
            ' 3. PREPARACIÓN DE DATOS SEGUROS (Evita errores de Foreign Key y NULLs)
            ' =========================================================
            ' Vendedor
            Dim idVend As Object = DBNull.Value
            If IsNumeric(TextBoxIdVendedor.Text) AndAlso Val(TextBoxIdVendedor.Text) > 0 Then
                idVend = Convert.ToInt32(TextBoxIdVendedor.Text)
            End If

            ' Presupuesto
            Dim idPresu As Object = DBNull.Value
            If Not String.IsNullOrWhiteSpace(TextBoxIdPresupuesto.Text) Then
                idPresu = TextBoxIdPresupuesto.Text.Trim()
            End If

            ' Forma de Pago
            Dim idFormaPago As Object = DBNull.Value
            If cboFormaPago.SelectedValue IsNot Nothing AndAlso cboFormaPago.SelectedIndex <> -1 Then
                idFormaPago = Convert.ToInt32(cboFormaPago.SelectedValue)
            End If

            ' Ruta
            Dim idRuta As Object = DBNull.Value
            If cboRuta.SelectedValue IsNot Nothing AndAlso cboRuta.SelectedIndex <> -1 Then
                idRuta = Convert.ToInt32(cboRuta.SelectedValue)
            End If

            ' =========================================================
            ' 4. GUARDAR CABECERA (INSERT o UPDATE)
            ' =========================================================
            Dim sqlCabecera As String = ""
            If esNuevo Then
                sqlCabecera = "INSERT INTO Pedidos (NumeroPedido, CodigoCliente, ID_Vendedor, Fecha, Estado, Observaciones, BaseImponible, Total, NumeroPresupuesto, FechaEntrega, ID_FormaPago, ID_Ruta) " &
                              "VALUES (@num, @cli, @vend, @fecha, @estado, @obs, @base, @totGen, @idPresu, @FEntrega, @formaPago, @ruta)"
            Else
                sqlCabecera = "UPDATE Pedidos SET CodigoCliente=@cli, ID_Vendedor=@vend, Fecha=@fecha, Estado=@estado, Observaciones=@obs, BaseImponible=@base, Total=@totGen, NumeroPresupuesto=@idPresu, FechaEntrega=@FEntrega, ID_FormaPago=@formaPago, ID_Ruta=@ruta WHERE NumeroPedido = @num"
            End If

            Using cmd As New SQLiteCommand(sqlCabecera, conexion)
                cmd.Transaction = transaccion
                cmd.Parameters.AddWithValue("@num", _numeroPedidoActual)
                cmd.Parameters.AddWithValue("@cli", TextBoxIdCliente.Text.Trim())
                cmd.Parameters.AddWithValue("@vend", idVend)

                ' Formateo seguro de fechas para SQLite
                Dim fecha As DateTime
                If DateTime.TryParse(TextBoxFecha.Text, fecha) Then
                    cmd.Parameters.AddWithValue("@fecha", fecha.ToString("yyyy-MM-dd HH:mm:ss"))
                Else
                    cmd.Parameters.AddWithValue("@fecha", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                End If

                cmd.Parameters.AddWithValue("@estado", TextBoxEstado.Text.Trim())
                cmd.Parameters.AddWithValue("@obs", TextBoxObservaciones.Text.Trim())
                cmd.Parameters.AddWithValue("@base", sumaBase)
                cmd.Parameters.AddWithValue("@totGen", sumaTotal)
                cmd.Parameters.AddWithValue("@idPresu", idPresu)
                cmd.Parameters.AddWithValue("@FEntrega", DateTimePickerFecha.Value.ToString("yyyy-MM-dd HH:mm:ss"))

                ' Parámetros Nuevos
                cmd.Parameters.AddWithValue("@formaPago", idFormaPago)
                cmd.Parameters.AddWithValue("@ruta", idRuta)

                cmd.ExecuteNonQuery()
            End Using

            ' =========================================================
            ' 5. GESTIÓN DE LÍNEAS (Borrar las eliminadas por el usuario)
            ' =========================================================
            For Each idDel In _idsParaBorrar
                Using cmdDel As New SQLiteCommand("DELETE FROM LineasPedido WHERE ID_Linea = @id", conexion)
                    cmdDel.Transaction = transaccion
                    cmdDel.Parameters.AddWithValue("@id", idDel)
                    cmdDel.ExecuteNonQuery()
                End Using
            Next
            _idsParaBorrar.Clear()

            ' =========================================================
            ' 6. GESTIÓN DE LÍNEAS (Insertar nuevas / Actualizar existentes)
            ' =========================================================
            If _dtLineas IsNot Nothing Then
                Dim orden As Integer = 1
                For Each row As DataRow In _dtLineas.Rows
                    If row.RowState = DataRowState.Deleted Then Continue For

                    Dim c As Decimal = 0 : Decimal.TryParse(row("Cantidad").ToString(), c)
                    Dim p As Decimal = 0 : Decimal.TryParse(row("PrecioUnitario").ToString(), p)
                    Dim d As Decimal = 0 : Decimal.TryParse(row("Descuento").ToString(), d)
                    Dim tot As Decimal = (c * p) * (1 - (d / 100))

                    Dim idLin = row("ID_Linea")
                    Dim idArt As Object = If(IsNumeric(row("ID_Articulo")) AndAlso Val(row("ID_Articulo")) > 0, row("ID_Articulo"), DBNull.Value)

                    If IsDBNull(idLin) OrElse Not IsNumeric(idLin) Then
                        ' ES UNA LÍNEA NUEVA (Hacemos INSERT)
                        Dim sqlIns As String = "INSERT INTO LineasPedido (NumeroPedido, NumeroOrden, ID_Articulo, Descripcion, Cantidad, PrecioUnitario, Descuento, Total) VALUES (@numP, @ord, @art, @desc, @cant, @prec, @dcto, @tot)"
                        Using cmdLin As New SQLiteCommand(sqlIns, conexion)
                            cmdLin.Transaction = transaccion
                            cmdLin.Parameters.AddWithValue("@numP", _numeroPedidoActual)
                            cmdLin.Parameters.AddWithValue("@ord", orden)
                            cmdLin.Parameters.AddWithValue("@art", idArt)
                            cmdLin.Parameters.AddWithValue("@desc", row("Descripcion"))
                            cmdLin.Parameters.AddWithValue("@cant", c)
                            cmdLin.Parameters.AddWithValue("@prec", p)
                            cmdLin.Parameters.AddWithValue("@dcto", d)
                            cmdLin.Parameters.AddWithValue("@tot", tot)
                            cmdLin.ExecuteNonQuery()
                        End Using
                    ElseIf row.RowState = DataRowState.Modified Then
                        ' ES UNA LÍNEA EXISTENTE (Hacemos UPDATE)
                        Dim sqlUpd As String = "UPDATE LineasPedido SET NumeroOrden=@ord, ID_Articulo=@art, Descripcion=@desc, Cantidad=@cant, PrecioUnitario=@prec, Descuento=@dcto, Total=@tot WHERE ID_Linea=@id"
                        Using cmdLin As New SQLiteCommand(sqlUpd, conexion)
                            cmdLin.Transaction = transaccion
                            cmdLin.Parameters.AddWithValue("@ord", orden)
                            cmdLin.Parameters.AddWithValue("@art", idArt)
                            cmdLin.Parameters.AddWithValue("@desc", row("Descripcion"))
                            cmdLin.Parameters.AddWithValue("@cant", c)
                            cmdLin.Parameters.AddWithValue("@prec", p)
                            cmdLin.Parameters.AddWithValue("@dcto", d)
                            cmdLin.Parameters.AddWithValue("@tot", tot)
                            cmdLin.Parameters.AddWithValue("@id", idLin)
                            cmdLin.ExecuteNonQuery()
                        End Using
                    End If
                    orden += 1
                Next
            End If

            ' =========================================================
            ' 7. CONFIRMAR Y RECARGAR EL FORMULARIO
            ' =========================================================
            transaccion.Commit()
            MessageBox.Show($"Pedido {_numeroPedidoActual} guardado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)

            CargarPedido(_numeroPedidoActual)

        Catch ex As Exception
            ' Si algo falla (como una validación de SQLite), deshacemos todo para no dejar datos a medias
            If transaccion IsNot Nothing Then transaccion.Rollback()
            MessageBox.Show("Error al guardar el pedido. Detalles técnicos:" & vbCrLf & ex.Message, "Error BD", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub ButtonNuevaLinea_Click(sender As Object, e As EventArgs) Handles ButtonNuevaLinea.Click
        ' 1. Seguridad: Si la tabla de memoria no existe, la inicializamos
        If _dtLineas Is Nothing Then
            _dtLineas = New DataTable()
        End If

        ' Aseguramos que tenga las columnas creadas (evita el error "Column does not belong to table")
        ConfigurarEstructuraDatos()

        ' 2. Creamos la nueva fila
        Dim nuevaFila As DataRow = _dtLineas.NewRow()

        ' 3. Calculamos el número de orden visual (para que salga 1, 2, 3...)
        Dim orden As Integer = 0
        For Each row As DataRow In _dtLineas.Rows
            If row.RowState <> DataRowState.Deleted Then orden += 1
        Next
        nuevaFila("NumeroOrden") = orden + 1

        ' 4. Valores por defecto (Vitales para que no falle la matemática luego)
        nuevaFila("ID_Pedido") = _numeroPedidoActual ' Si es 0 no pasa nada, se arregla al guardar
        nuevaFila("ID_Linea") = DBNull.Value     ' ESTO ES LO MÁS IMPORTANTE (Indica que es NUEVA)
        nuevaFila("ID_Articulo") = 0
        nuevaFila("Descripcion") = ""
        nuevaFila("Cantidad") = 1
        nuevaFila("PrecioUnitario") = 0
        nuevaFila("Descuento") = 0
        nuevaFila("Total") = 0

        ' 5. Añadimos a la tabla (se muestra en el Grid automáticamente)
        _dtLineas.Rows.Add(nuevaFila)

        ' 6. Truco de UX: Ponemos el foco en la celda del Artículo para escribir rápido
        If DataGridView1.Rows.Count > 0 Then
            Dim ultimaFilaIndex As Integer = DataGridView1.Rows.Count - 1
            ' Asegúrate de que la columna se llame "ID_Articulo" en tu grid
            DataGridView1.CurrentCell = DataGridView1.Rows(ultimaFilaIndex).Cells("ID_Articulo")
            DataGridView1.BeginEdit(True)
        End If
    End Sub





    Private Sub FrmPedidos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Estilos
        FrmPresupuestos.EstilizarGrid(DataGridView1)
        EstilizarFecha(DateTimePickerFecha)

        ' 1. Añadimos y cargamos los nuevos combos ANTES de cargar el pedido
        Me.Controls.Add(lblFormaPago) : Me.Controls.Add(cboFormaPago)
        Me.Controls.Add(lblRuta) : Me.Controls.Add(cboRuta)
        cboFormaPago.DropDownStyle = ComboBoxStyle.DropDownList
        cboRuta.DropDownStyle = ComboBoxStyle.DropDownList
        CargarDesplegables()

        ConfigurarGrid()
        ReorganizarControlesAutomaticamente()

        Dim ultimoNum As String = ObtenerUltimoNumeroPedido()
        If Not String.IsNullOrEmpty(ultimoNum) Then
            CargarPedido(ultimoNum)
        Else
            LimpiarFormulario()
        End If
    End Sub

    Private Sub CargarDesplegables()
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()

            ' Cargar Formas de Pago
            Dim daPago As New SQLiteDataAdapter("SELECT ID_FormaPago, Descripcion FROM FormasPago WHERE Activo=1", c)
            Dim dtPago As New DataTable()
            daPago.Fill(dtPago)
            cboFormaPago.DataSource = dtPago
            cboFormaPago.DisplayMember = "Descripcion"
            cboFormaPago.ValueMember = "ID_FormaPago"
            cboFormaPago.SelectedIndex = -1

            ' Cargar Rutas
            Dim daRuta As New SQLiteDataAdapter("SELECT ID_Ruta, NombreZona FROM Rutas WHERE Activo=1", c)
            Dim dtRuta As New DataTable()
            daRuta.Fill(dtRuta)
            cboRuta.DataSource = dtRuta
            cboRuta.DisplayMember = "NombreZona"
            cboRuta.ValueMember = "ID_Ruta"
            cboRuta.SelectedIndex = -1
        Catch ex As Exception
        End Try
    End Sub

    ' =========================================================
    ' 1. CONFIGURACIÓN (Estructura con NumeroPedido)
    ' =========================================================
    Private Sub ConfigurarEstructuraDatos()
        If _dtLineas Is Nothing Then _dtLineas = New DataTable()

        ' 1. Clave primaria de la línea (siempre necesaria)
        If Not _dtLineas.Columns.Contains("ID_Linea") Then _dtLineas.Columns.Add("ID_Linea", GetType(Object))

        ' 2. CAMBIO CLAVE: Usamos NumeroPedido (String) en lugar de ID_Pedido
        If Not _dtLineas.Columns.Contains("NumeroPedido") Then _dtLineas.Columns.Add("NumeroPedido", GetType(String))

        ' 3. Resto de columnas
        If Not _dtLineas.Columns.Contains("NumeroOrden") Then _dtLineas.Columns.Add("NumeroOrden", GetType(Integer))
        If Not _dtLineas.Columns.Contains("ID_Articulo") Then _dtLineas.Columns.Add("ID_Articulo", GetType(Object))
        If Not _dtLineas.Columns.Contains("Descripcion") Then _dtLineas.Columns.Add("Descripcion", GetType(String))

        ' Decimales
        If Not _dtLineas.Columns.Contains("Cantidad") Then _dtLineas.Columns.Add("Cantidad", GetType(Decimal))
        If Not _dtLineas.Columns.Contains("PrecioUnitario") Then _dtLineas.Columns.Add("PrecioUnitario", GetType(Decimal))
        If Not _dtLineas.Columns.Contains("Descuento") Then _dtLineas.Columns.Add("Descuento", GetType(Decimal))
        If Not _dtLineas.Columns.Contains("Total") Then _dtLineas.Columns.Add("Total", GetType(Decimal))
    End Sub

    Private Sub ConfigurarGrid()
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.Columns.Clear()

        ' FUNDAMENTAL: Desactivar auto-size global
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "ID_Linea", .DataPropertyName = "ID_Linea", .Visible = False})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "NumeroOrden", .DataPropertyName = "NumeroOrden", .HeaderText = "Nº", .Width = 40, .ReadOnly = True, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleCenter}})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "ID_Articulo", .DataPropertyName = "ID_Articulo", .HeaderText = "ID Art", .Width = 70})

        ' ¡AQUÍ ESTÁ LA MAGIA QUE ARREGLA EL HUECO BLANCO!
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {
            .Name = "Descripcion",
            .DataPropertyName = "Descripcion",
            .HeaderText = "Descripción",
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill ' <-- ESTO ES CLAVE
        })

        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "Cantidad", .DataPropertyName = "Cantidad", .HeaderText = "Cant.", .Width = 60, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "N2"}})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "PrecioUnitario", .DataPropertyName = "PrecioUnitario", .HeaderText = "Precio", .Width = 80, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "C2"}})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "Descuento", .DataPropertyName = "Descuento", .HeaderText = "% Dto", .Width = 60, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "N2"}})

        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {
            .Name = "Total",
            .DataPropertyName = "Total",
            .HeaderText = "Total",
            .ReadOnly = True,
            .Width = 90,
            .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "C2", .BackColor = Color.WhiteSmoke}
        })
    End Sub

    ' =========================================================
    ' 2. CARGA (Por String)
    ' =========================================================
    Private Sub CargarPedido(numeroPedido As String)
        Try
            Dim conexion = ConexionBD.GetConnection()
            If conexion.State <> ConnectionState.Open Then conexion.Open()

            Dim sql As String = "SELECT P.*, C.NombreFiscal AS NombreCliente, V.Nombre AS NombreVendedor " &
                                "FROM Pedidos P " &
                                "LEFT JOIN Clientes C ON P.CodigoCliente = C.CodigoCliente " &
                                "LEFT JOIN Vendedores V ON P.ID_Vendedor = V.ID_Vendedor " &
                                "WHERE P.NumeroPedido = @num"

            Using cmd As New SQLiteCommand(sql, conexion)
                cmd.Parameters.AddWithValue("@num", numeroPedido)
                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        _numeroPedidoActual = numeroPedido

                        ' 1. Cargar Textos Simples (Sin procesar nada raro)
                        TextBoxPedido.Text = reader("NumeroPedido").ToString()

                        ' 2. Fechas
                        If Not IsDBNull(reader("Fecha")) Then
                            TextBoxFecha.Text = Convert.ToDateTime(reader("Fecha")).ToShortDateString()
                        End If

                        ' 3. Estados y Observaciones
                        TextBoxObservaciones.Text = If(IsDBNull(reader("Observaciones")), "", reader("Observaciones").ToString())
                        TextBoxEstado.Text = If(IsDBNull(reader("Estado")), "Pendiente", reader("Estado").ToString())

                        ' 4. IDs y Nombres
                        TextBoxIdCliente.Text = reader("CodigoCliente").ToString()
                        TextBoxCliente.Text = If(IsDBNull(reader("NombreCliente")), "", reader("NombreCliente").ToString())
                        TextBoxCliente.Tag = reader("CodigoCliente")

                        TextBoxIdVendedor.Text = If(IsDBNull(reader("ID_Vendedor")), "", reader("ID_Vendedor").ToString())
                        TextBoxVendedor.Text = If(IsDBNull(reader("NombreVendedor")), "", reader("NombreVendedor").ToString())

                        ' 5. Totales (Si los tienes en pantalla)
                        If IsNumeric(reader("BaseImponible")) Then TextBoxBase.Text = Convert.ToDecimal(reader("BaseImponible")).ToString("C2")
                        If IsNumeric(reader("Total")) Then TextBoxTotalPed.Text = Convert.ToDecimal(reader("Total")).ToString("C2")

                        ' 6. Trazabilidad (Presupuesto Origen)
                        ' Verifica que el campo en BD sea ID_Presupuesto
                        If Not IsDBNull(reader("NumeroPresupuesto")) Then
                            TextBoxIdPresupuesto.Text = reader("NumeroPresupuesto").ToString()
                        End If
                        DateTimePickerFecha.Value = reader("FechaEntrega")
                    Else
                        MessageBox.Show("Pedido no encontrado.")
                        Return
                    End If
                End Using
            End Using

            ' Cargar las líneas después de la cabecera
            CargarLineas()

        Catch ex As Exception
            MessageBox.Show("Error al cargar pedido: " & ex.Message)
        End Try
    End Sub

    Private Sub CargarLineas()
        Try
            ' Buscamos por NumeroPedido (Texto)
            Dim sql As String = "SELECT ID_Linea, NumeroPedido, NumeroOrden, ID_Articulo, Descripcion, Cantidad, PrecioUnitario, Descuento, Total " &
                                "FROM LineasPedido WHERE NumeroPedido = @num ORDER BY NumeroOrden ASC"

            Dim conexion = ConexionBD.GetConnection()
            Using cmd As New SQLiteCommand(sql, conexion)
                cmd.Parameters.AddWithValue("@num", _numeroPedidoActual)
                Dim da As New SQLiteDataAdapter(cmd)
                _dtLineas = New DataTable()
                da.Fill(_dtLineas)

                ConfigurarEstructuraDatos()
                DataGridView1.DataSource = _dtLineas
            End Using

            CalcularTotalesGenerales()
            If DataGridView1.Columns.Contains("ID_Linea") Then DataGridView1.Columns("ID_Linea").Visible = False

        Catch ex As Exception
            MessageBox.Show("Error al cargar líneas: " & ex.Message)
        End Try
    End Sub



    ' =========================================================
    ' 4. ACCIONES DE BOTONES
    ' =========================================================

    Private Sub btnAnadirLin_Click(sender As Object, e As EventArgs) Handles ButtonNuevaLinea.Click
        If _dtLineas Is Nothing Then _dtLineas = New DataTable()
        ConfigurarEstructuraDatos()

        Dim nuevaFila As DataRow = _dtLineas.NewRow()
        Dim orden As Integer = 0
        For Each row As DataRow In _dtLineas.Rows
            If row.RowState <> DataRowState.Deleted Then orden += 1
        Next

        nuevaFila("NumeroOrden") = orden + 1
        nuevaFila("ID_Linea") = DBNull.Value
        nuevaFila("NumeroPedido") = _numeroPedidoActual
        nuevaFila("ID_Articulo") = 0
        nuevaFila("Descripcion") = ""
        nuevaFila("Cantidad") = 1
        nuevaFila("PrecioUnitario") = 0
        nuevaFila("Descuento") = 0
        nuevaFila("Total") = 0
        _dtLineas.Rows.Add(nuevaFila)

        If DataGridView1.Rows.Count Then
            DataGridView1.CurrentCell = DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells("ID_Articulo")
            DataGridView1.BeginEdit(True)
        End If

    End Sub

    Private Sub ButtonBorrarLineas_Click(sender As Object, e As EventArgs) Handles ButtonBorrarLineas.Click
        If DataGridView1.SelectedRows.Count = 0 Then Return
        If MessageBox.Show("¿Borrar?", "Confirmar", MessageBoxButtons.YesNo) = DialogResult.No Then Return

        Dim filas As New List(Of DataGridViewRow)
        For Each f As DataGridViewRow In DataGridView1.SelectedRows
            If Not f.IsNewRow Then filas.Add(f)
        Next

        For Each f In filas
            Dim idVal = f.Cells("ID_Linea").Value
            If IsNumeric(idVal) AndAlso Val(idVal) > 0 Then _idsParaBorrar.Add(CInt(idVal))
            DataGridView1.Rows.Remove(f)
        Next
        CalcularTotalesGenerales()
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles ButtonNuevoPresup.Click
        LimpiarFormulario()
    End Sub

    Private Sub LimpiarFormulario()
        _numeroPedidoActual = ""
        _idsParaBorrar.Clear()

        TextBoxPedido.Text = GenerarProximoNumeroPedido()
        TextBoxCliente.Text = "" : TextBoxIdCliente.Text = "" : TextBoxCliente.Tag = Nothing
        TextBoxVendedor.Text = "" : TextBoxIdVendedor.Text = "" : TextBoxVendedor.Tag = Nothing
        TextBoxObservaciones.Text = ""
        TextBoxFecha.Text = DateTime.Now.ToShortDateString()
        TextBoxEstado.Text = "Pendiente"
        TextBoxBase.Text = "0,00 €" : TextBoxIva.Text = "0,00 €" : TextBoxTotalPed.Text = "0,00 €" ' TextBoxTotalPed
        DateTimePickerFecha.Value = DateTime.Now.ToShortDateString()
        TextBoxIdPresupuesto.Text = ""
        ConfigurarGrid()
        _dtLineas = New DataTable()
        ConfigurarEstructuraDatos()
        DataGridView1.DataSource = _dtLineas
        ' Blindaje: Solo los limpia si ya han sido creados en memoria
        If cboFormaPago IsNot Nothing Then cboFormaPago.SelectedIndex = -1
        If cboRuta IsNot Nothing Then cboRuta.SelectedIndex = -1

        TextBoxIdCliente.Focus()
        If DataGridView1.Columns.Contains("ID_Linea") Then DataGridView1.Columns("ID_Linea").Visible = False
        TextBoxIdCliente.Focus()
    End Sub

    Private Sub btnBorrar_Click(sender As Object, e As EventArgs) Handles ButtonBorrar.Click
        If String.IsNullOrEmpty(_numeroPedidoActual) Then Return
        If MessageBox.Show("¿Eliminar pedido?", "Confirmar", MessageBoxButtons.YesNo) = DialogResult.No Then Return

        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()
            ' Borrado por TEXTO
            Dim cmd As New SQLiteCommand("PRAGMA foreign_keys = OFF; DELETE FROM LineasPedido WHERE NumeroPedido=@num; DELETE FROM Pedidos WHERE NumeroPedido=@num; PRAGMA foreign_keys = ON;", c)
            cmd.Parameters.AddWithValue("@num", _numeroPedidoActual)
            cmd.ExecuteNonQuery()

            MessageBox.Show("Eliminado.")
            LimpiarFormulario()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    ' =========================================================
    ' 5. IMPORTAR PRESUPUESTO (Adaptado a Texto)
    ' =========================================================
    Private Sub btnBuscarPresupuesto_Click(sender As Object, e As EventArgs) Handles btnBuscarPresupuesto.Click
        Using frm As New FrmBuscador()
            frm.TablaABuscar = "Presupuestos"
            If frm.ShowDialog() = DialogResult.OK Then
                Dim codPresu As String = frm.Resultado ' Ahora recibimos String
                If Not String.IsNullOrEmpty(codPresu) Then ImportarDatosPresupuesto(codPresu)
            End If
        End Using
    End Sub

    Private Sub ImportarDatosPresupuesto(codigoPresupuesto As String)
        Dim c = ConexionBD.GetConnection()
        Try
            If c.State <> ConnectionState.Open Then c.Open()

            ' 1. Cabecera (Buscamos por NumeroPresupuesto)
            Dim sqlCab As String = "SELECT P.*, C.NombreFiscal, V.Nombre AS NombreVend FROM Presupuestos P " &
                                   "LEFT JOIN Clientes C ON P.CodigoCliente=C.CodigoCliente " &
                                   "LEFT JOIN Vendedores V ON P.ID_Vendedor=V.ID_Vendedor " &
                                   "WHERE P.NumeroPresupuesto = @num"
            Using cmd As New SQLiteCommand(sqlCab, c)
                cmd.Parameters.AddWithValue("@num", codigoPresupuesto)
                Using r = cmd.ExecuteReader()
                    If r.Read() Then
                        LimpiarFormulario()
                        TextBoxIdCliente.Text = r("CodigoCliente").ToString()
                        TextBoxCliente.Text = r("NombreFiscal").ToString()
                        TextBoxIdVendedor.Text = If(IsDBNull(r("ID_Vendedor")), "", r("ID_Vendedor").ToString())
                        TextBoxVendedor.Text = If(IsDBNull(r("NombreVend")), "", r("NombreVend").ToString())

                        TextBoxObservaciones.Text = $"Desde Presupuesto {codigoPresupuesto}. "
                        If TextBoxIdPresupuesto IsNot Nothing Then
                            TextBoxIdPresupuesto.Text = codigoPresupuesto
                            TextBoxIdPresupuesto.Tag = codigoPresupuesto
                        End If
                    End If
                End Using
            End Using

            ' 2. Líneas (Buscamos por NumeroPresupuesto)
            Dim sqlLin As String = "SELECT * FROM LineasPresupuesto WHERE NumeroPresupuesto = @num ORDER BY NumeroOrden ASC"
            Dim dtOrigen As New DataTable()
            Using cmd As New SQLiteCommand(sqlLin, c)
                cmd.Parameters.AddWithValue("@num", codigoPresupuesto)
                Dim da As New SQLiteDataAdapter(cmd)
                da.Fill(dtOrigen)
            End Using

            _dtLineas.Rows.Clear()
            For Each rowOrig As DataRow In dtOrigen.Rows
                Dim rowNew As DataRow = _dtLineas.NewRow()

                ' --- CORRECCIÓN DEL ERROR ---
                ' Ya no existe ID_Pedido, ahora usamos NumeroPedido
                rowNew("NumeroPedido") = _numeroPedidoActual ' Asignamos el código actual (ej: PED-26-001)
                ' ----------------------------

                rowNew("ID_Linea") = DBNull.Value
                rowNew("NumeroOrden") = rowOrig("NumeroOrden")
                rowNew("ID_Articulo") = rowOrig("ID_Articulo")
                rowNew("Descripcion") = rowOrig("Descripcion")

                Dim ca As Decimal = 0 : Decimal.TryParse(rowOrig("Cantidad").ToString(), ca)
                Dim pr As Decimal = 0 : Decimal.TryParse(rowOrig("PrecioUnitario").ToString(), pr)
                Dim dt As Decimal = 0 : Decimal.TryParse(rowOrig("Descuento").ToString(), dt)

                rowNew("Cantidad") = ca : rowNew("PrecioUnitario") = pr : rowNew("Descuento") = dt
                rowNew("Total") = (ca * pr) * (1 - (dt / 100))

                _dtLineas.Rows.Add(rowNew)
            Next
            CalcularTotalesGenerales()
            MessageBox.Show("Presupuesto importado.")

        Catch ex As Exception
            MessageBox.Show("Error al importar: " & ex.Message)
        End Try
    End Sub

    ' ... (Mantén aquí tus funciones auxiliares: EstilizarFecha, CalcularTotalesGenerales, DataGridView events, etc.) ...
    ' Asegúrate de que CalcularTotalesGenerales itera sobre _dtLineas correctamente.

    Private Sub EstilizarFecha(dtp As DateTimePicker)
        dtp.Format = DateTimePickerFormat.Custom : dtp.CustomFormat = "dd/MM/yyyy"
        dtp.Font = New Font("Segoe UI", 10) : dtp.MinimumSize = New Size(0, 25)
    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim fila = DataGridView1.Rows(e.RowIndex)
        Dim colName = DataGridView1.Columns(e.ColumnIndex).Name
        If colName = "ID_Articulo" Then
            Dim idArt As String = fila.Cells("ID_Articulo").Value?.ToString()
            If String.IsNullOrWhiteSpace(idArt) Then
                fila.Cells("Descripcion").Value = "" : fila.Cells("PrecioUnitario").Value = 0 : fila.Tag = Nothing : Return
            End If
            Try
                Dim c = ConexionBD.GetConnection()
                Dim cmd As New SQLiteCommand("SELECT Descripcion, PrecioVenta, StockActual FROM Articulos WHERE ID_Articulo = @id", c)
                cmd.Parameters.AddWithValue("@id", idArt)
                If c.State <> ConnectionState.Open Then c.Open()
                Using r = cmd.ExecuteReader()
                    If r.Read() Then
                        fila.Cells("Descripcion").Value = r("Descripcion").ToString()
                        fila.Cells("PrecioUnitario").Value = Convert.ToDecimal(r("PrecioVenta"))
                        If Not IsDBNull(r("StockActual")) Then fila.Tag = Convert.ToDecimal(r("StockActual"))
                        If IsDBNull(fila.Cells("Cantidad").Value) OrElse Val(fila.Cells("Cantidad").Value) = 0 Then fila.Cells("Cantidad").Value = 1
                    Else
                        MessageBox.Show("Artículo no encontrado")
                    End If
                End Using
            Catch
            End Try
        End If
        If colName = "Cantidad" Or colName = "PrecioUnitario" Or colName = "Descuento" Then
            Dim c, p, d As Decimal
            Decimal.TryParse(fila.Cells("Cantidad").Value?.ToString(), c)
            Decimal.TryParse(fila.Cells("PrecioUnitario").Value?.ToString(), p)
            Decimal.TryParse(fila.Cells("Descuento").Value?.ToString(), d)
            fila.Cells("Total").Value = (c * p) * (1 - (d / 100))
            CalcularTotalesGenerales()
        End If
    End Sub

    Private Sub CalcularTotalesGenerales()
        Dim base As Decimal = 0
        If _dtLineas IsNot Nothing Then
            For Each row As DataRow In _dtLineas.Rows
                If row.RowState <> DataRowState.Deleted Then
                    Dim t As Decimal = 0
                    If _dtLineas.Columns.Contains("Total") Then Decimal.TryParse(row("Total").ToString(), t)
                    base += t
                End If
            Next
        End If
        TextBoxBase.Text = base.ToString("C2")
        TextBoxIva.Text = (base * 0.21D).ToString("C2")
        TextBoxTotalPed.Text = (base * 1.21D).ToString("C2")
    End Sub

    Private Function ObtenerUltimoNumeroPedido() As String
        Try
            Dim c = ConexionBD.GetConnection()
            Dim cmd As New SQLiteCommand("SELECT MAX(NumeroPedido) FROM Pedidos", c)
            Dim r = cmd.ExecuteScalar()
            Return If(IsDBNull(r), "", r.ToString())
        Catch
            Return ""
        End Try
    End Function

    Private Function GenerarProximoNumeroPedido() As String
        ' Formato deseado: PED-XXX
        Dim prefijo As String = "PED-"
        Dim nuevoNumero As String = "PED-001"

        Try
            Dim conexion = ConexionBD.GetConnection()
            ' Buscamos el último pedido. Ordenamos por ID para obtener el último creado realmente.
            Dim sql As String = "SELECT NumeroPedido FROM Pedidos WHERE NumeroPedido LIKE @patron ORDER BY NumeroPedido DESC LIMIT 1"

            Using cmd As New SQLiteCommand(sql, conexion)
                cmd.Parameters.AddWithValue("@patron", prefijo & "%")

                If conexion.State <> ConnectionState.Open Then conexion.Open()
                Dim resultado = cmd.ExecuteScalar()

                If resultado IsNot Nothing AndAlso Not IsDBNull(resultado) Then
                    Dim ultimoCodigo As String = resultado.ToString()
                    ' Ejemplo 1: PED-005  -> partes(0)="PED", partes(1)="005"
                    ' Ejemplo 2: PED-26-005 -> partes(0)="PED", partes(1)="26", partes(2)="005"

                    Dim partes As String() = ultimoCodigo.Split("-"c)

                    ' TRUCO: Siempre cogemos la ÚLTIMA parte (partes.Length - 1)
                    If partes.Length >= 2 Then
                        Dim ultimoSegmento As String = partes(partes.Length - 1)
                        Dim correlativo As Integer = 0

                        If Integer.TryParse(ultimoSegmento, correlativo) Then
                            correlativo += 1
                            nuevoNumero = $"{prefijo}{correlativo.ToString("D3")}"
                        End If
                    End If
                End If
            End Using
        Catch ex As Exception
            ' Si falla, generamos uno aleatorio para no bloquear
            nuevoNumero = "PED-" & DateTime.Now.Ticks.ToString().Substring(12)
        End Try

        Return nuevoNumero
    End Function

    Private Sub ButtonAnterior_Click(sender As Object, e As EventArgs) Handles ButtonAnterior.Click
        ' Lógica de navegación String (Ver Presupuestos)
        Try
            Dim conexion = ConexionBD.GetConnection()
            Dim numDestino As String = ""

            ' CAMBIO 7: Navegación alfanumérica (Funciona bien con formatos fijos tipo PRE-001)
            Dim sql As String = "SELECT MAX(NumeroPedido) FROM Pedidos WHERE NumeroPedido < @actual"
            Using cmd As New SQLiteCommand(sql, conexion)
                ' Si es nuevo, buscamos el último absoluto
                If String.IsNullOrEmpty(_numeroPedidoActual) Then
                    cmd.CommandText = "SELECT MAX(NumeroPedido) FROM Pedidos"
                Else
                    cmd.Parameters.AddWithValue("@actual", _numeroPedidoActual)
                End If

                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    numDestino = result.ToString()
                    CargarPedido(numDestino)
                Else
                    MessageBox.Show("Primer Pedido.")
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub ButtonSiguiente_Click(sender As Object, e As EventArgs) Handles ButtonSiguiente.Click
        ' Lógica de navegación String (Ver Presupuestos)
        If String.IsNullOrEmpty(_numeroPedidoActual) Then Return
        Try
            Dim conexion = ConexionBD.GetConnection()
            Dim sql As String = "SELECT MIN(NumeroPedido) FROM Pedidos WHERE NumeroPedido > @actual"
            Using cmd As New SQLiteCommand(sql, conexion)
                cmd.Parameters.AddWithValue("@actual", _numeroPedidoActual)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    CargarPedido(result.ToString())
                Else
                    MessageBox.Show("Último pedido.")
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub TextBoxIdCliente_Leave(sender As Object, e As EventArgs) Handles TextBoxIdCliente.Leave
        If String.IsNullOrWhiteSpace(TextBoxIdCliente.Text) Then TextBoxCliente.Text = "" : Return
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()
            ' Traemos también el ID_Ruta del cliente
            Dim cmd As New SQLiteCommand("SELECT NombreFiscal, ID_Ruta FROM Clientes WHERE CodigoCliente=@id", c)
            cmd.Parameters.AddWithValue("@id", TextBoxIdCliente.Text)

            Using r = cmd.ExecuteReader()
                If r.Read() Then
                    TextBoxCliente.Text = r("NombreFiscal").ToString()
                    ' Auto-seleccionar Ruta
                    Dim idRut = r("ID_Ruta")
                    If Not IsDBNull(idRut) Then cboRuta.SelectedValue = Convert.ToInt32(idRut)
                Else
                    TextBoxCliente.Text = "NO EXISTE"
                End If
            End Using
        Catch
        End Try
    End Sub
    Private Sub TextBoxIdVendedor_Leave(sender As Object, e As EventArgs) Handles TextBoxIdVendedor.Leave
        ' Tu lógica de buscar vendedor
        If String.IsNullOrWhiteSpace(TextBoxIdVendedor.Text) Then TextBoxVendedor.Text = "" : Return
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()
            Dim cmd As New SQLiteCommand("SELECT Nombre FROM Vendedores WHERE ID_Vendedor=@id", c)
            cmd.Parameters.AddWithValue("@id", Val(TextBoxIdVendedor.Text))
            Dim r = cmd.ExecuteScalar()
            TextBoxVendedor.Text = If(r IsNot Nothing, r.ToString(), "NO EXISTE")
        Catch
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub


#Region "Auto-Organización Visual (Pixel-Perfect) - PEDIDOS"
    ''' <summary>
    ''' Calcula matemáticamente la posición y tamaño de todos los controles del formulario de PEDIDOS.
    ''' </summary>
    Private Sub ReorganizarControlesAutomaticamente()
        ' 1. QUITAR ANCHORS
        For Each ctrl As Control In Me.Controls
            ctrl.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        Next

        ' 2. LA REJILLA MAESTRA (Todo agrupado a la izquierda)
        Dim margenIzq As Integer = 30
        Dim anchoForm As Integer = Me.ClientSize.Width
        Dim altoForm As Integer = Me.ClientSize.Height

        Dim yFila1 As Integer = 30
        Dim yFila2 As Integer = 55
        Dim yFila3 As Integer = 95
        Dim yFila4 As Integer = 120
        Dim yFila5 As Integer = 160 ' Nueva fila para etiquetas
        Dim yFila6 As Integer = 185 ' Nueva fila para controles

        lblFormaPago.Location = New Point(col1_X, yFila5)
        cboFormaPago.Bounds = New Rectangle(col1_X, yFila6, 200, 25)
        cboFormaPago.Font = New Font("Segoe UI", 10.5F)

        lblRuta.Location = New Point(col2_X + 25, yFila5)
        cboRuta.Bounds = New Rectangle(col2_X + 25, yFila6, 300, 25)
        cboRuta.Font = New Font("Segoe UI", 10.5F)
        ' ¡EL TRUCO! Coordenadas fijas
        Dim col1_X As Integer = margenIzq
        Dim col2_X As Integer = 190
        Dim col3_X As Integer = 750
        Dim col4_X As Integer = 920

        ' --- ETIQUETAS DINÁMICAS ---
        For Each ctrl As Control In Me.Controls
            If TypeOf ctrl Is Label AndAlso ctrl.Name <> "LineaTotales" Then
                ctrl.BringToFront()
                Dim texto As String = ctrl.Text.Trim().ToLower()
                Select Case texto
                    Case "pedido" : ctrl.Location = New Point(col1_X, yFila1)
                    Case "cliente" : ctrl.Location = New Point(col2_X, yFila1)
                    Case "fecha" : ctrl.Location = New Point(col3_X, yFila1)
                    Case "fecha entrega" : ctrl.Location = New Point(col4_X, yFila1)
                    Case "vendedor" : ctrl.Location = New Point(col1_X, yFila3)
                    Case "observaciones" : ctrl.Location = New Point(col2_X, yFila3)
                    Case "estado" : ctrl.Location = New Point(col3_X, yFila3)
                    Case "presupuesto" : ctrl.Location = New Point(col4_X, yFila3)
                End Select
            End If
        Next

        ' --- CAJAS DE TEXTO (Tamaños y posiciones fijos) ---
        ' Fila 1
        ' Encogemos la caja de pedido a 105px para que la lupa (en la posición 110) respire
        TextBoxPedido.Bounds = New Rectangle(col1_X, yFila2, 105, 25)
        btnBuscarPedido.Bounds = New Rectangle(col1_X + 110, yFila2, 30, 25)

        TextBoxIdCliente.Bounds = New Rectangle(col2_X, yFila2, 60, 25)
        TextBoxCliente.Bounds = New Rectangle(col2_X + 100, yFila2, 430, 25)

        TextBoxFecha.Bounds = New Rectangle(col3_X, yFila2, 140, 25)
        DateTimePickerFecha.Bounds = New Rectangle(col4_X, yFila2, 140, 25)

        ' Fila 2
        TextBoxIdVendedor.Bounds = New Rectangle(col1_X, yFila4, 50, 25)
        TextBoxVendedor.Bounds = New Rectangle(col1_X + 55, yFila4, 85, 25)

        TextBoxObservaciones.Bounds = New Rectangle(col2_X, yFila4, 530, 25)
        TextBoxEstado.Bounds = New Rectangle(col3_X, yFila4, 140, 25)

        TextBoxIdPresupuesto.Bounds = New Rectangle(col4_X, yFila4, 105, 25)
        btnBuscarPresupuesto.Bounds = New Rectangle(col4_X + 110, yFila4, 30, 25)

        ' =========================================================
        ' 1. SEPARADOR VISUAL CABECERA / DETALLE (Mejorado)
        ' =========================================================
        Dim lineaDivisoria As Label = Me.Controls.OfType(Of Label)().FirstOrDefault(Function(l) l.Name = "LineaDivisoria")
        If lineaDivisoria Is Nothing Then
            lineaDivisoria = New Label() With {.Name = "LineaDivisoria", .BackColor = Color.FromArgb(120, 130, 140), .Height = 2}
            Me.Controls.Add(lineaDivisoria)
        End If

        Dim yTabla As Integer = 230
        lineaDivisoria.Bounds = New Rectangle(margenIzq, yTabla - 20, anchoForm - (margenIzq * 2), 2)
        lineaDivisoria.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        lineaDivisoria.BringToFront()

        ' =========================================================
        ' 2. LA TABLA (Desplazada hacia abajo)
        ' =========================================================
        Dim altoTabla As Integer = altoForm - yTabla - 140
        DataGridView1.Bounds = New Rectangle(margenIzq, yTabla, anchoForm - (margenIzq * 2), altoTabla)
        DataGridView1.BackgroundColor = Me.BackColor
        DataGridView1.BorderStyle = BorderStyle.None

        ' =========================================================
        ' 3. TOTALES (Con "Caja" de resumen)
        ' =========================================================
        Dim xDerecha As Integer = DataGridView1.Right
        Dim yTotales As Integer = DataGridView1.Bottom + 10

        Dim panelTotales As Label = Me.Controls.OfType(Of Label)().FirstOrDefault(Function(l) l.Name = "PanelTotalesResumen")
        If panelTotales Is Nothing Then
            panelTotales = New Label() With {.Name = "PanelTotalesResumen", .BackColor = Color.FromArgb(25, 30, 40)}
            Me.Controls.Add(panelTotales)
            panelTotales.SendToBack()
        End If
        panelTotales.Bounds = New Rectangle(xDerecha - 320, yTotales, 320, 115)
        panelTotales.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right

        TextBoxBase.Bounds = New Rectangle(xDerecha - 140, yTotales + 10, 120, 25)
        TextBoxIva.Bounds = New Rectangle(xDerecha - 140, yTotales + 40, 120, 25)
        LabelBase.Bounds = New Rectangle(xDerecha - 300, yTotales + 10, 150, 25)
        LabelIva.Bounds = New Rectangle(xDerecha - 300, yTotales + 40, 150, 25)

        LabelBase.BackColor = Color.FromArgb(25, 30, 40) : LabelIva.BackColor = Color.FromArgb(25, 30, 40)
        LabelBase.TextAlign = ContentAlignment.MiddleRight : LabelIva.TextAlign = ContentAlignment.MiddleRight
        TextBoxBase.TextAlign = HorizontalAlignment.Right : TextBoxIva.TextAlign = HorizontalAlignment.Right

        TextBoxBase.BackColor = Color.FromArgb(25, 30, 40) : TextBoxBase.ForeColor = Color.White : TextBoxBase.BorderStyle = BorderStyle.None
        TextBoxIva.BackColor = Color.FromArgb(25, 30, 40) : TextBoxIva.ForeColor = Color.White : TextBoxIva.BorderStyle = BorderStyle.None

        Dim lineaTotal As Label = Me.Controls.OfType(Of Label)().FirstOrDefault(Function(l) l.Name = "LineaTotales")
        If lineaTotal Is Nothing Then
            lineaTotal = New Label() With {.Name = "LineaTotales", .BackColor = Color.FromArgb(100, 100, 100), .Height = 1}
            Me.Controls.Add(lineaTotal)
        End If
        lineaTotal.Bounds = New Rectangle(xDerecha - 300, yTotales + 70, 280, 1)
        lineaTotal.BringToFront()

        Dim colorAcento As Color = Color.FromArgb(0, 150, 255)
        Label7.Bounds = New Rectangle(xDerecha - 300, yTotales + 80, 150, 30)
        Label7.BackColor = Color.FromArgb(25, 30, 40) : Label7.TextAlign = ContentAlignment.MiddleRight
        Label7.Font = New Font("Segoe UI", 13, FontStyle.Bold) : Label7.ForeColor = colorAcento

        ' Usamos TextBoxTotalPed para la pantalla de pedidos
        TextBoxTotalPed.Bounds = New Rectangle(xDerecha - 140, yTotales + 80, 120, 30)
        TextBoxTotalPed.TextAlign = HorizontalAlignment.Right
        TextBoxTotalPed.Font = New Font("Segoe UI", 14, FontStyle.Bold) : TextBoxTotalPed.ForeColor = colorAcento
        TextBoxTotalPed.BackColor = Color.FromArgb(25, 30, 40) : TextBoxTotalPed.BorderStyle = BorderStyle.None

        ' =========================================================
        ' 4. BARRA DE HERRAMIENTAS INFERIOR (Organizada)
        ' =========================================================
        Dim yBotones As Integer = DataGridView1.Bottom + 45

        EstilizarBoton(ButtonGuardar, margenIzq, yBotones, Color.FromArgb(0, 120, 215), Color.White)
        EstilizarBoton(ButtonBorrar, margenIzq + 110, yBotones, Color.FromArgb(209, 52, 56), Color.White)
        EstilizarBoton(ButtonNuevoPresup, margenIzq + 220, yBotones, Color.FromArgb(0, 120, 215), Color.White)

        EstilizarBoton(ButtonBorrarLineas, margenIzq + 380, yBotones, Color.FromArgb(85, 85, 85), Color.White)
        ButtonBorrarLineas.Text = "- Quitar Línea"
        ButtonBorrarLineas.Width = 110

        EstilizarBoton(ButtonNuevaLinea, margenIzq + 500, yBotones, Color.FromArgb(40, 140, 90), Color.White)
        ButtonNuevaLinea.Text = "+ Añadir Línea"
        ButtonNuevaLinea.Width = 110

        EstilizarBoton(ButtonAnterior, xDerecha - 560, yBotones, Me.BackColor, Color.White)
        EstilizarBoton(ButtonSiguiente, xDerecha - 450, yBotones, Me.BackColor, Color.White)

        LabelStock.Location = New Point(margenIzq, DataGridView1.Bottom + 10)

        ' --- ANCHORS INTELIGENTES (Corregidos) ---
        DataGridView1.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right

        ' 1. Botones de la Izquierda (Se quedan clavados a la izquierda)
        Dim botonesAbajo As Control() = {ButtonGuardar, ButtonBorrar, ButtonNuevoPresup, ButtonBorrarLineas, ButtonNuevaLinea, LabelStock}
        For Each b In botonesAbajo : b.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left : Next

        ' 2. Botones de Navegación (Se anclan a la DERECHA para que no los pise la caja de totales)
        ButtonAnterior.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        ButtonSiguiente.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right

        ' 3. Tarjeta de Totales (A la derecha)
        TextBoxBase.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        TextBoxIva.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        TextBoxTotalPed.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        LabelBase.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        LabelIva.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        Label7.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        lineaTotal.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
    End Sub

    ''' <summary>
    ''' Configura los anclajes para que el formulario se adapte al maximizar.
    ''' </summary>
    Private Sub ConfigurarDiseñoResponsive()
        DataGridView1.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right

        ' Cabecera (Izquierda y central elásticas)
        TextBoxCliente.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        TextBoxObservaciones.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

        ' Cabecera derecha (Anclada a la derecha para no pisarse)
        TextBoxFecha.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        DateTimePickerFecha.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        TextBoxEstado.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        TextBoxIdPresupuesto.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        btnBuscarPresupuesto.Anchor = AnchorStyles.Top Or AnchorStyles.Right

        ' Auto-anclar las etiquetas correctamente (usando el mismo truco)
        For Each ctrl As Control In Me.Controls
            If TypeOf ctrl Is Label Then
                Dim texto = ctrl.Text.Trim().ToLower()
                If texto = "fecha" Or texto = "fecha entrega" Or texto = "estado" Or texto = "presupuesto" Then
                    ctrl.Anchor = AnchorStyles.Top Or AnchorStyles.Right
                ElseIf texto = "pedido" Or texto = "cliente" Or texto = "vendedor" Or texto = "observaciones" Then
                    ctrl.Anchor = AnchorStyles.Top Or AnchorStyles.Left
                End If
            End If
        Next

        ' Pie de página (Totales a la derecha)
        Dim controlesTotales As Control() = {LabelBase, TextBoxBase, LabelIva, TextBoxIva, Label7, TextBoxTotalPed}
        For Each ctrl In controlesTotales
            ctrl.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        Next

        ' Botones
        ButtonAnterior.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        ButtonSiguiente.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        LabelStock.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left

        Dim botonesIzq As Control() = {ButtonGuardar, ButtonBorrar, ButtonNuevoPresup, ButtonBorrarLineas, ButtonNuevaLinea}
        For Each ctrl In botonesIzq
            ctrl.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
        Next
    End Sub

    Private Sub EstilizarBoton(btn As Button, x As Integer, y As Integer, bg As Color, fg As Color, Optional ghost As Boolean = False)
        btn.Location = New Point(x, y)
        btn.Size = New Size(100, 35)
        btn.FlatStyle = FlatStyle.Flat
        btn.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        btn.Cursor = Cursors.Hand

        If ghost Then
            btn.BackColor = bg
            btn.ForeColor = fg
            btn.FlatAppearance.BorderColor = Color.White
            btn.FlatAppearance.BorderSize = 1
        Else
            btn.BackColor = bg
            btn.ForeColor = fg
            btn.FlatAppearance.BorderSize = 0
        End If
    End Sub
#End Region
    ' =========================================================
    ' FUNCIONES SALVAVIDAS PARA LEER DE LA BASE DE DATOS
    ' =========================================================
    Private Function LeerTexto(r As SQLiteDataReader, col As String) As String
        Try
            Dim idx As Integer = r.GetOrdinal(col)
            Return If(r.IsDBNull(idx), "", r.GetValue(idx).ToString())
        Catch ex As IndexOutOfRangeException
            MessageBox.Show($"Falta la columna '{col}' en la tabla de Pedidos. ¡Revisa tu BD!", "Aviso")
            Return ""
        End Try
    End Function

    Private Function LeerDecimal(r As SQLiteDataReader, col As String) As Decimal
        Try
            Dim idx As Integer = r.GetOrdinal(col)
            Return If(r.IsDBNull(idx), 0D, Convert.ToDecimal(r.GetValue(idx)))
        Catch ex As IndexOutOfRangeException
            MessageBox.Show($"Falta la columna '{col}' en la tabla de Pedidos.", "Aviso")
            Return 0D
        End Try
    End Function

    ' =========================================================
    ' FUNCIÓN PARA VOLCAR LOS DATOS DEL PEDIDO A LA PANTALLA
    ' =========================================================
    Private Sub CargarDatosDocumento(numeroDoc As String)
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()

            ' 1. CARGAR LA CABECERA DEL PEDIDO
            Dim sqlCabecera As String = "SELECT P.*, C.NombreFiscal AS NombreCli " &
                                        "FROM Pedidos P " &
                                        "LEFT JOIN Clientes C ON P.CodigoCliente = C.CodigoCliente " &
                                        "WHERE P.NumeroPedido = @num"

            Using cmd As New SQLiteCommand(sqlCabecera, c)
                cmd.Parameters.AddWithValue("@num", numeroDoc)
                Using r = cmd.ExecuteReader()
                    If r.Read() Then
                        TextBoxIdCliente.Text = LeerTexto(r, "CodigoCliente")
                        TextBoxCliente.Text = LeerTexto(r, "NombreCli")
                        TextBoxIdVendedor.Text = LeerTexto(r, "ID_Vendedor")

                        ' Si el pedido viene de un presupuesto, lo cargamos
                        TextBoxIdPresupuesto.Text = LeerTexto(r, "NumeroPresupuesto")

                        Dim fechaBD As String = LeerTexto(r, "Fecha")
                        If Not String.IsNullOrEmpty(fechaBD) Then
                            TextBoxFecha.Text = Convert.ToDateTime(fechaBD).ToString("dd/MM/yyyy")
                            Try : DateTimePickerFecha.Value = Convert.ToDateTime(fechaBD) : Catch : End Try
                        End If

                        TextBoxEstado.Text = LeerTexto(r, "Estado")
                        TextBoxObservaciones.Text = LeerTexto(r, "Observaciones")
                        ' Cargar Combos Nuevos
                        Dim idPago = r("ID_FormaPago")
                        If Not IsDBNull(idPago) Then cboFormaPago.SelectedValue = Convert.ToInt32(idPago) Else cboFormaPago.SelectedIndex = -1

                        Dim idRut = r("ID_Ruta")
                        If Not IsDBNull(idRut) Then cboRuta.SelectedValue = Convert.ToInt32(idRut) Else cboRuta.SelectedIndex = -1
                        ' Totales
                        Dim base As Decimal = LeerDecimal(r, "BaseImponible")
                        Dim iva As Decimal = LeerDecimal(r, "IVA") ' O ImporteIVA si lo llamaste así en tu BD
                        Dim total As Decimal = LeerDecimal(r, "Total")

                        TextBoxBase.Text = base.ToString("N2") & " €"
                        TextBoxIva.Text = iva.ToString("N2") & " €"
                        TextBoxTotalPed.Text = total.ToString("N2") & " €" ' Usamos TextBoxTotalPed
                    End If
                End Using
            End Using

            ' 2. CARGAR LAS LÍNEAS DEL PEDIDO
            Dim sqlLineas As String = "SELECT * FROM LineasPedido WHERE NumeroPedido = @num"
            Using cmdLineas As New SQLiteCommand(sqlLineas, c)
                cmdLineas.Parameters.AddWithValue("@num", numeroDoc)
                Dim da As New SQLiteDataAdapter(cmdLineas)
                Dim dt As New DataTable()
                da.Fill(dt)

                DataGridView1.DataSource = dt
            End Using
            ' =========================================================
            ' AJUSTAR EL ANCHO DE LAS COLUMNAS DE LA TABLA
            ' =========================================================
            If DataGridView1.Columns.Contains("Descripcion") Then
                DataGridView1.Columns("Descripcion").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            End If

            If DataGridView1.Columns.Contains("PrecioUnitario") Then
                DataGridView1.Columns("PrecioUnitario").Width = 140
            End If

            If DataGridView1.Columns.Contains("Total") Then
                DataGridView1.Columns("Total").Width = 170
            End If
        Catch ex As Exception
            MessageBox.Show("Error al cargar los datos del pedido: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ' =========================================================
    ' LUPA: BUSCADOR UNIVERSAL DE PEDIDOS
    ' =========================================================
    Private Sub btnBuscarPedido_Click(sender As Object, e As EventArgs) Handles btnBuscarPedido.Click
        Dim frmBuscar As New FrmBuscador()

        ' Tu buscador ya está programado para entender "Pedidos"
        frmBuscar.TablaABuscar = "Pedidos"

        If frmBuscar.ShowDialog() = DialogResult.OK Then

            ' 1. Escribimos el código devuelto (Ej: "PED-003") en la caja
            TextBoxPedido.Text = frmBuscar.Resultado

            ' 2. Llamamos a nuestra función para que lea y rellene todo
            CargarDatosDocumento(frmBuscar.Resultado)

        End If
    End Sub
    ' =========================================================
    ' LÓGICA DE STOCK EN TIEMPO REAL
    ' =========================================================
    ' =========================================================
    ' LÓGICA DE STOCK EN TIEMPO REAL (CON PREVISIÓN)
    ' =========================================================
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        ' Si la tabla está vacía o es una fila nueva, borramos el texto
        If DataGridView1.CurrentRow Is Nothing OrElse DataGridView1.CurrentRow.IsNewRow Then
            LabelStock.Text = "Stock disponible: -"
            LabelStock.ForeColor = Color.FromArgb(170, 175, 180)
            Return
        End If

        Try
            ' 1. Leer el ID del Artículo
            Dim idCelda As Object = DataGridView1.CurrentRow.Cells("ID_Articulo").Value

            If idCelda IsNot Nothing AndAlso Not DBNull.Value.Equals(idCelda) Then
                Dim idArticulo As Integer = Convert.ToInt32(idCelda)

                ' 2. Averiguar cómo se llama la columna de cantidad (En Albaranes a veces se llama "CantidadServida", en Pedidos "Cantidad")
                Dim nombreColumnaCant As String = "Cantidad"
                If DataGridView1.Columns.Contains("CantidadServida") Then nombreColumnaCant = "CantidadServida"

                ' 3. Leer la cantidad que el usuario ha escrito en la línea
                Dim cantidadEnLinea As Decimal = 0
                Dim cantCelda As Object = DataGridView1.CurrentRow.Cells(nombreColumnaCant).Value
                If cantCelda IsNot Nothing AndAlso Not DBNull.Value.Equals(cantCelda) Then
                    cantidadEnLinea = Convert.ToDecimal(cantCelda)
                End If

                ' 4. Consultar el Stock Real a la BD
                Dim stockActual As Decimal = ConsultarStock(idArticulo)

                ' 5. ¡LA MATEMÁTICA! (Lo que nos quedará después de este documento)
                Dim stockRestante As Decimal = stockActual - cantidadEnLinea

                ' 6. ESTILOS Y MENSAJES DINÁMICOS
                If stockRestante > 0 Then
                    ' Todo bien, hay stock de sobra
                    LabelStock.Text = $"Stock actual: {stockActual} (Te quedarán: {stockRestante})"
                    LabelStock.ForeColor = Color.FromArgb(40, 180, 90) ' Verde brillante

                ElseIf stockRestante = 0 Then
                    ' ¡Ojo! Nos quedamos a cero
                    LabelStock.Text = $"Stock actual: {stockActual} (¡Atención! Te quedarás a 0)"
                    LabelStock.ForeColor = Color.FromArgb(220, 160, 40) ' Naranja / Amarillo

                Else
                    ' ¡Alerta Roja! No hay suficiente para cubrir la línea
                    ' Usamos Math.Abs para que no ponga "Faltan -2", sino "Faltan 2"
                    LabelStock.Text = $"¡Stock insuficiente! Actual: {stockActual} (Te faltan: {Math.Abs(stockRestante)})"
                    LabelStock.ForeColor = Color.FromArgb(255, 80, 80) ' Rojo salmón
                End If

                LabelStock.Font = New Font("Segoe UI", 10.5F, FontStyle.Bold)
            End If
        Catch ex As Exception
            LabelStock.Text = "Stock disponible: -"
        End Try
    End Sub

    ' Función rápida para ir a SQLite a buscar el stock de ese artículo exacto
    Private Function ConsultarStock(idArticulo As Integer) As Decimal
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()

            Dim sql As String = "SELECT StockActual FROM Articulos WHERE ID_Articulo = @id"
            Using cmd As New System.Data.SQLite.SQLiteCommand(sql, c)
                cmd.Parameters.AddWithValue("@id", idArticulo)
                Dim resultado As Object = cmd.ExecuteScalar()

                If resultado IsNot Nothing AndAlso Not DBNull.Value.Equals(resultado) Then
                    Return Convert.ToDecimal(resultado)
                End If
            End Using
        Catch ex As Exception
        End Try

        Return 0
    End Function
End Class