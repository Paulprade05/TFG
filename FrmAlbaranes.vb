Imports System.Data.SQLite
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Text
Imports System.Security.Cryptography
Imports System.Runtime.InteropServices
Public Class FrmAlbaranes
    ' Clave principal TEXTO (Ej: ALB-26-001)
    Private _numeroAlbaranActual As String = ""
    Private _dtLineas As DataTable
    Private _idsParaBorrar As New List(Of Integer)
    ' --- NUEVOS DESPLEGABLES ---
    Private WithEvents cboFormaPago As New System.Windows.Forms.ComboBox()
    Private WithEvents cboRuta As New System.Windows.Forms.ComboBox()
    Private lblFormaPago As New Label() With {.Text = "Forma de Pago", .AutoSize = True, .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold), .ForeColor = Color.WhiteSmoke}
    Private lblRuta As New Label() With {.Text = "Ruta Asignada", .AutoSize = True, .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold), .ForeColor = Color.WhiteSmoke}
    Private Sub FrmAlbaranes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Estilos
        FrmPresupuestos.EstilizarGrid(DataGridView1)
        EstilizarFecha(DateTimePickerFecha)
        ' Cargar opciones de Portes por defecto
        If ComboBoxPortes.Items.Count = 0 Then
            ComboBoxPortes.Items.Add("Pagados")
            ComboBoxPortes.Items.Add("Debidos")
            ComboBoxPortes.SelectedIndex = 0
        End If

        ConfigurarGrid()
        Dim listaAgencias As List(Of Agencia) = ObtenerAgencias()
        ' Configuramos el ComboBox
        With cboAgencias
            .DataSource = listaAgencias
            .DisplayMember = "Nombre"    ' Lo que ve el usuario
            .ValueMember = "ID_Agencia"  ' El valor que usaremos para guardar en la DB
            .SelectedIndex = -1          ' Empezar sin ninguna seleccionada
        End With
        Me.Controls.Add(lblFormaPago) : Me.Controls.Add(cboFormaPago)
        Me.Controls.Add(lblRuta) : Me.Controls.Add(cboRuta)
        cboFormaPago.DropDownStyle = ComboBoxStyle.DropDownList
        cboRuta.DropDownStyle = ComboBoxStyle.DropDownList
        CargarDesplegables()
        ' Cargar último albarán
        Dim ultimoNum As String = ObtenerUltimoNumeroAlbaran()
        If Not String.IsNullOrEmpty(ultimoNum) Then
            CargarAlbaran(ultimoNum)
        Else
            LimpiarFormulario()
        End If
        TabControlModerno2.DrawMode = TabDrawMode.Normal
        TabControlModerno2.SizeMode = TabSizeMode.Normal
        TabControlModerno2.ItemSize = New Size(0, 0) ' (0,0) hace que Windows use el ancho real del texto
        If TextBox5 IsNot Nothing Then TextBox5.Visible = False
        If Button4 IsNot Nothing Then Button4.Visible = False
        If Label23 IsNot Nothing Then Label23.Visible = False
        ' --- AÑADE ESTAS DOS LÍNEAS AQUÍ ---
        ReorganizarControlesAutomaticamente()


    End Sub

    ' =========================================================
    ' 1. CONFIGURACIÓN DE DATOS Y GRID
    ' =========================================================
    Private Sub ConfigurarEstructuraDatos()
        If _dtLineas Is Nothing Then _dtLineas = New DataTable()

        ' Columnas de control
        If Not _dtLineas.Columns.Contains("ID_Linea") Then _dtLineas.Columns.Add("ID_Linea", GetType(Object))
        If Not _dtLineas.Columns.Contains("NumeroAlbaran") Then _dtLineas.Columns.Add("NumeroAlbaran", GetType(String))

        ' Datos de línea
        If Not _dtLineas.Columns.Contains("NumeroOrden") Then _dtLineas.Columns.Add("NumeroOrden", GetType(Integer))
        If Not _dtLineas.Columns.Contains("ID_Articulo") Then _dtLineas.Columns.Add("ID_Articulo", GetType(Object))
        If Not _dtLineas.Columns.Contains("Descripcion") Then _dtLineas.Columns.Add("Descripcion", GetType(String))

        ' Valores numéricos
        If Not _dtLineas.Columns.Contains("CantidadServida") Then _dtLineas.Columns.Add("CantidadServida", GetType(Decimal)) ' Ojo: CantidadServida
        If Not _dtLineas.Columns.Contains("PrecioUnitario") Then _dtLineas.Columns.Add("PrecioUnitario", GetType(Decimal))
        If Not _dtLineas.Columns.Contains("Descuento") Then _dtLineas.Columns.Add("Descuento", GetType(Decimal))
        If Not _dtLineas.Columns.Contains("Total") Then _dtLineas.Columns.Add("Total", GetType(Decimal))

        ' Trazabilidad con Pedido (Opcional si lo usas)
        If Not _dtLineas.Columns.Contains("ID_LineaPedido") Then _dtLineas.Columns.Add("ID_LineaPedido", GetType(Object))
    End Sub

    Private Sub ConfigurarGrid()
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.Columns.Clear()

        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "ID_Linea", .DataPropertyName = "ID_Linea", .Visible = False})

        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "NumeroOrden", .DataPropertyName = "NumeroOrden", .HeaderText = "Nº", .Width = 40, .ReadOnly = True, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleCenter}})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "ID_Articulo", .DataPropertyName = "ID_Articulo", .HeaderText = "ID Art", .Width = 70})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "Descripcion", .DataPropertyName = "Descripcion", .HeaderText = "Descripción", .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill})


        ' En Albaranes hablamos de "Cantidad Servida" (lo que se entrega)
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "CantidadServida", .DataPropertyName = "CantidadServida", .HeaderText = "Entregado", .Width = 70, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "N2", .BackColor = Color.Ivory}})

        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "PrecioUnitario", .DataPropertyName = "PrecioUnitario", .HeaderText = "Precio", .Width = 80, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "C2"}})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "Descuento", .DataPropertyName = "Descuento", .HeaderText = "% Dto", .Width = 60, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "N2"}})

        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "Total", .DataPropertyName = "Total", .HeaderText = "Total", .ReadOnly = True, .Width = 90, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "C2", .BackColor = Color.WhiteSmoke}})
    End Sub

    ' =========================================================
    ' 2. IMPORTAR PEDIDO (El botón mágico)
    ' =========================================================
    Private Sub btnImportarPedido_Click(sender As Object, e As EventArgs)
        Using frm As New FrmBuscador
            frm.TablaABuscar = "Pedidos"
            If frm.ShowDialog = DialogResult.OK Then
                Dim numPedido = frm.Resultado
                If Not String.IsNullOrEmpty(numPedido) Then ImportarDatosPedido(numPedido)
            End If
        End Using
    End Sub
    ' Busca el código del último albarán existente en la base de datos (Ej: "ALB-005")
    Private Function ObtenerUltimoNumeroAlbaran() As String
        Dim ultimoNumero As String = ""

        Try
            Dim conexion = ConexionBD.GetConnection()

            ' MAX funciona bien con textos si tienen el mismo formato y ceros a la izquierda (ALB-001, ALB-002...)
            Dim sql As String = "SELECT MAX(NumeroAlbaran) FROM Albaranes"

            Using cmd As New SQLiteCommand(sql, conexion)
                If conexion.State <> ConnectionState.Open Then conexion.Open()

                Dim resultado = cmd.ExecuteScalar()

                If resultado IsNot Nothing AndAlso Not IsDBNull(resultado) Then
                    ultimoNumero = resultado.ToString()
                End If
            End Using
        Catch ex As Exception
            ' Si hay error, devolvemos vacío y el formulario se pondrá en modo "Nuevo"
        End Try

        Return ultimoNumero
    End Function
    Private Sub ImportarDatosPedido(numeroPedido As String)
        Dim c = ConexionBD.GetConnection()
        Try
            If c.State <> ConnectionState.Open Then c.Open()

            ' 1. Cargar Cabecera del Pedido y Datos del Cliente
            ' IMPORTANTE: Traemos la dirección del cliente para congelarla en el albarán
            Dim sqlCab As String = "SELECT P.*, C.NombreFiscal, C.Direccion, C.Poblacion, C.CodigoPostal " &
                                   "FROM Pedidos P " &
                                   "LEFT JOIN Clientes C ON P.ID_Cliente = C.ID_Cliente " &
                                   "LEFT JOIN Vendedores V ON P.ID_Vendedor=V.ID_Vendedor " &
                                   "WHERE P.NumeroPedido = @num"

            Using cmd As New SQLiteCommand(sqlCab, c)
                cmd.Parameters.AddWithValue("@num", numeroPedido)
                Using r = cmd.ExecuteReader()
                    If r.Read() Then
                        LimpiarFormulario()

                        ' Referencia
                        TextBoxPedidoOrigen.Text = numeroPedido
                        TextBoxPedidoOrigen.Tag = numeroPedido ' Guardamos para BD

                        ' Cliente
                        TextBoxIdCliente.Text = r("ID_Cliente").ToString()
                        TextBoxCliente.Text = r("NombreFiscal").ToString()
                        TextBoxIdVendedor.Text = If(IsDBNull(r("ID_Vendedor")), "", r("ID_Vendedor").ToString())
                        TextBoxVendedor.Text = If(IsDBNull(r("NombreVend")), "", r("NombreVend").ToString())

                        ' Dirección de Envío (Se congela aquí)
                        TextBoxDireccion.Text = If(IsDBNull(r("Direccion")), "", r("Direccion").ToString())
                        TextBoxPoblacion.Text = If(IsDBNull(r("Poblacion")), "", r("Poblacion").ToString())
                        TextBoxCP.Text = If(IsDBNull(r("CodigoPostal")), "", r("CodigoPostal").ToString())

                        TextBoxObservaciones.Text = "Generado desde Pedido " & numeroPedido
                    End If
                End Using
            End Using

            ' 2. Cargar Líneas del Pedido
            ' OJO: Aquí mapeamos "Cantidad" del pedido a "CantidadServida" del albarán
            Dim sqlLin As String = "SELECT * FROM LineasPedido WHERE NumeroPedido = @num ORDER BY NumeroOrden ASC"
            Dim dtOrigen As New DataTable()
            Using cmd As New SQLiteCommand(sqlLin, c)
                cmd.Parameters.AddWithValue("@num", numeroPedido)
                Dim da As New SQLiteDataAdapter(cmd)
                da.Fill(dtOrigen)
            End Using

            _dtLineas.Rows.Clear()
            For Each rowOrig As DataRow In dtOrigen.Rows
                Dim rowNew As DataRow = _dtLineas.NewRow()

                rowNew("NumeroAlbaran") = _numeroAlbaranActual
                rowNew("ID_Linea") = DBNull.Value

                rowNew("NumeroOrden") = rowOrig("NumeroOrden")
                rowNew("ID_Articulo") = rowOrig("ID_Articulo")
                rowNew("Descripcion") = rowOrig("Descripcion")

                ' Mapeo de Cantidad
                Dim cant As Decimal = 0 : Decimal.TryParse(rowOrig("Cantidad").ToString(), cant)
                rowNew("CantidadServida") = cant ' Por defecto servimos todo lo pedido

                ' Precios
                rowNew("PrecioUnitario") = rowOrig("PrecioUnitario")
                rowNew("Descuento") = rowOrig("Descuento")
                rowNew("Total") = rowOrig("Total")

                ' Trazabilidad (Guardamos el ID de la línea original si lo tienes)
                rowNew("ID_LineaPedido") = rowOrig("ID_Linea")

                _dtLineas.Rows.Add(rowNew)
            Next

            CalcularTotales()
            MessageBox.Show("Datos del Pedido cargados.")

        Catch ex As Exception
            MessageBox.Show("Error al importar: " & ex.Message)
        End Try
    End Sub
    Private Sub CargarDesplegables()
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()
            ' Formas de pago
            Dim daPago As New SQLiteDataAdapter("SELECT ID_FormaPago, Descripcion FROM FormasPago WHERE Activo=1", c)
            Dim dtPago As New DataTable() : daPago.Fill(dtPago)
            cboFormaPago.DataSource = dtPago : cboFormaPago.DisplayMember = "Descripcion" : cboFormaPago.ValueMember = "ID_FormaPago" : cboFormaPago.SelectedIndex = -1
            ' Rutas
            Dim daRuta As New SQLiteDataAdapter("SELECT ID_Ruta, NombreZona FROM Rutas WHERE Activo=1", c)
            Dim dtRuta As New DataTable() : daRuta.Fill(dtRuta)
            cboRuta.DataSource = dtRuta : cboRuta.DisplayMember = "NombreZona" : cboRuta.ValueMember = "ID_Ruta" : cboRuta.SelectedIndex = -1
        Catch ex As Exception
        End Try
    End Sub
    ' =========================================================
    ' 3. GUARDAR (Insert/Update con nuevos campos)
    ' =========================================================
    ' =========================================================
    ' 3. GUARDAR (Con registro de movimientos automático)
    ' =========================================================
    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles ButtonGuardar.Click
        If String.IsNullOrWhiteSpace(TextBoxIdCliente.Text) Then MessageBox.Show("Falta el Cliente") : Return

        Dim esNuevo As Boolean = String.IsNullOrEmpty(_numeroAlbaranActual)
        If esNuevo Then
            TextBoxAlbaran.Text = GenerarProximoNumeroAlbaran()
            _numeroAlbaranActual = TextBoxAlbaran.Text
        End If

        Dim c = ConexionBD.GetConnection()
        Dim trans As SQLiteTransaction = Nothing

        Try
            If c.State <> ConnectionState.Open Then c.Open()
            trans = c.BeginTransaction()

            ' A) Calcular Totales
            Dim sumaBase As Decimal = 0
            For Each r As DataRow In _dtLineas.Rows
                If r.RowState <> DataRowState.Deleted Then
                    Dim cant As Decimal = 0 : Decimal.TryParse(r("CantidadServida").ToString(), cant)
                    Dim prec As Decimal = 0 : Decimal.TryParse(r("PrecioUnitario").ToString(), prec)
                    Dim desc As Decimal = 0 : Decimal.TryParse(r("Descuento").ToString(), desc)
                    Dim tot As Decimal = (cant * prec) * (1 - (desc / 100))

                    If _dtLineas.Columns.Contains("Total") Then r("Total") = tot
                    sumaBase += tot
                End If
            Next
            Dim sumaTotal As Decimal = sumaBase * 1.21D ' IVA 21%

            ' B) Guardar Cabecera
            Dim sql As String
            If esNuevo Then
                sql = "INSERT INTO Albaranes (NumeroAlbaran, NumeroPedido, Fecha, CodigoCliente, ID_Agencia, NumeroBultos, PesoTotal, CodigoSeguimiento, Portes, Estado, Observaciones, DireccionEnvio, Poblacion, CodigoPostal, FechaEntrega, BaseImponible, Total, ID_Vendedor) " &
                      "VALUES (@num, @ped, @fecha, @cli, @agencia, @bultos, @peso, @track, @portes, @estado, @obs, @dir, @pob, @cp, @FEntrega, @base, @total, @vend)"
            Else
                sql = "UPDATE Albaranes SET NumeroPedido=@ped, Fecha=@fecha, CodigoCliente=@cli, ID_Agencia=@agencia, NumeroBultos=@bultos, PesoTotal=@peso, CodigoSeguimiento=@track, Portes=@portes, Estado=@estado, Observaciones=@obs, DireccionEnvio=@dir, FechaEntrega=@FEntrega, Poblacion=@pob, CodigoPostal=@cp, BaseImponible=@base, Total=@total,ID_Vendedor=@vend WHERE NumeroAlbaran=@num"
            End If

            Dim fechaAlbaran As DateTime = DateTime.Now
            If DateTime.TryParse(TextBoxFecha.Text, fechaAlbaran) = False Then fechaAlbaran = DateTime.Now

            Using cmd As New SQLiteCommand(sql, c)
                cmd.Transaction = trans
                cmd.Parameters.AddWithValue("@num", _numeroAlbaranActual)

                Dim ped As Object = If(TextBoxPedidoOrigen.Tag IsNot Nothing, TextBoxPedidoOrigen.Tag, DBNull.Value)
                cmd.Parameters.AddWithValue("@ped", ped)
                cmd.Parameters.AddWithValue("@fecha", fechaAlbaran)
                cmd.Parameters.AddWithValue("@cli", TextBoxIdCliente.Text)

                cmd.Parameters.AddWithValue("@agencia", cboAgencias.SelectedValue)
                cmd.Parameters.AddWithValue("@bultos", If(IsNumeric(TextBoxBultos.Text), TextBoxBultos.Text, 1))
                cmd.Parameters.AddWithValue("@peso", If(IsNumeric(TextBoxPeso.Text), TextBoxPeso.Text, 0))
                cmd.Parameters.AddWithValue("@track", TextBoxTracking.Text)
                cmd.Parameters.AddWithValue("@portes", ComboBoxPortes.Text)
                cmd.Parameters.AddWithValue("@vend", TextBoxIdVendedor.Text)
                cmd.Parameters.AddWithValue("@estado", TextBoxEstado.Text)
                cmd.Parameters.AddWithValue("@obs", TextBoxObservaciones.Text)
                cmd.Parameters.AddWithValue("@FEntrega", DateTimePickerFecha.Value)
                cmd.Parameters.AddWithValue("@dir", TextBoxDireccion.Text)
                cmd.Parameters.AddWithValue("@pob", TextBoxPoblacion.Text)
                cmd.Parameters.AddWithValue("@cp", TextBoxCP.Text)
                cmd.Parameters.AddWithValue("@base", sumaBase)
                cmd.Parameters.AddWithValue("@total", sumaTotal)
                cmd.ExecuteNonQuery()
            End Using

            ' C) Guardar Líneas y PROCESAR MOVIMIENTOS

            ' C.1) Bajas de líneas (Se devuelve el stock al almacén)
            For Each idDel In _idsParaBorrar
                ' Recuperar qué articulo y cantidad era para devolverlo
                Dim idArtDel As Integer = 0
                Dim cantDel As Decimal = 0
                Using cmdInfo As New SQLiteCommand("SELECT ID_Articulo, CantidadServida FROM LineasAlbaran WHERE ID_Linea = @id", c, trans)
                    cmdInfo.Parameters.AddWithValue("@id", idDel)
                    Using reader = cmdInfo.ExecuteReader()
                        If reader.Read() Then
                            If Not IsDBNull(reader("ID_Articulo")) Then idArtDel = Convert.ToInt32(reader("ID_Articulo"))
                            cantDel = Convert.ToDecimal(reader("CantidadServida"))
                        End If
                    End Using
                End Using

                Using cmdDel As New SQLiteCommand("DELETE FROM LineasAlbaran WHERE ID_Linea = @id", c)
                    cmdDel.Transaction = trans : cmdDel.Parameters.AddWithValue("@id", idDel) : cmdDel.ExecuteNonQuery()
                End Using

                ' Devolver Stock (Variación negativa = ENTRADA)
                If idArtDel > 0 Then ActualizarStockYMovimiento(c, trans, _numeroAlbaranActual, idArtDel, -cantDel, fechaAlbaran)
            Next
            _idsParaBorrar.Clear()

            ' C.2) Inserciones y Actualizaciones
            If _dtLineas IsNot Nothing Then
                For Each r As DataRow In _dtLineas.Rows
                    If r.RowState = DataRowState.Deleted Then Continue For

                    Dim cant As Decimal = 0 : Decimal.TryParse(r("CantidadServida").ToString(), cant)
                    Dim prec As Decimal = 0 : Decimal.TryParse(r("PrecioUnitario").ToString(), prec)
                    Dim desc As Decimal = 0 : Decimal.TryParse(r("Descuento").ToString(), desc)
                    Dim tot As Decimal = (cant * prec) * (1 - (desc / 100))

                    Dim idLin = r("ID_Linea")
                    Dim idArt As Object = If(IsNumeric(r("ID_Articulo")) AndAlso Val(r("ID_Articulo")) > 0, r("ID_Articulo"), DBNull.Value)

                    If IsDBNull(idLin) OrElse Not IsNumeric(idLin) Then
                        ' INSERT
                        Dim sqlIns As String = "INSERT INTO LineasAlbaran (NumeroAlbaran, NumeroOrden, ID_Articulo, Descripcion, CantidadServida, PrecioUnitario, Descuento, Total) VALUES (@alb, @ord, @art, @desc, @cant, @prec, @dcto, @tot)"
                        Using cmdL As New SQLiteCommand(sqlIns, c)
                            cmdL.Transaction = trans
                            cmdL.Parameters.AddWithValue("@alb", _numeroAlbaranActual)
                            cmdL.Parameters.AddWithValue("@ord", r("NumeroOrden"))
                            cmdL.Parameters.AddWithValue("@art", idArt)
                            cmdL.Parameters.AddWithValue("@desc", r("Descripcion"))
                            cmdL.Parameters.AddWithValue("@cant", cant)
                            cmdL.Parameters.AddWithValue("@prec", prec)
                            cmdL.Parameters.AddWithValue("@dcto", desc)
                            cmdL.Parameters.AddWithValue("@tot", tot)
                            cmdL.ExecuteNonQuery()
                        End Using

                        ' MOVIMIENTO: Salida de almacén (Nueva línea)
                        If Not IsDBNull(idArt) Then ActualizarStockYMovimiento(c, trans, _numeroAlbaranActual, Convert.ToInt32(idArt), cant, fechaAlbaran)

                    ElseIf r.RowState = DataRowState.Modified Then
                        ' UPDATE: Calculamos la diferencia entre lo que había y lo que hay ahora
                        Dim cantAnterior As Decimal = 0
                        If r.HasVersion(DataRowVersion.Original) AndAlso Not IsDBNull(r("CantidadServida", DataRowVersion.Original)) Then
                            Decimal.TryParse(r("CantidadServida", DataRowVersion.Original).ToString(), cantAnterior)
                        End If
                        Dim variacion = cant - cantAnterior ' Ej: antes 2, ahora 5 -> Salen 3 más.

                        Dim sqlUpd As String = "UPDATE LineasAlbaran SET ID_Articulo=@art, Descripcion=@desc, CantidadServida=@cant, PrecioUnitario=@prec, Descuento=@dcto, Total=@tot WHERE ID_Linea=@id"
                        Using cmdL As New SQLiteCommand(sqlUpd, c)
                            cmdL.Transaction = trans
                            cmdL.Parameters.AddWithValue("@art", idArt)
                            cmdL.Parameters.AddWithValue("@desc", r("Descripcion"))
                            cmdL.Parameters.AddWithValue("@cant", cant)
                            cmdL.Parameters.AddWithValue("@prec", prec)
                            cmdL.Parameters.AddWithValue("@dcto", desc)
                            cmdL.Parameters.AddWithValue("@tot", tot)
                            cmdL.Parameters.AddWithValue("@id", idLin)
                            cmdL.ExecuteNonQuery()
                        End Using

                        ' MOVIMIENTO: Ajuste por la diferencia modificada
                        If Not IsDBNull(idArt) AndAlso variacion <> 0 Then
                            ActualizarStockYMovimiento(c, trans, _numeroAlbaranActual, Convert.ToInt32(idArt), variacion, fechaAlbaran)
                        End If
                    End If
                Next
            End If

            trans.Commit()
            MessageBox.Show("Albarán Guardado y Movimientos generados.")
            CargarAlbaran(_numeroAlbaranActual)

        Catch ex As Exception
            If trans IsNot Nothing Then trans.Rollback()
            MessageBox.Show("Error al guardar: " & ex.Message)
        End Try
    End Sub

    ' =========================================================
    ' 4. FUNCIONES AUXILIARES (Carga, Limpieza, Generación)
    ' =========================================================
    Private Sub CargarAlbaran(num As String)

        Try
            Dim conexion = ConexionBD.GetConnection()
            If conexion.State <> ConnectionState.Open Then conexion.Open()

            Dim sql As String = "SELECT A.*, C.NombreFiscal AS NombreCliente, V.Nombre AS NombreVendedor, T.Nombre As NombreAgencia," &
                                "C.Poblacion, C.CodigoPostal " &
                                "FROM Albaranes A " &
                                "LEFT JOIN Clientes C ON A.CodigoCliente = C.CodigoCliente " &
                                "LEFT JOIN Vendedores V ON A.ID_Vendedor = V.ID_Vendedor " &
                                "LEFT JOIN Agencias T ON A.ID_Agencia = T.ID_Agencia " &
                                "WHERE A.NumeroAlbaran= @num"

            Using cmd As New SQLiteCommand(sql, conexion)
                cmd.Parameters.AddWithValue("@num", num)
                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        _numeroAlbaranActual = num

                        ' 1. Cargar Textos Simples (Sin procesar nada raro)
                        TextBoxAlbaran.Text = reader("NumeroAlbaran").ToString()

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
                        TextBoxDireccion.Text = reader("DireccionEnvio").ToString()
                        TextBoxPoblacion.Text = reader("Poblacion").ToString()
                        TextBoxCP.Text = reader("CodigoPostal").ToString()

                        TextBoxIdVendedor.Text = If(IsDBNull(reader("ID_Vendedor")), "", reader("ID_Vendedor").ToString())
                        TextBoxVendedor.Text = If(IsDBNull(reader("NombreVendedor")), "", reader("NombreVendedor").ToString())

                        If Not IsDBNull(reader("ID_Agencia")) Then
                            ' Al poner el ID aquí, el ComboBox busca el nombre y lo selecciona solo
                            cboAgencias.SelectedValue = Convert.ToInt32(reader("ID_Agencia"))
                        Else
                            ' Si no tiene agencia, limpiamos la selección
                            cboAgencias.SelectedIndex = -1
                        End If
                        ' 5. Totales (Si los tienes en pantalla)
                        If IsNumeric(reader("BaseImponible")) Then TextBoxBase.Text = Convert.ToDecimal(reader("BaseImponible")).ToString("C2")
                        If IsNumeric(reader("Total")) Then TextBoxTotalAlb.Text = Convert.ToDecimal(reader("Total")).ToString("C2")

                        ' 6. Trazabilidad (Presupuesto Origen)
                        ' Verifica que el campo en BD sea ID_Presupuesto
                        If Not IsDBNull(reader("NumeroPedido")) Then
                            TextBoxPedidoOrigen.Text = reader("NumeroPedido").ToString()
                        End If
                        DateTimePickerFecha.Value = Convert.ToDateTime(reader("FechaEntrega")).ToShortDateString
                        TextBoxBultos.Text = reader("NumeroBultos")
                        TextBoxPeso.Text = reader("PesoTotal")
                        TextBoxTracking.Text = reader("CodigoSeguimiento")
                        ComboBoxPortes.ValueMember = reader("Portes")
                    Else
                        MessageBox.Show("Albaran no encontrado.")
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
            Dim sql As String = "SELECT ID_Linea, NumeroAlbaran, NumeroOrden, ID_Articulo, Descripcion, CantidadServida, PrecioUnitario, Descuento, Total " &
                                "FROM LineasAlbaran WHERE NumeroAlbaran = @num ORDER BY NumeroOrden ASC"

            Dim conexion = ConexionBD.GetConnection()
            Using cmd As New SQLiteCommand(sql, conexion)
                cmd.Parameters.AddWithValue("@num", _numeroAlbaranActual)
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
        TextBoxTotalAlb.Text = (base * 1.21D).ToString("C2")
    End Sub
    Private Sub LimpiarFormulario()
        _numeroAlbaranActual = ""
        _idsParaBorrar.Clear()
        TextBoxAlbaran.Text = GenerarProximoNumeroAlbaran()
        TextBoxIdCliente.Text = "" : TextBoxCliente.Text = ""
        TextBoxPedidoOrigen.Text = "" : TextBoxPedidoOrigen.Tag = Nothing
        TextBoxIdVendedor.Text = "" : TextBoxVendedor.Text = ""
        cboAgencias.SelectedValue = ""
        ' Limpiar Logística
        TextBoxBultos.Text = "1" : TextBoxPeso.Text = "0"
        TextBoxDireccion.Text = "" : TextBoxPoblacion.Text = "" : TextBoxCP.Text = ""
        TextBoxTracking.Text = GenerarCodigoSeguimiento()
        TextBoxFecha.Text = DateTime.Now.ToShortDateString()
        DateTimePickerFecha.Value = DateTime.Now.ToShortDateString()
        TextBoxBase.Text = "0,00" : TextBoxTotalAlb.Text = "0,00"
        TextBoxEstado.Text = "Pendiente"
        ConfigurarGrid()
        _dtLineas = New DataTable()
        ConfigurarEstructuraDatos()
        DataGridView1.DataSource = _dtLineas
        If DataGridView1.Columns.Contains("ID_Linea") Then DataGridView1.Columns("ID_Linea").Visible = False
    End Sub

    Private Function GenerarProximoNumeroAlbaran() As String
        Dim prefijo As String = "ALB-"
        Dim nuevo As String = "ALB-001"
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()
            Dim cmd As New SQLiteCommand("SELECT NumeroAlbaran FROM Albaranes WHERE NumeroAlbaran LIKE @pat ORDER BY NumeroAlbaran DESC LIMIT 1", c)
            cmd.Parameters.AddWithValue("@pat", prefijo & "%")
            Dim res = cmd.ExecuteScalar()
            If res IsNot Nothing AndAlso Not IsDBNull(res) Then
                Dim parts = res.ToString().Split("-"c)
                If parts.Length >= 2 Then
                    nuevo = prefijo & (Convert.ToInt32(parts(parts.Length - 1)) + 1).ToString("D3")
                End If
            End If
        Catch
            nuevo = "ALB-" & DateTime.Now.Ticks.ToString().Substring(12)
        End Try
        Return nuevo
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Using frm As New FrmBuscador()
            frm.TablaABuscar = "Albaranes"
            If frm.ShowDialog() = DialogResult.OK Then
                CargarAlbaran(frm.Resultado)
            End If
        End Using
    End Sub
    Private Sub CalcularTotales()
        Dim base As Decimal = 0
        If _dtLineas IsNot Nothing Then
            For Each r As DataRow In _dtLineas.Rows
                If r.RowState <> DataRowState.Deleted Then
                    Dim c As Decimal = 0 : Decimal.TryParse(r("CantidadServida").ToString(), c)
                    Dim p As Decimal = 0 : Decimal.TryParse(r("PrecioUnitario").ToString(), p)
                    Dim d As Decimal = 0 : Decimal.TryParse(r("Descuento").ToString(), d)
                    base += (c * p) * (1 - (d / 100))
                End If
            Next
        End If
        TextBoxBase.Text = base.ToString("C2")
        TextBoxTotalAlb.Text = (base * 1.21D).ToString("C2")
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
    ' Agrega tus eventos de Celda (CellEndEdit) y Botones Borrar/Nuevo/Lineas aquí (son idénticos a Pedidos)
    Private Sub ButtonAnterior_Click(sender As Object, e As EventArgs) Handles ButtonAnterior.Click
        ' Lógica de navegación String (Ver Presupuestos)
        Try
            Dim conexion = ConexionBD.GetConnection()
            Dim numDestino As String = ""

            ' CAMBIO 7: Navegación alfanumérica (Funciona bien con formatos fijos tipo PRE-001)
            Dim sql As String = "SELECT MAX(NumeroAlbaran) FROM Albaranes WHERE NumeroAlbaran < @actual"
            Using cmd As New SQLiteCommand(sql, conexion)
                ' Si es nuevo, buscamos el último absoluto
                If String.IsNullOrEmpty(_numeroAlbaranActual) Then
                    cmd.CommandText = "SELECT MAX(NumeroAlbaran) FROM Albaranes"
                Else
                    cmd.Parameters.AddWithValue("@actual", _numeroAlbaranActual)
                End If

                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    numDestino = result.ToString()
                    CargarAlbaran(numDestino)
                Else
                    MessageBox.Show("Primer albaran.")
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub ButtonSiguiente_Click(sender As Object, e As EventArgs) Handles ButtonSiguiente.Click
        ' Lógica de navegación String (Ver Presupuesto)
        If String.IsNullOrEmpty(_numeroAlbaranActual) Then Return
        Try
            Dim conexion = ConexionBD.GetConnection()
            Dim sql As String = "SELECT MIN(NumeroAlbaran) FROM Albaranes WHERE NumeroAlbaran > @actual"
            Using cmd As New SQLiteCommand(sql, conexion)
                cmd.Parameters.AddWithValue("@actual", _numeroAlbaranActual)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    CargarAlbaran(result.ToString())
                Else
                    MessageBox.Show("Ultimo albaran")
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub


    Private Sub ButtonNuevo_Click(sender As Object, e As EventArgs) Handles ButtonNuevoPed.Click
        'Simplemente llamamos al metodo de limpieza que ya se encarga de todo
        LimpiarFormulario()
    End Sub
    Private Sub ButtonBorrar_Click(sender As Object, e As EventArgs) Handles ButtonBorrar.Click
        If _numeroAlbaranActual = "" Then
            MessageBox.Show("No puedes borrar un albaran que aún no existe (es nuevo).")
            Return
        End If

        If MessageBox.Show("¿Estás seguro de ELIMINAR este albaran? Se devolverá el stock al almacén.", "Confirmar Borrado", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No Then
            Return
        End If

        Dim conexion = ConexionBD.GetConnection()
        Dim transaccion As SQLiteTransaction = Nothing

        Try
            If conexion.State <> ConnectionState.Open Then conexion.Open()
            transaccion = conexion.BeginTransaction()

            ' PASO A: Recuperar el stock de las líneas antes de borrarlas
            Dim sqlRecuperar As String = "SELECT ID_Articulo, CantidadServida FROM LineasAlbaran WHERE NumeroAlbaran = @id"
            Dim lineasABorrar As New List(Of Tuple(Of Integer, Decimal))
            Using cmdRec As New SQLiteCommand(sqlRecuperar, conexion, transaccion)
                cmdRec.Parameters.AddWithValue("@id", _numeroAlbaranActual)
                Using reader = cmdRec.ExecuteReader()
                    While reader.Read()
                        If Not IsDBNull(reader("ID_Articulo")) Then
                            lineasABorrar.Add(New Tuple(Of Integer, Decimal)(Convert.ToInt32(reader("ID_Articulo")), Convert.ToDecimal(reader("CantidadServida"))))
                        End If
                    End While
                End Using
            End Using

            ' Devolvemos el stock (Variación negativa = ENTRADA)
            For Each linea In lineasABorrar
                ActualizarStockYMovimiento(conexion, transaccion, _numeroAlbaranActual & " (Borrado)", linea.Item1, -linea.Item2, DateTime.Now)
            Next

            ' PASO B: Borrar las LÍNEAS
            Dim sqlLineas As String = "DELETE FROM LineasAlbaran WHERE NumeroAlbaran= @id"
            Using cmd As New SQLiteCommand(sqlLineas, conexion)
                cmd.Transaction = transaccion
                cmd.Parameters.AddWithValue("@id", _numeroAlbaranActual)
                cmd.ExecuteNonQuery()
            End Using

            ' PASO C: Borrar la CABECERA
            Dim sqlCabecera As String = "DELETE FROM Albaranes WHERE NumeroAlbaran = @id"
            Using cmd As New SQLiteCommand(sqlCabecera, conexion)
                cmd.Transaction = transaccion
                cmd.Parameters.AddWithValue("@id", _numeroAlbaranActual)
                cmd.ExecuteNonQuery()
            End Using

            transaccion.Commit()
            MessageBox.Show("Albarán eliminado correctamente. Stock restaurado.", "Eliminado", MessageBoxButtons.OK, MessageBoxIcon.Information)
            LimpiarFormulario()

        Catch ex As Exception
            If transaccion IsNot Nothing Then transaccion.Rollback()
            MessageBox.Show("Error crítico al borrar: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
        nuevaFila("NumeroAlbaran") = _numeroAlbaranActual ' Si es 0 no pasa nada, se arregla al guardar
        nuevaFila("ID_Linea") = DBNull.Value     ' ESTO ES LO MÁS IMPORTANTE (Indica que es NUEVA)
        nuevaFila("ID_Articulo") = 0
        nuevaFila("Descripcion") = ""
        nuevaFila("CantidadServida") = 1
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



    Private Sub TextBoxIdCliente_Leave(sender As Object, e As EventArgs) Handles TextBoxIdCliente.Leave
        ' Tu lógica de buscar cliente
        If String.IsNullOrWhiteSpace(TextBoxIdCliente.Text) Then TextBoxCliente.Text = "" : Return
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()
            Dim cmd As New SQLiteCommand("SELECT NombreFiscal, Direccion, Poblacion, CodigoPostal FROM Clientes WHERE CodigoCliente=@id", c)
            cmd.Parameters.AddWithValue("@id", TextBoxIdCliente.Text)
            Dim r = cmd.ExecuteReader()
            If r.Read() Then
                TextBoxCliente.Text = r("NombreFiscal")
                TextBoxDireccion.Text = r("Direccion")
                TextBoxPoblacion.Text = r("Poblacion")
                TextBoxCP.Text = r("CodigoPostal")
            End If

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

    Private Sub btnImportarPedido_Click_1(sender As Object, e As EventArgs) Handles btnImportarPedido.Click
        Using frm As New FrmBuscador
            frm.TablaABuscar = "Pedidos"
            If frm.ShowDialog = DialogResult.OK Then
                Dim codPresu = frm.Resultado ' Ahora recibimos String
                If Not String.IsNullOrEmpty(codPresu) Then ImportarDatosPresupuesto(codPresu)
            End If
        End Using

    End Sub
    Private Sub ImportarDatosPresupuesto(CodigoPedido As String)
        Dim c = ConexionBD.GetConnection()
        Try
            If c.State <> ConnectionState.Open Then c.Open()

            ' 1. Cabecera (Buscamos por NumeroPresupuesto)
            Dim sqlCab As String = "SELECT P.*, C.NombreFiscal, V.Nombre AS NombreVend FROM Pedidos P " &
                                   "LEFT JOIN Clientes C ON P.CodigoCliente=C.CodigoCliente " &
                                   "LEFT JOIN Vendedores V ON P.ID_Vendedor=V.ID_Vendedor " &
                                   "WHERE P.NumeroPedido = @num"
            Using cmd As New SQLiteCommand(sqlCab, c)
                cmd.Parameters.AddWithValue("@num", CodigoPedido)
                Using r = cmd.ExecuteReader()
                    If r.Read() Then
                        LimpiarFormulario()
                        TextBoxIdCliente.Text = r("CodigoCliente").ToString()
                        TextBoxCliente.Text = r("NombreFiscal").ToString()
                        TextBoxIdVendedor.Text = If(IsDBNull(r("ID_Vendedor")), "", r("ID_Vendedor").ToString())
                        TextBoxVendedor.Text = If(IsDBNull(r("NombreVend")), "", r("NombreVend").ToString())
                        DateTimePickerFecha.Value = r("FechaEntrega")
                        TextBoxObservaciones.Text = $"Desde Presupuesto {CodigoPedido}. "
                        If TextBoxPedidoOrigen IsNot Nothing Then
                            TextBoxPedidoOrigen.Text = CodigoPedido
                            TextBoxPedidoOrigen.Tag = CodigoPedido
                        End If
                    End If
                End Using
            End Using

            ' 2. Líneas (Buscamos por NumeroPresupuesto)
            Dim sqlLin As String = "SELECT * FROM LineasPedido WHERE NumeroPedido = @num ORDER BY NumeroOrden ASC"
            Dim dtOrigen As New DataTable()
            Using cmd As New SQLiteCommand(sqlLin, c)
                cmd.Parameters.AddWithValue("@num", CodigoPedido)
                Dim da As New SQLiteDataAdapter(cmd)
                da.Fill(dtOrigen)
            End Using

            _dtLineas.Rows.Clear()
            For Each rowOrig As DataRow In dtOrigen.Rows
                Dim rowNew As DataRow = _dtLineas.NewRow()

                ' --- CORRECCIÓN DEL ERROR ---
                ' Ya no existe ID_Pedido, ahora usamos NumeroPedido
                rowNew("NumeroAlbaran") = _numeroAlbaranActual ' Asignamos el código actual (ej: PED-26-001)
                ' ----------------------------

                rowNew("ID_Linea") = DBNull.Value
                rowNew("NumeroOrden") = rowOrig("NumeroOrden")
                rowNew("ID_Articulo") = rowOrig("ID_Articulo")
                rowNew("Descripcion") = rowOrig("Descripcion")

                Dim ca As Decimal = 0 : Decimal.TryParse(rowOrig("Cantidad").ToString(), ca)
                Dim pr As Decimal = 0 : Decimal.TryParse(rowOrig("PrecioUnitario").ToString(), pr)
                Dim dt As Decimal = 0 : Decimal.TryParse(rowOrig("Descuento").ToString(), dt)

                rowNew("CantidadServida") = ca : rowNew("PrecioUnitario") = pr : rowNew("Descuento") = dt
                rowNew("Total") = (ca * pr) * (1 - (dt / 100))

                _dtLineas.Rows.Add(rowNew)
            Next
            CalcularTotalesGenerales()
            MessageBox.Show("Pedido importado.")

            For Each rowOrig As DataRow In dtOrigen.Rows
                Dim rowNew As DataRow = _dtLineas.NewRow()
                rowNew("NumeroAlbaran") = rowOrig("NumeroOrden")
                rowNew("ID_Linea") = DBNull.Value


            Next
        Catch ex As Exception
            MessageBox.Show("Error al importar: " & ex.Message)
        End Try
    End Sub
    Private Sub EstilizarFecha(dtp As DateTimePicker)
        dtp.Format = DateTimePickerFormat.Custom : dtp.CustomFormat = "dd/MM/yyyy"
        dtp.Font = New Font("Segoe UI", 10) : dtp.MinimumSize = New Size(0, 25)
    End Sub




    Private Sub ReorganizarControlesAutomaticamente()
        If Me.ClientSize.Width < 100 Then Return

        ' 1. Quitar anclajes antiguos
        For Each ctrl As Control In Me.Controls
            ctrl.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        Next

        ' =========================================================
        ' DECLARACIÓN DE TODAS LAS COORDENADAS (¡Arriba del todo!)
        ' =========================================================
        Dim margenIzq As Integer = 30
        Dim anchoForm As Integer = Me.ClientSize.Width
        Dim altoForm As Integer = Me.ClientSize.Height
        Dim yTabla As Integer = 245

        Dim yFila1 As Integer = 30
        Dim yFila2 As Integer = 55
        Dim yFila3 As Integer = 95
        Dim yFila4 As Integer = 120
        Dim yFila5 As Integer = 160 ' Fila para Formas de Pago y Rutas
        Dim yFila6 As Integer = 185

        Dim col1_X As Integer = margenIzq
        Dim col2_X As Integer = 190
        Dim col3_X As Integer = 380

        ' =========================================================
        ' 2. CABECERA IZQUIERDA (Distancias compactas)
        ' =========================================================
        ' Alinear Etiquetas
        For Each ctrl As Control In Me.Controls
            If TypeOf ctrl Is Label AndAlso Not ctrl.Name.Contains("Linea") AndAlso ctrl.Parent Is Me Then
                ctrl.BringToFront()
                ctrl.ForeColor = Color.WhiteSmoke
                ctrl.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)

                Select Case ctrl.Text.Trim().ToLower()
                    Case "factura", "documento", "pedido", "albaran", "albarán" : ctrl.Location = New Point(col1_X, yFila1)
                    Case "cliente" : ctrl.Location = New Point(col2_X, yFila1)
                    Case "fecha" : ctrl.Location = New Point(col3_X, yFila1)
                    Case "fecha entrega" : ctrl.Location = New Point(col3_X + 130, yFila1) ' <--- CORRECCIÓN DE POSICIÓN
                    Case "vendedor" : ctrl.Location = New Point(col1_X, yFila3)
                    Case "estado" : ctrl.Location = New Point(col2_X, yFila3)
                    Case "agencia" : ctrl.Location = New Point(col3_X, yFila3)
                End Select
            End If
        Next
       
        ' Alinear Cajas Fila 1
        If TextBoxAlbaran IsNot Nothing Then TextBoxAlbaran.Bounds = New Rectangle(col1_X, yFila2, 100, 25)
        If Button1 IsNot Nothing Then Button1.Bounds = New Rectangle(col1_X + 105, yFila2, 30, 25)

        If TextBoxIdCliente IsNot Nothing Then TextBoxIdCliente.Bounds = New Rectangle(col2_X, yFila2, 50, 25)
        If TextBoxCliente IsNot Nothing Then TextBoxCliente.Bounds = New Rectangle(col2_X + 55, yFila2, 130, 25)

        ' --- CAJAS DE FECHA ---
        If TextBoxFecha IsNot Nothing Then TextBoxFecha.Bounds = New Rectangle(col3_X, yFila2, 120, 25)

        If DateTimePickerFecha IsNot Nothing Then
            ' Le damos la posición nueva y lo ensanchamos a 120
            DateTimePickerFecha.Bounds = New Rectangle(col3_X + 130, yFila2, 120, 25)
            DateTimePickerFecha.BringToFront()
        End If

        ' Alinear Cajas Fila 2
        If TextBoxIdVendedor IsNot Nothing Then TextBoxIdVendedor.Bounds = New Rectangle(col1_X, yFila4, 40, 25)
        If TextBoxVendedor IsNot Nothing Then TextBoxVendedor.Bounds = New Rectangle(col1_X + 45, yFila4, 110, 25)

        If TextBoxEstado IsNot Nothing Then TextBoxEstado.Bounds = New Rectangle(col2_X, yFila4, 185, 25)
        If cboAgencias IsNot Nothing Then cboAgencias.Bounds = New Rectangle(col3_X, yFila4, 130, 25)

        ' --- COMBOS NUEVOS (Forma de pago y Ruta) ---
        If lblFormaPago IsNot Nothing Then lblFormaPago.Location = New Point(col1_X, yFila5)
        If cboFormaPago IsNot Nothing Then
            cboFormaPago.Bounds = New Rectangle(col1_X, yFila6, 140, 25)
            cboFormaPago.Font = New Font("Segoe UI", 10.5F)
        End If

        If lblRuta IsNot Nothing Then lblRuta.Location = New Point(col2_X, yFila5)
        If cboRuta IsNot Nothing Then
            cboRuta.Bounds = New Rectangle(col2_X, yFila6, 320, 25) ' Ancho de 320 para que no pise la logística
            cboRuta.Font = New Font("Segoe UI", 10.5F)
        End If

        ' =========================================================
        ' 3. CABECERA DERECHA (TabControl Logística)
        ' =========================================================
        Dim anchoLogistica As Integer = 650

        If TabControlModerno2 IsNot Nothing Then
            TabControlModerno2.Visible = True
            TabControlModerno2.ItemSize = New Size(180, 30)
            TabControlModerno2.Bounds = New Rectangle(anchoForm - margenIzq - anchoLogistica, 25, anchoLogistica, 195)
            TabControlModerno2.BringToFront()

            For Each tab As TabPage In TabControlModerno2.TabPages
                tab.BackColor = Me.BackColor
            Next

            Dim MoverInterno = Sub(nombre As String, x As Integer, y As Integer, w As Integer, h As Integer)
                                   Dim encontrados = TabControlModerno2.Controls.Find(nombre, True)
                                   If encontrados.Length > 0 Then
                                       encontrados(0).Bounds = New Rectangle(x, y, w, h)
                                       If TypeOf encontrados(0) Is Label Then
                                           encontrados(0).ForeColor = Color.WhiteSmoke
                                           encontrados(0).BackColor = Color.Transparent
                                           encontrados(0).Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
                                       End If
                                   End If
                               End Sub

            MoverInterno("Label25", 15, 10, 150, 20) : MoverInterno("TextBoxDireccion", 15, 36, 250, 25)
            MoverInterno("Label27", 280, 10, 150, 20) : MoverInterno("TextBoxPoblacion", 280, 36, 210, 25)
            MoverInterno("Label26", 510, 10, 120, 20) : MoverInterno("TextBoxCP", 510, 36, 100, 25)

            MoverInterno("Label31", 15, 85, 200, 20) : MoverInterno("TextBoxTracking", 15, 111, 250, 25)
            MoverInterno("Label28", 280, 85, 100, 20) : MoverInterno("ComboBoxPortes", 280, 111, 130, 25)
            MoverInterno("Label30", 430, 85, 60, 20) : MoverInterno("TextBoxBultos", 430, 111, 70, 25)
            MoverInterno("Label29", 510, 85, 60, 20) : MoverInterno("TextBoxPeso", 510, 111, 100, 25)

            MoverInterno("Label20", 15, 10, 100, 20) : MoverInterno("TextBoxAlbaranOrigen", 15, 36, 130, 25)
            MoverInterno("btnImportarAlbaran", 150, 36, 30, 25)
            MoverInterno("Label21", 200, 10, 150, 20) : MoverInterno("TextBoxObservaciones", 200, 36, 410, 120)

            Dim ocultarDuplicados = Sub(nombre As String)
                                        Dim encontrados = TabControlModerno2.Controls.Find(nombre, True)
                                        If encontrados.Length > 0 Then encontrados(0).Visible = False
                                    End Sub
            ocultarDuplicados("TextBox5") : ocultarDuplicados("Button4") : ocultarDuplicados("Label23")
        End If

        ' =========================================================
        ' 4. LÍNEA DIVISORIA Y TABLA 
        ' =========================================================
        Dim lineaDivisoria As Label = Me.Controls.OfType(Of Label)().FirstOrDefault(Function(l) l.Name = "LineaDivisoria")
        If lineaDivisoria Is Nothing Then
            lineaDivisoria = New Label() With {.Name = "LineaDivisoria", .BackColor = Color.FromArgb(120, 130, 140), .Height = 2}
            Me.Controls.Add(lineaDivisoria)
        End If
        lineaDivisoria.Bounds = New Rectangle(margenIzq, yTabla - 15, anchoForm - (margenIzq * 2), 2)
        lineaDivisoria.BringToFront()

        DataGridView1.Parent = Me
        DataGridView1.MaximumSize = New Size(0, 0)
        DataGridView1.Dock = DockStyle.None
        DataGridView1.Bounds = New Rectangle(margenIzq, yTabla, anchoForm - (margenIzq * 2), altoForm - yTabla - 150)
        DataGridView1.BackgroundColor = Me.BackColor
        DataGridView1.BorderStyle = BorderStyle.None

        If DataGridView1.Columns.Contains("Descripcion") Then DataGridView1.Columns("Descripcion").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        ' =========================================================
        ' 5. TOTALES 
        ' =========================================================
        Dim xDerecha As Integer = anchoForm - margenIzq
        Dim yTotales As Integer = DataGridView1.Bottom + 10
        Dim colorFondoTotales As Color = Color.FromArgb(25, 30, 40)

        Dim panelTotales As Label = Me.Controls.OfType(Of Label)().FirstOrDefault(Function(l) l.Name = "PanelTotalesResumen")
        If panelTotales Is Nothing Then
            panelTotales = New Label() With {.Name = "PanelTotalesResumen", .BackColor = colorFondoTotales}
            Me.Controls.Add(panelTotales)
        End If
        panelTotales.Bounds = New Rectangle(xDerecha - 320, yTotales, 320, 115)
        panelTotales.SendToBack()

        If TextBoxBase IsNot Nothing Then TextBoxBase.Bounds = New Rectangle(xDerecha - 135, yTotales + 10, 120, 25)
        If TextBoxIva IsNot Nothing Then TextBoxIva.Bounds = New Rectangle(xDerecha - 135, yTotales + 40, 120, 25)
        If TextBoxTotalAlb IsNot Nothing Then TextBoxTotalAlb.Bounds = New Rectangle(xDerecha - 135, yTotales + 80, 120, 30)

        If LabelBase IsNot Nothing Then LabelBase.Location = New Point(xDerecha - 310, yTotales + 12)
        If LabelIva IsNot Nothing Then LabelIva.Location = New Point(xDerecha - 310, yTotales + 42)
        If Label7 IsNot Nothing Then Label7.Location = New Point(xDerecha - 310, yTotales + 85)

        If LabelBase IsNot Nothing Then LabelBase.BackColor = colorFondoTotales : LabelBase.ForeColor = Color.White
        If LabelIva IsNot Nothing Then LabelIva.BackColor = colorFondoTotales : LabelIva.ForeColor = Color.White
        If Label7 IsNot Nothing Then Label7.BackColor = colorFondoTotales : Label7.ForeColor = Color.FromArgb(0, 150, 255) : Label7.Font = New Font("Segoe UI", 13, FontStyle.Bold)

        If TextBoxBase IsNot Nothing Then TextBoxBase.BackColor = colorFondoTotales : TextBoxBase.ForeColor = Color.White : TextBoxBase.BorderStyle = BorderStyle.None
        If TextBoxIva IsNot Nothing Then TextBoxIva.BackColor = colorFondoTotales : TextBoxIva.ForeColor = Color.White : TextBoxIva.BorderStyle = BorderStyle.None
        If TextBoxTotalAlb IsNot Nothing Then
            TextBoxTotalAlb.BackColor = colorFondoTotales : TextBoxTotalAlb.ForeColor = Color.FromArgb(0, 150, 255) : TextBoxTotalAlb.BorderStyle = BorderStyle.None : TextBoxTotalAlb.Font = New Font("Segoe UI", 14, FontStyle.Bold)
        End If

        Dim lineaTotal As Label = Me.Controls.OfType(Of Label)().FirstOrDefault(Function(l) l.Name = "LineaTotales")
        If lineaTotal Is Nothing Then
            lineaTotal = New Label() With {.Name = "LineaTotales", .BackColor = Color.FromArgb(100, 100, 100), .Height = 1}
            Me.Controls.Add(lineaTotal)
        End If
        lineaTotal.Bounds = New Rectangle(xDerecha - 300, yTotales + 70, 280, 1)
        lineaTotal.BringToFront()

        ' =========================================================
        ' 6. BOTONES
        ' =========================================================
        Dim yBotones As Integer = DataGridView1.Bottom + 45

        EstilizarBoton(ButtonGuardar, margenIzq, yBotones, Color.FromArgb(0, 120, 215), Color.White)
        EstilizarBoton(ButtonBorrar, margenIzq + 115, yBotones, Color.FromArgb(209, 52, 56), Color.White)
        If ButtonNuevoPed IsNot Nothing Then EstilizarBoton(ButtonNuevoPed, margenIzq + 230, yBotones, Color.FromArgb(0, 120, 215), Color.White)

        EstilizarBoton(ButtonBorrarLineas, margenIzq + 380, yBotones, Color.FromArgb(85, 85, 85), Color.White)
        If ButtonBorrarLineas IsNot Nothing Then ButtonBorrarLineas.Text = "- Quitar Línea" : ButtonBorrarLineas.Width = 110

        EstilizarBoton(ButtonNuevaLinea, margenIzq + 500, yBotones, Color.FromArgb(40, 140, 90), Color.White)
        If ButtonNuevaLinea IsNot Nothing Then ButtonNuevaLinea.Text = "+ Añadir Línea" : ButtonNuevaLinea.Width = 110

        EstilizarBoton(ButtonAnterior, xDerecha - 580, yBotones, Me.BackColor, Color.White)
        EstilizarBoton(ButtonSiguiente, xDerecha - 465, yBotones, Me.BackColor, Color.White)

        Dim lblStock As Label = Me.Controls.OfType(Of Label)().FirstOrDefault(Function(l) l.Name = "LabelStock")
        If lblStock IsNot Nothing Then lblStock.Location = New Point(margenIzq, DataGridView1.Bottom + 10)

        ' =========================================================
        ' 7. RE-APLICAR ANCLAJES 
        ' =========================================================
        DataGridView1.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        If TabControlModerno2 IsNot Nothing Then TabControlModerno2.Anchor = AnchorStyles.Top Or AnchorStyles.Right

        Dim anclajeAbajoIzq As Control() = {ButtonGuardar, ButtonBorrar, ButtonNuevoPed, ButtonBorrarLineas, ButtonNuevaLinea, lblStock}
        For Each c In anclajeAbajoIzq
            If c IsNot Nothing Then c.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
        Next

        Dim anclajeAbajoDer As Control() = {ButtonAnterior, ButtonSiguiente, TextBoxBase, TextBoxIva, TextBoxTotalAlb, LabelBase, LabelIva, Label7, panelTotales, lineaTotal}
        For Each c In anclajeAbajoDer
            If c IsNot Nothing Then c.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        Next
    End Sub

    ' Le ponemos "System.Windows.Forms.Button" para que Visual Studio no se confunda
    Private Sub EstilizarBoton(btn As System.Windows.Forms.Button, x As Integer, y As Integer, bg As Color, fg As Color)
        If btn IsNot Nothing Then
            btn.Location = New Point(x, y)
            btn.BackColor = bg : btn.ForeColor = fg
            btn.FlatStyle = FlatStyle.Flat : btn.FlatAppearance.BorderSize = 0
            btn.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
            btn.Cursor = Cursors.Hand : btn.Height = 30
        End If
    End Sub




    ' =========================================================
    ' STOCK EN TIEMPO REAL (Adaptado a CantidadServida)
    ' =========================================================
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        Dim lblStock As Label = Me.Controls.OfType(Of Label)().FirstOrDefault(Function(l) l.Name = "LabelStock")
        If lblStock Is Nothing Then Return

        If DataGridView1.CurrentRow Is Nothing OrElse DataGridView1.CurrentRow.IsNewRow Then
            lblStock.Text = "Stock disponible: -"
            lblStock.ForeColor = Color.FromArgb(170, 175, 180)
            Return
        End If

        Try
            Dim idCelda As Object = DataGridView1.CurrentRow.Cells("ID_Articulo").Value

            If idCelda IsNot Nothing AndAlso Not DBNull.Value.Equals(idCelda) Then
                Dim idArticulo As Integer = Convert.ToInt32(idCelda)

                Dim cantidadEnLinea As Decimal = 0
                ' En Albaranes leemos CantidadServida
                Dim cantCelda As Object = DataGridView1.CurrentRow.Cells("CantidadServida").Value
                If cantCelda IsNot Nothing AndAlso Not DBNull.Value.Equals(cantCelda) Then
                    cantidadEnLinea = Convert.ToDecimal(cantCelda)
                End If

                Dim stockActual As Decimal = ConsultarStock(idArticulo)
                Dim stockRestante As Decimal = stockActual - cantidadEnLinea

                If stockRestante > 0 Then
                    lblStock.Text = $"Stock actual: {stockActual} (Te quedarán: {stockRestante})"
                    lblStock.ForeColor = Color.FromArgb(40, 180, 90)
                ElseIf stockRestante = 0 Then
                    lblStock.Text = $"Stock actual: {stockActual} (¡Atención! Te quedarás a 0)"
                    lblStock.ForeColor = Color.FromArgb(220, 160, 40)
                Else
                    lblStock.Text = $"¡Stock insuficiente! Actual: {stockActual} (Te faltan: {Math.Abs(stockRestante)})"
                    lblStock.ForeColor = Color.FromArgb(255, 80, 80)
                End If

                lblStock.Font = New Font("Segoe UI", 10.5F, FontStyle.Bold)
            End If
        Catch ex As Exception
            lblStock.Text = "Stock disponible: -"
        End Try
    End Sub

    Private Function ConsultarStock(idArticulo As Integer) As Decimal
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()
            Dim sql As String = "SELECT StockActual FROM Articulos WHERE ID_Articulo = @id"
            Using cmd As New SQLiteCommand(sql, c)
                cmd.Parameters.AddWithValue("@id", idArticulo)
                Dim res = cmd.ExecuteScalar()
                If res IsNot Nothing AndAlso Not DBNull.Value.Equals(res) Then
                    Return Convert.ToDecimal(res)
                End If
            End Using
        Catch
        End Try
        Return 0
    End Function
    Protected Overrides ReadOnly Property CreateParams As System.Windows.Forms.CreateParams
        Get
            ' Forzamos el uso del tipo especifico de windows forms
            Dim cp As System.Windows.Forms.CreateParams = MyBase.CreateParams
            'Aplicamos el estilo extendido para evitar parpadeos (WS_EX_COMPOSITED)
            cp.ExStyle = cp.ExStyle Or &H2000000
            Return cp
        End Get
    End Property
    ' =========================================================
    ' MOTOR DE STOCK Y MOVIMIENTOS
    ' =========================================================
    Private Sub ActualizarStockYMovimiento(c As SQLiteConnection, trans As SQLiteTransaction, albaran As String, idArticulo As Integer, variacionSalida As Decimal, fecha As DateTime)
        ' Si la variación es 0, no hay nada que hacer
        If variacionSalida = 0 Then Return

        ' 1. Consultar el stock actual del artículo
        Dim stockActual As Decimal = 0
        Using cmdStock As New SQLiteCommand("SELECT StockActual FROM Articulos WHERE ID_Articulo = @id", c, trans)
            cmdStock.Parameters.AddWithValue("@id", idArticulo)
            Dim res = cmdStock.ExecuteScalar()
            If res IsNot Nothing AndAlso Not IsDBNull(res) Then stockActual = Convert.ToDecimal(res)
        End Using

        ' 2. Calcular el nuevo stock (Si variacionSalida es positivo, resta. Si es negativo, suma)
        Dim nuevoStock As Decimal = stockActual - variacionSalida

        ' 3. Actualizar el stock en la tabla Articulos
        Using cmdUpd As New SQLiteCommand("UPDATE Articulos SET StockActual = @nuevo WHERE ID_Articulo = @id", c, trans)
            cmdUpd.Parameters.AddWithValue("@nuevo", nuevoStock)
            cmdUpd.Parameters.AddWithValue("@id", idArticulo)
            cmdUpd.ExecuteNonQuery()
        End Using

        ' 4. Registrar en MovimientosAlmacen
        Dim tipoMov As String = If(variacionSalida > 0, "SALIDA", "ENTRADA")
        Dim cantidadMov As Decimal = Math.Abs(variacionSalida)

        Dim sqlMov As String = "INSERT INTO MovimientosAlmacen (Fecha, ID_Articulo, TipoMovimiento, Cantidad, StockResultante, DocumentoReferencia, ID_Usuario) " &
                               "VALUES (@fecha, @idArt, @tipo, @cant, @stockRes, @doc, 1)" ' Ponemos 1 por defecto al ID_Usuario de momento
        Using cmdMov As New SQLiteCommand(sqlMov, c, trans)
            cmdMov.Parameters.AddWithValue("@fecha", fecha.ToString("yyyy-MM-dd HH:mm:ss"))
            cmdMov.Parameters.AddWithValue("@idArt", idArticulo)
            cmdMov.Parameters.AddWithValue("@tipo", tipoMov)
            cmdMov.Parameters.AddWithValue("@cant", cantidadMov)
            cmdMov.Parameters.AddWithValue("@stockRes", nuevoStock)
            cmdMov.Parameters.AddWithValue("@doc", albaran)
            cmdMov.ExecuteNonQuery()
        End Using
    End Sub
End Class
Public Module GeneradorUtilidades

    ''' <summary>
    ''' Genera un código de seguimiento: 7 letras (A-Z, a-z) y 4 números (0-9)
    ''' </summary>
    Public Function GenerarCodigoSeguimiento() As String
        Dim letras As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
        Dim numeros As String = "0123456789"
        Dim rand As New Random()

        ' Usamos StringBuilder para cumplir con las mejores prácticas de gestión de memoria 
        Dim resultado As New StringBuilder(11)

        ' 1. Generar 7 letras aleatorias
        For i As Integer = 1 To 7
            Dim indice As Integer = rand.Next(0, letras.Length)
            resultado.Append(letras(indice))
        Next

        ' 2. Generar 4 números aleatorios
        For i As Integer = 1 To 4
            Dim indice As Integer = rand.Next(0, numeros.Length)
            resultado.Append(numeros(indice))
        Next

        Return resultado.ToString()
    End Function
    Public Function ObtenerAgencias() As List(Of Agencia)
        Dim lista As New List(Of Agencia)()
        Dim sql As String = "SELECT ID_Agencia, Nombre FROM Agencias ORDER BY Nombre ASC"

        ' Asegúrate de que la ruta a tu DB sea correcta
        Dim conn As SQLiteConnection = ConexionBD.GetConnection


        Try

            Using cmd As New SQLiteCommand(sql, conn)
                Using reader As SQLiteDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim ag As New Agencia()
                        ag.ID_Agencia = Convert.ToInt32(reader("ID_Agencia"))
                        ag.Nombre = reader("Nombre").ToString()
                        lista.Add(ag)
                    End While
                End Using
            End Using
        Catch ex As Exception
            MsgBox("Error al cargar agencias: " & ex.Message)
        End Try

        Return lista
    End Function

End Module
Public Class Agencia
    Public Property ID_Agencia As Integer
    Public Property Nombre As String

    ' Esto permite que el ComboBox sepa qué texto mostrar por defecto
    Public Overrides Function ToString() As String
        Return Nombre
    End Function
End Class