Imports System.Data.SQLite
Imports System.Text
Imports System.Security.Cryptography
Imports System.Runtime.InteropServices
Public Class FrmFacturas
    Private _numeroFacturaActual As String = ""
    Private _dtLineas As DataTable
    Private _idsParaBorrar As New List(Of Integer)
    Private WithEvents cboFormaPago As New System.Windows.Forms.ComboBox()
    Private WithEvents cboRuta As New System.Windows.Forms.ComboBox()
    Private lblFormaPago As New Label() With {.Text = "Forma de Pago", .AutoSize = True, .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold), .ForeColor = Color.WhiteSmoke}
    Private lblRuta As New Label() With {.Text = "Ruta Asignada", .AutoSize = True, .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold), .ForeColor = Color.WhiteSmoke}
    Private Sub FrmFacturas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FrmPresupuestos.EstilizarGrid(DataGridView1)
        EstilizarFecha(DateTimePickerFecha)
        If ComboBoxPortes.Items.Count = 0 Then
            ComboBoxPortes.Items.Add("Pagados")
            ComboBoxPortes.Items.Add("Debidos")
            ComboBoxPortes.SelectedIndex = 0
        End If
        ConfigurarGrid()
        Dim listaAgencias As List(Of Agencia) = ObtenerAgencias()
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
        Dim ultimoNum As String = ObtenerUltimoNumeroFactura()
        If Not String.IsNullOrEmpty(ultimoNum) Then
            CargarFactura(ultimoNum)
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
        If Not _dtLineas.Columns.Contains("NumeroFactura") Then _dtLineas.Columns.Add("NumeroFactura", GetType(String))

        ' Datos de línea
        If Not _dtLineas.Columns.Contains("NumeroOrden") Then _dtLineas.Columns.Add("NumeroOrden", GetType(Integer))
        If Not _dtLineas.Columns.Contains("ID_Articulo") Then _dtLineas.Columns.Add("ID_Articulo", GetType(Object))
        If Not _dtLineas.Columns.Contains("Descripcion") Then _dtLineas.Columns.Add("Descripcion", GetType(String))

        ' Valores numéricos
        If Not _dtLineas.Columns.Contains("Cantidad") Then _dtLineas.Columns.Add("Cantidad", GetType(Decimal)) ' Ojo: CantidadServida
        If Not _dtLineas.Columns.Contains("PrecioUnitario") Then _dtLineas.Columns.Add("PrecioUnitario", GetType(Decimal))
        If Not _dtLineas.Columns.Contains("Descuento") Then _dtLineas.Columns.Add("Descuento", GetType(Decimal))
        If Not _dtLineas.Columns.Contains("PorcentajeIVA") Then _dtLineas.Columns.Add("PorcentajeIVA", GetType(Decimal))
        If Not _dtLineas.Columns.Contains("Total") Then _dtLineas.Columns.Add("Total", GetType(Decimal))

        ' Trazabilidad con Pedido (Opcional si lo usas)
        If Not _dtLineas.Columns.Contains("ID_LineaPedido") Then _dtLineas.Columns.Add("ID_LineaPedido", GetType(Object))
        ' Nuevas columnas calculadas para el IVA
        If Not _dtLineas.Columns.Contains("PrecioConIVA") Then _dtLineas.Columns.Add("PrecioConIVA", GetType(Decimal))
        If Not _dtLineas.Columns.Contains("TotalConIVA") Then _dtLineas.Columns.Add("TotalConIVA", GetType(Decimal))
    End Sub
    Private Sub ConfigurarGrid()
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.Columns.Clear()

        ' --- Columna técnica oculta ---
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "ID_Linea", .DataPropertyName = "ID_Linea", .Visible = False})

        ' --- 1. Identificación ---
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "NumeroOrden", .DataPropertyName = "NumeroOrden", .HeaderText = "Nº", .Width = 40, .ReadOnly = True, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleCenter}})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "ID_Articulo", .DataPropertyName = "ID_Articulo", .HeaderText = "ID Art", .Width = 60})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "Descripcion", .DataPropertyName = "Descripcion", .HeaderText = "Descripción", .Width = 250})

        ' --- 2. Cantidad y Precio ---
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "Cantidad", .DataPropertyName = "Cantidad", .HeaderText = "Cantidad", .Width = 75, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "N2", .BackColor = Color.Ivory}})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "PrecioUnitario", .DataPropertyName = "PrecioUnitario", .HeaderText = "Precio Base", .Width = 85, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "C2"}})

        ' --- 3. Modificadores (% Dto y % IVA) ---
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "Descuento", .DataPropertyName = "Descuento", .HeaderText = "% Dto", .Width = 60, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "N2"}})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "PorcentajeIVA", .DataPropertyName = "PorcentajeIVA", .HeaderText = "% IVA", .Width = 60, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "N0"}})

        ' --- 4. Totales de la línea ---
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "Total", .DataPropertyName = "Total", .HeaderText = "Total Base", .ReadOnly = True, .Width = 90, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "C2", .BackColor = Color.WhiteSmoke}})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "TotalConIVA", .DataPropertyName = "TotalConIVA", .HeaderText = "Total (+IVA)", .ReadOnly = True, .Width = 100, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "C2", .BackColor = Color.FromArgb(230, 240, 250)}})
    End Sub

    Private Function ObtenerUltimoNumeroFactura() As String
        Dim ultimoNumero As String = ""
        Try
            Dim conexion = ConexionBD.GetConnection()
            Dim sql As String = "SELECT MAX(NumeroFactura) FROM FacturasVenta"
            Using cmd As New SQLiteCommand(sql, conexion)
                If conexion.State <> ConnectionState.Open Then conexion.Open()

                Dim resultado = cmd.ExecuteScalar()

                If resultado IsNot Nothing AndAlso Not IsDBNull(resultado) Then
                    ultimoNumero = resultado.ToString()
                End If
            End Using
        Catch ex As Exception

        End Try
        Return ultimoNumero
    End Function
    Private Sub ImportarDatosPedido(numeroAlbaran As String)
        Dim c = ConexionBD.GetConnection()
        Try
            If c.State <> ConnectionState.Open Then c.Open()
            Dim sqlCab As String = "SELECT A.*, C.NombreFiscal, C.Direccion, C.Poblacion, C.CodigoPostal " &
                                   "FROM Albaranes A " &
                                   "LEFT JOIN Clientes C ON A.ID_Cliente = C.ID_Cliente " &
                                   "LEFT JOIN Vendedores V ON A.ID_Vendedor = V.ID_Vendedor " &
                                   "WHERE A.NumeroAlbaran = @num"
            Using cmd As New SQLiteCommand(sqlCab, c)
                cmd.Parameters.AddWithValue("@num", numeroAlbaran)
                Using r = cmd.ExecuteReader()
                    If r.Read() Then
                        LimpiarFormulario()

                        ' Referencia
                        TextBoxAlbaranOrigen.Text = numeroAlbaran
                        TextBoxAlbaranOrigen.Tag = numeroAlbaran ' Guardamos para BD

                        ' Cliente
                        TextBoxIdCliente.Text = r("ID_Cliente").ToString()
                        TextBoxCliente.Text = r("NombreFiscal").ToString()
                        TextBoxIdVendedor.Text = If(IsDBNull(r("ID_Vendedor")), "", r("ID_Vendedor").ToString())
                        TextBoxVendedor.Text = If(IsDBNull(r("NombreVend")), "", r("NombreVend").ToString())

                        ' Dirección de Envío (Se congela aquí)
                        TextBoxDireccion.Text = If(IsDBNull(r("Direccion")), "", r("Direccion").ToString())
                        TextBoxPoblacion.Text = If(IsDBNull(r("Poblacion")), "", r("Poblacion").ToString())
                        TextBoxCP.Text = If(IsDBNull(r("CodigoPostal")), "", r("CodigoPostal").ToString())

                        TextBoxObservaciones.Text = "Generado desde Pedido " & numeroAlbaran
                    End If
                End Using
            End Using

            Dim sqlLin As String = "SELECT * FROM LineasAlbaran Where NumeroAlbaran = @num order by NumeroOrden ASC"
            Dim dtOrigen As New DataTable()
            Using cmd As New SQLiteCommand(sqlLin, c)
                cmd.Parameters.AddWithValue("@num", numeroAlbaran)
                Dim da As New SQLiteDataAdapter(cmd)
                da.Fill(dtOrigen)
            End Using
            _dtLineas.Rows.Clear()
            For Each rowOrig As DataRow In dtOrigen.Rows
                Dim rowNew As DataRow = _dtLineas.NewRow()

                rowNew("NumeroFactura") = _numeroFacturaActual
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
    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles ButtonGuardar.Click
        If String.IsNullOrWhiteSpace(TextBoxIdCliente.Text) Then MessageBox.Show("Falta el Cliente") : Return

        Dim esNuevo As Boolean = String.IsNullOrEmpty(_numeroFacturaActual)
        If esNuevo Then
            TextBoxFactura.Text = GenerarProximoNumeroFactura()
            _numeroFacturaActual = TextBoxFactura.Text
        End If

        Dim c = ConexionBD.GetConnection()
        Dim trans As SQLiteTransaction = Nothing

        Try
            If c.State <> ConnectionState.Open Then c.Open()
            trans = c.BeginTransaction()

            ' ==========================================
            ' A) Calcular Totales
            ' ==========================================
            Dim sumaBase As Decimal = 0
            Dim sumaIva As Decimal = 0

            For Each r As DataRow In _dtLineas.Rows
                If r.RowState <> DataRowState.Deleted Then
                    ' IMPORTANTE: Leemos de la columna "Cantidad" (no CantidadServida)
                    Dim cant As Decimal = 0 : Decimal.TryParse(r("Cantidad").ToString(), cant)
                    Dim prec As Decimal = 0 : Decimal.TryParse(r("PrecioUnitario").ToString(), prec)
                    Dim desc As Decimal = 0 : Decimal.TryParse(r("Descuento").ToString(), desc)
                    Dim ivaLinea As Decimal = 21 : Decimal.TryParse(r("PorcentajeIVA").ToString(), ivaLinea)

                    Dim totLinea As Decimal = (cant * prec) * (1 - (desc / 100))

                    If _dtLineas.Columns.Contains("Total") Then r("Total") = totLinea

                    sumaBase += totLinea
                    sumaIva += totLinea * (ivaLinea / 100)
                End If
            Next

            Dim sumaTotal As Decimal = sumaBase + sumaIva

            ' ==========================================
            ' B) Guardar Cabecera (FacturasVenta)
            ' ==========================================
            Dim sql As String
            If esNuevo Then
                sql = "INSERT INTO FacturasVenta (NumeroFactura, Fecha, CodigoCliente, NumeroAlbaran, BaseImponible, ImporteIVA, TotalFactura, Cobrada, ID_FormaPago, ID_Vendedor, FechaVencimiento, NombreFiscal, CIF, Direccion, Observaciones, ID_Ruta, ID_Agencia, NumeroBultos, PesoTotal, CodigoSeguimiento, Portes, Poblacion, CodigoPostal, Estado) " &
                      "VALUES (@num, @fecha, @cli, @alb, @base, @iva, @total, @cobrada, @fpago, @vend, @fvenc, @fiscal, @cif, @dir, @obs, @ruta, @agencia, @bultos, @peso, @track, @portes, @pob, @cp, @estado)"
            Else
                sql = "UPDATE FacturasVenta SET Fecha=@fecha, CodigoCliente=@cli, NumeroAlbaran=@alb, BaseImponible=@base, ImporteIVA=@iva, TotalFactura=@total, Cobrada=@cobrada, ID_FormaPago=@fpago, ID_Vendedor=@vend, FechaVencimiento=@fvenc, NombreFiscal=@fiscal, CIF=@cif, Direccion=@dir, Observaciones=@obs, ID_Ruta=@ruta, ID_Agencia=@agencia, NumeroBultos=@bultos, PesoTotal=@peso, CodigoSeguimiento=@track, Portes=@portes, Poblacion=@pob, CodigoPostal=@cp, Estado=@estado " &
                      "WHERE NumeroFactura=@num"
            End If

            Dim fechaFactura As DateTime = DateTime.Now
            If DateTime.TryParse(TextBoxFecha.Text, fechaFactura) = False Then fechaFactura = DateTime.Now

            Using cmd As New SQLiteCommand(sql, c)
                cmd.Transaction = trans

                cmd.Parameters.AddWithValue("@num", _numeroFacturaActual)
                cmd.Parameters.AddWithValue("@fecha", fechaFactura)
                cmd.Parameters.AddWithValue("@cli", TextBoxIdCliente.Text)

                Dim alb As Object = If(TextBoxAlbaranOrigen.Tag IsNot Nothing, TextBoxAlbaranOrigen.Tag, DBNull.Value)
                cmd.Parameters.AddWithValue("@alb", alb)

                cmd.Parameters.AddWithValue("@base", sumaBase)
                cmd.Parameters.AddWithValue("@iva", sumaIva)
                cmd.Parameters.AddWithValue("@total", sumaTotal)

                cmd.Parameters.AddWithValue("@cobrada", 0) ' Por defecto 0 (Falso)
                cmd.Parameters.AddWithValue("@fpago", If(cboFormaPago.SelectedIndex <> -1, cboFormaPago.SelectedValue, DBNull.Value))
                cmd.Parameters.AddWithValue("@fvenc", DateTimePickerFecha.Value)

                cmd.Parameters.AddWithValue("@fiscal", TextBoxCliente.Text)
                cmd.Parameters.AddWithValue("@cif", DBNull.Value) ' Lo ideal es cargarlo desde el cliente
                cmd.Parameters.AddWithValue("@dir", TextBoxDireccion.Text)
                cmd.Parameters.AddWithValue("@pob", TextBoxPoblacion.Text)
                cmd.Parameters.AddWithValue("@cp", TextBoxCP.Text)

                cmd.Parameters.AddWithValue("@vend", If(String.IsNullOrWhiteSpace(TextBoxIdVendedor.Text), DBNull.Value, TextBoxIdVendedor.Text))
                cmd.Parameters.AddWithValue("@ruta", If(cboRuta.SelectedIndex <> -1, cboRuta.SelectedValue, DBNull.Value))
                cmd.Parameters.AddWithValue("@agencia", If(cboAgencias.SelectedIndex <> -1, cboAgencias.SelectedValue, DBNull.Value))

                cmd.Parameters.AddWithValue("@bultos", If(IsNumeric(TextBoxBultos.Text), CInt(TextBoxBultos.Text), 1))
                cmd.Parameters.AddWithValue("@peso", If(IsNumeric(TextBoxPeso.Text), CDbl(TextBoxPeso.Text), 0))
                cmd.Parameters.AddWithValue("@track", TextBoxTracking.Text)
                cmd.Parameters.AddWithValue("@portes", ComboBoxPortes.Text)
                cmd.Parameters.AddWithValue("@estado", TextBoxEstado.Text)
                cmd.Parameters.AddWithValue("@obs", TextBoxObservaciones.Text)

                cmd.ExecuteNonQuery()
            End Using

            ' ==========================================
            ' C) Bajas de Líneas (Devolver Stock)
            ' ==========================================
            For Each idDel In _idsParaBorrar
                Dim idArtDel As Integer = 0
                Dim cantDel As Decimal = 0

                Using cmdInfo As New SQLiteCommand("SELECT ID_Articulo, Cantidad FROM LineasFacturaVenta WHERE ID_Linea = @id", c, trans)
                    cmdInfo.Parameters.AddWithValue("@id", idDel)
                    Using reader = cmdInfo.ExecuteReader()
                        If reader.Read() Then
                            If Not IsDBNull(reader("ID_Articulo")) Then idArtDel = Convert.ToInt32(reader("ID_Articulo"))
                            cantDel = Convert.ToDecimal(reader("Cantidad"))
                        End If
                    End Using
                End Using

                Using cmdDel As New SQLiteCommand("DELETE FROM LineasFacturaVenta WHERE ID_Linea = @id", c, trans)
                    cmdDel.Parameters.AddWithValue("@id", idDel)
                    cmdDel.ExecuteNonQuery()
                End Using

            Next
            _idsParaBorrar.Clear()

            ' ==========================================
            ' D) Inserciones y Actualizaciones de Líneas
            ' ==========================================
            If _dtLineas IsNot Nothing Then
                For Each r As DataRow In _dtLineas.Rows
                    If r.RowState = DataRowState.Deleted Then Continue For

                    Dim cant As Decimal = 0 : Decimal.TryParse(r("Cantidad").ToString(), cant)
                    Dim prec As Decimal = 0 : Decimal.TryParse(r("PrecioUnitario").ToString(), prec)
                    Dim desc As Decimal = 0 : Decimal.TryParse(r("Descuento").ToString(), desc)
                    Dim ivaPorc As Decimal = 21 : Decimal.TryParse(r("PorcentajeIVA").ToString(), ivaPorc)
                    Dim tot As Decimal = (cant * prec) * (1 - (desc / 100))

                    ' IMPORTANTE: Se llama ID_Linea en BD, pero asegurate de que en tu DataTable local (ConfigurarEstructuraDatos) también se llame así.
                    Dim idLin = r("ID_Linea")
                    Dim idArt As Object = If(IsNumeric(r("ID_Articulo")) AndAlso Val(r("ID_Articulo")) > 0, r("ID_Articulo"), DBNull.Value)

                    If IsDBNull(idLin) OrElse Not IsNumeric(idLin) Then
                        ' INSERT
                        Dim sqlIns As String = "INSERT INTO LineasFacturaVenta (NumeroFactura, NumeroOrden, ID_Articulo, Descripcion, Cantidad, PrecioUnitario, Descuento, PorcentajeIVA, Total) " &
                                               "VALUES (@idFac, @ord, @art, @desc, @cant, @prec, @dcto, @iva, @tot)"

                        Using cmdL As New SQLiteCommand(sqlIns, c, trans)
                            cmdL.Parameters.AddWithValue("@idFac", _numeroFacturaActual)
                            cmdL.Parameters.AddWithValue("@ord", r("NumeroOrden"))
                            cmdL.Parameters.AddWithValue("@art", idArt)
                            cmdL.Parameters.AddWithValue("@desc", r("Descripcion"))
                            cmdL.Parameters.AddWithValue("@cant", CDbl(cant))
                            cmdL.Parameters.AddWithValue("@prec", CDbl(prec))
                            cmdL.Parameters.AddWithValue("@dcto", CDbl(desc))
                            cmdL.Parameters.AddWithValue("@iva", CDbl(ivaPorc))
                            cmdL.Parameters.AddWithValue("@tot", CDbl(tot))
                            cmdL.ExecuteNonQuery()
                        End Using


                    ElseIf r.RowState = DataRowState.Modified Then
                        ' UPDATE
                        Dim cantAnterior As Decimal = 0
                        If r.HasVersion(DataRowVersion.Original) AndAlso Not IsDBNull(r("Cantidad", DataRowVersion.Original)) Then
                            Decimal.TryParse(r("Cantidad", DataRowVersion.Original).ToString(), cantAnterior)
                        End If
                        Dim variacion = cant - cantAnterior

                        Dim sqlUpd As String = "UPDATE LineasFacturaVenta SET " &
                                               "ID_Articulo=@art, Descripcion=@desc, Cantidad=@cant, PrecioUnitario=@prec, Descuento=@dcto, PorcentajeIVA=@iva, Total=@tot " &
                                               "WHERE ID_Linea=@id"

                        Using cmdL As New SQLiteCommand(sqlUpd, c, trans)
                            cmdL.Parameters.AddWithValue("@art", idArt)
                            cmdL.Parameters.AddWithValue("@desc", r("Descripcion"))
                            cmdL.Parameters.AddWithValue("@cant", CDbl(cant))
                            cmdL.Parameters.AddWithValue("@prec", CDbl(prec))
                            cmdL.Parameters.AddWithValue("@dcto", CDbl(desc))
                            cmdL.Parameters.AddWithValue("@iva", CDbl(ivaPorc))
                            cmdL.Parameters.AddWithValue("@tot", CDbl(tot))
                            cmdL.Parameters.AddWithValue("@id", idLin)
                            cmdL.ExecuteNonQuery()
                        End Using

                        ' Ajuste de almacén

                    End If
                Next
            End If

            trans.Commit()
            MessageBox.Show("Factura Guardada y Registrada correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
            CargarFactura(_numeroFacturaActual)

        Catch ex As Exception
            If trans IsNot Nothing Then trans.Rollback()
            MessageBox.Show("Error crítico al guardar la factura: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub CargarFactura(num As String)
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()

            Dim sql As String = "SELECT F.*, C.NombreFiscal AS NombreCliBD, V.Nombre AS NombreVendedor " &
                                "FROM FacturasVenta F " &
                                "LEFT JOIN Clientes C ON F.CodigoCliente = C.CodigoCliente " &
                                "LEFT JOIN Vendedores V ON F.ID_Vendedor = V.ID_Vendedor " &
                                "WHERE F.NumeroFactura= @num"

            Using cmd As New SQLiteCommand(sql, c)
                cmd.Parameters.AddWithValue("@num", num)

                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        _numeroFacturaActual = num
                        TextBoxFactura.Text = reader("NumeroFactura").ToString()

                        ' --- BLINDAJE TOTAL DE FECHAS (Sin Convert.ToDateTime) ---
                        If Not IsDBNull(reader("Fecha")) AndAlso reader("Fecha").ToString().Trim() <> "" Then
                            Dim dF As DateTime
                            If DateTime.TryParse(reader("Fecha").ToString(), dF) Then
                                TextBoxFecha.Text = dF.ToShortDateString()
                            End If
                        End If

                        If Not IsDBNull(reader("FechaVencimiento")) AndAlso reader("FechaVencimiento").ToString().Trim() <> "" Then
                            Dim dV As DateTime
                            If DateTime.TryParse(reader("FechaVencimiento").ToString(), dV) Then
                                DateTimePickerFecha.Value = dV
                            End If
                        End If
                        ' ---------------------------------------------------------

                        ' Textos Simples
                        TextBoxEstado.Text = If(IsDBNull(reader("Estado")), "Emitida", reader("Estado").ToString())
                        TextBoxObservaciones.Text = If(IsDBNull(reader("Observaciones")), "", reader("Observaciones").ToString())

                        ' Cliente y Dirección
                        TextBoxIdCliente.Text = reader("CodigoCliente").ToString()
                        Dim nombreCongelado As String = If(IsDBNull(reader("NombreFiscal")), "", reader("NombreFiscal").ToString())
                        Dim nombreActualBD As String = If(IsDBNull(reader("NombreCliBD")), "", reader("NombreCliBD").ToString())

                        TextBoxCliente.Text = If(nombreCongelado <> "", nombreCongelado, nombreActualBD)
                        TextBoxDireccion.Text = reader("Direccion").ToString()
                        TextBoxPoblacion.Text = reader("Poblacion").ToString()
                        TextBoxCP.Text = reader("CodigoPostal").ToString()

                        ' Vendedor
                        TextBoxIdVendedor.Text = If(IsDBNull(reader("ID_Vendedor")), "", reader("ID_Vendedor").ToString())
                        TextBoxVendedor.Text = If(IsDBNull(reader("NombreVendedor")), "", reader("NombreVendedor").ToString())

                        ' Combos (Desplegables)
                        If Not IsDBNull(reader("ID_Agencia")) Then cboAgencias.SelectedValue = Convert.ToInt32(reader("ID_Agencia")) Else cboAgencias.SelectedIndex = -1
                        If Not IsDBNull(reader("ID_FormaPago")) Then cboFormaPago.SelectedValue = Convert.ToInt32(reader("ID_FormaPago")) Else cboFormaPago.SelectedIndex = -1
                        If Not IsDBNull(reader("ID_Ruta")) Then cboRuta.SelectedValue = Convert.ToInt32(reader("ID_Ruta")) Else cboRuta.SelectedIndex = -1

                        If Not IsDBNull(reader("NumeroAlbaran")) AndAlso reader("NumeroAlbaran").ToString() <> "" Then
                            TextBoxAlbaranOrigen.Text = reader("NumeroAlbaran").ToString()
                            TextBoxAlbaranOrigen.Tag = reader("NumeroAlbaran").ToString()
                        Else
                            TextBoxAlbaranOrigen.Text = ""
                            TextBoxAlbaranOrigen.Tag = Nothing
                        End If

                        TextBoxBultos.Text = reader("NumeroBultos").ToString()
                        TextBoxPeso.Text = reader("PesoTotal").ToString()
                        TextBoxTracking.Text = reader("CodigoSeguimiento").ToString()
                        ComboBoxPortes.Text = reader("Portes").ToString()
                    Else
                        MessageBox.Show("Factura no encontrada.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If
                End Using
            End Using

            ' Cargamos las líneas de la factura
            CargarLineas()

        Catch ex As Exception
            MessageBox.Show("Error al cargar la factura: " & ex.Message, "Error Crítico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub CargarLineas()
        Try
            Dim c = ConexionBD.GetConnection()

            Using cmd As New SQLiteCommand("SELECT * FROM LineasFacturaVenta WHERE NumeroFactura = @num ORDER BY NumeroOrden ASC", c)
                cmd.Parameters.AddWithValue("@num", _numeroFacturaActual)

                Dim da As New SQLiteDataAdapter(cmd)
                _dtLineas = New DataTable()
                da.Fill(_dtLineas)

                ' Esto añade las columnas virtuales (PrecioConIVA y TotalConIVA) vacías
                ConfigurarEstructuraDatos()

                ' ==============================================================
                ' MAGIA: Calculamos el IVA de las facturas viejas al cargarlas
                ' ==============================================================
                For Each row As DataRow In _dtLineas.Rows
                    ' Evitamos errores si alguna celda antigua está nula en la BD
                    Dim cant As Decimal = If(IsDBNull(row("Cantidad")), 0, Convert.ToDecimal(row("Cantidad")))
                    Dim prec As Decimal = If(IsDBNull(row("PrecioUnitario")), 0, Convert.ToDecimal(row("PrecioUnitario")))
                    Dim dto As Decimal = If(IsDBNull(row("Descuento")), 0, Convert.ToDecimal(row("Descuento")))
                    Dim iva As Decimal = If(IsDBNull(row("PorcentajeIVA")), 21, Convert.ToDecimal(row("PorcentajeIVA")))

                    ' Aseguramos que el % de IVA se muestre (por si en la BD era nulo)
                    row("PorcentajeIVA") = iva

                    ' Calculamos los precios virtuales
                    row("PrecioConIVA") = prec * (1 + (iva / 100))

                    Dim totalSinIva As Decimal = (cant * prec) * (1 - (dto / 100))
                    row("Total") = totalSinIva ' Refrescamos por si acaso
                    row("TotalConIVA") = totalSinIva * (1 + (iva / 100))
                Next
                ' ==============================================================

                DataGridView1.DataSource = _dtLineas
            End Using

            CalcularTotalesGenerales()

            ' Ocultamos las columnas técnicas (oculto las dos por si tienes un nombre u otro)
            If DataGridView1.Columns.Contains("ID_LineaFactura") Then DataGridView1.Columns("ID_LineaFactura").Visible = False
            If DataGridView1.Columns.Contains("ID_Linea") Then DataGridView1.Columns("ID_Linea").Visible = False

        Catch ex As Exception
            MessageBox.Show("Error al cargar líneas: " & ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub
    Private Sub CalcularTotalesGenerales()
        Dim base As Decimal = 0
        Dim sumaIva As Decimal = 0

        If _dtLineas IsNot Nothing Then
            For Each row As DataRow In _dtLineas.Rows
                If row.RowState <> DataRowState.Deleted Then
                    Dim t As Decimal = 0
                    Dim ivaLinea As Decimal = 21 ' Por defecto 21 por si acaso

                    ' Leemos el total de la línea y su IVA
                    If _dtLineas.Columns.Contains("Total") Then Decimal.TryParse(row("Total").ToString(), t)
                    If _dtLineas.Columns.Contains("PorcentajeIVA") AndAlso Not IsDBNull(row("PorcentajeIVA")) Then
                        Decimal.TryParse(row("PorcentajeIVA").ToString(), ivaLinea)
                    End If

                    ' Sumamos a los montones generales
                    base += t
                    sumaIva += t * (ivaLinea / 100)
                End If
            Next
        End If

        ' Actualizamos las cajas (Asegúrate de que tus cajas se llamen así)
        If TextBoxBase IsNot Nothing Then TextBoxBase.Text = base.ToString("C2")
        If TextBoxIva IsNot Nothing Then TextBoxIva.Text = sumaIva.ToString("C2")
        If TextBoxTotalAlb IsNot Nothing Then TextBoxTotalAlb.Text = (base + sumaIva).ToString("C2")
    End Sub

    Private Sub LimpiarFormulario()
        _numeroFacturaActual = ""
        _idsParaBorrar.Clear()
        TextBoxFactura.Text = GenerarProximoNumeroFactura()
        TextBoxIdCliente.Text = "" : TextBoxCliente.Text = ""
        TextBoxAlbaranOrigen.Text = "" : TextBoxAlbaranOrigen.Tag = Nothing
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
    Private Function GenerarProximoNumeroFactura() As String
        Dim prefijo As String = "FAC-"
        Dim nuevo As String = "FAC-001"
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()
            Dim cmd As New SQLiteCommand("SELECT NumeroFactura FROM FacturasVenta WHERE NumeroFactura LIKE @pat ORDER BY NumeroFactura DESC LIMIT 1", c)
            cmd.Parameters.AddWithValue("@pat", prefijo & "%")
            Dim res = cmd.ExecuteScalar()
            If res IsNot Nothing AndAlso Not IsDBNull(res) Then
                Dim parts = res.ToString().Split("-"c)
                If parts.Length >= 2 Then
                    nuevo = prefijo & (Convert.ToInt32(parts(parts.Length - 1)) + 1).ToString("D3")
                End If
            End If
        Catch
            nuevo = "FAC-" & DateTime.Now.Ticks.ToString().Substring(12)
        End Try
        Return nuevo
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Using frm As New FrmBuscador()
            frm.TablaABuscar = "Facturas"
            If frm.ShowDialog() = DialogResult.OK Then
                CargarFactura(frm.Resultado)
            End If
        End Using
    End Sub
    Private Sub CalcularTotales()
        Dim base As Decimal = 0
        If _dtLineas IsNot Nothing Then
            For Each r As DataRow In _dtLineas.Rows
                If r.RowState <> DataRowState.Deleted Then
                    Dim c As Decimal = 0 : Decimal.TryParse(r("Cantidad").ToString(), c)
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

        ' =================================================================
        ' 1. CUANDO EL USUARIO METE EL CÓDIGO DEL ARTÍCULO
        ' =================================================================
        If colName = "ID_Articulo" Then
            Dim idArt As String = fila.Cells("ID_Articulo").Value?.ToString()
            Dim idAntiguo As String = fila.Cells("ID_Articulo").Tag?.ToString()

            ' ---> EL FRENO MÁGICO <---
            ' Si sales de la celda pero el código es EXACTAMENTE el mismo que había antes,
            ' abortamos la operación. Así no destruimos los datos congelados de la factura.
            If idArt = idAntiguo Then Return

            If String.IsNullOrWhiteSpace(idArt) Then
                fila.Cells("Descripcion").Value = ""
                fila.Cells("PrecioUnitario").Value = 0
                fila.Cells("PorcentajeIVA").Value = 21
                fila.Tag = Nothing
                Return
            End If

            Try
                Dim c = ConexionBD.GetConnection()
                If c.State <> ConnectionState.Open Then c.Open()

                Dim sql As String = "SELECT Descripcion, PrecioVenta, StockActual, TipoIVA FROM Articulos WHERE ID_Articulo = @id"

                Using cmd As New SQLiteCommand(sql, c)
                    cmd.Parameters.AddWithValue("@id", idArt)
                    Using r = cmd.ExecuteReader()
                        If r.Read() Then
                            fila.Cells("Descripcion").Value = r("Descripcion").ToString()
                            fila.Cells("PrecioUnitario").Value = Convert.ToDecimal(r("PrecioVenta"))
                            fila.Cells("PorcentajeIVA").Value = If(IsDBNull(r("TipoIVA")), 21, Convert.ToDecimal(r("TipoIVA")))

                            If Not IsDBNull(r("StockActual")) Then fila.Tag = Convert.ToDecimal(r("StockActual"))
                            If IsDBNull(fila.Cells("Cantidad").Value) OrElse Val(fila.Cells("Cantidad").Value) = 0 Then fila.Cells("Cantidad").Value = 1

                            ' Actualizamos el bolsillo secreto con el nuevo ID válido
                            fila.Cells("ID_Articulo").Tag = idArt
                        Else
                            MessageBox.Show("Artículo no encontrado", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show("Error al buscar artículo: " & ex.Message)
            End Try
        End If
        ' =================================================================
        ' 2. RECALCULAR TOTALES AL CAMBIAR CANTIDAD, PRECIO, DTO O IVA
        ' =================================================================
        If colName = "Cantidad" Or colName = "PrecioUnitario" Or colName = "Descuento" Or colName = "PorcentajeIVA" Then
            Dim cant As Decimal = 0, prec As Decimal = 0, dto As Decimal = 0, iva As Decimal = 21

            Decimal.TryParse(fila.Cells("Cantidad").Value?.ToString(), cant)
            Decimal.TryParse(fila.Cells("PrecioUnitario").Value?.ToString(), prec)
            Decimal.TryParse(fila.Cells("Descuento").Value?.ToString(), dto)
            Decimal.TryParse(fila.Cells("PorcentajeIVA").Value?.ToString(), iva)

            ' Matemáticas puras
            Dim precioConIva As Decimal = prec * (1 + (iva / 100))
            Dim totalSinIva As Decimal = (cant * prec) * (1 - (dto / 100))
            Dim totalConIva As Decimal = totalSinIva * (1 + (iva / 100))

            ' Metemos los datos a la pantalla

            fila.Cells("Total").Value = totalSinIva
            fila.Cells("TotalConIVA").Value = totalConIva

            ' Avisamos al resumen de abajo para que también se actualice
            CalcularTotalesGenerales()
        End If
    End Sub

    Private Sub ButtonAnterior_Click(sender As Object, e As EventArgs) Handles ButtonAnterior.Click
        ' Lógica de navegación String (Ver Presupuestos)
        Try
            Dim conexion = ConexionBD.GetConnection()
            Dim numDestino As String = ""

            ' CAMBIO 7: Navegación alfanumérica (Funciona bien con formatos fijos tipo PRE-001)
            Dim sql As String = "SELECT MAX(NumeroFactura) FROM FacturasVenta WHERE NumeroFactura < @actual"
            Using cmd As New SQLiteCommand(sql, conexion)
                ' Si es nuevo, buscamos el último absoluto
                If String.IsNullOrEmpty(_numeroFacturaActual) Then
                    cmd.CommandText = "SELECT MAX(NumeroFactura) FROM FacturasVenta"
                Else
                    cmd.Parameters.AddWithValue("@actual", _numeroFacturaActual)
                End If

                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    numDestino = result.ToString()
                    CargarFactura(numDestino)
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
        If String.IsNullOrEmpty(_numeroFacturaActual) Then Return
        Try
            Dim conexion = ConexionBD.GetConnection()
            Dim sql As String = "SELECT MIN(NumeroFactura) FROM FacturasVenta WHERE NumeroFactura > @actual"
            Using cmd As New SQLiteCommand(sql, conexion)
                cmd.Parameters.AddWithValue("@actual", _numeroFacturaActual)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    CargarFactura(result.ToString())
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
        '-----------------------------------------------------------------------------------------------
        'HACER MAS TARDE
        '-----------------------------------------------------------------------------------------------
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
        nuevaFila("NumeroFactura") = _numeroFacturaActual ' Si es 0 no pasa nada, se arregla al guardar
        nuevaFila("ID_Linea") = DBNull.Value     ' ESTO ES LO MÁS IMPORTANTE (Indica que es NUEVA)
        nuevaFila("ID_Articulo") = 0
        nuevaFila("Descripcion") = ""
        nuevaFila("Cantidad") = 1
        nuevaFila("PrecioUnitario") = 0
        nuevaFila("Descuento") = 0
        nuevaFila("PorcentajeIVA") = 21
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
    ' =================================================================
    ' RECORDAR EL VALOR ANTES DE EDITAR (Para no sobreescribir facturas viejas)
    ' =================================================================
    Private Sub DataGridView1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        Dim colName = DataGridView1.Columns(e.ColumnIndex).Name
        If colName = "ID_Articulo" Then
            Dim celda = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex)
            ' Guardamos el ID que hay ahora mismo en el "bolsillo" (Tag) de la celda
            celda.Tag = celda.Value?.ToString()
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
    Private Sub btnImportarAlbaran_Click_1(sender As Object, e As EventArgs) Handles btnImportarAlbaran.Click
        Using frm As New FrmBuscador
            frm.TablaABuscar = "Pedidos"
            If frm.ShowDialog = DialogResult.OK Then
                Dim codAlba = frm.Resultado ' Ahora recibimos String
                If Not String.IsNullOrEmpty(codAlba) Then ImportarDatosAlbaran(codAlba)
            End If
        End Using

    End Sub
    Private Sub ImportarDatosAlbaran(CodigoAlbaran)
        Dim c = ConexionBD.GetConnection()
        Try
            If c.State <> ConnectionState.Open Then c.Open()
            Dim sqlCab As String = "SELECT A.*, C.NombreFiscal, V.Nombre AS NombreVend FROM Albaranes A " &
                                   "LEFT JOIN Clientes C ON A.CodigoCliente=C.CodigoCliente " &
                                   "LEFT JOIN Vendedores V ON A.ID_Vendedor=V.ID_Vendedor " &
                                   "WHERE A.NumeroAlbaran= @num"
            Using cmd As New SQLiteCommand(sqlCab, c)
                cmd.Parameters.AddWithValue("@num", CodigoAlbaran)
                Using r = cmd.ExecuteReader()
                    If r.Read() Then
                        LimpiarFormulario()
                        TextBoxIdCliente.Text = r("CodigoCliente").ToString()
                        TextBoxCliente.Text = r("NombreFiscal").ToString()
                        TextBoxIdVendedor.Text = If(IsDBNull(r("ID_Vendedor")), "", r("ID_Vendedor").ToString())
                        TextBoxVendedor.Text = If(IsDBNull(r("NombreVend")), "", r("NombreVend").ToString())
                        DateTimePickerFecha.Value = r("FechaEntrega")
                        TextBoxObservaciones.Text = $"Desde Presupuesto {CodigoAlbaran}. "
                        If TextBoxAlbaranOrigen IsNot Nothing Then
                            TextBoxAlbaranOrigen.Text = CodigoAlbaran
                            TextBoxAlbaranOrigen.Tag = CodigoAlbaran
                        End If
                    End If
                End Using
            End Using
            Dim sqlLin As String = "SELECT * FROM LineasAlbaran Where Numeroalbaran = @num order by NumeroOrden ASC"
            Dim dtOrigen As New DataTable()
            Using cmd As New SQLiteCommand(sqlLin, c)
                cmd.Parameters.AddWithValue("@num", CodigoAlbaran)
                Dim da As New SQLiteDataAdapter(cmd)
                da.Fill(dtOrigen)
            End Using
            _dtLineas.Rows.Clear()
            For Each rowOrig As DataRow In dtOrigen.Rows
                Dim rowNew As DataRow = _dtLineas.NewRow()

                ' --- CORRECCIÓN DEL ERROR ---
                ' Ya no existe ID_Pedido, ahora usamos NumeroPedido
                rowNew("NumeroFactura") = _numeroFacturaActual ' Asignamos el código actual (ej: PED-26-001)
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
                    Case "factura", "documento", "pedido", "albaran", "albarán"
                        ctrl.Text = "Factura" ' Forzamos el nombre correcto
                        ctrl.Location = New Point(col1_X, yFila1)
                    Case "cliente" : ctrl.Location = New Point(col2_X, yFila1)
                    Case "fecha" : ctrl.Location = New Point(col3_X, yFila1)
                    Case "fecha entrega", "vencimiento"
                        ctrl.Text = "Vencimiento" ' Forzamos el nombre correcto
                        ctrl.Location = New Point(col3_X + 130, yFila1)
                    Case "vendedor" : ctrl.Location = New Point(col1_X, yFila3)
                    Case "estado" : ctrl.Location = New Point(col2_X, yFila3)
                    Case "agencia" : ctrl.Location = New Point(col3_X, yFila3)
                End Select
            End If
        Next

        ' Alinear Cajas Fila 1
        If TextBoxFactura IsNot Nothing Then TextBoxFactura.Bounds = New Rectangle(col1_X, yFila2, 100, 25)
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
    Private Sub EstilizarBoton(btn As System.Windows.Forms.Button, x As Integer, y As Integer, bg As Color, fg As Color)
        If btn IsNot Nothing Then
            btn.Location = New Point(x, y)
            btn.BackColor = bg : btn.ForeColor = fg
            btn.FlatStyle = FlatStyle.Flat : btn.FlatAppearance.BorderSize = 0
            btn.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
            btn.Cursor = Cursors.Hand : btn.Height = 30
        End If
    End Sub
    Protected Overrides ReadOnly Property CreateParams As System.Windows.Forms.CreateParams
        Get
            ' Forzamos el uso del tipo especifico de windows forms
            Dim cp As System.Windows.Forms.CreateParams = MyBase.CreateParams
            'Aplicamos el estilo extendido para evitar parpadeos (WS_EX_COMPOSITED)
            cp.ExStyle = cp.ExStyle Or &H2000000
            Return cp
        End Get
    End Property
End Class
