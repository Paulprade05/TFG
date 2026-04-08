Imports System.Data.SQLite

Public Class FrmPedidos

#Region "1. Variables Globales y Propiedades"
    Private _numeroPedidoActual As String = ""
    Private _dtLineas As DataTable
    Private _idsParaBorrar As New List(Of Integer)
    ' --- MOTOR DE IMPRESIÓN DEL DOCUMENTO ---
    Private WithEvents btnImprimir As New Button()
    Private WithEvents docImprimir As New Printing.PrintDocument()
    Private _filaActualImpresion As Integer = 0
    Private WithEvents cboFormaPago As New ComboBox()
    Private WithEvents cboRuta As New ComboBox()
    Private lblFormaPago As New Label() With {.Text = "Forma de Pago", .AutoSize = True, .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold), .ForeColor = Color.WhiteSmoke}
    Private lblRuta As New Label() With {.Text = "Ruta Asignada", .AutoSize = True, .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold), .ForeColor = Color.WhiteSmoke}
    Private WithEvents cboEstado As New ComboBox()
#End Region

#Region "2. Eventos de Inicialización (Load)"
    Private Sub FrmPedidos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.SuspendLayout()

        ConfigurarDiseñoResponsive()
        EstilizarGrid(DataGridView1)

        ' Combos
        Me.Controls.Add(lblFormaPago) : Me.Controls.Add(cboFormaPago)
        Me.Controls.Add(lblRuta) : Me.Controls.Add(cboRuta)
        cboFormaPago.DropDownStyle = ComboBoxStyle.DropDownList
        cboRuta.DropDownStyle = ComboBoxStyle.DropDownList
        CargarDesplegables()
        Me.Controls.Add(cboEstado)
        cboEstado.DropDownStyle = ComboBoxStyle.DropDownList
        cboEstado.Items.Clear()
        cboEstado.Items.AddRange(New String() {"Pendiente", "En Preparación", "Servido"})
        ReorganizarControlesAutomaticamente()
        ConfigurarColumnasGrid()
        ' CONFIGURACIÓN DEL BOTÓN IMPRIMIR (Exportar PDF)
        btnImprimir.Text = "Exportar PDF"
        btnImprimir.Size = New Size(120, 35)

        ' Lo anclamos debajo del botón Guardar (Asegúrate de que tu botón se llame así)
        btnImprimir.Location = New Point(ButtonGuardar.Location.X, ButtonGuardar.Location.Y + ButtonGuardar.Height + 10)

        btnImprimir.BackColor = Color.FromArgb(40, 140, 90) ' Verde corporativo
        btnImprimir.ForeColor = Color.White
        btnImprimir.FlatStyle = FlatStyle.Flat
        btnImprimir.FlatAppearance.BorderSize = 0
        btnImprimir.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        btnImprimir.Cursor = Cursors.Hand
        Me.Controls.Add(btnImprimir)
        btnImprimir.BringToFront()
        Dim ultimoNum As String = ObtenerUltimoNumeroPedido()
        If Not String.IsNullOrEmpty(ultimoNum) Then
            CargarPedido(ultimoNum)
        Else
            LimpiarFormulario()
        End If
        Me.ResumeLayout(True)
    End Sub

    Private Sub CargarDesplegables()
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()

            Dim daPago As New SQLiteDataAdapter("SELECT ID_FormaPago, Descripcion FROM FormasPago WHERE Activo=1", c)
            Dim dtPago As New DataTable() : daPago.Fill(dtPago)
            cboFormaPago.DataSource = dtPago : cboFormaPago.DisplayMember = "Descripcion" : cboFormaPago.ValueMember = "ID_FormaPago" : cboFormaPago.SelectedIndex = -1

            Dim daRuta As New SQLiteDataAdapter("SELECT ID_Ruta, NombreZona FROM Rutas WHERE Activo=1", c)
            Dim dtRuta As New DataTable() : daRuta.Fill(dtRuta)
            cboRuta.DataSource = dtRuta : cboRuta.DisplayMember = "NombreZona" : cboRuta.ValueMember = "ID_Ruta" : cboRuta.SelectedIndex = -1
        Catch ex As Exception
        End Try
    End Sub
#End Region

#Region "3. Configuración UI y Diseño"
    Private Sub ConfigurarDiseñoResponsive()
        DataGridView1.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        Dim controlesTotales As Control() = {LabelBase, TextBoxBase, LabelIva, TextBoxIva, Label7, TextBoxTotalPed, LabelStock}
        For Each ctrl In controlesTotales : ctrl.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right : Next
        ButtonAnterior.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        ButtonSiguiente.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        Dim botonesIzq As Control() = {ButtonGuardar, ButtonBorrar, ButtonNuevoPresup, ButtonBorrarLineas, ButtonNuevaLinea}
        For Each ctrl In botonesIzq : ctrl.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left : Next
        If LabelStock IsNot Nothing Then LabelStock.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
    End Sub

    Private Sub ConfigurarColumnasGrid()
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.Columns.Clear()

        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "ID_Linea", .DataPropertyName = "ID_Linea", .Visible = False})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "NumeroOrden", .DataPropertyName = "NumeroOrden", .HeaderText = "Nº", .Width = 40, .ReadOnly = True, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleCenter}})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "ID_Articulo", .DataPropertyName = "ID_Articulo", .HeaderText = "ID Art", .Width = 60})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "Descripcion", .DataPropertyName = "Descripcion", .HeaderText = "Descripción", .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill})

        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "Cantidad", .DataPropertyName = "Cantidad", .HeaderText = "Cantidad", .Width = 75, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "N2", .BackColor = Color.Ivory}})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "PrecioUnitario", .DataPropertyName = "PrecioUnitario", .HeaderText = "Precio Base", .Width = 85, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "C2"}})

        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "Descuento", .DataPropertyName = "Descuento", .HeaderText = "% Dto", .Width = 60, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "N2"}})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "PorcentajeIVA", .DataPropertyName = "PorcentajeIVA", .HeaderText = "% IVA", .Width = 60, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "N0"}})

        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "Total", .DataPropertyName = "Total", .HeaderText = "Total Base", .ReadOnly = True, .Width = 90, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "C2", .BackColor = Color.WhiteSmoke}})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "TotalConIVA", .DataPropertyName = "TotalConIVA", .HeaderText = "Total (+IVA)", .ReadOnly = True, .Width = 100, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "C2", .BackColor = Color.FromArgb(230, 240, 250)}})
    End Sub

    Public Shared Sub EstilizarGrid(dgv As DataGridView)
        dgv.BackgroundColor = Color.White : dgv.BorderStyle = BorderStyle.None : dgv.CellBorderStyle = DataGridViewCellBorderStyle.None
        dgv.EnableHeadersVisualStyles = False : dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
        dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 55, 65)
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        dgv.ColumnHeadersHeight = 40
        dgv.DefaultCellStyle.BackColor = Color.White : dgv.DefaultCellStyle.ForeColor = Color.Black
        dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(200, 230, 255) : dgv.DefaultCellStyle.SelectionForeColor = Color.Black
        dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245)
        dgv.AlternatingRowsDefaultCellStyle.SelectionBackColor = Color.FromArgb(200, 230, 255)
        dgv.RowHeadersVisible = False : dgv.RowTemplate.Height = 35 : dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect : dgv.MultiSelect = False
    End Sub
#End Region

#Region "4. Lógica de Negocio y Cálculos"
    Private Sub CalcularTotalesGenerales()
        Dim base As Decimal = 0
        Dim sumaIva As Decimal = 0

        If _dtLineas IsNot Nothing Then
            For Each row As DataRow In _dtLineas.Rows
                If row.RowState <> DataRowState.Deleted Then
                    Dim t As Decimal = 0
                    Dim ivaLinea As Decimal = 21

                    If _dtLineas.Columns.Contains("Total") Then Decimal.TryParse(row("Total").ToString(), t)
                    If _dtLineas.Columns.Contains("PorcentajeIVA") AndAlso Not IsDBNull(row("PorcentajeIVA")) Then
                        Decimal.TryParse(row("PorcentajeIVA").ToString(), ivaLinea)
                    End If

                    base += t
                    sumaIva += t * (ivaLinea / 100)
                End If
            Next
        End If

        If TextBoxBase IsNot Nothing Then TextBoxBase.Text = base.ToString("C2")
        If TextBoxIva IsNot Nothing Then TextBoxIva.Text = sumaIva.ToString("C2")
        If TextBoxTotalPed IsNot Nothing Then TextBoxTotalPed.Text = (base + sumaIva).ToString("C2")
    End Sub

    Private Sub LimpiarFormulario()
        _numeroPedidoActual = ""
        _idsParaBorrar.Clear()
        TextBoxPedido.Text = GenerarProximoNumeroPedido()

        TextBoxCliente.Text = "" : TextBoxIdCliente.Text = "" : TextBoxCliente.Tag = Nothing
        TextBoxVendedor.Text = "" : TextBoxIdVendedor.Text = ""
        TextBoxObservaciones.Text = ""
        TextBoxIdPresupuesto.Text = ""
        TextBoxFecha.Text = DateTime.Now.ToShortDateString()
        DateTimePickerFecha.Value = DateTime.Now
        cboEstado.Text = "Pendiente"
        TextBoxBase.Text = "0,00 €" : TextBoxIva.Text = "0,00 €" : TextBoxTotalPed.Text = "0,00 €"

        _dtLineas = New DataTable()
        ConfigurarEstructuraDataTable()
        DataGridView1.DataSource = _dtLineas
        If cboFormaPago IsNot Nothing Then cboFormaPago.SelectedIndex = -1
        If cboRuta IsNot Nothing Then cboRuta.SelectedIndex = -1
        TextBoxIdCliente.Focus()
    End Sub

    Private Sub ConfigurarEstructuraDataTable()
        With _dtLineas.Columns
            If Not .Contains("ID_Linea") Then .Add("ID_Linea", GetType(Object))
            If Not .Contains("NumeroPedido") Then .Add("NumeroPedido", GetType(String))
            If Not .Contains("NumeroOrden") Then .Add("NumeroOrden", GetType(Integer))
            If Not .Contains("ID_Articulo") Then .Add("ID_Articulo", GetType(Object))
            If Not .Contains("Descripcion") Then .Add("Descripcion", GetType(String))
            If Not .Contains("Cantidad") Then .Add("Cantidad", GetType(Decimal))
            If Not .Contains("PrecioUnitario") Then .Add("PrecioUnitario", GetType(Decimal))
            If Not .Contains("Descuento") Then .Add("Descuento", GetType(Decimal))

            ' --- COLUMNAS IVA Y VIRTUALES ---
            If Not .Contains("PorcentajeIVA") Then .Add("PorcentajeIVA", GetType(Decimal))
            If Not .Contains("PrecioConIVA") Then .Add("PrecioConIVA", GetType(Decimal))
            If Not .Contains("Total") Then .Add("Total", GetType(Decimal))
            If Not .Contains("TotalConIVA") Then .Add("TotalConIVA", GetType(Decimal))
        End With
    End Sub

    Private Function GenerarProximoNumeroPedido() As String
        Dim prefijo As String = "PED-"
        Dim nuevoNumero As String = $"{prefijo}001"
        Try
            Dim sql As String = "SELECT NumeroPedido FROM Pedidos WHERE NumeroPedido LIKE @patron ORDER BY NumeroPedido DESC LIMIT 1"

            ' --- SOLUCIÓN: Usamos Dim en lugar de Using para NO destruir la conexión compartida ---
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()

            Using cmd As New SQLiteCommand(sql, c)
                cmd.Parameters.AddWithValue("@patron", prefijo & "%")
                Dim resultado = cmd.ExecuteScalar()
                If resultado IsNot Nothing Then
                    Dim partes As String() = resultado.ToString().Split("-"c)
                    If partes.Length >= 2 AndAlso IsNumeric(partes(1)) Then
                        nuevoNumero = $"{prefijo}{(CInt(partes(1)) + 1).ToString("D3")}"
                    End If
                End If
            End Using
        Catch
            nuevoNumero = $"PED-{DateTime.Now:HHmmss}"
        End Try
        Return nuevoNumero
    End Function
#End Region

#Region "5. Persistencia (Base de Datos - CRUD)"
    Private Sub CargarPedido(numeroPedido As String)
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()

            ' A. CARGAR CABECERA
            Dim sqlCab As String = "SELECT P.*, C.NombreFiscal AS NombreCliente, V.Nombre AS NombreVendedor " &
                                   "FROM Pedidos P " &
                                   "LEFT JOIN Clientes C ON P.CodigoCliente = C.CodigoCliente " &
                                   "LEFT JOIN Vendedores V ON P.ID_Vendedor = V.ID_Vendedor " &
                                   "WHERE P.NumeroPedido = @num"

            Using cmd As New SQLiteCommand(sqlCab, c)
                cmd.Parameters.AddWithValue("@num", numeroPedido)
                Using r = cmd.ExecuteReader()
                    If r.Read() Then
                        _numeroPedidoActual = numeroPedido
                        TextBoxPedido.Text = r("NumeroPedido").ToString()
                        TextBoxFecha.Text = If(IsDBNull(r("Fecha")), "", Convert.ToDateTime(r("Fecha")).ToShortDateString())
                        If Not IsDBNull(r("FechaEntrega")) Then DateTimePickerFecha.Value = Convert.ToDateTime(r("FechaEntrega"))

                        TextBoxObservaciones.Text = r("Observaciones").ToString()
                        cboEstado.Text = r("Estado").ToString()
                        TextBoxIdCliente.Text = r("CodigoCliente").ToString()
                        TextBoxCliente.Text = r("NombreCliente").ToString()
                        TextBoxIdVendedor.Text = r("ID_Vendedor").ToString()
                        TextBoxVendedor.Text = r("NombreVendedor").ToString()
                        TextBoxIdPresupuesto.Text = r("NumeroPresupuesto").ToString()

                        Dim idPago = r("ID_FormaPago")
                        If Not IsDBNull(idPago) Then cboFormaPago.SelectedValue = Convert.ToInt32(idPago) Else cboFormaPago.SelectedIndex = -1

                        Dim idRut = r("ID_Ruta")
                        If Not IsDBNull(idRut) Then cboRuta.SelectedValue = Convert.ToInt32(idRut) Else cboRuta.SelectedIndex = -1
                    Else
                        MessageBox.Show("Pedido no encontrado.") : Return
                    End If
                End Using
            End Using

            ' B. CARGAR LÍNEAS
            Dim sqlLin As String = "SELECT * FROM LineasPedido WHERE NumeroPedido = @num ORDER BY NumeroOrden ASC"
            Using cmd As New SQLiteCommand(sqlLin, c)
                cmd.Parameters.AddWithValue("@num", numeroPedido)
                Dim da As New SQLiteDataAdapter(cmd)
                _dtLineas = New DataTable()
                da.Fill(_dtLineas)
                ConfigurarEstructuraDataTable()

                ' Magia para calcular IVA al vuelo
                For Each row As DataRow In _dtLineas.Rows
                    Dim cant As Decimal = If(IsDBNull(row("Cantidad")), 0, Convert.ToDecimal(row("Cantidad")))
                    Dim prec As Decimal = If(IsDBNull(row("PrecioUnitario")), 0, Convert.ToDecimal(row("PrecioUnitario")))
                    Dim dto As Decimal = If(IsDBNull(row("Descuento")), 0, Convert.ToDecimal(row("Descuento")))
                    Dim iva As Decimal = If(IsDBNull(row("PorcentajeIVA")), 21, Convert.ToDecimal(row("PorcentajeIVA")))

                    row("PorcentajeIVA") = iva
                    row("PrecioConIVA") = prec * (1 + (iva / 100))

                    Dim totalSinIva As Decimal = (cant * prec) * (1 - (dto / 100))
                    row("Total") = totalSinIva
                    row("TotalConIVA") = totalSinIva * (1 + (iva / 100))
                Next

                DataGridView1.DataSource = _dtLineas
                CalcularTotalesGenerales()
            End Using

        Catch ex As Exception
            MessageBox.Show("Error al cargar: " & ex.Message)
        End Try
    End Sub

    Private Sub GuardarPedido()
        If String.IsNullOrWhiteSpace(TextBoxIdCliente.Text) Then MessageBox.Show("Falta el Cliente") : Return

        Dim esNuevo As Boolean = String.IsNullOrEmpty(_numeroPedidoActual)
        If esNuevo Then
            TextBoxPedido.Text = GenerarProximoNumeroPedido()
            _numeroPedidoActual = TextBoxPedido.Text
        End If

        Dim c = ConexionBD.GetConnection()
        Dim trans As SQLiteTransaction = Nothing
        Try
            If c.State <> ConnectionState.Open Then c.Open()
            trans = c.BeginTransaction()

            ' 0. Calcular sumas para la cabecera
            Dim sumaBase As Decimal = 0
            Dim sumaIva As Decimal = 0
            For Each r As DataRow In _dtLineas.Rows
                If r.RowState <> DataRowState.Deleted Then
                    Dim tot As Decimal = 0 : Decimal.TryParse(r("Total").ToString(), tot)
                    Dim ivaLinea As Decimal = 21 : Decimal.TryParse(r("PorcentajeIVA").ToString(), ivaLinea)
                    sumaBase += tot
                    sumaIva += tot * (ivaLinea / 100)
                End If
            Next

            ' 1. Guardar Cabecera 
            Dim idFormaPago As Object = If(cboFormaPago.SelectedValue IsNot Nothing AndAlso cboFormaPago.SelectedIndex <> -1, cboFormaPago.SelectedValue, DBNull.Value)
            Dim idRuta As Object = If(cboRuta.SelectedValue IsNot Nothing AndAlso cboRuta.SelectedIndex <> -1, cboRuta.SelectedValue, DBNull.Value)
            Dim idVend As Object = If(IsNumeric(TextBoxIdVendedor.Text) AndAlso Val(TextBoxIdVendedor.Text) > 0, Convert.ToInt32(TextBoxIdVendedor.Text), DBNull.Value)

            Dim sql As String = ""
            If esNuevo Then
                sql = "INSERT INTO Pedidos (NumeroPedido, NumeroPresupuesto, CodigoCliente, ID_Vendedor, Fecha, FechaEntrega, Observaciones, Estado, ID_FormaPago, ID_Ruta, BaseImponible, ImporteIVA, Total) " &
                  "VALUES (@num, @numpres, @cli, @vend, @fecha, @fechaEnt, @obs, @est, @formaPago, @ruta, @base, @iva, @total)"
            Else
                sql = "UPDATE Pedidos SET NumeroPresupuesto=@numpres, CodigoCliente=@cli, ID_Vendedor=@vend, Fecha=@fecha, FechaEntrega=@fechaEnt, Observaciones=@obs, Estado=@est, ID_FormaPago=@formaPago, ID_Ruta=@ruta, BaseImponible=@base, ImporteIVA=@iva, Total=@total " &
                  "WHERE NumeroPedido = @num"
            End If

            Using cmd As New SQLiteCommand(sql, c)
                cmd.Transaction = trans
                cmd.Parameters.AddWithValue("@num", _numeroPedidoActual)
                cmd.Parameters.AddWithValue("@cli", TextBoxIdCliente.Text.Trim())
                cmd.Parameters.AddWithValue("@vend", idVend)

                Dim idPresu As Object = If(String.IsNullOrWhiteSpace(TextBoxIdPresupuesto.Text), DBNull.Value, TextBoxIdPresupuesto.Text.Trim())
                cmd.Parameters.AddWithValue("@numpres", idPresu)

                ' --- CORRECCIÓN FECHAS ---
                Dim fecha As DateTime
                If DateTime.TryParse(TextBoxFecha.Text, fecha) Then
                    cmd.Parameters.AddWithValue("@fecha", fecha.ToString("yyyy-MM-dd HH:mm:ss"))
                Else
                    cmd.Parameters.AddWithValue("@fecha", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                End If

                ' Aquí estaba el error, ahora asigamos @fechaEnt correctamente al DateTimePicker
                cmd.Parameters.AddWithValue("@fechaEnt", DateTimePickerFecha.Value.ToString("yyyy-MM-dd HH:mm:ss"))
                ' -------------------------

                cmd.Parameters.AddWithValue("@obs", TextBoxObservaciones.Text.Trim())
                cmd.Parameters.AddWithValue("@est", cboEstado.Text.Trim())
                cmd.Parameters.AddWithValue("@formaPago", idFormaPago)
                cmd.Parameters.AddWithValue("@ruta", idRuta)

                cmd.Parameters.AddWithValue("@base", sumaBase)
                cmd.Parameters.AddWithValue("@iva", sumaIva)
                cmd.Parameters.AddWithValue("@total", sumaBase + sumaIva)
                cmd.ExecuteNonQuery()
            End Using

            ' 2. Borrar líneas eliminadas
            For Each idDel In _idsParaBorrar
                Using cmdDel As New SQLiteCommand("DELETE FROM LineasPedido WHERE ID_Linea = @id", c)
                    cmdDel.Transaction = trans : cmdDel.Parameters.AddWithValue("@id", idDel) : cmdDel.ExecuteNonQuery()
                End Using
            Next
            _idsParaBorrar.Clear()

            ' 3. Guardar Líneas
            Dim orden As Integer = 1
            For Each row As DataRow In _dtLineas.Rows
                If row.RowState = DataRowState.Deleted Then Continue For

                Dim idLin = row("ID_Linea")
                Dim idArt = If(Val(row("ID_Articulo")) > 0, row("ID_Articulo"), DBNull.Value)
                Dim sqlLine As String = ""

                If IsDBNull(idLin) OrElse Val(idLin) = 0 Then
                    sqlLine = "INSERT INTO LineasPedido (NumeroPedido, NumeroOrden, ID_Articulo, Descripcion, Cantidad, PrecioUnitario, Descuento, PorcentajeIVA, Total) " &
                              "VALUES (@num, @ord, @art, @desc, @cant, @prec, @dcto, @iva, @tot)"
                Else
                    sqlLine = "UPDATE LineasPedido SET NumeroOrden=@ord, ID_Articulo=@art, Descripcion=@desc, Cantidad=@cant, PrecioUnitario=@prec, Descuento=@dcto, PorcentajeIVA=@iva, Total=@tot WHERE ID_Linea=@id"
                End If

                Using cmdL As New SQLiteCommand(sqlLine, c)
                    cmdL.Transaction = trans
                    cmdL.Parameters.AddWithValue("@num", _numeroPedidoActual)
                    cmdL.Parameters.AddWithValue("@ord", orden)
                    cmdL.Parameters.AddWithValue("@art", idArt)
                    cmdL.Parameters.AddWithValue("@desc", row("Descripcion"))
                    cmdL.Parameters.AddWithValue("@cant", row("Cantidad"))
                    cmdL.Parameters.AddWithValue("@prec", row("PrecioUnitario"))
                    cmdL.Parameters.AddWithValue("@dcto", row("Descuento"))
                    cmdL.Parameters.AddWithValue("@iva", If(IsDBNull(row("PorcentajeIVA")), 21, row("PorcentajeIVA")))
                    cmdL.Parameters.AddWithValue("@tot", row("Total"))
                    If Not String.IsNullOrEmpty(sqlLine) AndAlso sqlLine.Contains("WHERE") Then cmdL.Parameters.AddWithValue("@id", idLin)
                    cmdL.ExecuteNonQuery()
                End Using
                orden += 1
            Next
            ' =========================================================
            ' MAGIA DE ESTADOS: Marcar Presupuesto como Convertido
            ' =========================================================
            If Not String.IsNullOrWhiteSpace(TextBoxIdPresupuesto.Text) Then
                Dim sqlEstado As String = "UPDATE Presupuestos SET Estado = 'Convertido' WHERE NumeroPresupuesto = @idPresu"
                Using cmdEst As New SQLiteCommand(sqlEstado, c)
                    cmdEst.Transaction = trans
                    cmdEst.Parameters.AddWithValue("@idPresu", TextBoxIdPresupuesto.Text.Trim())
                    cmdEst.ExecuteNonQuery()
                End Using
            End If
            trans.Commit()
            MessageBox.Show("Guardado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
            CargarPedido(_numeroPedidoActual)

        Catch ex As Exception
            If trans IsNot Nothing Then trans.Rollback()
            MessageBox.Show("Error al guardar: " & ex.Message)
        End Try
    End Sub

    Private Function ObtenerUltimoNumeroPedido() As String
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()
            Dim cmd As New SQLiteCommand("SELECT MAX(NumeroPedido) FROM Pedidos", c)
            Dim r = cmd.ExecuteScalar()
            Return If(IsDBNull(r), "", r.ToString())
        Catch
            Return ""
        End Try
    End Function
#End Region

#Region "6. IMPORTAR PRESUPUESTO"
    Private Sub btnBuscarPresupuesto_Click(sender As Object, e As EventArgs) Handles btnBuscarPresupuesto.Click
        Using frm As New FrmBuscador()
            frm.TablaABuscar = "Presupuestos"
            If frm.ShowDialog() = DialogResult.OK Then
                Dim codPresu As String = frm.Resultado
                If Not String.IsNullOrEmpty(codPresu) Then ImportarDatosPresupuesto(codPresu)
            End If
        End Using
    End Sub

    Private Sub ImportarDatosPresupuesto(codigoPresupuesto As String)
        Dim c = ConexionBD.GetConnection()
        Try
            If c.State <> ConnectionState.Open Then c.Open()

            ' 1. Cabecera 
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
                        TextBoxObservaciones.Text = $"Generado desde Presupuesto {codigoPresupuesto}."
                        Dim idFPago = r("ID_FormaPago")
                        If Not IsDBNull(idFPago) AndAlso idFPago IsNot Nothing AndAlso idFPago.ToString() <> "" Then
                            cboFormaPago.SelectedValue = Convert.ToInt32(idFPago)
                        Else
                            cboFormaPago.SelectedIndex = -1 ' Se queda en blanco si el presupuesto no tenía
                        End If
                        TextBoxIdPresupuesto.Text = codigoPresupuesto
                    End If
                End Using
            End Using

            ' 2. Líneas
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

                rowNew("NumeroPedido") = _numeroPedidoActual
                rowNew("ID_Linea") = DBNull.Value
                rowNew("NumeroOrden") = rowOrig("NumeroOrden")
                rowNew("ID_Articulo") = rowOrig("ID_Articulo")
                rowNew("Descripcion") = rowOrig("Descripcion")

                Dim ca As Decimal = 0 : Decimal.TryParse(rowOrig("Cantidad").ToString(), ca)
                Dim pr As Decimal = 0 : Decimal.TryParse(rowOrig("PrecioUnitario").ToString(), pr)
                Dim dt As Decimal = 0 : Decimal.TryParse(rowOrig("Descuento").ToString(), dt)

                Dim iva As Decimal = 21
                If dtOrigen.Columns.Contains("PorcentajeIVA") AndAlso Not IsDBNull(rowOrig("PorcentajeIVA")) Then
                    Decimal.TryParse(rowOrig("PorcentajeIVA").ToString(), iva)
                End If

                rowNew("Cantidad") = ca
                rowNew("PrecioUnitario") = pr
                rowNew("Descuento") = dt
                rowNew("PorcentajeIVA") = iva
                rowNew("PrecioConIVA") = pr * (1 + (iva / 100))

                Dim totalSinIva As Decimal = (ca * pr) * (1 - (dt / 100))
                rowNew("Total") = totalSinIva
                rowNew("TotalConIVA") = totalSinIva * (1 + (iva / 100))

                _dtLineas.Rows.Add(rowNew)
            Next

            CalcularTotalesGenerales()
            MessageBox.Show("Presupuesto importado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Error al importar: " & ex.Message)
        End Try
    End Sub
#End Region

#Region "7. Manejadores de Eventos (Botones y Grid)"

    Private Sub ButtonGuardar_Click(sender As Object, e As EventArgs) Handles ButtonGuardar.Click
        GuardarPedido()
    End Sub

    Private Sub ButtonNuevoPresup_Click(sender As Object, e As EventArgs) Handles ButtonNuevoPresup.Click
        LimpiarFormulario()
    End Sub

    Private Sub ButtonBorrar_Click(sender As Object, e As EventArgs) Handles ButtonBorrar.Click
        If String.IsNullOrEmpty(_numeroPedidoActual) Then Return
        If MessageBox.Show("¿Seguro que deseas eliminar este pedido?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes Then
            Try
                Dim c = ConexionBD.GetConnection()
                If c.State <> ConnectionState.Open Then c.Open()
                Dim cmd As New SQLiteCommand("DELETE FROM LineasPedido WHERE NumeroPedido=@num; DELETE FROM Pedidos WHERE NumeroPedido=@num;", c)
                cmd.Parameters.AddWithValue("@num", _numeroPedidoActual)
                cmd.ExecuteNonQuery()
                LimpiarFormulario()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub

    Private Sub btnBuscarPedido_Click(sender As Object, e As EventArgs) Handles btnBuscarPedido.Click
        Dim frmBuscar As New FrmBuscador()
        frmBuscar.TablaABuscar = "Pedidos"
        If frmBuscar.ShowDialog() = DialogResult.OK Then
            CargarPedido(frmBuscar.Resultado)
        End If
    End Sub

    Private Sub Navegar(direccion As String)
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()
            Dim sql As String = If(direccion = "ANT",
                                   "SELECT MAX(NumeroPedido) FROM Pedidos WHERE NumeroPedido < @act",
                                   "SELECT MIN(NumeroPedido) FROM Pedidos WHERE NumeroPedido > @act")

            If String.IsNullOrEmpty(_numeroPedidoActual) And direccion = "ANT" Then
                sql = "SELECT MAX(NumeroPedido) FROM Pedidos"
            End If

            Using cmd As New SQLiteCommand(sql, c)
                cmd.Parameters.AddWithValue("@act", _numeroPedidoActual)
                Dim res = cmd.ExecuteScalar()
                If res IsNot Nothing AndAlso Not IsDBNull(res) Then CargarPedido(res.ToString())
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub ButtonAnterior_Click(sender As Object, e As EventArgs) Handles ButtonAnterior.Click
        Navegar("ANT")
    End Sub

    Private Sub ButtonSiguiente_Click(sender As Object, e As EventArgs) Handles ButtonSiguiente.Click
        Navegar("SIG")
    End Sub

    Private Sub ButtonNuevaLinea_Click(sender As Object, e As EventArgs) Handles ButtonNuevaLinea.Click
        If _dtLineas Is Nothing Then
            _dtLineas = New DataTable()
            ConfigurarEstructuraDataTable()
        End If

        Dim fila = _dtLineas.NewRow()
        fila("NumeroOrden") = _dtLineas.Rows.Count + 1
        fila("NumeroPedido") = _numeroPedidoActual
        fila("Cantidad") = 1
        fila("PrecioUnitario") = 0
        fila("Descuento") = 0
        fila("PorcentajeIVA") = 21
        fila("Total") = 0
        _dtLineas.Rows.Add(fila)

        If DataGridView1.Rows.Count > 0 Then
            DataGridView1.CurrentCell = DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells("ID_Articulo")
            DataGridView1.BeginEdit(True)
        End If
    End Sub

    Private Sub ButtonBorrarLineas_Click(sender As Object, e As EventArgs) Handles ButtonBorrarLineas.Click
        If DataGridView1.SelectedRows.Count = 0 Then Return
        For Each f As DataGridViewRow In DataGridView1.SelectedRows
            If Not f.IsNewRow Then
                Dim idVal = f.Cells("ID_Linea").Value
                If IsNumeric(idVal) AndAlso Val(idVal) > 0 Then _idsParaBorrar.Add(CInt(idVal))
                DataGridView1.Rows.Remove(f)
            End If
        Next
        CalcularTotalesGenerales()
    End Sub

    ' EL FRENO MÁGICO
    Private Sub DataGridView1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        Dim colName = DataGridView1.Columns(e.ColumnIndex).Name
        If colName = "ID_Articulo" Then
            Dim celda = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex)
            celda.Tag = celda.Value?.ToString()
        End If
    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim fila = DataGridView1.Rows(e.RowIndex)
        Dim colName = DataGridView1.Columns(e.ColumnIndex).Name

        ' A. Artículo
        If colName = "ID_Articulo" Then
            Dim idArt As String = fila.Cells("ID_Articulo").Value?.ToString()
            Dim idAntiguo As String = fila.Cells("ID_Articulo").Tag?.ToString()

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
                            fila.Cells("ID_Articulo").Tag = idArt
                        Else
                            MessageBox.Show("Artículo no encontrado", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End Using
                End Using
            Catch ex As Exception
            End Try
        End If

        ' B. Cálculos
        If colName = "Cantidad" Or colName = "PrecioUnitario" Or colName = "Descuento" Or colName = "PorcentajeIVA" Then
            Dim cant As Decimal = 0, prec As Decimal = 0, dto As Decimal = 0, iva As Decimal = 21
            Decimal.TryParse(fila.Cells("Cantidad").Value?.ToString(), cant)
            Decimal.TryParse(fila.Cells("PrecioUnitario").Value?.ToString(), prec)
            Decimal.TryParse(fila.Cells("Descuento").Value?.ToString(), dto)
            Decimal.TryParse(fila.Cells("PorcentajeIVA").Value?.ToString(), iva)

            Dim precioConIva As Decimal = prec * (1 + (iva / 100))
            Dim totalSinIva As Decimal = (cant * prec) * (1 - (dto / 100))
            Dim totalConIva As Decimal = totalSinIva * (1 + (iva / 100))

            fila.Cells("Total").Value = totalSinIva
            fila.Cells("TotalConIVA").Value = totalConIva

            CalcularTotalesGenerales()
        End If
    End Sub

    ' LÓGICA DE STOCK EN TIEMPO REAL (Respetada y adaptada)
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        If LabelStock Is Nothing Then Return
        If DataGridView1.CurrentRow Is Nothing OrElse DataGridView1.CurrentRow.IsNewRow Then
            LabelStock.Text = "Stock disponible: -"
            LabelStock.ForeColor = Color.FromArgb(170, 175, 180)
            Return
        End If

        Try
            Dim idCelda As Object = DataGridView1.CurrentRow.Cells("ID_Articulo").Value
            If idCelda IsNot Nothing AndAlso Not DBNull.Value.Equals(idCelda) Then
                Dim idArticulo As Integer = Convert.ToInt32(idCelda)

                Dim cantidadEnLinea As Decimal = 0
                Dim cantCelda As Object = DataGridView1.CurrentRow.Cells("Cantidad").Value
                If cantCelda IsNot Nothing AndAlso Not DBNull.Value.Equals(cantCelda) Then
                    cantidadEnLinea = Convert.ToDecimal(cantCelda)
                End If

                Dim stockActual As Decimal = ConsultarStock(idArticulo)
                Dim stockRestante As Decimal = stockActual - cantidadEnLinea

                If stockRestante > 0 Then
                    LabelStock.Text = $"Stock actual: {stockActual} (Te quedarán: {stockRestante})"
                    LabelStock.ForeColor = Color.FromArgb(40, 180, 90)
                ElseIf stockRestante = 0 Then
                    LabelStock.Text = $"Stock actual: {stockActual} (¡Atención! Te quedarás a 0)"
                    LabelStock.ForeColor = Color.FromArgb(220, 160, 40)
                Else
                    LabelStock.Text = $"¡Stock insuficiente! Actual: {stockActual} (Te faltan: {Math.Abs(stockRestante)})"
                    LabelStock.ForeColor = Color.FromArgb(255, 80, 80)
                End If
                LabelStock.Font = New Font("Segoe UI", 10.5F, FontStyle.Bold)
            End If
        Catch ex As Exception
            LabelStock.Text = "Stock disponible: -"
        End Try
    End Sub

    Private Function ConsultarStock(idArticulo As Integer) As Decimal
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()
            Dim cmd As New System.Data.SQLite.SQLiteCommand("SELECT StockActual FROM Articulos WHERE ID_Articulo = @id", c)
            cmd.Parameters.AddWithValue("@id", idArticulo)
            Dim resultado As Object = cmd.ExecuteScalar()
            If resultado IsNot Nothing AndAlso Not DBNull.Value.Equals(resultado) Then Return Convert.ToDecimal(resultado)
        Catch
        End Try
        Return 0
    End Function

    Private Sub DataGridView1_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles DataGridView1.RowsRemoved
        CalcularTotalesGenerales()
    End Sub

    Private Sub TextBoxIdCliente_Leave(sender As Object, e As EventArgs) Handles TextBoxIdCliente.Leave
        If String.IsNullOrWhiteSpace(TextBoxIdCliente.Text) Then TextBoxCliente.Text = "" : Return
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()
            Dim cmd As New SQLiteCommand("SELECT NombreFiscal, ID_Ruta FROM Clientes WHERE CodigoCliente=@id", c)
            cmd.Parameters.AddWithValue("@id", TextBoxIdCliente.Text)
            Using r = cmd.ExecuteReader()
                If r.Read() Then
                    TextBoxCliente.Text = r("NombreFiscal").ToString()
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
        If String.IsNullOrWhiteSpace(TextBoxIdVendedor.Text) Then TextBoxVendedor.Text = "" : Return
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()
            Dim r = New SQLiteCommand("SELECT Nombre FROM Vendedores WHERE ID_Vendedor='" & TextBoxIdVendedor.Text & "'", c).ExecuteScalar()
            TextBoxVendedor.Text = If(r IsNot Nothing, r.ToString(), "NO EXISTE")
        Catch
        End Try
    End Sub
#End Region

#Region "Auto-Organización Visual (Pixel-Perfect)"
    Private Sub ReorganizarControlesAutomaticamente()
        For Each ctrl As Control In Me.Controls : ctrl.Anchor = AnchorStyles.Top Or AnchorStyles.Left : Next

        Dim margenIzq As Integer = 30
        Dim anchoForm As Integer = Me.ClientSize.Width
        Dim altoForm As Integer = Me.ClientSize.Height

        Dim col1_X As Integer = margenIzq
        Dim col2_X As Integer = 190
        Dim col3_X As Integer = 750
        Dim col4_X As Integer = 920

        Dim yFila1 As Integer = 30
        Dim yFila2 As Integer = 55
        Dim yFila3 As Integer = 95
        Dim yFila4 As Integer = 120
        Dim yFila5 As Integer = 160
        Dim yFila6 As Integer = 185

        For Each ctrl As Control In Me.Controls
            If TypeOf ctrl Is Label AndAlso ctrl.Name <> "LineaTotales" AndAlso ctrl.Name <> "PanelTotalesResumen" Then
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

        TextBoxPedido.Bounds = New Rectangle(col1_X, yFila2, 105, 25)
        btnBuscarPedido.Bounds = New Rectangle(col1_X + 110, yFila2, 30, 25)
        TextBoxIdCliente.Bounds = New Rectangle(col2_X, yFila2, 60, 25)
        TextBoxCliente.Bounds = New Rectangle(col2_X + 70, yFila2, 460, 25)
        TextBoxFecha.Bounds = New Rectangle(col3_X, yFila2, 140, 25)
        If DateTimePickerFecha IsNot Nothing Then DateTimePickerFecha.Bounds = New Rectangle(col4_X, yFila2, 140, 25)

        TextBoxIdVendedor.Bounds = New Rectangle(col1_X, yFila4, 50, 25)
        TextBoxVendedor.Bounds = New Rectangle(col1_X + 55, yFila4, 85, 25)
        TextBoxObservaciones.Bounds = New Rectangle(col2_X, yFila4, 530, 25)
        cboEstado.Bounds = New Rectangle(col3_X, yFila4, 140, 25)
        If TextBoxIdPresupuesto IsNot Nothing Then TextBoxIdPresupuesto.Bounds = New Rectangle(col4_X, yFila4, 105, 25)
        If btnBuscarPresupuesto IsNot Nothing Then btnBuscarPresupuesto.Bounds = New Rectangle(col4_X + 110, yFila4, 30, 25)

        lblFormaPago.Location = New Point(col1_X, yFila5)
        cboFormaPago.Bounds = New Rectangle(col1_X, yFila6, 140, 25)
        cboFormaPago.Font = New Font("Segoe UI", 10.5F)

        lblRuta.Location = New Point(col2_X, yFila5)
        cboRuta.Bounds = New Rectangle(col2_X, yFila6, 530, 25)
        cboRuta.Font = New Font("Segoe UI", 10.5F)

        Dim lineaDivisoria As Label = Me.Controls.OfType(Of Label)().FirstOrDefault(Function(l) l.Name = "LineaDivisoria")
        If lineaDivisoria Is Nothing Then
            lineaDivisoria = New Label() With {.Name = "LineaDivisoria", .BackColor = Color.FromArgb(120, 130, 140), .Height = 2}
            Me.Controls.Add(lineaDivisoria)
        End If

        Dim yTabla As Integer = 240
        lineaDivisoria.Bounds = New Rectangle(margenIzq, yTabla - 20, anchoForm - (margenIzq * 2), 2)
        lineaDivisoria.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        lineaDivisoria.BringToFront()

        Dim altoTabla As Integer = altoForm - yTabla - 140
        DataGridView1.Bounds = New Rectangle(margenIzq, yTabla, anchoForm - (margenIzq * 2), altoTabla)
        DataGridView1.BackgroundColor = Me.BackColor
        DataGridView1.BorderStyle = BorderStyle.None

        Dim xDerecha As Integer = DataGridView1.Right
        Dim yTotales As Integer = DataGridView1.Bottom + 10

        Dim panelTotales As Label = Me.Controls.OfType(Of Label)().FirstOrDefault(Function(l) l.Name = "PanelTotalesResumen")
        If panelTotales Is Nothing Then
            panelTotales = New Label() With {.Name = "PanelTotalesResumen", .BackColor = Color.FromArgb(25, 30, 40)}
            Me.Controls.Add(panelTotales)
            panelTotales.SendToBack()
        End If
        panelTotales.Bounds = New Rectangle(xDerecha - 320, yTotales, 150, 115)
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
        If Label7 IsNot Nothing Then
            Label7.Bounds = New Rectangle(xDerecha - 300, yTotales + 80, 150, 30)
            Label7.BackColor = Color.FromArgb(25, 30, 40) : Label7.TextAlign = ContentAlignment.MiddleRight
            Label7.Font = New Font("Segoe UI", 13, FontStyle.Bold) : Label7.ForeColor = colorAcento
        End If

        If TextBoxTotalPed IsNot Nothing Then
            TextBoxTotalPed.Bounds = New Rectangle(xDerecha - 140, yTotales + 80, 120, 30)
            TextBoxTotalPed.TextAlign = HorizontalAlignment.Right
            TextBoxTotalPed.Font = New Font("Segoe UI", 14, FontStyle.Bold) : TextBoxTotalPed.ForeColor = colorAcento
            TextBoxTotalPed.BackColor = Color.FromArgb(25, 30, 40) : TextBoxTotalPed.BorderStyle = BorderStyle.None
        End If

        Dim panelWidth As Integer = TextBoxBase.Right - LabelBase.Left + 20
        panelTotales.Bounds = New Rectangle(xDerecha - panelWidth, yTotales, panelWidth, 115)

        Dim yBotones As Integer = DataGridView1.Bottom + 45

        EstilizarBoton(ButtonGuardar, margenIzq, yBotones, Color.FromArgb(0, 120, 215), Color.White)
        EstilizarBoton(ButtonBorrar, margenIzq + 110, yBotones, Color.FromArgb(209, 52, 56), Color.White)
        EstilizarBoton(ButtonNuevoPresup, margenIzq + 220, yBotones, Color.FromArgb(0, 120, 215), Color.White)

        EstilizarBoton(ButtonBorrarLineas, margenIzq + 380, yBotones, Color.FromArgb(85, 85, 85), Color.White)
        ButtonBorrarLineas.Text = "- Quitar Línea" : ButtonBorrarLineas.Width = 110

        EstilizarBoton(ButtonNuevaLinea, margenIzq + 500, yBotones, Color.FromArgb(40, 140, 90), Color.White)
        ButtonNuevaLinea.Text = "+ Añadir Línea" : ButtonNuevaLinea.Width = 110

        EstilizarBoton(ButtonAnterior, xDerecha - 560, yBotones, Me.BackColor, Color.White)
        EstilizarBoton(ButtonSiguiente, xDerecha - 450, yBotones, Me.BackColor, Color.White)

        If LabelStock IsNot Nothing Then LabelStock.Location = New Point(margenIzq, DataGridView1.Bottom + 10)

        DataGridView1.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        Dim botonesAbajo As Control() = {ButtonGuardar, ButtonBorrar, ButtonNuevoPresup, ButtonBorrarLineas, ButtonNuevaLinea, ButtonAnterior, ButtonSiguiente}

        For Each b In botonesAbajo
            If b IsNot Nothing Then
                b.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
            End If
        Next
        TextBoxBase.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
            TextBoxIva.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
            If TextBoxTotalPed IsNot Nothing Then TextBoxTotalPed.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
            LabelBase.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
            LabelIva.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
            If Label7 IsNot Nothing Then Label7.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
            lineaTotal.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
    End Sub

    Private Sub EstilizarBoton(btn As Button, x As Integer, y As Integer, bg As Color, fg As Color)
        btn.Location = New Point(x, y)
        btn.Size = New Size(100, 35)
        btn.FlatStyle = FlatStyle.Flat
        btn.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        btn.Cursor = Cursors.Hand
        btn.BackColor = bg
        btn.ForeColor = fg
        btn.FlatAppearance.BorderSize = 0
    End Sub
#End Region

    ' =========================================================================
    ' MOTOR DE EXPORTACIÓN A PDF (PEDIDOS)
    ' =========================================================================
    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        ' OJO: Cambia TextBoxPedido por el nombre real de tu TextBox del ID del pedido
        If String.IsNullOrWhiteSpace(TextBoxPedido.Text) Then
            MessageBox.Show("No hay ningún documento cargado para imprimir.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End If

        docImprimir.DefaultPageSettings.Landscape = False
        docImprimir.DefaultPageSettings.PaperSize = New Printing.PaperSize("A4", 827, 1169)
        docImprimir.DocumentName = "Pedido_" & TextBoxPedido.Text

        Dim ppd As New PrintPreviewDialog()
        ppd.Document = docImprimir
        ppd.Width = 900 : ppd.Height = 700
        CType(ppd, Form).WindowState = FormWindowState.Maximized
        ppd.ShowDialog()
    End Sub

    Private Sub docImprimir_BeginPrint(sender As Object, e As Printing.PrintEventArgs) Handles docImprimir.BeginPrint
        _filaActualImpresion = 0
    End Sub

    Private Sub docImprimir_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles docImprimir.PrintPage
        Dim g As Graphics = e.Graphics

        ' --- FUENTES PROFESIONALES ---
        Dim fTituloDoc As New Font("Segoe UI", 24, FontStyle.Bold)
        Dim fEmpresa As New Font("Segoe UI", 14, FontStyle.Bold)
        Dim fCabecera As New Font("Segoe UI", 10, FontStyle.Bold)
        Dim fClienteNombre As New Font("Segoe UI", 12, FontStyle.Bold)
        Dim fNormal As New Font("Segoe UI", 10, FontStyle.Regular)
        Dim fFila As New Font("Segoe UI", 9, FontStyle.Regular)
        Dim fTotalGordo As New Font("Segoe UI", 14, FontStyle.Bold)

        ' --- COLORES Y PINCELES ---
        Dim bNegro As New SolidBrush(Color.Black)
        Dim bGrisOscuro As New SolidBrush(Color.FromArgb(70, 75, 80))
        Dim bAzulCorporativo As New SolidBrush(Color.FromArgb(40, 50, 70))
        Dim bBlanco As New SolidBrush(Color.White)
        Dim bGrisClaro As New SolidBrush(Color.FromArgb(245, 245, 245))

        Dim lapizFino As New Pen(Color.FromArgb(220, 220, 220), 1)
        Dim lapizGrueso As New Pen(Color.FromArgb(40, 50, 70), 2)

        Dim margenIzq As Integer = 50
        Dim margenDer As Integer = 777
        Dim anchoPagina As Integer = 727
        Dim yPos As Integer = 50

        Dim formatoIzquierda As New StringFormat() With {.Alignment = StringAlignment.Near, .LineAlignment = StringAlignment.Center}
        Dim formatoDerecha As New StringFormat() With {.Alignment = StringAlignment.Far, .LineAlignment = StringAlignment.Center}
        Dim formatoCentro As New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}

        ' ===========================================================
        ' 1. CABECERA DEL DOCUMENTO
        ' ===========================================================
        If _filaActualImpresion = 0 Then
            ' A) EMPRESA
            Dim empNombre As String = "EMPRESA NO CONFIGURADA", empCIF As String = "", empDireccion As String = "", empCP_Pob As String = "", empTelefono As String = ""
            Try
                Dim c = ConexionBD.GetConnection()
                If c.State <> ConnectionState.Open Then c.Open()
                Dim sql As String = "SELECT NombreFiscal, CIF, Direccion, Poblacion, CodigoPostal, Telefono FROM Empresa LIMIT 1"
                Using cmd As New SQLiteCommand(sql, c)
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            empNombre = If(IsDBNull(reader("NombreFiscal")), "", reader("NombreFiscal").ToString())
                            empCIF = If(IsDBNull(reader("CIF")), "", reader("CIF").ToString())
                            empDireccion = If(IsDBNull(reader("Direccion")), "", reader("Direccion").ToString())
                            Dim cp As String = If(IsDBNull(reader("CodigoPostal")), "", reader("CodigoPostal").ToString())
                            Dim pob As String = If(IsDBNull(reader("Poblacion")), "", reader("Poblacion").ToString())
                            empCP_Pob = (cp & " " & pob).Trim()
                            empTelefono = If(IsDBNull(reader("Telefono")), "", reader("Telefono").ToString())
                        End If
                    End Using
                End Using
            Catch ex As Exception
            End Try

            g.DrawString(empNombre, fEmpresa, bAzulCorporativo, margenIzq, yPos)
            g.DrawString("CIF: " & empCIF, fNormal, bGrisOscuro, margenIzq, yPos + 25)
            Dim bloqueContacto As String = empDireccion
            If empCP_Pob <> "" Then bloqueContacto &= vbCrLf & empCP_Pob
            If empTelefono <> "" Then bloqueContacto &= vbCrLf & "Tlf: " & empTelefono
            g.DrawString(bloqueContacto, fNormal, bGrisOscuro, margenIzq, yPos + 45)
            ' B) DATOS DEL DOCUMENTO (PEDIDO)
            g.DrawString("PEDIDO", fTituloDoc, bAzulCorporativo, margenDer, yPos, formatoDerecha)

            ' Datos principales
            g.DrawString("Nº Documento: " & TextBoxPedido.Text, fCabecera, bNegro, margenDer, yPos + 40, formatoDerecha)
            g.DrawString("Fecha: " & TextBoxFecha.Text, fNormal, bGrisOscuro, margenDer, yPos + 60, formatoDerecha)

            ' Fecha de Entrega (Destacada en azul)
            ' OJO: Pon el nombre real de tu DateTimePicker de la fecha de entrega
            g.DrawString("Fecha Entrega: " & DateTimePickerFecha.Value.ToShortDateString(), fCabecera, bAzulCorporativo, margenDer, yPos + 80, formatoDerecha)

            ' Trazabilidad y Vendedor
            ' OJO: Pon el nombre real del TextBox donde pone PRE-024
            If TextBoxIdPresupuesto IsNot Nothing AndAlso TextBoxIdPresupuesto.Text <> "" Then
                g.DrawString("Origen: Presupuesto " & TextBoxIdPresupuesto.Text, fNormal, bGrisOscuro, margenDer, yPos + 105, formatoDerecha)
            End If
            g.DrawString("Comercial: " & TextBoxVendedor.Text, fNormal, bGrisOscuro, margenDer, yPos + 125, formatoDerecha)

            yPos += 145 ' Bajamos la línea para que respire
            g.DrawLine(lapizGrueso, margenIzq, yPos, margenDer, yPos)
            yPos += 20

            ' C) CLIENTE
            Dim cliNombre As String = TextBoxCliente.Text
            Dim cliCIF As String = "", cliDireccion As String = "", cliPoblacion As String = "", cliContacto As String = ""

            Try
                Dim c = ConexionBD.GetConnection()
                If c.State <> ConnectionState.Open Then c.Open()
                Dim sqlCli As String = "SELECT NombreFiscal, CIF, Direccion, Poblacion, Provincia, Telefono, Email FROM Clientes WHERE NombreFiscal = @filtro OR CodigoCliente = @filtro LIMIT 1"
                Using cmd As New SQLiteCommand(sqlCli, c)
                    cmd.Parameters.AddWithValue("@filtro", TextBoxCliente.Text.Trim())
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            cliNombre = If(IsDBNull(reader("NombreFiscal")), TextBoxCliente.Text, reader("NombreFiscal").ToString())
                            cliCIF = If(IsDBNull(reader("CIF")), "", "CIF: " & reader("CIF").ToString())
                            cliDireccion = If(IsDBNull(reader("Direccion")), "", reader("Direccion").ToString())
                            Dim pob As String = If(IsDBNull(reader("Poblacion")), "", reader("Poblacion").ToString())
                            Dim prov As String = If(IsDBNull(reader("Provincia")), "", reader("Provincia").ToString())
                            cliPoblacion = If(pob <> "" And prov <> "", pob & " (" & prov & ")", pob & prov)
                            Dim tel As String = If(IsDBNull(reader("Telefono")), "", "Tlf: " & reader("Telefono").ToString())
                            Dim mail As String = If(IsDBNull(reader("Email")), "", reader("Email").ToString())
                            cliContacto = If(tel <> "" And mail <> "", tel & " | " & mail, tel & mail)
                        End If
                    End Using
                End Using
            Catch ex As Exception
            End Try

            g.FillRectangle(bGrisClaro, margenIzq, yPos, anchoPagina, 145)
            g.DrawRectangle(lapizFino, margenIzq, yPos, anchoPagina, 145)
            g.DrawString("DATOS DEL CLIENTE:", fCabecera, bAzulCorporativo, margenIzq + 15, yPos + 10)
            g.DrawString(cliNombre, fClienteNombre, bNegro, margenIzq + 15, yPos + 35)

            Dim xDatos As Integer = margenIzq + 15
            Dim yDatos As Integer = yPos + 62
            Dim interlineado As Integer = 18

            If cliCIF <> "" Then g.DrawString(cliCIF, fNormal, bGrisOscuro, xDatos, yDatos) : yDatos += interlineado
            If cliDireccion <> "" Then g.DrawString(cliDireccion, fNormal, bGrisOscuro, xDatos, yDatos) : yDatos += interlineado
            If cliPoblacion <> "" Then g.DrawString(cliPoblacion, fNormal, bGrisOscuro, xDatos, yDatos) : yDatos += interlineado
            If cliContacto <> "" Then g.DrawString(cliContacto, fNormal, bGrisOscuro, xDatos, yDatos)

            yPos += 165
        Else
            yPos += 30
        End If

        ' ===========================================================
        ' 2. CABECERA DE LA TABLA
        ' ===========================================================
        Dim rectCant As New Rectangle(margenIzq, yPos, 60, 30)
        Dim rectDesc As New Rectangle(margenIzq + 60, yPos, 340, 30)
        Dim rectPrecio As New Rectangle(margenIzq + 400, yPos, 110, 30)
        Dim rectDto As New Rectangle(margenIzq + 510, yPos, 70, 30)
        Dim rectTotal As New Rectangle(margenIzq + 580, yPos, 147, 30)

        g.FillRectangle(bAzulCorporativo, margenIzq, yPos, anchoPagina, 30)

        g.DrawString("Cant.", fCabecera, bBlanco, rectCant, formatoCentro)
        g.DrawString("Descripción", fCabecera, bBlanco, rectDesc, formatoIzquierda)
        g.DrawString("Precio", fCabecera, bBlanco, rectPrecio, formatoDerecha)
        g.DrawString("% Dto", fCabecera, bBlanco, rectDto, formatoDerecha)
        g.DrawString("Total", fCabecera, bBlanco, rectTotal, formatoDerecha)

        yPos += 35

        ' ===========================================================
        ' 3. RECORRER LÍNEAS DEL DATAGRIDVIEW
        ' ===========================================================
        While _filaActualImpresion < DataGridView1.Rows.Count
            Dim row As DataGridViewRow = DataGridView1.Rows(_filaActualImpresion)
            If row.IsNewRow Then Exit While

            rectCant.Y = yPos : rectDesc.Y = yPos : rectPrecio.Y = yPos : rectDto.Y = yPos : rectTotal.Y = yPos

            ' ASEGÚRATE DE QUE LOS NOMBRES INTERNOS DE ESTAS COLUMNAS SEAN LOS MISMOS QUE EN PRESUPUESTOS
            Dim cant As String = If(row.Cells("Cantidad").Value IsNot Nothing, row.Cells("Cantidad").Value.ToString(), "0")
            Dim desc As String = If(row.Cells("Descripcion").Value IsNot Nothing, row.Cells("Descripcion").Value.ToString(), "")
            Dim precio As String = If(row.Cells("PrecioUnitario").Value IsNot Nothing, Convert.ToDecimal(row.Cells("PrecioUnitario").Value).ToString("N2") & " €", "0,00 €")
            Dim dto As String = If(row.Cells("Descuento").Value IsNot Nothing, row.Cells("Descuento").Value.ToString() & " %", "0 %")
            Dim totalLinea As String = If(row.Cells("Total").Value IsNot Nothing, Convert.ToDecimal(row.Cells("Total").Value).ToString("N2") & " €", "0,00 €")

            If desc.Length > 55 Then desc = desc.Substring(0, 52) & "..."

            g.DrawString(cant, fFila, bNegro, rectCant, formatoCentro)
            g.DrawString(desc, fFila, bNegro, rectDesc, formatoIzquierda)
            g.DrawString(precio, fFila, bNegro, rectPrecio, formatoDerecha)
            g.DrawString(dto, fFila, bNegro, rectDto, formatoDerecha)
            g.DrawString(totalLinea, fFila, bNegro, rectTotal, formatoDerecha)

            yPos += 25
            g.DrawLine(lapizFino, margenIzq, yPos, margenDer, yPos)
            yPos += 5

            _filaActualImpresion += 1

            If yPos > 1000 AndAlso _filaActualImpresion < DataGridView1.Rows.Count Then
                e.HasMorePages = True
                Return
            End If
        End While

        ' ===========================================================
        ' 4. TOTALES
        ' ===========================================================
        If _filaActualImpresion >= DataGridView1.Rows.Count Then
            yPos += 20

            ' Asegúrate de que las cajas de texto de los totales se llaman así
            Dim totalBase As String = TextBoxBase.Text.Replace("Base imponible :", "").Trim()
            Dim totalIVA As String = TextBoxIva.Text.Replace("I.V.A :", "").Trim()
            Dim totalDoc As String = TextBoxTotalPed.Text.Replace("TOTAL :", "").Trim()

            Dim anchoCajaTotales As Integer = 300
            Dim xCaja As Integer = margenDer - anchoCajaTotales
            ' --- CONDICIONES Y OBSERVACIONES (Abajo a la izquierda) ---
            Dim yNotas As Integer = yPos

            ' Forma de Pago
            If cboFormaPago.Text <> "" Then
                g.DrawString("Forma de pago:", fCabecera, bAzulCorporativo, margenIzq, yNotas)
                g.DrawString(cboFormaPago.Text, fNormal, bNegro, margenIzq + 110, yNotas)
                yNotas += 35
            End If

            ' Observaciones del pedido
            If TextBoxObservaciones.Text.Trim() <> "" Then
                g.DrawString("Observaciones:", fCabecera, bAzulCorporativo, margenIzq, yNotas)

                ' Rectángulo para que haga saltos de línea automáticos si el texto es largo
                Dim rectObs As New Rectangle(margenIzq, yNotas + 20, anchoCajaTotales - 20, 60)
                g.DrawString(TextBoxObservaciones.Text, fNormal, bGrisOscuro, rectObs)
            End If
            g.DrawString("Base Imponible:", fNormal, bGrisOscuro, xCaja, yPos)
            g.DrawString(totalBase, fNormal, bNegro, margenDer, yPos, formatoDerecha)
            yPos += 25

            g.DrawString("Impuestos (IVA):", fNormal, bGrisOscuro, xCaja, yPos)
            g.DrawString(totalIVA, fNormal, bNegro, margenDer, yPos, formatoDerecha)
            yPos += 30

            g.FillRectangle(bAzulCorporativo, xCaja, yPos, anchoCajaTotales, 45)

            Dim rectTxtTotal As New Rectangle(xCaja + 15, yPos, 150, 45)
            Dim rectValTotal As New Rectangle(xCaja + 150, yPos, anchoCajaTotales - 165, 45)

            g.DrawString("TOTAL A PAGAR", fCabecera, bBlanco, rectTxtTotal, formatoIzquierda)
            g.DrawString(totalDoc, fTotalGordo, bBlanco, rectValTotal, formatoDerecha)

            yPos += 80
            g.DrawString("Gracias por confiar en nuestros servicios.", fFila, bGrisOscuro, margenIzq, yPos)

            e.HasMorePages = False
        End If
    End Sub
End Class