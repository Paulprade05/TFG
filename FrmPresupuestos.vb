Imports System.Data.SQLite

Public Class FrmPresupuestos

#Region "1. Variables Globales y Propiedades"
    Private _numeroPresupuestoActual As String = ""
    Private _dtLineas As DataTable
    Private _idsParaBorrar As New List(Of Integer)

    Private WithEvents cboFormaPago As New ComboBox()
    'Private WithEvents cboRuta As New ComboBox()
    Private lblFormaPago As New Label() With {.Text = "Forma de Pago", .AutoSize = True, .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold), .ForeColor = Color.WhiteSmoke}
    'Private lblRuta As New Label() With {.Text = "Ruta Asignada", .AutoSize = True, .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold), .ForeColor = Color.WhiteSmoke}
#End Region

#Region "2. Eventos de Inicialización (Load)"
    Private Sub FrmPresupuestos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ConfigurarDiseñoResponsive()
        EstilizarGrid(DataGridView1)

        Me.Controls.Add(lblFormaPago) : Me.Controls.Add(cboFormaPago)
        'Me.Controls.Add(lblRuta) : Me.Controls.Add(cboRuta)
        cboFormaPago.DropDownStyle = ComboBoxStyle.DropDownList
        'cboRuta.DropDownStyle = ComboBoxStyle.DropDownList
        CargarDesplegables()

        ReorganizarControlesAutomaticamente()
        ConfigurarColumnasGrid()

        Dim ultimoNum As String = ObtenerUltimoNumeroPresupuesto()
        If Not String.IsNullOrEmpty(ultimoNum) Then
            CargarPresupuesto(ultimoNum)
        Else
            LimpiarFormulario()
        End If
    End Sub

    Private Sub CargarDesplegables()
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()
            Dim daPago As New SQLiteDataAdapter("SELECT ID_FormaPago, Descripcion FROM FormasPago WHERE Activo=1", c)
            Dim dtPago As New DataTable() : daPago.Fill(dtPago)
            cboFormaPago.DataSource = dtPago : cboFormaPago.DisplayMember = "Descripcion" : cboFormaPago.ValueMember = "ID_FormaPago" : cboFormaPago.SelectedIndex = -1

            'Dim daRuta As New SQLiteDataAdapter("SELECT ID_Ruta, NombreZona FROM Rutas WHERE Activo=1", c)
            'Dim dtRuta As New DataTable() : daRuta.Fill(dtRuta)
            'cboRuta.DataSource = dtRuta : cboRuta.DisplayMember = "NombreZona" : cboRuta.ValueMember = "ID_Ruta" : cboRuta.SelectedIndex = -1
        Catch ex As Exception
        End Try
    End Sub
#End Region

#Region "3. Configuración UI y Diseño"
    Private Sub ConfigurarDiseñoResponsive()
        DataGridView1.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        Dim controlesTotales As Control() = {LabelBase, TextBoxBase, LabelIva, TextBoxIva, Label7, TextBoxTotalPresup, LabelStock}
        For Each ctrl In controlesTotales : ctrl.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right : Next
        ButtonAnterior.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        ButtonSiguiente.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        Dim botonesIzq As Control() = {ButtonGuardar, ButtonBorrar, ButtonNuevoPresup, ButtonBorrarLineas, ButtonNuevaLinea}
        For Each ctrl In botonesIzq : ctrl.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left : Next
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
        If TextBoxTotalPresup IsNot Nothing Then TextBoxTotalPresup.Text = (base + sumaIva).ToString("C2")
    End Sub

    Private Sub LimpiarFormulario()
        _numeroPresupuestoActual = ""
        _idsParaBorrar.Clear()
        TextBoxPresupuesto.Text = GenerarProximoNumero()

        TextBoxCliente.Text = "" : TextBoxIdCliente.Text = "" : TextBoxCliente.Tag = Nothing
        TextBoxVendedor.Text = "" : TextBoxIdVendedor.Text = ""
        TextBoxObservaciones.Text = ""
        TextBoxFecha.Text = DateTime.Now.ToShortDateString()
        TextBoxEstado.Text = "Pendiente"
        TextBoxBase.Text = "0,00 €" : TextBoxIva.Text = "0,00 €" : TextBoxTotalPresup.Text = "0,00 €"

        _dtLineas = New DataTable()
        ConfigurarEstructuraDataTable()
        DataGridView1.DataSource = _dtLineas
        If cboFormaPago IsNot Nothing Then cboFormaPago.SelectedIndex = -1
        'If cboRuta IsNot Nothing Then cboRuta.SelectedIndex = -1
        TextBoxIdCliente.Focus()
    End Sub


    Private Sub ConfigurarEstructuraDataTable()
        With _dtLineas.Columns
            If Not .Contains("ID_Linea") Then .Add("ID_Linea", GetType(Object))
            If Not .Contains("NumeroPresupuesto") Then .Add("NumeroPresupuesto", GetType(String))
            If Not .Contains("NumeroOrden") Then .Add("NumeroOrden", GetType(Integer))
            If Not .Contains("ID_Articulo") Then .Add("ID_Articulo", GetType(Object))
            If Not .Contains("Descripcion") Then .Add("Descripcion", GetType(String))
            If Not .Contains("Cantidad") Then .Add("Cantidad", GetType(Decimal))
            If Not .Contains("PrecioUnitario") Then .Add("PrecioUnitario", GetType(Decimal))
            If Not .Contains("Descuento") Then .Add("Descuento", GetType(Decimal))

            ' --- NUEVAS COLUMNAS IVA Y VIRTUALES ---
            If Not .Contains("PorcentajeIVA") Then .Add("PorcentajeIVA", GetType(Decimal))
            If Not .Contains("PrecioConIVA") Then .Add("PrecioConIVA", GetType(Decimal))
            If Not .Contains("Total") Then .Add("Total", GetType(Decimal))
            If Not .Contains("TotalConIVA") Then .Add("TotalConIVA", GetType(Decimal))
        End With
    End Sub

    Private Function GenerarProximoNumero() As String
        Dim prefijo As String = "PRE-"
        Dim nuevoNumero As String = $"{prefijo}001"
        Try
            Dim sql As String = "SELECT NumeroPresupuesto FROM Presupuestos WHERE NumeroPresupuesto LIKE @patron ORDER BY NumeroPresupuesto DESC LIMIT 1"

            ' ---> EL CAMBIO CLAVE: Dim en lugar de Using <---
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
            ' Aquí hemos quitado el "End Using" de la conexión para que NO la destruya

        Catch
            nuevoNumero = $"PRE-{DateTime.Now:HHmmss}"
        End Try
        Return nuevoNumero
    End Function
#End Region

#Region "5. Persistencia (Base de Datos - CRUD)"
    Private Sub CargarPresupuesto(numeroPresupuesto As String)
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()

            ' A. CARGAR CABECERA
            Dim sqlCab As String = "SELECT P.*, C.NombreFiscal AS NombreCliente, V.Nombre AS NombreVendedor " &
                                   "FROM Presupuestos P " &
                                   "LEFT JOIN Clientes C ON P.CodigoCliente = C.CodigoCliente " &
                                   "LEFT JOIN Vendedores V ON P.ID_Vendedor = V.ID_Vendedor " &
                                   "WHERE P.NumeroPresupuesto = @num"

            Using cmd As New SQLiteCommand(sqlCab, c)
                cmd.Parameters.AddWithValue("@num", numeroPresupuesto)
                Using r = cmd.ExecuteReader()
                    If r.Read() Then
                        _numeroPresupuestoActual = numeroPresupuesto
                        TextBoxPresupuesto.Text = r("NumeroPresupuesto").ToString()
                        TextBoxFecha.Text = If(IsDBNull(r("Fecha")), "", Convert.ToDateTime(r("Fecha")).ToShortDateString())
                        TextBoxObservaciones.Text = r("Observaciones").ToString()
                        TextBoxEstado.Text = r("Estado").ToString()
                        TextBoxIdCliente.Text = r("CodigoCliente").ToString()
                        TextBoxCliente.Text = r("NombreCliente").ToString()
                        TextBoxIdVendedor.Text = r("ID_Vendedor").ToString()
                        TextBoxVendedor.Text = r("NombreVendedor").ToString()

                        Dim idPago = r("ID_FormaPago")
                        If Not IsDBNull(idPago) Then cboFormaPago.SelectedValue = Convert.ToInt32(idPago) Else cboFormaPago.SelectedIndex = -1

                        'Dim idRut = r("ID_Ruta")
                        'If Not IsDBNull(idRut) Then cboRuta.SelectedValue = Convert.ToInt32(idRut) Else cboRuta.SelectedIndex = -1
                    Else
                        MessageBox.Show("Presupuesto no encontrado.") : Return
                    End If
                End Using
            End Using

            ' B. CARGAR LÍNEAS
            Dim sqlLin As String = "SELECT * FROM LineasPresupuesto WHERE NumeroPresupuesto = @num ORDER BY NumeroOrden ASC"
            Using cmd As New SQLiteCommand(sqlLin, c)
                cmd.Parameters.AddWithValue("@num", numeroPresupuesto)
                Dim da As New SQLiteDataAdapter(cmd)
                _dtLineas = New DataTable()
                da.Fill(_dtLineas)
                ConfigurarEstructuraDataTable()

                ' Magia para calcular IVA al vuelo en presupuestos viejos
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

    Private Sub GuardarPresupuesto()
        If String.IsNullOrWhiteSpace(TextBoxIdCliente.Text) Then MessageBox.Show("Falta el Cliente") : Return

        Dim esNuevo As Boolean = String.IsNullOrEmpty(_numeroPresupuestoActual)
        If esNuevo Then
            TextBoxPresupuesto.Text = GenerarProximoNumero()
            _numeroPresupuestoActual = TextBoxPresupuesto.Text
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
            'Dim idRuta As Object = If(cboRuta.SelectedValue IsNot Nothing AndAlso cboRuta.SelectedIndex <> -1, cboRuta.SelectedValue, DBNull.Value)
            Dim idVend As Object = If(IsNumeric(TextBoxIdVendedor.Text) AndAlso Val(TextBoxIdVendedor.Text) > 0, Convert.ToInt32(TextBoxIdVendedor.Text), DBNull.Value)

            Dim sql As String = ""
            If esNuevo Then
                sql = "INSERT INTO Presupuestos (NumeroPresupuesto, CodigoCliente, ID_Vendedor, Fecha, Observaciones, Estado, ID_FormaPago, BaseImponible, ImporteIVA, TotalPresupuesto) " &
                      "VALUES (@num, @cli, @vend, @fecha, @obs, @est, @formaPago, @base, @iva, @total)"
            Else
                sql = "UPDATE Presupuestos SET CodigoCliente=@cli, ID_Vendedor=@vend, Fecha=@fecha, Observaciones=@obs, Estado=@est, ID_FormaPago=@formaPago, BaseImponible=@base, ImporteIVA=@iva, TotalPresupuesto=@total " &
                      "WHERE NumeroPresupuesto = @num"
            End If

            Using cmd As New SQLiteCommand(sql, c)
                cmd.Transaction = trans
                cmd.Parameters.AddWithValue("@num", _numeroPresupuestoActual)
                cmd.Parameters.AddWithValue("@cli", TextBoxIdCliente.Text.Trim())
                cmd.Parameters.AddWithValue("@vend", idVend)

                Dim fecha As DateTime
                If DateTime.TryParse(TextBoxFecha.Text, fecha) Then cmd.Parameters.AddWithValue("@fecha", fecha.ToString("yyyy-MM-dd HH:mm:ss")) Else cmd.Parameters.AddWithValue("@fecha", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

                cmd.Parameters.AddWithValue("@obs", TextBoxObservaciones.Text.Trim())
                cmd.Parameters.AddWithValue("@est", TextBoxEstado.Text.Trim())
                cmd.Parameters.AddWithValue("@formaPago", idFormaPago)
                'cmd.Parameters.AddWithValue("@ruta", idRuta)

                cmd.Parameters.AddWithValue("@base", sumaBase)
                cmd.Parameters.AddWithValue("@iva", sumaIva)
                cmd.Parameters.AddWithValue("@total", sumaBase + sumaIva)
                cmd.ExecuteNonQuery()
            End Using

            ' 2. Borrar líneas eliminadas
            For Each idDel In _idsParaBorrar
                Using cmdDel As New SQLiteCommand("DELETE FROM LineasPresupuesto WHERE ID_Linea = @id", c)
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
                    sqlLine = "INSERT INTO LineasPresupuesto (NumeroPresupuesto, NumeroOrden, ID_Articulo, Descripcion, Cantidad, PrecioUnitario, Descuento, PorcentajeIVA, Total) " &
                              "VALUES (@num, @ord, @art, @desc, @cant, @prec, @dcto, @iva, @tot)"
                Else
                    sqlLine = "UPDATE LineasPresupuesto SET NumeroOrden=@ord, ID_Articulo=@art, Descripcion=@desc, Cantidad=@cant, PrecioUnitario=@prec, Descuento=@dcto, PorcentajeIVA=@iva, Total=@tot WHERE ID_Linea=@id"
                End If

                Using cmdL As New SQLiteCommand(sqlLine, c)
                    cmdL.Transaction = trans
                    cmdL.Parameters.AddWithValue("@num", _numeroPresupuestoActual)
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

            trans.Commit()
            MessageBox.Show("Guardado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
            CargarPresupuesto(_numeroPresupuestoActual)

        Catch ex As Exception
            If trans IsNot Nothing Then trans.Rollback()
            MessageBox.Show("Error al guardar: " & ex.Message)
        End Try
    End Sub

    Private Function ObtenerUltimoNumeroPresupuesto() As String
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()
            Dim cmd As New SQLiteCommand("SELECT MAX(NumeroPresupuesto) FROM Presupuestos", c)
            Dim r = cmd.ExecuteScalar()
            Return If(IsDBNull(r), "", r.ToString())
        Catch
            Return ""
        End Try
    End Function
#End Region

#Region "6. Manejadores de Eventos (Botones y Grid)"

    Private Sub ButtonGuardar_Click(sender As Object, e As EventArgs) Handles ButtonGuardar.Click
        GuardarPresupuesto()
    End Sub

    Private Sub ButtonNuevoPresup_Click(sender As Object, e As EventArgs) Handles ButtonNuevoPresup.Click
        LimpiarFormulario()
    End Sub

    Private Sub ButtonBorrar_Click(sender As Object, e As EventArgs) Handles ButtonBorrar.Click
        If String.IsNullOrEmpty(_numeroPresupuestoActual) Then Return
        If MessageBox.Show("¿Seguro que deseas eliminar este presupuesto?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes Then
            Try
                Dim c = ConexionBD.GetConnection()
                If c.State <> ConnectionState.Open Then c.Open()
                Dim cmd As New SQLiteCommand("DELETE FROM LineasPresupuesto WHERE NumeroPresupuesto=@num; DELETE FROM Presupuestos WHERE NumeroPresupuesto=@num;", c)
                cmd.Parameters.AddWithValue("@num", _numeroPresupuestoActual)
                cmd.ExecuteNonQuery()
                LimpiarFormulario()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub

    Private Sub btnBuscarPresupuesto_Click(sender As Object, e As EventArgs) Handles btnBuscarPresupuesto.Click
        Dim frmBuscar As New FrmBuscador()
        frmBuscar.TablaABuscar = "Presupuestos"

        If frmBuscar.ShowDialog() = DialogResult.OK Then
            CargarPresupuesto(frmBuscar.Resultado)
        End If
    End Sub

    Private Sub Navegar(direccion As String)
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()
            Dim sql As String = If(direccion = "ANT",
                                   "SELECT MAX(NumeroPresupuesto) FROM Presupuestos WHERE NumeroPresupuesto < @act",
                                   "SELECT MIN(NumeroPresupuesto) FROM Presupuestos WHERE NumeroPresupuesto > @act")

            If String.IsNullOrEmpty(_numeroPresupuestoActual) And direccion = "ANT" Then
                sql = "SELECT MAX(NumeroPresupuesto) FROM Presupuestos"
            End If

            Using cmd As New SQLiteCommand(sql, c)
                cmd.Parameters.AddWithValue("@act", _numeroPresupuestoActual)
                Dim res = cmd.ExecuteScalar()
                If res IsNot Nothing AndAlso Not IsDBNull(res) Then CargarPresupuesto(res.ToString())
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
        fila("NumeroPresupuesto") = _numeroPresupuestoActual
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

    Private Sub DataGridView1_RowEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.RowEnter
        If DataGridView1.Rows(e.RowIndex).Tag IsNot Nothing Then
            Dim stock As Decimal = Val(DataGridView1.Rows(e.RowIndex).Tag)
            LabelStock.Text = If(stock <= 0, "⚠️ ¡SIN STOCK!", $"Stock Disponible: {stock}")
            LabelStock.ForeColor = If(stock <= 0, Color.Red, Color.DarkGreen)
        Else
            LabelStock.Text = ""
        End If
    End Sub

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
                    'Dim idRut = r("ID_Ruta")
                    'If Not IsDBNull(idRut) Then cboRuta.SelectedValue = Convert.ToInt32(idRut)
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
                    Case "presupuesto" : ctrl.Location = New Point(col1_X, yFila1)
                    Case "cliente" : ctrl.Location = New Point(col2_X, yFila1)
                    Case "fecha" : ctrl.Location = New Point(col3_X, yFila1)
                    Case "vendedor" : ctrl.Location = New Point(col1_X, yFila3)
                    Case "observaciones" : ctrl.Location = New Point(col2_X, yFila3)
                    Case "estado" : ctrl.Location = New Point(col3_X, yFila3)
                End Select
            End If
        Next

        TextBoxPresupuesto.Bounds = New Rectangle(col1_X, yFila2, 105, 25)
        btnBuscarPresupuesto.Bounds = New Rectangle(col1_X + 110, yFila2, 30, 25)
        TextBoxIdCliente.Bounds = New Rectangle(col2_X, yFila2, 60, 25)
        TextBoxCliente.Bounds = New Rectangle(col2_X + 70, yFila2, 460, 25)
        TextBoxFecha.Bounds = New Rectangle(col3_X, yFila2, 140, 25)

        TextBoxIdVendedor.Bounds = New Rectangle(col1_X, yFila4, 50, 25)
        TextBoxVendedor.Bounds = New Rectangle(col1_X + 55, yFila4, 85, 25)
        TextBoxObservaciones.Bounds = New Rectangle(col2_X, yFila4, 530, 25)
        TextBoxEstado.Bounds = New Rectangle(col3_X, yFila4, 140, 25)

        lblFormaPago.Location = New Point(col1_X, yFila5)
        cboFormaPago.Bounds = New Rectangle(col1_X, yFila6, 140, 25)
        cboFormaPago.Font = New Font("Segoe UI", 10.5F)

        'lblRuta.Location = New Point(col2_X, yFila5)
        ' cboRuta.Bounds = New Rectangle(col2_X, yFila6, 530, 25)
        ' cboRuta.Font = New Font("Segoe UI", 10.5F)

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
        panelTotales.Bounds = New Rectangle(xDerecha - 320, yTotales, 150, 115) ' Ajustado dinamicamente luego
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

        If TextBoxTotalPresup IsNot Nothing Then
            TextBoxTotalPresup.Bounds = New Rectangle(xDerecha - 140, yTotales + 80, 120, 30)
            TextBoxTotalPresup.TextAlign = HorizontalAlignment.Right
            TextBoxTotalPresup.Font = New Font("Segoe UI", 14, FontStyle.Bold) : TextBoxTotalPresup.ForeColor = colorAcento
            TextBoxTotalPresup.BackColor = Color.FromArgb(25, 30, 40) : TextBoxTotalPresup.BorderStyle = BorderStyle.None
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

        LabelStock.Location = New Point(margenIzq, DataGridView1.Bottom + 10)

        DataGridView1.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        Dim botonesAbajo As Control() = {ButtonGuardar, ButtonBorrar, ButtonNuevoPresup, ButtonBorrarLineas, ButtonNuevaLinea, ButtonAnterior, ButtonSiguiente, LabelStock}
        For Each b In botonesAbajo : b.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left : Next

        TextBoxBase.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        TextBoxIva.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        If TextBoxTotalPresup IsNot Nothing Then TextBoxTotalPresup.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
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

End Class