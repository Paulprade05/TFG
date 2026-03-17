Imports System.Data.SQLite

Public Class FrmPresupuestos

#Region "1. Variables Globales y Propiedades"
    Private _numeroPresupuestoActual As String = ""
    Private _dtLineas As DataTable
    Private _idsParaBorrar As New List(Of Integer)

    ' --- NUEVOS DESPLEGABLES ---
    Private WithEvents cboFormaPago As New ComboBox()
    Private WithEvents cboRuta As New ComboBox()
    Private lblFormaPago As New Label() With {.Text = "Forma de Pago", .AutoSize = True, .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold), .ForeColor = Color.WhiteSmoke}
    Private lblRuta As New Label() With {.Text = "Ruta Asignada", .AutoSize = True, .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold), .ForeColor = Color.WhiteSmoke}
#End Region

#Region "2. Eventos de Inicialización (Load)"
    Private Sub FrmPresupuestos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ConfigurarDiseñoResponsive()
        EstilizarGrid(DataGridView1)

        ' Añadimos los combos al formulario ANTES de cargar datos
        Me.Controls.Add(lblFormaPago) : Me.Controls.Add(cboFormaPago)
        Me.Controls.Add(lblRuta) : Me.Controls.Add(cboRuta)
        cboFormaPago.DropDownStyle = ComboBoxStyle.DropDownList
        cboRuta.DropDownStyle = ComboBoxStyle.DropDownList
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
#End Region

#Region "3. Configuración UI y Diseño (Responsive)"
    ''' <summary>
    ''' Configura los anclajes (Anchors) para que el formulario se adapte a cualquier pantalla.
    ''' </summary>
    Private Sub ConfigurarDiseñoResponsive()
        ' 1. TABLA: Se estira hacia todos lados
        DataGridView1.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right

        ' 2. CABECERA: Todo anclado a la izquierda (Top, Left) para que NO se choquen.
        ' Quitamos el AnchorRight para evitar que los campos se superpongan.
        'TextBoxCliente.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        'TextBoxObservaciones.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        'TextBoxEstado.Anchor = AnchorStyles.Top Or AnchorStyles.Left
        'Label6.Anchor = AnchorStyles.Top Or AnchorStyles.Left

        ' 3. TOTALES: Se quedan pegados abajo a la derecha
        Dim controlesTotales As Control() = {LabelBase, TextBoxBase, LabelIva, TextBoxIva, Label7, TextBoxTotalPresup, LabelStock}
        For Each ctrl In controlesTotales
            ctrl.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        Next

        ' 4. BOTONES NAVEGACIÓN: Abajo a la derecha (junto a los totales)
        ButtonAnterior.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        ButtonSiguiente.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right

        ' 5. BOTONES ACCIÓN: Abajo a la izquierda
        Dim botonesIzq As Control() = {ButtonGuardar, ButtonBorrar, ButtonNuevoPresup, ButtonBorrarLineas, ButtonNuevaLinea}
        For Each ctrl In botonesIzq
            ctrl.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
        Next
    End Sub

    ''' <summary>
    ''' Define las columnas del Grid. Descripción se estira, importes fijos.
    ''' </summary>
    Private Sub ConfigurarColumnasGrid()
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.Columns.Clear()
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None ' Control manual total

        ' Columna ID Oculta
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "ID_Linea", .DataPropertyName = "ID_Linea", .Visible = False})

        ' Columna Nº (Pequeña y centrada)
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {
            .Name = "NumeroOrden", .DataPropertyName = "NumeroOrden", .HeaderText = "Nº", .Width = 40, .ReadOnly = True,
            .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleCenter}
        })

        ' Columna ID Art
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "ID_Articulo", .DataPropertyName = "ID_Articulo", .HeaderText = "ID Art", .Width = 70})

        ' Columna Descripción (FILL - Ocupa todo el espacio sobrante)
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {
            .Name = "Descripcion", .DataPropertyName = "Descripcion", .HeaderText = "Descripción",
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        })

        ' Columnas Numéricas (Alineadas a la derecha con formato)
        DataGridView1.Columns.Add(CrearColumnaNumerica("Cantidad", "Cant.", 60, "N2"))
        DataGridView1.Columns.Add(CrearColumnaNumerica("PrecioUnitario", "Precio", 80, "C2"))
        DataGridView1.Columns.Add(CrearColumnaNumerica("Descuento", "% Dto", 60, "N2"))

        ' Columna Total (Lectura, gris y negrita)
        Dim colTotal = CrearColumnaNumerica("Total", "Total", 100, "C2")
        colTotal.ReadOnly = True
        colTotal.DefaultCellStyle.BackColor = Color.WhiteSmoke
        colTotal.DefaultCellStyle.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        DataGridView1.Columns.Add(colTotal)
    End Sub

    ' Helper para crear columnas rápido y limpio
    Private Function CrearColumnaNumerica(name As String, header As String, width As Integer, format As String) As DataGridViewTextBoxColumn
        Return New DataGridViewTextBoxColumn() With {
            .Name = name, .DataPropertyName = name, .HeaderText = header, .Width = width,
            .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = format}
        }
    End Function

    Public Shared Sub EstilizarGrid(dgv As DataGridView)
        ' 1. CONFIGURACIÓN BASE
        dgv.BackgroundColor = Color.White
        dgv.BorderStyle = BorderStyle.None
        dgv.CellBorderStyle = DataGridViewCellBorderStyle.None

        ' 2. CABECERA (VOLVEMOS AL COLOR ANTERIOR)
        dgv.EnableHeadersVisualStyles = False
        dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None

        ' Tono Gris Azulado (El que te gustaba, a juego con el menú)
        dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 55, 65)
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        dgv.ColumnHeadersHeight = 40

        ' 3. FILAS Y SELECCIÓN (Mantenemos Cebra y Azul Suave)
        dgv.DefaultCellStyle.BackColor = Color.White
        dgv.DefaultCellStyle.ForeColor = Color.Black
        dgv.DefaultCellStyle.Font = New Font("Segoe UI", 10, FontStyle.Regular)

        ' Selección Azul Pastel Suave
        dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(200, 230, 255)
        dgv.DefaultCellStyle.SelectionForeColor = Color.Black

        ' 4. EFECTO CEBRA (Filas alternas)
        ' Gris muy muy suave para que se note la línea pero no ensucie
        dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245)
        dgv.AlternatingRowsDefaultCellStyle.ForeColor = Color.Black

        ' La selección en las filas alternas también debe ser azul pastel
        dgv.AlternatingRowsDefaultCellStyle.SelectionBackColor = Color.FromArgb(200, 230, 255)
        dgv.AlternatingRowsDefaultCellStyle.SelectionForeColor = Color.Black

        ' 5. AJUSTES FINALES
        dgv.RowHeadersVisible = False
        dgv.RowTemplate.Height = 35
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.MultiSelect = False
    End Sub
#End Region

#Region "4. Lógica de Negocio y Cálculos"
    Private Sub RecalcularFila(ByRef fila As DataGridViewRow)
        If fila.IsNewRow Then Return
        Dim c As Decimal = 0 : Decimal.TryParse(fila.Cells("Cantidad").Value?.ToString(), c)
        Dim p As Decimal = 0 : Decimal.TryParse(fila.Cells("PrecioUnitario").Value?.ToString(), p)
        Dim d As Decimal = 0 : Decimal.TryParse(fila.Cells("Descuento").Value?.ToString(), d)
        fila.Cells("Total").Value = (c * p) * (1 - (d / 100))
    End Sub

    Private Sub CalcularTotalesGenerales()
        Dim base As Decimal = 0
        If _dtLineas IsNot Nothing Then
            For Each row As DataRow In _dtLineas.Rows
                If row.RowState <> DataRowState.Deleted Then
                    Dim tot As Decimal = 0
                    ' Intentamos usar la columna calculada, si no, recalculamos
                    If _dtLineas.Columns.Contains("Total") AndAlso IsNumeric(row("Total")) Then
                        tot = Convert.ToDecimal(row("Total"))
                    Else
                        Dim c As Decimal = Val(row("Cantidad").ToString())
                        Dim p As Decimal = Val(row("PrecioUnitario").ToString())
                        Dim d As Decimal = Val(row("Descuento").ToString())
                        tot = (c * p) * (1 - (d / 100))
                    End If
                    base += tot
                End If
            Next
        End If
        TextBoxBase.Text = base.ToString("C2")
        TextBoxIva.Text = (base * 0.21D).ToString("C2")
        TextBoxTotalPresup.Text = (base * 1.21D).ToString("C2")
    End Sub

    Private Sub LimpiarFormulario()
        _numeroPresupuestoActual = ""
        _idsParaBorrar.Clear()
        TextBoxPresupuesto.Text = GenerarProximoNumero()

        ' Limpieza de campos
        TextBoxCliente.Text = "" : TextBoxIdCliente.Text = "" : TextBoxCliente.Tag = Nothing
        TextBoxVendedor.Text = "" : TextBoxIdVendedor.Text = ""
        TextBoxObservaciones.Text = ""
        TextBoxFecha.Text = DateTime.Now.ToShortDateString()
        TextBoxEstado.Text = "Pendiente"
        TextBoxBase.Text = "0,00 €" : TextBoxIva.Text = "0,00 €" : TextBoxTotalPresup.Text = "0,00 €"

        ' Reiniciar DataTable
        _dtLineas = New DataTable()
        ConfigurarEstructuraDataTable()
        DataGridView1.DataSource = _dtLineas
        If cboFormaPago IsNot Nothing Then cboFormaPago.SelectedIndex = -1
        If cboRuta IsNot Nothing Then cboRuta.SelectedIndex = -1
        TextBoxIdCliente.Focus()
    End Sub

    Private Sub ConfigurarEstructuraDataTable()
        ' Definimos el esquema de la tabla en memoria
        With _dtLineas.Columns
            If Not .Contains("ID_Linea") Then .Add("ID_Linea", GetType(Object))
            If Not .Contains("NumeroPresupuesto") Then .Add("NumeroPresupuesto", GetType(String))
            If Not .Contains("NumeroOrden") Then .Add("NumeroOrden", GetType(Integer))
            If Not .Contains("ID_Articulo") Then .Add("ID_Articulo", GetType(Object))
            If Not .Contains("Descripcion") Then .Add("Descripcion", GetType(String))
            If Not .Contains("Cantidad") Then .Add("Cantidad", GetType(Decimal))
            If Not .Contains("PrecioUnitario") Then .Add("PrecioUnitario", GetType(Decimal))
            If Not .Contains("Descuento") Then .Add("Descuento", GetType(Decimal))
            If Not .Contains("Total") Then .Add("Total", GetType(Decimal))
        End With
    End Sub

    Private Function GenerarProximoNumero() As String
        Dim prefijo As String = "PRE-"
        Dim nuevoNumero As String = $"{prefijo}001"
        Try
            Dim sql As String = "SELECT NumeroPresupuesto FROM Presupuestos WHERE NumeroPresupuesto LIKE @patron ORDER BY NumeroPresupuesto DESC LIMIT 1"
            Using c = ConexionBD.GetConnection()
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
            End Using
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
                        ' --- CARGAR COMBOS ---
                        Dim idPago = r("ID_FormaPago")
                        If Not IsDBNull(idPago) Then cboFormaPago.SelectedValue = Convert.ToInt32(idPago) Else cboFormaPago.SelectedIndex = -1

                        Dim idRut = r("ID_Ruta")
                        If Not IsDBNull(idRut) Then cboRuta.SelectedValue = Convert.ToInt32(idRut) Else cboRuta.SelectedIndex = -1
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
                ConfigurarEstructuraDataTable() ' Asegura tipos correctos
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

            ' 1. Guardar Cabecera (Blindada y con Combos)
            Dim idFormaPago As Object = If(cboFormaPago.SelectedValue IsNot Nothing AndAlso cboFormaPago.SelectedIndex <> -1, cboFormaPago.SelectedValue, DBNull.Value)
            Dim idRuta As Object = If(cboRuta.SelectedValue IsNot Nothing AndAlso cboRuta.SelectedIndex <> -1, cboRuta.SelectedValue, DBNull.Value)
            Dim idVend As Object = If(IsNumeric(TextBoxIdVendedor.Text) AndAlso Val(TextBoxIdVendedor.Text) > 0, Convert.ToInt32(TextBoxIdVendedor.Text), DBNull.Value)

            Dim sql As String = ""
            If esNuevo Then
                sql = "INSERT INTO Presupuestos (NumeroPresupuesto, CodigoCliente, ID_Vendedor, Fecha, Observaciones, Estado, ID_FormaPago, ID_Ruta) VALUES (@num, @cli, @vend, @fecha, @obs, @est, @formaPago, @ruta)"
            Else
                sql = "UPDATE Presupuestos SET CodigoCliente=@cli, ID_Vendedor=@vend, Fecha=@fecha, Observaciones=@obs, Estado=@est, ID_FormaPago=@formaPago, ID_Ruta=@ruta WHERE NumeroPresupuesto = @num"
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
                cmd.Parameters.AddWithValue("@ruta", idRuta)
                cmd.ExecuteNonQuery()
            End Using

            ' 2. Borrar líneas eliminadas
            For Each idDel In _idsParaBorrar
                Using cmdDel As New SQLiteCommand("DELETE FROM LineasPresupuesto WHERE ID_Linea = @id", c)
                    cmdDel.Transaction = trans : cmdDel.Parameters.AddWithValue("@id", idDel) : cmdDel.ExecuteNonQuery()
                End Using
            Next
            _idsParaBorrar.Clear()

            ' 3. Guardar Líneas (Insert/Update)
            Dim orden As Integer = 1
            For Each row As DataRow In _dtLineas.Rows
                If row.RowState = DataRowState.Deleted Then Continue For

                ' Preparar valores
                Dim idLin = row("ID_Linea")
                Dim idArt = If(Val(row("ID_Articulo")) > 0, row("ID_Articulo"), DBNull.Value)
                Dim sqlLine As String = ""

                If IsDBNull(idLin) OrElse Val(idLin) = 0 Then
                    sqlLine = "INSERT INTO LineasPresupuesto (NumeroPresupuesto, NumeroOrden, ID_Articulo, Descripcion, Cantidad, PrecioUnitario, Descuento, Total) VALUES (@num, @ord, @art, @desc, @cant, @prec, @dcto, @tot)"
                Else
                    sqlLine = "UPDATE LineasPresupuesto SET NumeroOrden=@ord, ID_Articulo=@art, Descripcion=@desc, Cantidad=@cant, PrecioUnitario=@prec, Descuento=@dcto, Total=@tot WHERE ID_Linea=@id"
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
                    cmdL.Parameters.AddWithValue("@tot", row("Total"))
                    If Not String.IsNullOrEmpty(sqlLine) AndAlso sqlLine.Contains("WHERE") Then cmdL.Parameters.AddWithValue("@id", idLin)
                    cmdL.ExecuteNonQuery()
                End Using
                orden += 1
            Next

            trans.Commit()
            MessageBox.Show("Guardado correctamente.")
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

    ' --- BOTONERA PRINCIPAL ---
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

    ' --- NAVEGACIÓN ---
    Private Sub Navegar(direccion As String)
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()
            Dim sql As String = If(direccion = "ANT",
                                   "SELECT MAX(NumeroPresupuesto) FROM Presupuestos WHERE NumeroPresupuesto < @act",
                                   "SELECT MIN(NumeroPresupuesto) FROM Presupuestos WHERE NumeroPresupuesto > @act")

            ' Si es nuevo, al dar atrás buscamos el último de todos
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

    ' --- GESTIÓN DE LÍNEAS ---
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
        fila("Total") = 0
        _dtLineas.Rows.Add(fila)

        ' Ir a la celda del artículo para escribir rápido
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

    ' --- EVENTOS DEL GRID (Lógica de Artículo y Cálculos) ---
    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim fila = DataGridView1.Rows(e.RowIndex)
        Dim colName = DataGridView1.Columns(e.ColumnIndex).Name

        ' A. Si cambió el artículo -> Buscar datos
        If colName = "ID_Articulo" Then
            Dim idArt = fila.Cells("ID_Articulo").Value?.ToString()
            If Not String.IsNullOrEmpty(idArt) Then
                Try
                    Dim c = ConexionBD.GetConnection()
                    If c.State <> ConnectionState.Open Then c.Open()
                    Dim cmd As New SQLiteCommand("SELECT Descripcion, PrecioVenta, StockActual FROM Articulos WHERE ID_Articulo = @id", c)
                    cmd.Parameters.AddWithValue("@id", idArt)
                    Using r = cmd.ExecuteReader()
                        If r.Read() Then
                            fila.Cells("Descripcion").Value = r("Descripcion")
                            fila.Cells("PrecioUnitario").Value = r("PrecioVenta")
                            fila.Tag = r("StockActual") ' Guardamos stock en Tag
                        End If
                    End Using
                Catch
                End Try
            End If
        End If

        ' B. Si cambió cantidad/precio -> Recalcular
        If {"Cantidad", "PrecioUnitario", "Descuento", "ID_Articulo"}.Contains(colName) Then
            RecalcularFila(fila)
            CalcularTotalesGenerales()
        End If
    End Sub

    ' --- Validar Stock visualmente ---
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

    ' --- BUSCADORES DE TEXTBOX ---
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

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
#End Region
#Region "Auto-Organización Visual (Pixel-Perfect)"
    ''' <summary>
    ''' Calcula matemáticamente la posición y tamaño de todos los controles
    ''' para asegurar un diseño perfecto y alineado.
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

        ' ¡EL TRUCO! Coordenadas fijas
        Dim col1_X As Integer = margenIzq
        Dim col2_X As Integer = 190
        Dim col3_X As Integer = 750
        Dim col4_X As Integer = 920

        Dim yFila1 As Integer = 30
        Dim yFila2 As Integer = 55
        Dim yFila3 As Integer = 95
        Dim yFila4 As Integer = 120
        Dim yFila5 As Integer = 160 ' Nueva fila etiquetas
        Dim yFila6 As Integer = 185 ' Nueva fila controles

        ' =========================================================
        ' 3. COLOCACIÓN DE CABECERA (Etiquetas y Cajas)
        ' =========================================================
        ' --- ETIQUETAS DINÁMICAS ---
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

        ' --- CAJAS DE TEXTO ---
        ' Fila 1 (Presupuesto, Cliente, Fecha)
        TextBoxPresupuesto.Bounds = New Rectangle(col1_X, yFila2, 105, 25)
        btnBuscarPresupuesto.Bounds = New Rectangle(col1_X + 110, yFila2, 30, 25)

        TextBoxIdCliente.Bounds = New Rectangle(col2_X, yFila2, 60, 25)
        TextBoxCliente.Bounds = New Rectangle(col2_X + 70, yFila2, 460, 25)
        TextBoxFecha.Bounds = New Rectangle(col3_X, yFila2, 140, 25)

        ' Fila 2 (Vendedor, Observaciones, Estado)
        TextBoxIdVendedor.Bounds = New Rectangle(col1_X, yFila4, 50, 25)
        TextBoxVendedor.Bounds = New Rectangle(col1_X + 55, yFila4, 85, 25)
        TextBoxObservaciones.Bounds = New Rectangle(col2_X, yFila4, 530, 25)
        TextBoxEstado.Bounds = New Rectangle(col3_X, yFila4, 140, 25)

        ' =========================================================
        ' 4. COLOCACIÓN DE NUEVOS COMBOS (Logística)
        ' =========================================================
        ' 1. FORMA DE PAGO (Columna 1)
        lblFormaPago.Location = New Point(col1_X, yFila5)
        cboFormaPago.Bounds = New Rectangle(col1_X, yFila6, 140, 25)
        cboFormaPago.Font = New Font("Segoe UI", 10.5F)

        ' 2. RUTA (Columna 2 - Mismo ancho que las Observaciones)
        lblRuta.Location = New Point(col2_X, yFila5)
        cboRuta.Bounds = New Rectangle(col2_X, yFila6, 530, 25)
        cboRuta.Font = New Font("Segoe UI", 10.5F)

        ' =========================================================
        ' 5. SEPARADOR VISUAL CABECERA / DETALLE
        ' =========================================================
        Dim lineaDivisoria As Label = Me.Controls.OfType(Of Label)().FirstOrDefault(Function(l) l.Name = "LineaDivisoria")
        If lineaDivisoria Is Nothing Then
            lineaDivisoria = New Label() With {.Name = "LineaDivisoria", .BackColor = Color.FromArgb(120, 130, 140), .Height = 2}
            Me.Controls.Add(lineaDivisoria)
        End If

        ' Bajamos la tabla a 240 para que quepan los combos holgados
        Dim yTabla As Integer = 240
        lineaDivisoria.Bounds = New Rectangle(margenIzq, yTabla - 20, anchoForm - (margenIzq * 2), 2)
        lineaDivisoria.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        lineaDivisoria.BringToFront()

        ' =========================================================
        ' 6. LA TABLA (Desplazada hacia abajo)
        ' =========================================================
        Dim altoTabla As Integer = altoForm - yTabla - 140
        DataGridView1.Bounds = New Rectangle(margenIzq, yTabla, anchoForm - (margenIzq * 2), altoTabla)
        DataGridView1.BackgroundColor = Me.BackColor
        DataGridView1.BorderStyle = BorderStyle.None

        ' =========================================================
        ' 7. TOTALES (Con "Caja" de resumen)
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

        ' =========================================================
        ' 8. BARRA DE HERRAMIENTAS INFERIOR (Organizada)
        ' =========================================================
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

        ' --- ANCHORS INTELIGENTES ---
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

    ''' <summary>
    ''' Aplica un diseño plano y moderno a los botones de forma automatizada.
    ''' </summary>
    Private Sub EstilizarBoton(btn As Button, x As Integer, y As Integer, bg As Color, fg As Color, Optional ghost As Boolean = False)
        btn.Location = New Point(x, y)
        btn.Size = New Size(100, 35)
        btn.FlatStyle = FlatStyle.Flat
        btn.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        btn.Cursor = Cursors.Hand

        If ghost Then
            ' Botones secundarios (solo borde)
            btn.BackColor = bg
            btn.ForeColor = fg
            btn.FlatAppearance.BorderColor = Color.White
            btn.FlatAppearance.BorderSize = 1
        Else
            ' Botones principales (color sólido)
            btn.BackColor = bg
            btn.ForeColor = fg
            btn.FlatAppearance.BorderSize = 0
        End If
    End Sub
#End Region
    Private Sub btnBuscarPresupuesto_Click(sender As Object, e As EventArgs) Handles btnBuscarPresupuesto.Click
        Dim frmBuscar As New FrmBuscador()
        frmBuscar.TablaABuscar = "Presupuestos"

        If frmBuscar.ShowDialog() = DialogResult.OK Then

            ' 1. Escribimos el código devuelto (Ej: "PRE-003") en la caja
            TextBoxPresupuesto.Text = frmBuscar.Resultado

            ' 2. ¡LA MAGIA! Llamamos a la función para que lea la base de datos
            CargarDatosDocumento(frmBuscar.Resultado)

        End If
    End Sub
    ' =========================================================
    ' FUNCIONES SALVAVIDAS PARA LEER DE LA BASE DE DATOS
    ' =========================================================
    Private Function LeerTexto(r As SQLiteDataReader, col As String) As String
        Try
            Dim idx As Integer = r.GetOrdinal(col)
            Return If(r.IsDBNull(idx), "", r.GetValue(idx).ToString())
        Catch ex As IndexOutOfRangeException
            MessageBox.Show($"El código busca la columna '{col}', pero no existe en tu tabla. ¡Revisa cómo se llama en DBeaver/SQLite!", "Fallo de columna detectado")
            Return ""
        End Try
    End Function

    Private Function LeerDecimal(r As SQLiteDataReader, col As String) As Decimal
        Try
            Dim idx As Integer = r.GetOrdinal(col)
            Return If(r.IsDBNull(idx), 0D, Convert.ToDecimal(r.GetValue(idx)))
        Catch ex As IndexOutOfRangeException
            MessageBox.Show($"El código busca la columna '{col}', pero no existe en tu tabla. ¡Revisa cómo se llama en DBeaver/SQLite!", "Fallo de columna detectado")
            Return 0D
        End Try
    End Function

    ' =========================================================
    ' FUNCIÓN PARA VOLCAR LOS DATOS DE LA BASE DE DATOS A LA PANTALLA
    ' =========================================================
    Private Sub CargarDatosDocumento(numeroDoc As String)
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()

            ' 1. CARGAR LA CABECERA
            Dim sqlCabecera As String = "SELECT P.*, C.NombreFiscal AS NombreCli " &
                                        "FROM Presupuestos P " &
                                        "LEFT JOIN Clientes C ON P.CodigoCliente = C.CodigoCliente " &
                                        "WHERE P.NumeroPresupuesto = @num"

            Using cmd As New SQLiteCommand(sqlCabecera, c)
                cmd.Parameters.AddWithValue("@num", numeroDoc)
                Using r = cmd.ExecuteReader()
                    If r.Read() Then
                        ' Usamos nuestras nuevas funciones seguras
                        TextBoxIdCliente.Text = LeerTexto(r, "CodigoCliente")
                        TextBoxCliente.Text = LeerTexto(r, "NombreCli") ' Esta no fallará porque la generamos en el JOIN

                        TextBoxIdVendedor.Text = LeerTexto(r, "ID_Vendedor")
                        ' --- CARGAR COMBOS ---
                        Dim idPago = r("ID_FormaPago")
                        If Not IsDBNull(idPago) Then cboFormaPago.SelectedValue = Convert.ToInt32(idPago) Else cboFormaPago.SelectedIndex = -1

                        Dim idRut = r("ID_Ruta")
                        If Not IsDBNull(idRut) Then cboRuta.SelectedValue = Convert.ToInt32(idRut) Else cboRuta.SelectedIndex = -1
                        ' Fecha
                        Dim fechaBD As String = LeerTexto(r, "Fecha")
                        If Not String.IsNullOrEmpty(fechaBD) Then
                            TextBoxFecha.Text = Convert.ToDateTime(fechaBD).ToString("dd/MM/yyyy")
                        End If

                        TextBoxEstado.Text = LeerTexto(r, "Estado")
                        TextBoxObservaciones.Text = LeerTexto(r, "Observaciones")

                        ' Totales
                        Dim base As Decimal = LeerDecimal(r, "BaseImponible")
                        Dim iva As Decimal = LeerDecimal(r, "IVA")
                        Dim total As Decimal = LeerDecimal(r, "Total")

                        TextBoxBase.Text = base.ToString("N2") & " €"
                        TextBoxIva.Text = iva.ToString("N2") & " €"
                        TextBoxTotalPresup.Text = total.ToString("N2") & " €"
                    End If
                End Using
            End Using

            ' 2. CARGAR LAS LÍNEAS
            ' Si tu tabla de líneas se llama distinto, avísame.
            Dim sqlLineas As String = "SELECT * FROM LineasPresupuesto WHERE NumeroPresupuesto = @num"
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
            ' 1. Nos aseguramos de que la descripción ocupe el espacio sobrante
            If DataGridView1.Columns.Contains("Descripcion") Then
                DataGridView1.Columns("Descripcion").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            End If

            ' 2. Le damos más anchura al Precio y al Total (Ajusta los números a tu gusto)
            ' Nota: Pongo "PrecioUnitario" porque vi en tu base de datos que así se llama la columna.
            ' Si en el código le cambiaste el nombre a "Precio", pon "Precio" entre las comillas.
            If DataGridView1.Columns.Contains("PrecioUnitario") Then
                DataGridView1.Columns("PrecioUnitario").Width = 150 ' Por defecto suele ser 100
            End If

            If DataGridView1.Columns.Contains("Total") Then
                DataGridView1.Columns("Total").Width = 170 ' Le damos bastante espacio para números grandes
            End If

            ' (Opcional) Si quieres centrar o alinear a la derecha los números
            ' DataGridView1.Columns("Total").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Catch ex As Exception
            MessageBox.Show("Error general al cargar los datos: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class