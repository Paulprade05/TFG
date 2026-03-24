Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data.SQLite

Public Class FrmArticulos

    ' =========================================================
    ' 1. DECLARACIÓN DE CONTROLES (100% Código)
    ' =========================================================
    Private WithEvents txtCodigo As New TextBox()
    Private WithEvents txtCodigoBarras As New TextBox()
    Private WithEvents txtDescripcion As New TextBox()

    Private WithEvents cboFamilia As New ComboBox()
    Private WithEvents cboProveedor As New ComboBox()

    Private WithEvents txtPrecioCompra As New TextBox()
    Private WithEvents txtPrecioVenta As New TextBox()
    Private WithEvents cboIVA As New ComboBox()

    Private WithEvents txtStockActual As New TextBox()
    Private WithEvents txtStockMinimo As New TextBox()
    Private WithEvents txtObservaciones As New TextBox()

    Private WithEvents btnGuardar As New Button()
    Private WithEvents btnBorrar As New Button()
    Private WithEvents btnNuevo As New Button()

    Private WithEvents dgvArticulos As New DataGridView()
    Private _idArticuloActual As Integer = 0
    Private _filtrarSinStockAlAbrir As Boolean = False

    ' =========================================================
    ' NUEVO CONSTRUCTOR: Recibe la orden del Dashboard
    ' =========================================================
    Public Sub New(Optional soloSinStock As Boolean = False)
        InitializeComponent()
        _filtrarSinStockAlAbrir = soloSinStock
    End Sub

    ' =========================================================
    ' 2. INICIALIZACIÓN (UNIFICADA)
    ' =========================================================
    Private Sub FrmArticulos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Gestión de Artículos"
        Me.BackColor = Color.FromArgb(70, 75, 80) ' Fondo premium
        Me.MinimumSize = New Size(950, 650)

        ConstruirInterfaz()
        ConfigurarGrid()
        CargarDesplegables()

        ' Aquí decidimos qué cargar dependiendo de si venimos del Dashboard o no
        If _filtrarSinStockAlAbrir Then
            ' Columna StockActual es la que tienes en la base de datos
            CargarArticulos("WHERE StockActual <= 0")
        Else
            CargarArticulos()
        End If
    End Sub
    ' =========================================================
    ' 3. CONSTRUCTOR DE LA INTERFAZ (Bloque compacto y ordenado)
    ' =========================================================
    Private Sub ConstruirInterfaz()
        Dim margenIzq As Integer = 30

        Dim yFila1 As Integer = 25
        Dim yFila2 As Integer = 85
        Dim yFila3 As Integer = 145

        ' --- Fila 1: Identificación ---
        CrearCampo("Cód. Referencia", txtCodigo, margenIzq, yFila1, 130)
        CrearCampo("Cód. Barras", txtCodigoBarras, 180, yFila1, 150)
        ' Tamaño fijo y elegante para la descripción (460px), sin estirarse al infinito
        CrearCampo("Descripción del Artículo", txtDescripcion, 350, yFila1, 460)

        ' --- Fila 2: Categorización y Precios ---
        CrearCampo("Familia", cboFamilia, margenIzq, yFila2, 180)
        cboFamilia.DropDownStyle = ComboBoxStyle.DropDownList

        CrearCampo("Proveedor", cboProveedor, 230, yFila2, 200)
        cboProveedor.DropDownStyle = ComboBoxStyle.DropDownList

        CrearCampo("Coste (€)", txtPrecioCompra, 450, yFila2, 100)
        CrearCampo("P.V.P. (€)", txtPrecioVenta, 570, yFila2, 100)

        CrearCampo("% IVA", cboIVA, 690, yFila2, 120)
        cboIVA.Items.AddRange({"21", "10", "4", "0"})
        cboIVA.DropDownStyle = ComboBoxStyle.DropDownList

        ' --- Fila 3: Stock y Detalles ---
        CrearCampo("Stock Actual", txtStockActual, margenIzq, yFila3, 100)
        CrearCampo("Stock Mín.", txtStockMinimo, 150, yFila3, 100)

        ' Tamaño fijo para observaciones (540px) alineado con el borde de la fila de arriba
        CrearCampo("Observaciones / Detalles Técnicos", txtObservaciones, 270, yFila3, 540)
        txtObservaciones.Height = 45 : txtObservaciones.Multiline = True

        ' --- Línea Divisoria ---
        Dim yLinea As Integer = 215
        Dim linea As New Label() With {
            .Bounds = New Rectangle(margenIzq, yLinea, Me.ClientSize.Width - (margenIzq * 2), 2),
            .BackColor = Color.FromArgb(100, 110, 120),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        }
        Me.Controls.Add(linea)

        ' --- Tabla (Con margen inferior para los botones) ---
        Dim yTabla As Integer = 235
        Dim altoTabla As Integer = Me.ClientSize.Height - yTabla - 80
        dgvArticulos.Bounds = New Rectangle(margenIzq, yTabla, Me.ClientSize.Width - (margenIzq * 2), altoTabla)
        dgvArticulos.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        Me.Controls.Add(dgvArticulos)

        ' --- Botones (Clavados abajo a la izquierda) ---
        Dim yBotones As Integer = dgvArticulos.Bottom + 20
        ConfigurarBoton(btnGuardar, "Guardar", margenIzq, yBotones, Color.FromArgb(0, 120, 215))
        ConfigurarBoton(btnBorrar, "Borrar", margenIzq + 115, yBotones, Color.FromArgb(209, 52, 56))
        ConfigurarBoton(btnNuevo, "Nuevo", margenIzq + 230, yBotones, Color.FromArgb(0, 120, 215))

        btnGuardar.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
        btnBorrar.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
        btnNuevo.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
    End Sub

    ' Método de creación de campos simplificado y con anclaje fijo
    Private Sub CrearCampo(textoLabel As String, ctrl As Control, x As Integer, y As Integer, w As Integer)
        Dim lbl As New Label() With {
            .Text = textoLabel,
            .Location = New Point(x, y),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold),
            .ForeColor = Color.WhiteSmoke
        }
        Me.Controls.Add(lbl)

        ctrl.Bounds = New Rectangle(x, y + 23, w, 27)
        ctrl.Font = New Font("Segoe UI", 10.5F)

        If TypeOf ctrl Is TextBox Then
            DirectCast(ctrl, TextBox).BorderStyle = BorderStyle.FixedSingle
        End If

        ' Todas las cajas se quedan ancladas a la izquierda, sin estirarse
        ctrl.Anchor = AnchorStyles.Top Or AnchorStyles.Left

        Me.Controls.Add(ctrl)
    End Sub

    ' =========================================================
    ' 4. ESTILOS DEL GRID (Diseño unificado del ERP)
    ' =========================================================
    Private Sub ConfigurarGrid()
        ' 1. Aplicamos tu diseño oficial
        Try
            FrmPresupuestos.EstilizarGrid(dgvArticulos)
        Catch ex As Exception
        End Try

        ' 2. Ajustes básicos
        dgvArticulos.AutoGenerateColumns = False
        dgvArticulos.AllowUserToAddRows = False
        dgvArticulos.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvArticulos.ReadOnly = True ' Se edita desde las cajas de arriba, no directamente en la celda

        ' 3. Mapeo de columnas
        dgvArticulos.Columns.Clear()

        dgvArticulos.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "ID_Articulo", .DataPropertyName = "ID_Articulo", .Visible = False})
        dgvArticulos.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "CodigoReferencia", .DataPropertyName = "CodigoReferencia", .HeaderText = "Ref.", .Width = 100})
        dgvArticulos.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "CodigoBarras", .DataPropertyName = "CodigoBarras", .HeaderText = "EAN / Barras", .Width = 120})
        dgvArticulos.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "Descripcion", .DataPropertyName = "Descripcion", .HeaderText = "Descripción", .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill})
        dgvArticulos.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "PrecioCoste", .DataPropertyName = "PrecioCoste", .HeaderText = "Coste", .Width = 90, .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "C2", .Alignment = DataGridViewContentAlignment.MiddleRight}})
        dgvArticulos.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "PrecioVenta", .DataPropertyName = "PrecioVenta", .HeaderText = "P.V.P.", .Width = 90, .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "C2", .Alignment = DataGridViewContentAlignment.MiddleRight}})
        dgvArticulos.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "StockActual", .DataPropertyName = "StockActual", .HeaderText = "Stock", .Width = 80, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleCenter}})

        ' Columnas ocultas (para recuperar los datos al hacer clic en la fila)
        dgvArticulos.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "ID_Familia", .DataPropertyName = "ID_Familia", .Visible = False})
        dgvArticulos.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "ID_ProveedorHabitual", .DataPropertyName = "ID_ProveedorHabitual", .Visible = False})
        dgvArticulos.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "TipoIVA", .DataPropertyName = "TipoIVA", .Visible = False})
        dgvArticulos.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "StockMinimo", .DataPropertyName = "StockMinimo", .Visible = False})
        dgvArticulos.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "Observaciones", .DataPropertyName = "Observaciones", .Visible = False})
    End Sub
    Private Sub ConfigurarBoton(btn As Button, texto As String, x As Integer, y As Integer, colorFondo As Color)
        btn.Text = texto : btn.Bounds = New Rectangle(x, y, 100, 32)
        btn.BackColor = colorFondo : btn.ForeColor = Color.White
        btn.FlatStyle = FlatStyle.Flat : btn.FlatAppearance.BorderSize = 0
        btn.Font = New Font("Segoe UI", 10, FontStyle.Bold) : btn.Cursor = Cursors.Hand
        Me.Controls.Add(btn)
    End Sub
    ' =========================================================
    ' MAGIA VISUAL: Ajustar el alto de la tabla al contenido
    ' =========================================================
    Private Sub AjustarAltoTabla()
        If dgvArticulos Is Nothing Then Return

        ' 1. Calculamos lo que ocupan la cabecera y todas las filas
        Dim altoNecesario As Integer = dgvArticulos.ColumnHeadersHeight
        For Each fila As DataGridViewRow In dgvArticulos.Rows
            altoNecesario += fila.Height
        Next

        ' Le sumamos un pelín para el borde
        altoNecesario += 3

        ' 2. Calculamos el tope máximo (dejando 80px abajo para los botones)
        Dim altoMaximo As Integer = Me.ClientSize.Height - dgvArticulos.Top - 80

        ' 3. Ajustamos el alto real del control
        If altoNecesario > altoMaximo Then
            dgvArticulos.Height = altoMaximo ' Si hay muchos, llega al tope y pone scroll
        Else
            dgvArticulos.Height = altoNecesario ' Si hay pocos, se encoge
        End If

        ' Pintamos el fondo que sobra del mismo color que el formulario
        dgvArticulos.BackgroundColor = Me.BackColor
    End Sub

    ' Para que se recalcule si maximizas la pantalla
    Private Sub FrmArticulos_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        If dgvArticulos IsNot Nothing AndAlso dgvArticulos.RowCount > 0 Then
            AjustarAltoTabla()
        End If
    End Sub

    ' =========================================================
    ' 5. LÓGICA DE BASE DE DATOS Y EVENTOS
    ' =========================================================


    ' Ahora admite un filtro opcional
    Private Sub CargarArticulos(Optional filtroSQL As String = "")
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()

            ' Si pasamos un filtro (ej: WHERE StockActual <= 0), se lo añade a la consulta
            Dim sql As String = "SELECT * FROM Articulos " & filtroSQL & " ORDER BY ID_Articulo ASC"

            Using da As New SQLiteDataAdapter(sql, c)
                Dim dt As New DataTable()
                da.Fill(dt)
                dgvArticulos.DataSource = dt
                AjustarAltoTabla()
            End Using
        Catch ex As Exception
            MessageBox.Show("Error al cargar artículos: " & ex.Message, "Error BD", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub CargarDesplegables()
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()

            ' --- Cargar Familias (Dejo esto por si lo tienes junto) ---
            Dim daFam As New SQLiteDataAdapter("SELECT ID_Familia, Nombre FROM Familias", c)
            Dim dtFam As New DataTable()
            daFam.Fill(dtFam)
            cboFamilia.DataSource = dtFam
            cboFamilia.DisplayMember = "Nombre"
            cboFamilia.ValueMember = "ID_Familia"

            ' =========================================================
            ' --- CARGAR PROVEEDORES (AQUÍ ESTÁ LA CLAVE) ---
            ' =========================================================
            ' Leemos el Código y el Nombre de tu tabla Proveedores
            Dim sqlProv As String = "SELECT CodigoProveedor, NombreFiscal FROM Proveedores ORDER BY NombreFiscal"
            Using daProv As New SQLiteDataAdapter(sqlProv, c)
                Dim dtProv As New DataTable()
                daProv.Fill(dtProv)

                cboProveedor.DataSource = dtProv
                cboProveedor.DisplayMember = "NombreFiscal"      ' El texto que ves (Nvidia)
                cboProveedor.ValueMember = "CodigoProveedor" ' El código oculto (PROV-004) <--- ESTO FALLABA
                cboProveedor.SelectedIndex = -1
            End Using

        Catch ex As Exception
            MessageBox.Show("Error al cargar combos: " & ex.Message)
        End Try
    End Sub
    Private Sub dgvArticulos_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvArticulos.CellClick
        If e.RowIndex >= 0 Then
            Try
                Dim fila As DataGridViewRow = dgvArticulos.Rows(e.RowIndex)

                _idArticuloActual = Convert.ToInt32(fila.Cells("ID_Articulo").Value)

                txtCodigo.Text = If(IsDBNull(fila.Cells("CodigoReferencia").Value), "", fila.Cells("CodigoReferencia").Value?.ToString())
                txtCodigoBarras.Text = If(IsDBNull(fila.Cells("CodigoBarras").Value), "", fila.Cells("CodigoBarras").Value?.ToString())
                txtDescripcion.Text = If(IsDBNull(fila.Cells("Descripcion").Value), "", fila.Cells("Descripcion").Value?.ToString())
                txtPrecioCompra.Text = If(IsDBNull(fila.Cells("PrecioCoste").Value), "", fila.Cells("PrecioCoste").Value?.ToString())
                txtPrecioVenta.Text = If(IsDBNull(fila.Cells("PrecioVenta").Value), "", fila.Cells("PrecioVenta").Value?.ToString())
                txtStockActual.Text = If(IsDBNull(fila.Cells("StockActual").Value), "", fila.Cells("StockActual").Value?.ToString())
                txtStockMinimo.Text = If(IsDBNull(fila.Cells("StockMinimo").Value), "0", fila.Cells("StockMinimo").Value?.ToString())
                txtObservaciones.Text = If(IsDBNull(fila.Cells("Observaciones").Value), "", fila.Cells("Observaciones").Value?.ToString())

                Dim iva = If(IsDBNull(fila.Cells("TipoIVA").Value), "21", fila.Cells("TipoIVA").Value?.ToString())
                cboIVA.Text = iva

                Dim idFam = fila.Cells("ID_Familia").Value
                If Not IsDBNull(idFam) AndAlso idFam IsNot Nothing Then
                    cboFamilia.SelectedValue = Convert.ToInt32(idFam)
                Else
                    cboFamilia.SelectedIndex = -1
                End If

                ' =========================================================
                ' CARGAR PROVEEDOR (Método blindado y definitivo)
                ' =========================================================
                Dim idProv = fila.Cells("ID_ProveedorHabitual").Value

                If Not IsDBNull(idProv) AndAlso idProv IsNot Nothing AndAlso idProv.ToString().Trim() <> "" Then
                    Dim valorProv As String = idProv.ToString().Trim()

                    ' Intento 1: Asignación directa
                    cboProveedor.SelectedValue = valorProv

                    ' Intento 2 (El Infalible): Si WinForms se pone tonto y sigue en blanco, lo buscamos a la fuerza
                    If cboProveedor.SelectedIndex = -1 Then
                        For i As Integer = 0 To cboProveedor.Items.Count - 1
                            Dim filaVista As DataRowView = TryCast(cboProveedor.Items(i), DataRowView)
                            ' Comprobamos si el código de esta fila coincide con "PROV-004"
                            If filaVista IsNot Nothing AndAlso filaVista("CodigoProveedor").ToString() = valorProv Then
                                cboProveedor.SelectedIndex = i
                                Exit For ' Lo hemos pillado, paramos de buscar
                            End If
                        Next
                    End If
                Else
                    cboProveedor.SelectedIndex = -1
                End If

            Catch ex As Exception
                MessageBox.Show("Error al seleccionar la celda: " & ex.Message)
            End Try
        End If
    End Sub
    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        If String.IsNullOrWhiteSpace(txtCodigo.Text) OrElse String.IsNullOrWhiteSpace(txtDescripcion.Text) Then
            MessageBox.Show("La Referencia y la Descripción son obligatorias.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()

            Using cmdFK As New SQLiteCommand("PRAGMA foreign_keys = OFF;", c)
                cmdFK.ExecuteNonQuery()
            End Using

            Dim pCoste As Decimal = 0 : Decimal.TryParse(txtPrecioCompra.Text, pCoste)
            Dim pVenta As Decimal = 0 : Decimal.TryParse(txtPrecioVenta.Text, pVenta)
            Dim stock As Decimal = 0 : Decimal.TryParse(txtStockActual.Text, stock)
            Dim sMinimo As Decimal = 0 : Decimal.TryParse(txtStockMinimo.Text, sMinimo)
            Dim iva As Integer = 21 : Integer.TryParse(cboIVA.Text, iva)

            Dim idFamilia As Object = If(cboFamilia.SelectedValue IsNot Nothing, cboFamilia.SelectedValue, DBNull.Value)
            Dim idProveedor As Object = If(cboProveedor.SelectedValue IsNot Nothing, cboProveedor.SelectedValue, DBNull.Value)

            Dim sql As String
            If _idArticuloActual = 0 Then
                sql = "INSERT INTO Articulos (CodigoReferencia, CodigoBarras, Descripcion, ID_Familia, ID_ProveedorHabitual, PrecioCoste, PrecioVenta, TipoIVA, StockActual, StockMinimo, Observaciones) " &
                      "VALUES (@cod, @barras, @desc, @fam, @prov, @pcoste, @pventa, @iva, @stock, @smin, @obs)"
            Else
                sql = "UPDATE Articulos SET CodigoReferencia=@cod, CodigoBarras=@barras, Descripcion=@desc, ID_Familia=@fam, ID_ProveedorHabitual=@prov, PrecioCoste=@pcoste, PrecioVenta=@pventa, TipoIVA=@iva, StockActual=@stock, StockMinimo=@smin, Observaciones=@obs " &
                      "WHERE ID_Articulo=@id"
            End If

            Using cmd As New SQLiteCommand(sql, c)
                cmd.Parameters.AddWithValue("@cod", txtCodigo.Text.Trim())
                cmd.Parameters.AddWithValue("@barras", txtCodigoBarras.Text.Trim())
                cmd.Parameters.AddWithValue("@desc", txtDescripcion.Text.Trim())
                cmd.Parameters.AddWithValue("@fam", idFamilia)
                cmd.Parameters.AddWithValue("@prov", idProveedor)
                cmd.Parameters.AddWithValue("@pcoste", pCoste)
                cmd.Parameters.AddWithValue("@pventa", pVenta)
                cmd.Parameters.AddWithValue("@iva", iva)
                cmd.Parameters.AddWithValue("@stock", stock)
                cmd.Parameters.AddWithValue("@smin", sMinimo)
                cmd.Parameters.AddWithValue("@obs", txtObservaciones.Text)

                If _idArticuloActual > 0 Then cmd.Parameters.AddWithValue("@id", _idArticuloActual)
                cmd.ExecuteNonQuery()
            End Using

            MessageBox.Show("Artículo guardado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnNuevo.PerformClick()
            CargarArticulos()

        Catch ex As Exception
            MessageBox.Show("Error al guardar: " & ex.Message, "Error BD", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnBorrar_Click(sender As Object, e As EventArgs) Handles btnBorrar.Click
        If _idArticuloActual = 0 Then
            MessageBox.Show("Selecciona un artículo de la tabla primero.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        If MessageBox.Show("¿Eliminar definitivamente el artículo seleccionado?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes Then
            Try
                Dim c = ConexionBD.GetConnection()
                If c.State <> ConnectionState.Open Then c.Open()

                Using cmd As New SQLiteCommand("DELETE FROM Articulos WHERE ID_Articulo = @id", c)
                    cmd.Parameters.AddWithValue("@id", _idArticuloActual)
                    cmd.ExecuteNonQuery()
                End Using

                btnNuevo.PerformClick()
                CargarArticulos()
            Catch ex As Exception
                MessageBox.Show("Error al borrar: " & ex.Message, "Error BD", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        _idArticuloActual = 0
        txtCodigo.Clear() : txtCodigoBarras.Clear() : txtDescripcion.Clear()
        txtPrecioCompra.Clear() : txtPrecioVenta.Clear() : txtStockActual.Clear()
        txtStockMinimo.Clear() : txtObservaciones.Clear()
        cboIVA.SelectedIndex = 0
        cboFamilia.SelectedIndex = -1
        cboProveedor.SelectedIndex = -1
        txtCodigo.Focus()
        dgvArticulos.ClearSelection()
    End Sub

    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or &H2000000
            Return cp
        End Get
    End Property
End Class