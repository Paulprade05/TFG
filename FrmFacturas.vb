Imports System.Data.SQLite
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Text

Public Class FrmFacturas
    ' =========================================================
    ' 1. DECLARACIÓN DE CONTROLES (100% Por Código)
    ' =========================================================
    Private _numeroFacturaActual As String = ""
    Private _dtLineas As DataTable
    Private _idsParaBorrar As New List(Of Integer)

    ' Cabecera Izquierda
    Private TextBoxFactura As New TextBox()
    Private WithEvents btnBuscarFactura As New Button() With {.Text = "🔍"}
    Private WithEvents TextBoxIdCliente As New TextBox()
    Private TextBoxCliente As New TextBox()
    Private TextBoxFecha As New TextBox()
    Private DateTimePickerFecha As New DateTimePicker()
    Private WithEvents TextBoxIdVendedor As New TextBox()
    Private TextBoxVendedor As New TextBox()
    Private TextBoxEstado As New TextBox()
    Private cboAgencias As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList}
    Private WithEvents cboFormaPago As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList}
    Private WithEvents cboRuta As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList}

    ' Logística (Ahora irán dentro de un TabControl)
    Private TextBoxDireccion As New TextBox()
    Private TextBoxPoblacion As New TextBox()
    Private TextBoxCP As New TextBox()
    Private TextBoxTracking As New TextBox()
    Private ComboBoxPortes As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList}
    Private TextBoxBultos As New TextBox()
    Private TextBoxPeso As New TextBox()
    Private TextBoxAlbaranOrigen As New TextBox()
    Private WithEvents btnImportarAlbaran As New Button() With {.Text = "📥 Importar"}
    Private TextBoxObservaciones As New TextBox()

    ' Grid y Totales
    Private WithEvents DataGridView1 As New DataGridView()
    Private TextBoxBase As New TextBox()
    Private TextBoxIva As New TextBox()
    Private TextBoxTotalFac As New TextBox()
    Private LabelBase As New Label()
    Private LabelIva As New Label()
    Private Label7 As New Label()

    ' Botones Inferiores
    Private WithEvents ButtonGuardar As New Button()
    Private WithEvents ButtonBorrar As New Button()
    Private WithEvents ButtonNuevoFac As New Button()
    Private WithEvents ButtonBorrarLineas As New Button()
    Private WithEvents ButtonNuevaLinea As New Button()
    Private WithEvents ButtonAnterior As New Button()
    Private WithEvents ButtonSiguiente As New Button()
    Private LabelStock As New Label()

    ' Evitar parpadeos de dibujado
    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or &H2000000 ' WS_EX_COMPOSITED
            Return cp
        End Get
    End Property

    ' =========================================================
    ' 2. INICIO Y DIBUJADO DE LA INTERFAZ
    ' =========================================================
    Private Sub FrmFacturas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GenerarInterfazVisual()

        Try : FrmPresupuestos.EstilizarGrid(DataGridView1) : Catch : End Try
        EstilizarFecha(DateTimePickerFecha)

        If ComboBoxPortes.Items.Count = 0 Then
            ComboBoxPortes.Items.AddRange(New Object() {"Pagados", "Debidos"})
            ComboBoxPortes.SelectedIndex = 0
        End If

        ConfigurarGrid()
        CargarDesplegables()

        Dim ultimoNum As String = ObtenerUltimoNumeroFactura()
        If Not String.IsNullOrEmpty(ultimoNum) Then
            CargarFactura(ultimoNum)
        Else
            LimpiarFormulario()
        End If
    End Sub

    Private Sub GenerarInterfazVisual()
        ' Usamos el tamaño REAL del formulario para que los Anclajes (Anchor) no se vuelvan locos al maximizar
        Dim anchoForm As Integer = Me.ClientSize.Width
        Dim altoForm As Integer = Me.ClientSize.Height

        ' Seguro contra ejecuciones minimizadas
        If anchoForm < 100 Then anchoForm = 1200
        If altoForm < 100 Then altoForm = 800

        Me.BackColor = Color.FromArgb(45, 55, 75) ' Color Albaranes
        Me.ForeColor = Color.WhiteSmoke
        Dim fontLbl As New Font("Segoe UI", 9.5F, FontStyle.Bold)
        Dim fontTxt As New Font("Segoe UI", 10.0F)

        Dim CrearLabel = Function(t As String, x As Integer, y As Integer) As Label
                             Dim l As New Label With {.Text = t, .Location = New Point(x, y), .AutoSize = True, .Font = fontLbl, .ForeColor = Color.WhiteSmoke}
                             Me.Controls.Add(l)
                             Return l
                         End Function

        ' --- A. CABECERA IZQUIERDA ---
        Dim mIzq As Integer = 20, col2 As Integer = 190, col3 As Integer = 390, col4 As Integer = 530

        CrearLabel("Factura", mIzq, 15) : TextBoxFactura.Bounds = New Rectangle(mIzq, 40, 100, 25) : Me.Controls.Add(TextBoxFactura)
        btnBuscarFactura.Bounds = New Rectangle(mIzq + 105, 40, 30, 25) : EstilizarBoton(btnBuscarFactura, Color.FromArgb(85, 85, 85), Color.White) : Me.Controls.Add(btnBuscarFactura)

        CrearLabel("Cliente", col2, 15) : TextBoxIdCliente.Bounds = New Rectangle(col2, 40, 50, 25) : Me.Controls.Add(TextBoxIdCliente)
        TextBoxCliente.Bounds = New Rectangle(col2 + 55, 40, 130, 25) : Me.Controls.Add(TextBoxCliente)

        CrearLabel("Fecha", col3, 15) : TextBoxFecha.Bounds = New Rectangle(col3, 40, 120, 25) : Me.Controls.Add(TextBoxFecha)
        CrearLabel("Vencimiento", col4, 15) : DateTimePickerFecha.Bounds = New Rectangle(col4, 40, 120, 25) : DateTimePickerFecha.Format = DateTimePickerFormat.Short : Me.Controls.Add(DateTimePickerFecha)

        CrearLabel("Vendedor", mIzq, 75) : TextBoxIdVendedor.Bounds = New Rectangle(mIzq, 100, 40, 25) : Me.Controls.Add(TextBoxIdVendedor)
        TextBoxVendedor.Bounds = New Rectangle(mIzq + 45, 100, 90, 25) : Me.Controls.Add(TextBoxVendedor)

        CrearLabel("Estado", col2, 75) : TextBoxEstado.Bounds = New Rectangle(col2, 100, 185, 25) : Me.Controls.Add(TextBoxEstado)
        CrearLabel("Agencia", col3, 75) : cboAgencias.Bounds = New Rectangle(col3, 100, 130, 25) : Me.Controls.Add(cboAgencias)

        CrearLabel("Forma de Pago", mIzq, 135) : cboFormaPago.Bounds = New Rectangle(mIzq, 160, 140, 25) : Me.Controls.Add(cboFormaPago)
        CrearLabel("Ruta Asignada", col2, 135) : cboRuta.Bounds = New Rectangle(col2, 160, 310, 25) : Me.Controls.Add(cboRuta)

        ' --- B. TABCONTROL DE LOGÍSTICA ---
        Dim tabWidth As Integer = 630
        Dim tabCtrl As New TabControl() With {.ItemSize = New Size(140, 25), .Bounds = New Rectangle(anchoForm - tabWidth - 20, 20, tabWidth, 175), .Anchor = AnchorStyles.Top Or AnchorStyles.Right}
        Dim tab1 As New TabPage("Logística y Envío") With {.BackColor = Color.FromArgb(55, 65, 85)}
        Dim tab2 As New TabPage("Detalles y Origen") With {.BackColor = Color.FromArgb(55, 65, 85)}
        tabCtrl.TabPages.Add(tab1) : tabCtrl.TabPages.Add(tab2)
        Me.Controls.Add(tabCtrl)

        Dim CrearLblTab = Sub(tb As TabPage, t As String, x As Integer, y As Integer)
                              tb.Controls.Add(New Label With {.Text = t, .Location = New Point(x, y), .AutoSize = True, .Font = fontLbl, .ForeColor = Color.WhiteSmoke})
                          End Sub

        ' Tab 1
        CrearLblTab(tab1, "Dirección", 15, 15) : TextBoxDireccion.Bounds = New Rectangle(15, 40, 240, 25) : tab1.Controls.Add(TextBoxDireccion)
        CrearLblTab(tab1, "Población", 270, 15) : TextBoxPoblacion.Bounds = New Rectangle(270, 40, 210, 25) : tab1.Controls.Add(TextBoxPoblacion)
        CrearLblTab(tab1, "C.P.", 495, 15) : TextBoxCP.Bounds = New Rectangle(495, 40, 110, 25) : tab1.Controls.Add(TextBoxCP)

        CrearLblTab(tab1, "Código Seguimiento", 15, 80) : TextBoxTracking.Bounds = New Rectangle(15, 105, 240, 25) : tab1.Controls.Add(TextBoxTracking)
        CrearLblTab(tab1, "Portes", 270, 80) : ComboBoxPortes.Bounds = New Rectangle(270, 105, 120, 25) : tab1.Controls.Add(ComboBoxPortes)
        CrearLblTab(tab1, "Bultos", 405, 80) : TextBoxBultos.Bounds = New Rectangle(405, 105, 70, 25) : tab1.Controls.Add(TextBoxBultos)
        CrearLblTab(tab1, "Peso (Kg)", 495, 80) : TextBoxPeso.Bounds = New Rectangle(495, 105, 110, 25) : tab1.Controls.Add(TextBoxPeso)

        ' Tab 2
        CrearLblTab(tab2, "Albarán Origen", 15, 15) : TextBoxAlbaranOrigen.Bounds = New Rectangle(15, 40, 130, 25) : tab2.Controls.Add(TextBoxAlbaranOrigen)
        btnImportarAlbaran.Bounds = New Rectangle(160, 40, 115, 25) : EstilizarBoton(btnImportarAlbaran, Color.FromArgb(0, 120, 215), Color.White) : tab2.Controls.Add(btnImportarAlbaran)
        CrearLblTab(tab2, "Observaciones", 15, 80) : TextBoxObservaciones.Bounds = New Rectangle(15, 105, 590, 25) : tab2.Controls.Add(TextBoxObservaciones)

        ' Fuente global para cajas
        For Each ctrl As Control In Me.Controls
            If TypeOf ctrl Is TextBox OrElse TypeOf ctrl Is ComboBox Then ctrl.Font = fontTxt
        Next
        For Each ctrl As Control In tab1.Controls
            If TypeOf ctrl Is TextBox OrElse TypeOf ctrl Is ComboBox Then ctrl.Font = fontTxt
        Next
        For Each ctrl As Control In tab2.Controls
            If TypeOf ctrl Is TextBox OrElse TypeOf ctrl Is ComboBox Then ctrl.Font = fontTxt
        Next

        ' --- C. GRID Y SEPARADOR ---
        Dim yTabla As Integer = 220
        Me.Controls.Add(New Label With {.BackColor = Color.FromArgb(80, 90, 100), .Bounds = New Rectangle(mIzq, yTabla - 10, anchoForm - (mIzq * 2), 2), .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right})

        ' CORRECCIÓN IMPORTANTE: Aumentamos a 160 el margen inferior para que el Grid nunca tape a los botones
        DataGridView1.Bounds = New Rectangle(mIzq, yTabla, anchoForm - (mIzq * 2), altoForm - yTabla - 160)
        DataGridView1.BackgroundColor = Me.BackColor : DataGridView1.BorderStyle = BorderStyle.None
        DataGridView1.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        Me.Controls.Add(DataGridView1)

        ' --- D. PANEL TOTALES (DERECHA) ---
        Dim yTotales As Integer = altoForm - 145 ' Subimos el panel 
        Dim PanelTotales As New Panel() With {.BackColor = Color.FromArgb(20, 25, 35), .Bounds = New Rectangle(anchoForm - mIzq - 320, yTotales, 320, 130), .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right}
        Me.Controls.Add(PanelTotales)

        LabelBase.Text = "Base Imponible:" : LabelBase.Bounds = New Rectangle(10, 15, 140, 20) : LabelBase.ForeColor = Color.WhiteSmoke : LabelBase.Font = fontLbl
        TextBoxBase.Bounds = New Rectangle(160, 13, 140, 25) : TextBoxBase.BackColor = PanelTotales.BackColor : TextBoxBase.ForeColor = Color.White : TextBoxBase.BorderStyle = BorderStyle.None : TextBoxBase.TextAlign = HorizontalAlignment.Right : TextBoxBase.Font = fontTxt

        LabelIva.Text = "Importe IVA:" : LabelIva.Bounds = New Rectangle(10, 45, 140, 20) : LabelIva.ForeColor = Color.WhiteSmoke : LabelIva.Font = fontLbl
        TextBoxIva.Bounds = New Rectangle(160, 43, 140, 25) : TextBoxIva.BackColor = PanelTotales.BackColor : TextBoxIva.ForeColor = Color.White : TextBoxIva.BorderStyle = BorderStyle.None : TextBoxIva.TextAlign = HorizontalAlignment.Right : TextBoxIva.Font = fontTxt

        Label7.Text = "TOTAL:" : Label7.Bounds = New Rectangle(10, 85, 140, 25) : Label7.ForeColor = Color.FromArgb(0, 150, 255) : Label7.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        TextBoxTotalFac.Bounds = New Rectangle(160, 83, 140, 30) : TextBoxTotalFac.BackColor = PanelTotales.BackColor : TextBoxTotalFac.ForeColor = Color.FromArgb(0, 150, 255) : TextBoxTotalFac.BorderStyle = BorderStyle.None : TextBoxTotalFac.TextAlign = HorizontalAlignment.Right : TextBoxTotalFac.Font = New Font("Segoe UI", 14, FontStyle.Bold)

        PanelTotales.Controls.AddRange({LabelBase, TextBoxBase, LabelIva, TextBoxIva, New Label() With {.BackColor = Color.FromArgb(100, 110, 120), .Bounds = New Rectangle(10, 75, 300, 1)}, Label7, TextBoxTotalFac})

        ' --- E. BOTONES INFERIORES Y STOCK ---
        Dim yBotones As Integer = altoForm - 55 ' Posición fija relativa a la ventana actual

        ButtonGuardar.Text = "💾 Guardar" : ButtonGuardar.Bounds = New Rectangle(mIzq, yBotones, 110, 35) : EstilizarBoton(ButtonGuardar, Color.FromArgb(0, 120, 215), Color.White) : Me.Controls.Add(ButtonGuardar)
        ButtonBorrar.Text = "🗑 Borrar" : ButtonBorrar.Bounds = New Rectangle(mIzq + 120, yBotones, 100, 35) : EstilizarBoton(ButtonBorrar, Color.FromArgb(209, 52, 56), Color.White) : Me.Controls.Add(ButtonBorrar)
        ButtonNuevoFac.Text = "📄 Nueva" : ButtonNuevoFac.Bounds = New Rectangle(mIzq + 230, yBotones, 100, 35) : EstilizarBoton(ButtonNuevoFac, Color.FromArgb(0, 120, 215), Color.White) : Me.Controls.Add(ButtonNuevoFac)
        ButtonBorrarLineas.Text = "- Quitar Línea" : ButtonBorrarLineas.Bounds = New Rectangle(mIzq + 340, yBotones, 130, 35) : EstilizarBoton(ButtonBorrarLineas, Color.FromArgb(85, 85, 85), Color.White) : Me.Controls.Add(ButtonBorrarLineas)
        ButtonNuevaLinea.Text = "+ Añadir Línea" : ButtonNuevaLinea.Bounds = New Rectangle(mIzq + 480, yBotones, 130, 35) : EstilizarBoton(ButtonNuevaLinea, Color.FromArgb(40, 140, 90), Color.White) : Me.Controls.Add(ButtonNuevaLinea)

        ButtonAnterior.Text = "◀ Anterior" : ButtonAnterior.Bounds = New Rectangle(anchoForm - mIzq - 580, yBotones, 110, 35) : EstilizarBoton(ButtonAnterior, Me.BackColor, Color.White) : Me.Controls.Add(ButtonAnterior)
        ButtonSiguiente.Text = "Siguiente ▶" : ButtonSiguiente.Bounds = New Rectangle(anchoForm - mIzq - 460, yBotones, 110, 35) : EstilizarBoton(ButtonSiguiente, Me.BackColor, Color.White) : Me.Controls.Add(ButtonSiguiente)

        LabelStock.Bounds = New Rectangle(mIzq, yBotones - 25, 400, 25) : LabelStock.ForeColor = Color.FromArgb(40, 180, 90) : LabelStock.Font = fontLbl : Me.Controls.Add(LabelStock)

        Dim botones As Control() = {ButtonGuardar, ButtonBorrar, ButtonNuevoFac, ButtonBorrarLineas, ButtonNuevaLinea, ButtonAnterior, ButtonSiguiente, LabelStock}
        For Each b In botones
            If b Is ButtonAnterior Or b Is ButtonSiguiente Then b.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right Else b.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
        Next
    End Sub

    Private Sub EstilizarBoton(btn As Button, bg As Color, fg As Color)
        btn.BackColor = bg : btn.ForeColor = fg
        btn.FlatStyle = FlatStyle.Flat : btn.FlatAppearance.BorderSize = 0
        btn.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        btn.Cursor = Cursors.Hand
    End Sub

    Private Sub EstilizarFecha(dtp As DateTimePicker)
        dtp.Format = DateTimePickerFormat.Custom : dtp.CustomFormat = "dd/MM/yyyy"
        dtp.Font = New Font("Segoe UI", 10) : dtp.MinimumSize = New Size(0, 25)
    End Sub

    ' =========================================================
    ' 3. ESTRUCTURA Y GRID
    ' =========================================================
    Private Sub ConfigurarEstructuraDatos()
        If _dtLineas Is Nothing Then _dtLineas = New DataTable()
        If Not _dtLineas.Columns.Contains("ID_LineaFactura") Then _dtLineas.Columns.Add("ID_LineaFactura", GetType(Object))
        If Not _dtLineas.Columns.Contains("NumeroFactura") Then _dtLineas.Columns.Add("NumeroFactura", GetType(String))
        If Not _dtLineas.Columns.Contains("NumeroOrden") Then _dtLineas.Columns.Add("NumeroOrden", GetType(Integer))
        If Not _dtLineas.Columns.Contains("ID_Articulo") Then _dtLineas.Columns.Add("ID_Articulo", GetType(Object))
        If Not _dtLineas.Columns.Contains("Descripcion") Then _dtLineas.Columns.Add("Descripcion", GetType(String))
        If Not _dtLineas.Columns.Contains("Cantidad") Then _dtLineas.Columns.Add("Cantidad", GetType(Decimal))
        If Not _dtLineas.Columns.Contains("PrecioUnitario") Then _dtLineas.Columns.Add("PrecioUnitario", GetType(Decimal))
        If Not _dtLineas.Columns.Contains("Descuento") Then _dtLineas.Columns.Add("Descuento", GetType(Decimal))
        If Not _dtLineas.Columns.Contains("PorcentajeIVA") Then _dtLineas.Columns.Add("PorcentajeIVA", GetType(Decimal))
        If Not _dtLineas.Columns.Contains("TotalLinea") Then _dtLineas.Columns.Add("TotalLinea", GetType(Decimal))
    End Sub

    Private Sub ConfigurarGrid()
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "ID_LineaFactura", .DataPropertyName = "ID_LineaFactura", .Visible = False})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "NumeroOrden", .DataPropertyName = "NumeroOrden", .HeaderText = "Nº", .Width = 40, .ReadOnly = True, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleCenter}})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "ID_Articulo", .DataPropertyName = "ID_Articulo", .HeaderText = "ID Art", .Width = 70})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "Descripcion", .DataPropertyName = "Descripcion", .HeaderText = "Descripción", .Width = 250, .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "Cantidad", .DataPropertyName = "Cantidad", .HeaderText = "Cant.", .Width = 70, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "N2", .BackColor = Color.Ivory}})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "PrecioUnitario", .DataPropertyName = "PrecioUnitario", .HeaderText = "Precio", .Width = 80, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "C2"}})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "Descuento", .DataPropertyName = "Descuento", .HeaderText = "% Dto", .Width = 60, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "N2"}})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "PorcentajeIVA", .DataPropertyName = "PorcentajeIVA", .HeaderText = "% IVA", .Width = 60, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "N0"}})
        DataGridView1.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "TotalLinea", .DataPropertyName = "TotalLinea", .HeaderText = "Total", .ReadOnly = True, .Width = 90, .DefaultCellStyle = New DataGridViewCellStyle With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "C2", .BackColor = Color.WhiteSmoke}})
    End Sub

    Private Sub CargarDesplegables()
        Try
            Dim c = ConexionBD.GetConnection() : If c.State <> ConnectionState.Open Then c.Open()
            ' Agencias (si tienes el módulo, de lo contrario comenta estas dos líneas)
            ' cboAgencias.DataSource = GeneradorUtilidades.ObtenerAgencias()
            ' cboAgencias.DisplayMember = "Nombre" : cboAgencias.ValueMember = "ID_Agencia" : cboAgencias.SelectedIndex = -1

            Dim dtPago As New DataTable() : Dim daPago As New SQLiteDataAdapter("SELECT ID_FormaPago, Descripcion FROM FormasPago WHERE Activo=1", c)
            daPago.Fill(dtPago) : cboFormaPago.DataSource = dtPago : cboFormaPago.DisplayMember = "Descripcion" : cboFormaPago.ValueMember = "ID_FormaPago" : cboFormaPago.SelectedIndex = -1

            Dim dtRuta As New DataTable() : Dim daRuta As New SQLiteDataAdapter("SELECT ID_Ruta, NombreZona FROM Rutas WHERE Activo=1", c)
            daRuta.Fill(dtRuta) : cboRuta.DataSource = dtRuta : cboRuta.DisplayMember = "NombreZona" : cboRuta.ValueMember = "ID_Ruta" : cboRuta.SelectedIndex = -1
        Catch : End Try
    End Sub

    ' =========================================================
    ' 4. LÓGICA DE DATOS (Facturas)
    ' =========================================================
    Private Sub btnBuscarFactura_Click(sender As Object, e As EventArgs) Handles btnBuscarFactura.Click
        Using frm As New FrmBuscador
            frm.TablaABuscar = "Facturas"
            If frm.ShowDialog = DialogResult.OK Then CargarFactura(frm.Resultado)
        End Using
    End Sub

    Private Sub btnImportarAlbaran_Click(sender As Object, e As EventArgs) Handles btnImportarAlbaran.Click
        Using frm As New FrmBuscador
            frm.TablaABuscar = "Albaranes"
            If frm.ShowDialog = DialogResult.OK AndAlso Not String.IsNullOrEmpty(frm.Resultado) Then ImportarDatosAlbaran(frm.Resultado)
        End Using
    End Sub

    Private Sub ImportarDatosAlbaran(numeroAlbaran As String)
        Dim c = ConexionBD.GetConnection()
        Try
            If c.State <> ConnectionState.Open Then c.Open()
            Dim sqlCab As String = "SELECT A.*, C.NombreFiscal, C.CIF, C.Direccion, C.Poblacion, C.CodigoPostal, V.Nombre AS NombreVend FROM Albaranes A LEFT JOIN Clientes C ON A.CodigoCliente = C.CodigoCliente LEFT JOIN Vendedores V ON A.ID_Vendedor=V.ID_Vendedor WHERE A.NumeroAlbaran = @num"
            Using cmd As New SQLiteCommand(sqlCab, c)
                cmd.Parameters.AddWithValue("@num", numeroAlbaran)
                Using r = cmd.ExecuteReader()
                    If r.Read() Then
                        LimpiarFormulario()
                        TextBoxAlbaranOrigen.Text = numeroAlbaran : TextBoxAlbaranOrigen.Tag = numeroAlbaran
                        TextBoxIdCliente.Text = r("CodigoCliente").ToString() : TextBoxCliente.Text = If(IsDBNull(r("NombreFiscal")), "", r("NombreFiscal").ToString())
                        TextBoxIdVendedor.Text = If(IsDBNull(r("ID_Vendedor")), "", r("ID_Vendedor").ToString()) : TextBoxVendedor.Text = If(IsDBNull(r("NombreVend")), "", r("NombreVend").ToString())
                        TextBoxDireccion.Text = If(IsDBNull(r("DireccionEnvio")), r("Direccion").ToString(), r("DireccionEnvio").ToString())
                        TextBoxPoblacion.Text = If(IsDBNull(r("Poblacion")), "", r("Poblacion").ToString()) : TextBoxCP.Text = If(IsDBNull(r("CodigoPostal")), "", r("CodigoPostal").ToString())
                        TextBoxObservaciones.Text = "Facturado desde Albarán: " & numeroAlbaran
                    End If
                End Using
            End Using

            Dim dtOrigen As New DataTable()
            Using cmd As New SQLiteCommand("SELECT * FROM LineasAlbaran WHERE NumeroAlbaran = @num ORDER BY NumeroOrden ASC", c)
                cmd.Parameters.AddWithValue("@num", numeroAlbaran)
                Dim da As New SQLiteDataAdapter(cmd) : da.Fill(dtOrigen)
            End Using

            _dtLineas.Rows.Clear()
            For Each rowOrig As DataRow In dtOrigen.Rows
                Dim rowNew As DataRow = _dtLineas.NewRow()
                rowNew("NumeroFactura") = _numeroFacturaActual : rowNew("ID_LineaFactura") = DBNull.Value
                rowNew("NumeroOrden") = rowOrig("NumeroOrden") : rowNew("ID_Articulo") = rowOrig("ID_Articulo") : rowNew("Descripcion") = rowOrig("Descripcion")
                Dim cant As Decimal = 0 : Decimal.TryParse(rowOrig("CantidadServida").ToString(), cant)
                Dim prec As Decimal = 0 : Decimal.TryParse(rowOrig("PrecioUnitario").ToString(), prec)
                Dim desc As Decimal = 0 : Decimal.TryParse(rowOrig("Descuento").ToString(), desc)
                rowNew("Cantidad") = cant : rowNew("PrecioUnitario") = prec : rowNew("Descuento") = desc : rowNew("PorcentajeIVA") = 21
                rowNew("TotalLinea") = (cant * prec) * (1 - (desc / 100))
                _dtLineas.Rows.Add(rowNew)
            Next
            CalcularTotalesGenerales()
            MessageBox.Show("Albarán importado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error al importar: " & ex.Message)
        End Try
    End Sub

    Private Sub ButtonGuardar_Click(sender As Object, e As EventArgs) Handles ButtonGuardar.Click
        If String.IsNullOrWhiteSpace(TextBoxIdCliente.Text) Then MessageBox.Show("Falta el Cliente") : Return
        Dim esNuevo As Boolean = String.IsNullOrEmpty(_numeroFacturaActual)
        If esNuevo Then TextBoxFactura.Text = GenerarProximoNumeroFactura() : _numeroFacturaActual = TextBoxFactura.Text

        Dim c = ConexionBD.GetConnection() : Dim trans As SQLiteTransaction = Nothing
        Try
            If c.State <> ConnectionState.Open Then c.Open()
            trans = c.BeginTransaction()

            Dim sumaBase As Decimal = 0, sumaIva As Decimal = 0
            For Each r As DataRow In _dtLineas.Rows
                If r.RowState <> DataRowState.Deleted Then
                    Dim cant As Decimal = 0 : Decimal.TryParse(r("Cantidad").ToString(), cant)
                    Dim prec As Decimal = 0 : Decimal.TryParse(r("PrecioUnitario").ToString(), prec)
                    Dim desc As Decimal = 0 : Decimal.TryParse(r("Descuento").ToString(), desc)
                    Dim ivaPorc As Decimal = 21 : Decimal.TryParse(r("PorcentajeIVA").ToString(), ivaPorc)
                    Dim totLinea As Decimal = (cant * prec) * (1 - (desc / 100))
                    If _dtLineas.Columns.Contains("TotalLinea") Then r("TotalLinea") = totLinea
                    sumaBase += totLinea : sumaIva += totLinea * (ivaPorc / 100)
                End If
            Next

            Dim sql As String = If(esNuevo,
                "INSERT INTO FacturasVenta (NumeroFactura, NumeroAlbaran, Fecha, FechaVencimiento, ID_Cliente, NombreFiscal, CIF, Direccion, Poblacion, CodigoPostal, ID_Agencia, NumeroBultos, PesoTotal, CodigoSeguimiento, Portes, Estado, Observaciones, BaseImponible, ImporteIVA, TotalFactura, ID_Vendedor, ID_FormaPago, ID_Ruta) VALUES (@num, @alb, @fecha, @venc, @cli, @nomFis, @cif, @dir, @pob, @cp, @agencia, @bultos, @peso, @track, @portes, @estado, @obs, @base, @iva, @total, @vend, @pago, @ruta)",
                "UPDATE FacturasVenta SET NumeroAlbaran=@alb, Fecha=@fecha, FechaVencimiento=@venc, ID_Cliente=@cli, NombreFiscal=@nomFis, CIF=@cif, Direccion=@dir, Poblacion=@pob, CodigoPostal=@cp, ID_Agencia=@agencia, NumeroBultos=@bultos, PesoTotal=@peso, CodigoSeguimiento=@track, Portes=@portes, Estado=@estado, Observaciones=@obs, BaseImponible=@base, ImporteIVA=@iva, TotalFactura=@total, ID_Vendedor=@vend, ID_FormaPago=@pago, ID_Ruta=@ruta WHERE NumeroFactura=@num")

            Using cmd As New SQLiteCommand(sql, c, trans)
                cmd.Parameters.AddWithValue("@num", _numeroFacturaActual)
                cmd.Parameters.AddWithValue("@alb", If(TextBoxAlbaranOrigen.Tag IsNot Nothing, TextBoxAlbaranOrigen.Tag, DBNull.Value))
                Dim dF As DateTime = DateTime.Now : DateTime.TryParse(TextBoxFecha.Text, dF) : cmd.Parameters.AddWithValue("@fecha", dF)
                cmd.Parameters.AddWithValue("@venc", DateTimePickerFecha.Value)
                cmd.Parameters.AddWithValue("@cli", TextBoxIdCliente.Text)
                cmd.Parameters.AddWithValue("@nomFis", TextBoxCliente.Text) : cmd.Parameters.AddWithValue("@cif", DBNull.Value)
                cmd.Parameters.AddWithValue("@dir", TextBoxDireccion.Text) : cmd.Parameters.AddWithValue("@pob", TextBoxPoblacion.Text) : cmd.Parameters.AddWithValue("@cp", TextBoxCP.Text)
                cmd.Parameters.AddWithValue("@agencia", If(cboAgencias.SelectedIndex <> -1, cboAgencias.SelectedValue, DBNull.Value))
                cmd.Parameters.AddWithValue("@bultos", If(IsNumeric(TextBoxBultos.Text), TextBoxBultos.Text, 1))
                cmd.Parameters.AddWithValue("@peso", If(IsNumeric(TextBoxPeso.Text), TextBoxPeso.Text, 0))
                cmd.Parameters.AddWithValue("@track", TextBoxTracking.Text) : cmd.Parameters.AddWithValue("@portes", ComboBoxPortes.Text)
                cmd.Parameters.AddWithValue("@vend", If(String.IsNullOrEmpty(TextBoxIdVendedor.Text), DBNull.Value, TextBoxIdVendedor.Text))
                cmd.Parameters.AddWithValue("@pago", If(cboFormaPago.SelectedIndex <> -1, cboFormaPago.SelectedValue, DBNull.Value))
                cmd.Parameters.AddWithValue("@ruta", If(cboRuta.SelectedIndex <> -1, cboRuta.SelectedValue, DBNull.Value))
                cmd.Parameters.AddWithValue("@estado", TextBoxEstado.Text) : cmd.Parameters.AddWithValue("@obs", TextBoxObservaciones.Text)
                cmd.Parameters.AddWithValue("@base", sumaBase) : cmd.Parameters.AddWithValue("@iva", sumaIva) : cmd.Parameters.AddWithValue("@total", sumaBase + sumaIva)
                cmd.ExecuteNonQuery()
            End Using

            ' Bajas (devolver stock)
            For Each idDel In _idsParaBorrar
                Dim idArtDel As Integer = 0, cantDel As Decimal = 0
                Using cmdInfo As New SQLiteCommand("SELECT ID_Articulo, Cantidad FROM LineasFacturaVenta WHERE ID_LineaFactura = @id", c, trans)
                    cmdInfo.Parameters.AddWithValue("@id", idDel)
                    Using reader = cmdInfo.ExecuteReader()
                        If reader.Read() Then
                            If Not IsDBNull(reader("ID_Articulo")) Then idArtDel = Convert.ToInt32(reader("ID_Articulo"))
                            cantDel = Convert.ToDecimal(reader("Cantidad"))
                        End If
                    End Using
                End Using
                Using cmdDel As New SQLiteCommand("DELETE FROM LineasFacturaVenta WHERE ID_LineaFactura = @id", c, trans)
                    cmdDel.Parameters.AddWithValue("@id", idDel) : cmdDel.ExecuteNonQuery()
                End Using
                If idArtDel > 0 Then ActualizarStockYMovimiento(c, trans, _numeroFacturaActual, idArtDel, -cantDel, DateTime.Now)
            Next
            _idsParaBorrar.Clear()

            ' Inserciones y Actualizaciones
            If _dtLineas IsNot Nothing Then
                For Each r As DataRow In _dtLineas.Rows
                    If r.RowState = DataRowState.Deleted Then Continue For
                    Dim cant As Decimal = 0 : Decimal.TryParse(r("Cantidad").ToString(), cant)
                    Dim prec As Decimal = 0 : Decimal.TryParse(r("PrecioUnitario").ToString(), prec)
                    Dim desc As Decimal = 0 : Decimal.TryParse(r("Descuento").ToString(), desc)
                    Dim ivaPorc As Decimal = 21 : Decimal.TryParse(r("PorcentajeIVA").ToString(), ivaPorc)
                    Dim totLinea As Decimal = (cant * prec) * (1 - (desc / 100))
                    Dim idLin = r("ID_LineaFactura")
                    Dim idArt As Object = If(IsNumeric(r("ID_Articulo")) AndAlso Val(r("ID_Articulo")) > 0, r("ID_Articulo"), DBNull.Value)

                    If IsDBNull(idLin) OrElse Not IsNumeric(idLin) Then
                        Using cmdL As New SQLiteCommand("INSERT INTO LineasFacturaVenta (NumeroFactura, NumeroOrden, ID_Articulo, Descripcion, Cantidad, PrecioUnitario, Descuento, PorcentajeIVA, TotalLinea) VALUES (@fac, @ord, @art, @desc, @cant, @prec, @dcto, @iva, @tot)", c, trans)
                            cmdL.Parameters.AddWithValue("@fac", _numeroFacturaActual) : cmdL.Parameters.AddWithValue("@ord", r("NumeroOrden"))
                            cmdL.Parameters.AddWithValue("@art", idArt) : cmdL.Parameters.AddWithValue("@desc", r("Descripcion"))
                            cmdL.Parameters.AddWithValue("@cant", cant) : cmdL.Parameters.AddWithValue("@prec", prec) : cmdL.Parameters.AddWithValue("@dcto", desc)
                            cmdL.Parameters.AddWithValue("@iva", ivaPorc) : cmdL.Parameters.AddWithValue("@tot", totLinea)
                            cmdL.ExecuteNonQuery()
                        End Using
                        If Not IsDBNull(idArt) Then ActualizarStockYMovimiento(c, trans, _numeroFacturaActual, Convert.ToInt32(idArt), cant, DateTime.Now)
                    ElseIf r.RowState = DataRowState.Modified Then
                        Dim cantAnt As Decimal = 0
                        If r.HasVersion(DataRowVersion.Original) AndAlso Not IsDBNull(r("Cantidad", DataRowVersion.Original)) Then Decimal.TryParse(r("Cantidad", DataRowVersion.Original).ToString(), cantAnt)
                        Using cmdL As New SQLiteCommand("UPDATE LineasFacturaVenta SET ID_Articulo=@art, Descripcion=@desc, Cantidad=@cant, PrecioUnitario=@prec, Descuento=@dcto, PorcentajeIVA=@iva, TotalLinea=@tot WHERE ID_LineaFactura=@id", c, trans)
                            cmdL.Parameters.AddWithValue("@art", idArt) : cmdL.Parameters.AddWithValue("@desc", r("Descripcion"))
                            cmdL.Parameters.AddWithValue("@cant", cant) : cmdL.Parameters.AddWithValue("@prec", prec) : cmdL.Parameters.AddWithValue("@dcto", desc)
                            cmdL.Parameters.AddWithValue("@iva", ivaPorc) : cmdL.Parameters.AddWithValue("@tot", totLinea) : cmdL.Parameters.AddWithValue("@id", idLin)
                            cmdL.ExecuteNonQuery()
                        End Using
                        If Not IsDBNull(idArt) AndAlso (cant - cantAnt) <> 0 Then ActualizarStockYMovimiento(c, trans, _numeroFacturaActual, Convert.ToInt32(idArt), (cant - cantAnt), DateTime.Now)
                    End If
                Next
            End If
            trans.Commit() : MessageBox.Show("Factura Guardada y Registrada.") : CargarFactura(_numeroFacturaActual)
        Catch ex As Exception
            If trans IsNot Nothing Then trans.Rollback()
            MessageBox.Show("Error al guardar: " & ex.Message)
        End Try
    End Sub

    Private Sub ButtonBorrar_Click(sender As Object, e As EventArgs) Handles ButtonBorrar.Click
        If _numeroFacturaActual = "" Then Return
        If MessageBox.Show("¿Estás seguro de ELIMINAR esta factura? Se devolverá el stock.", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No Then Return
        Dim c = ConexionBD.GetConnection() : Dim trans As SQLiteTransaction = Nothing
        Try
            If c.State <> ConnectionState.Open Then c.Open()
            trans = c.BeginTransaction()
            Dim lineasABorrar As New List(Of Tuple(Of Integer, Decimal))
            Using cmdRec As New SQLiteCommand("SELECT ID_Articulo, Cantidad FROM LineasFacturaVenta WHERE NumeroFactura = @id", c, trans)
                cmdRec.Parameters.AddWithValue("@id", _numeroFacturaActual)
                Using reader = cmdRec.ExecuteReader()
                    While reader.Read()
                        If Not IsDBNull(reader("ID_Articulo")) Then lineasABorrar.Add(New Tuple(Of Integer, Decimal)(Convert.ToInt32(reader("ID_Articulo")), Convert.ToDecimal(reader("Cantidad"))))
                    End While
                End Using
            End Using
            For Each linea In lineasABorrar : ActualizarStockYMovimiento(c, trans, _numeroFacturaActual & " (Borrado)", linea.Item1, -linea.Item2, DateTime.Now) : Next
            Using cmd As New SQLiteCommand("DELETE FROM LineasFacturaVenta WHERE NumeroFactura= @id", c, trans)
                cmd.Parameters.AddWithValue("@id", _numeroFacturaActual) : cmd.ExecuteNonQuery()
            End Using
            Using cmd As New SQLiteCommand("DELETE FROM FacturasVenta WHERE NumeroFactura = @id", c, trans)
                cmd.Parameters.AddWithValue("@id", _numeroFacturaActual) : cmd.ExecuteNonQuery()
            End Using
            trans.Commit() : MessageBox.Show("Factura eliminada.") : LimpiarFormulario()
        Catch ex As Exception
            If trans IsNot Nothing Then trans.Rollback()
        End Try
    End Sub

    Private Sub CargarFactura(num As String)
        Try
            Dim c = ConexionBD.GetConnection() : If c.State <> ConnectionState.Open Then c.Open()
            Dim sql As String = "SELECT F.*, C.NombreFiscal AS NombreCliente, V.Nombre AS NombreVendedor FROM FacturasVenta F LEFT JOIN Clientes C ON F.ID_Cliente = C.CodigoCliente LEFT JOIN Vendedores V ON F.ID_Vendedor = V.ID_Vendedor WHERE F.NumeroFactura= @num"
            Using cmd As New SQLiteCommand(sql, c)
                cmd.Parameters.AddWithValue("@num", num)
                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        _numeroFacturaActual = num : TextBoxFactura.Text = num
                        If Not IsDBNull(reader("Fecha")) Then TextBoxFecha.Text = Convert.ToDateTime(reader("Fecha")).ToShortDateString()
                        If Not IsDBNull(reader("FechaVencimiento")) Then DateTimePickerFecha.Value = Convert.ToDateTime(reader("FechaVencimiento"))
                        TextBoxEstado.Text = If(IsDBNull(reader("Estado")), "Emitida", reader("Estado").ToString())
                        TextBoxObservaciones.Text = If(IsDBNull(reader("Observaciones")), "", reader("Observaciones").ToString())
                        TextBoxIdCliente.Text = reader("ID_Cliente").ToString() : TextBoxCliente.Text = If(IsDBNull(reader("NombreCliente")), "", reader("NombreCliente").ToString())
                        TextBoxDireccion.Text = reader("Direccion").ToString() : TextBoxPoblacion.Text = reader("Poblacion").ToString() : TextBoxCP.Text = reader("CodigoPostal").ToString()
                        TextBoxIdVendedor.Text = If(IsDBNull(reader("ID_Vendedor")), "", reader("ID_Vendedor").ToString()) : TextBoxVendedor.Text = If(IsDBNull(reader("NombreVendedor")), "", reader("NombreVendedor").ToString())
                        If Not IsDBNull(reader("ID_Agencia")) Then cboAgencias.SelectedValue = Convert.ToInt32(reader("ID_Agencia")) Else cboAgencias.SelectedIndex = -1
                        cboFormaPago.SelectedValue = If(Not IsDBNull(reader("ID_FormaPago")), Convert.ToInt32(reader("ID_FormaPago")), -1)
                        cboRuta.SelectedValue = If(Not IsDBNull(reader("ID_Ruta")), Convert.ToInt32(reader("ID_Ruta")), -1)
                        If Not IsDBNull(reader("NumeroAlbaran")) Then TextBoxAlbaranOrigen.Text = reader("NumeroAlbaran").ToString() : TextBoxAlbaranOrigen.Tag = reader("NumeroAlbaran").ToString()
                        TextBoxBultos.Text = reader("NumeroBultos").ToString() : TextBoxPeso.Text = reader("PesoTotal").ToString() : TextBoxTracking.Text = reader("CodigoSeguimiento").ToString() : ComboBoxPortes.Text = reader("Portes").ToString()
                    End If
                End Using
            End Using
            CargarLineas()
        Catch : End Try
    End Sub

    Private Sub CargarLineas()
        Try
            Dim c = ConexionBD.GetConnection()
            Using cmd As New SQLiteCommand("SELECT * FROM LineasFacturaVenta WHERE NumeroFactura = @num ORDER BY NumeroOrden ASC", c)
                cmd.Parameters.AddWithValue("@num", _numeroFacturaActual)
                Dim da As New SQLiteDataAdapter(cmd) : _dtLineas = New DataTable() : da.Fill(_dtLineas)
                ConfigurarEstructuraDatos() : DataGridView1.DataSource = _dtLineas
            End Using
            CalcularTotalesGenerales()
            If DataGridView1.Columns.Contains("ID_LineaFactura") Then DataGridView1.Columns("ID_LineaFactura").Visible = False
        Catch : End Try
    End Sub

    Private Sub CalcularTotalesGenerales()
        Dim base As Decimal = 0, iva As Decimal = 0
        If _dtLineas IsNot Nothing Then
            For Each row As DataRow In _dtLineas.Rows
                If row.RowState <> DataRowState.Deleted Then
                    Dim t As Decimal = 0 : Decimal.TryParse(row("TotalLinea").ToString(), t)
                    Dim pIva As Decimal = 21 : Decimal.TryParse(row("PorcentajeIVA").ToString(), pIva)
                    base += t : iva += t * (pIva / 100)
                End If
            Next
        End If
        TextBoxBase.Text = base.ToString("C2") : TextBoxIva.Text = iva.ToString("C2") : TextBoxTotalFac.Text = (base + iva).ToString("C2")
    End Sub

    Private Sub LimpiarFormulario()
        _numeroFacturaActual = "" : _idsParaBorrar.Clear()
        TextBoxFactura.Text = GenerarProximoNumeroFactura()
        TextBoxIdCliente.Text = "" : TextBoxCliente.Text = "" : TextBoxAlbaranOrigen.Text = "" : TextBoxAlbaranOrigen.Tag = Nothing
        TextBoxIdVendedor.Text = "" : TextBoxVendedor.Text = "" : cboAgencias.SelectedIndex = -1 : cboFormaPago.SelectedIndex = -1 : cboRuta.SelectedIndex = -1
        TextBoxBultos.Text = "1" : TextBoxPeso.Text = "0" : TextBoxDireccion.Text = "" : TextBoxPoblacion.Text = "" : TextBoxCP.Text = ""
        TextBoxTracking.Text = "TRK-" & DateTime.Now.Ticks.ToString().Substring(10)
        TextBoxFecha.Text = DateTime.Now.ToShortDateString() : DateTimePickerFecha.Value = DateTime.Now.AddDays(30)
        TextBoxEstado.Text = "Emitida" : TextBoxObservaciones.Text = "" : TextBoxBase.Text = "0,00" : TextBoxTotalFac.Text = "0,00" : TextBoxIva.Text = "0,00"
        ConfigurarGrid() : _dtLineas = New DataTable() : ConfigurarEstructuraDatos() : DataGridView1.DataSource = _dtLineas
    End Sub

    Private Function GenerarProximoNumeroFactura() As String
        Dim nuevo As String = "FAC-" & DateTime.Now.ToString("yy") & "-001"
        Try
            Dim c = ConexionBD.GetConnection() : If c.State <> ConnectionState.Open Then c.Open()
            Using cmd As New SQLiteCommand("SELECT NumeroFactura FROM FacturasVenta WHERE NumeroFactura LIKE @pat ORDER BY NumeroFactura DESC LIMIT 1", c)
                cmd.Parameters.AddWithValue("@pat", "FAC-" & DateTime.Now.ToString("yy") & "-%")
                Dim res = cmd.ExecuteScalar()
                If res IsNot Nothing AndAlso Not IsDBNull(res) Then
                    Dim parts = res.ToString().Split("-"c)
                    If parts.Length >= 3 Then nuevo = "FAC-" & DateTime.Now.ToString("yy") & "-" & (Convert.ToInt32(parts(parts.Length - 1)) + 1).ToString("D3")
                End If
            End Using
        Catch : End Try
        Return nuevo
    End Function

    Private Function ObtenerUltimoNumeroFactura() As String
        Try
            Dim c = ConexionBD.GetConnection() : If c.State <> ConnectionState.Open Then c.Open()
            Dim res = (New SQLiteCommand("SELECT MAX(NumeroFactura) FROM FacturasVenta", c)).ExecuteScalar()
            If res IsNot Nothing AndAlso Not IsDBNull(res) Then Return res.ToString()
        Catch : End Try
        Return ""
    End Function

    Private Sub ButtonAnterior_Click(sender As Object, e As EventArgs) Handles ButtonAnterior.Click
        Try
            Dim c = ConexionBD.GetConnection() : Dim sql = "SELECT MAX(NumeroFactura) FROM FacturasVenta WHERE NumeroFactura < @actual"
            Using cmd As New SQLiteCommand(sql, c)
                If String.IsNullOrEmpty(_numeroFacturaActual) Then cmd.CommandText = "SELECT MAX(NumeroFactura) FROM FacturasVenta" Else cmd.Parameters.AddWithValue("@actual", _numeroFacturaActual)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then CargarFactura(result.ToString())
            End Using
        Catch : End Try
    End Sub

    Private Sub ButtonSiguiente_Click(sender As Object, e As EventArgs) Handles ButtonSiguiente.Click
        If String.IsNullOrEmpty(_numeroFacturaActual) Then Return
        Try
            Dim c = ConexionBD.GetConnection() : Using cmd As New SQLiteCommand("SELECT MIN(NumeroFactura) FROM FacturasVenta WHERE NumeroFactura > @actual", c)
                cmd.Parameters.AddWithValue("@actual", _numeroFacturaActual)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then CargarFactura(result.ToString())
            End Using
        Catch : End Try
    End Sub

    Private Sub ButtonNuevoFac_Click(sender As Object, e As EventArgs) Handles ButtonNuevoFac.Click
        LimpiarFormulario()
    End Sub

    Private Sub ButtonNuevaLinea_Click(sender As Object, e As EventArgs) Handles ButtonNuevaLinea.Click
        If _dtLineas Is Nothing Then _dtLineas = New DataTable()
        ConfigurarEstructuraDatos()
        Dim nuevaFila = _dtLineas.NewRow()
        nuevaFila("NumeroOrden") = _dtLineas.Rows.Count + 1 : nuevaFila("NumeroFactura") = _numeroFacturaActual : nuevaFila("ID_LineaFactura") = DBNull.Value : nuevaFila("ID_Articulo") = DBNull.Value : nuevaFila("Descripcion") = "" : nuevaFila("Cantidad") = 1 : nuevaFila("PrecioUnitario") = 0 : nuevaFila("Descuento") = 0 : nuevaFila("PorcentajeIVA") = 21 : nuevaFila("TotalLinea") = 0
        _dtLineas.Rows.Add(nuevaFila)
        If DataGridView1.Rows.Count > 0 Then DataGridView1.CurrentCell = DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells("ID_Articulo") : DataGridView1.BeginEdit(True)
    End Sub

    Private Sub ButtonBorrarLineas_Click(sender As Object, e As EventArgs) Handles ButtonBorrarLineas.Click
        If DataGridView1.CurrentRow IsNot Nothing Then
            Dim vistaFila As DataRowView = TryCast(DataGridView1.CurrentRow.DataBoundItem, DataRowView)
            If vistaFila IsNot Nothing Then
                Dim row As DataRow = vistaFila.Row
                If Not IsDBNull(row("ID_LineaFactura")) AndAlso IsNumeric(row("ID_LineaFactura")) Then _idsParaBorrar.Add(Convert.ToInt32(row("ID_LineaFactura")))
                row.Delete() : CalcularTotalesGenerales()
            End If
        End If
    End Sub

    Private Sub TextBoxIdCliente_Leave(sender As Object, e As EventArgs) Handles TextBoxIdCliente.Leave
        If String.IsNullOrWhiteSpace(TextBoxIdCliente.Text) Then TextBoxCliente.Text = "" : Return
        Try
            Dim c = ConexionBD.GetConnection() : If c.State <> ConnectionState.Open Then c.Open()
            Using cmd As New SQLiteCommand("SELECT NombreFiscal, Direccion, Poblacion, CodigoPostal FROM Clientes WHERE CodigoCliente=@id", c)
                cmd.Parameters.AddWithValue("@id", TextBoxIdCliente.Text)
                Using r = cmd.ExecuteReader()
                    If r.Read() Then TextBoxCliente.Text = r("NombreFiscal").ToString() : TextBoxDireccion.Text = r("Direccion").ToString() : TextBoxPoblacion.Text = r("Poblacion").ToString() : TextBoxCP.Text = r("CodigoPostal").ToString()
                End Using
            End Using
        Catch : End Try
    End Sub

    Private Sub TextBoxIdVendedor_Leave(sender As Object, e As EventArgs) Handles TextBoxIdVendedor.Leave
        If String.IsNullOrWhiteSpace(TextBoxIdVendedor.Text) Then TextBoxVendedor.Text = "" : Return
        Try
            Dim c = ConexionBD.GetConnection() : If c.State <> ConnectionState.Open Then c.Open()
            Using cmd As New SQLiteCommand("SELECT Nombre FROM Vendedores WHERE ID_Vendedor=@id", c)
                cmd.Parameters.AddWithValue("@id", Val(TextBoxIdVendedor.Text))
                Dim r = cmd.ExecuteScalar() : TextBoxVendedor.Text = If(r IsNot Nothing, r.ToString(), "NO EXISTE")
            End Using
        Catch : End Try
    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim fila = DataGridView1.Rows(e.RowIndex)
        Dim colName = DataGridView1.Columns(e.ColumnIndex).Name
        If colName = "ID_Articulo" Then
            Dim idArt As String = fila.Cells("ID_Articulo").Value?.ToString()
            If String.IsNullOrWhiteSpace(idArt) Then fila.Cells("Descripcion").Value = "" : fila.Cells("PrecioUnitario").Value = 0 : fila.Tag = Nothing : Return
            Try
                Dim c = ConexionBD.GetConnection() : If c.State <> ConnectionState.Open Then c.Open()
                Using cmd As New SQLiteCommand("SELECT Descripcion, PrecioVenta, StockActual FROM Articulos WHERE ID_Articulo = @id", c)
                    cmd.Parameters.AddWithValue("@id", idArt)
                    Using r = cmd.ExecuteReader()
                        If r.Read() Then
                            fila.Cells("Descripcion").Value = r("Descripcion").ToString() : fila.Cells("PrecioUnitario").Value = Convert.ToDecimal(r("PrecioVenta")) : fila.Cells("PorcentajeIVA").Value = 21
                            If Not IsDBNull(r("StockActual")) Then fila.Tag = Convert.ToDecimal(r("StockActual"))
                            If IsDBNull(fila.Cells("Cantidad").Value) OrElse Val(fila.Cells("Cantidad").Value) = 0 Then fila.Cells("Cantidad").Value = 1
                        End If
                    End Using
                End Using
            Catch : End Try
        End If
        If colName = "ID_Articulo" Or colName = "Cantidad" Or colName = "PrecioUnitario" Or colName = "Descuento" Or colName = "PorcentajeIVA" Then
            Dim c_cant As Decimal = 0, p As Decimal = 0, d As Decimal = 0
            Decimal.TryParse(fila.Cells("Cantidad").Value?.ToString(), c_cant) : Decimal.TryParse(fila.Cells("PrecioUnitario").Value?.ToString(), p) : Decimal.TryParse(fila.Cells("Descuento").Value?.ToString(), d)
            fila.Cells("TotalLinea").Value = (c_cant * p) * (1 - (d / 100))
            CalcularTotalesGenerales()
            DataGridView1_SelectionChanged(Nothing, Nothing)
        End If
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        If DataGridView1.CurrentRow Is Nothing OrElse DataGridView1.CurrentRow.IsNewRow Then
            LabelStock.Text = "Stock disponible: -" : LabelStock.ForeColor = Color.FromArgb(170, 175, 180) : Return
        End If
        Try
            Dim idCelda As Object = DataGridView1.CurrentRow.Cells("ID_Articulo").Value
            If idCelda IsNot Nothing AndAlso Not DBNull.Value.Equals(idCelda) Then
                Dim idArticulo As Integer = Convert.ToInt32(idCelda)
                Dim cantidadEnLinea As Decimal = 0
                Dim cantCelda As Object = DataGridView1.CurrentRow.Cells("Cantidad").Value
                If cantCelda IsNot Nothing AndAlso Not DBNull.Value.Equals(cantCelda) Then cantidadEnLinea = Convert.ToDecimal(cantCelda)

                Dim stockActual As Decimal = 0
                Dim c = ConexionBD.GetConnection() : If c.State <> ConnectionState.Open Then c.Open()
                Using cmd As New SQLiteCommand("SELECT StockActual FROM Articulos WHERE ID_Articulo = @id", c)
                    cmd.Parameters.AddWithValue("@id", idArticulo)
                    Dim res = cmd.ExecuteScalar() : If res IsNot Nothing AndAlso Not DBNull.Value.Equals(res) Then stockActual = Convert.ToDecimal(res)
                End Using

                Dim stockRestante As Decimal = stockActual - cantidadEnLinea
                If stockRestante > 0 Then
                    LabelStock.Text = $"Stock actual: {stockActual} (Te quedarán: {stockRestante})" : LabelStock.ForeColor = Color.FromArgb(40, 180, 90)
                ElseIf stockRestante = 0 Then
                    LabelStock.Text = $"Stock actual: {stockActual} (¡Atención! Te quedarás a 0)" : LabelStock.ForeColor = Color.FromArgb(220, 160, 40)
                Else
                    LabelStock.Text = $"¡Stock insuficiente! Actual: {stockActual} (Te faltan: {Math.Abs(stockRestante)})" : LabelStock.ForeColor = Color.FromArgb(255, 80, 80)
                End If
            End If
        Catch : End Try
    End Sub

    Private Sub ActualizarStockYMovimiento(c As SQLiteConnection, trans As SQLiteTransaction, refDoc As String, idArticulo As Integer, variacionSalida As Decimal, fecha As DateTime)
        If variacionSalida = 0 Then Return
        Dim stockActual As Decimal = 0
        Using cmdStock As New SQLiteCommand("SELECT StockActual FROM Articulos WHERE ID_Articulo = @id", c, trans)
            cmdStock.Parameters.AddWithValue("@id", idArticulo)
            Dim res = cmdStock.ExecuteScalar() : If res IsNot Nothing AndAlso Not IsDBNull(res) Then stockActual = Convert.ToDecimal(res)
        End Using
        Dim nuevoStock As Decimal = stockActual - variacionSalida
        Using cmdUpd As New SQLiteCommand("UPDATE Articulos SET StockActual = @nuevo WHERE ID_Articulo = @id", c, trans)
            cmdUpd.Parameters.AddWithValue("@nuevo", nuevoStock) : cmdUpd.Parameters.AddWithValue("@id", idArticulo) : cmdUpd.ExecuteNonQuery()
        End Using
        Using cmdMov As New SQLiteCommand("INSERT INTO MovimientosAlmacen (Fecha, ID_Articulo, TipoMovimiento, Cantidad, StockResultante, DocumentoReferencia, ID_Usuario) VALUES (@fecha, @idArt, @tipo, @cant, @stockRes, @doc, 1)", c, trans)
            cmdMov.Parameters.AddWithValue("@fecha", fecha.ToString("yyyy-MM-dd HH:mm:ss")) : cmdMov.Parameters.AddWithValue("@idArt", idArticulo) : cmdMov.Parameters.AddWithValue("@tipo", If(variacionSalida > 0, "SALIDA", "ENTRADA")) : cmdMov.Parameters.AddWithValue("@cant", Math.Abs(variacionSalida)) : cmdMov.Parameters.AddWithValue("@stockRes", nuevoStock) : cmdMov.Parameters.AddWithValue("@doc", refDoc)
            cmdMov.ExecuteNonQuery()
        End Using
    End Sub
End Class