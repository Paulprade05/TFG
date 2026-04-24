<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmAlbaranesCompra
    Inherits System.Windows.Forms.Form

    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        components = New ComponentModel.Container()
        LabelStock = New Label()
        Label7 = New Label()
        LabelIva = New Label()
        LabelBase = New Label()
        TextBoxTotalAlb = New TextBox()
        TextBoxIva = New TextBox()
        TextBoxBase = New TextBox()
        ButtonNuevaLinea = New Button()
        ButtonBorrarLineas = New Button()
        ButtonSiguiente = New Button()
        ButtonAnterior = New Button()
        ButtonNuevoPed = New Button()
        ButtonBorrar = New Button()
        ButtonGuardar = New Button()
        Button1 = New Button()
        DataGridView1 = New DataGridView()
        TextBoxComprador = New TextBox()
        TextBoxIdComprador = New TextBox()
        Label4 = New Label()
        TextBoxObservaciones = New TextBox()
        Label6 = New Label()
        Label5 = New Label()
        TextBoxFecha = New TextBox()
        TextBoxProveedor = New TextBox()
        TextBoxIdProveedor = New TextBox()
        TextBoxAlbaran = New TextBox()
        Label3 = New Label()
        Label2 = New Label()
        Label1 = New Label()
        TextBoxPedidoOrigen = New TextBox()
        Label8 = New Label()
        DateTimePickerFecha = New DateTimePicker()
        Label9 = New Label()
        btnImportarPedido = New Button()
        TextBoxDireccion = New TextBox()
        TextBoxPoblacion = New TextBox()
        TextBoxCP = New TextBox()
        TextBoxBultos = New TextBox()
        TextBoxPeso = New TextBox()
        ComboBoxPortes = New ComboBox()
        LabelDireccion = New Label()
        LabelPoblacion = New Label()
        LabelCP = New Label()
        LabelBultos = New Label()
        LabelPeso = New Label()
        LabelPortes = New Label()
        CType(DataGridView1, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' LabelStock
        ' 
        LabelStock.AutoSize = True
        LabelStock.Font = New Font("Segoe UI Black", 9.75F, FontStyle.Bold)
        LabelStock.Location = New Point(1250, 281)
        LabelStock.Name = "LabelStock"
        LabelStock.Size = New Size(120, 17)
        LabelStock.Text = "Stock disponible:-"
        ' 
        ' Label7, LabelIva, LabelBase
        ' 
        Label7.AutoSize = True : Label7.Font = New Font("Segoe UI Black", 12.0F, FontStyle.Bold) : Label7.Location = New Point(1243, 532) : Label7.Name = "Label7" : Label7.Text = "TOTAL : "
        LabelIva.AutoSize = True : LabelIva.Font = New Font("Segoe UI", 11.25F, FontStyle.Bold) : LabelIva.Location = New Point(1243, 486) : LabelIva.Name = "LabelIva" : LabelIva.Text = "I.V.A :"
        LabelBase.AutoSize = True : LabelBase.Font = New Font("Segoe UI", 11.25F, FontStyle.Bold) : LabelBase.Location = New Point(1243, 443) : LabelBase.Name = "LabelBase" : LabelBase.Text = "Base imponible :"
        ' 
        ' TextBox totales
        ' 
        TextBoxTotalAlb.BackColor = Color.FromArgb(70, 75, 80) : TextBoxTotalAlb.BorderStyle = BorderStyle.None : TextBoxTotalAlb.Font = New Font("Segoe UI Black", 12.0F, FontStyle.Bold) : TextBoxTotalAlb.ForeColor = Color.WhiteSmoke : TextBoxTotalAlb.Location = New Point(1273, 531) : TextBoxTotalAlb.Name = "TextBoxTotalAlb" : TextBoxTotalAlb.ReadOnly = True : TextBoxTotalAlb.Size = New Size(256, 22) : TextBoxTotalAlb.TextAlign = HorizontalAlignment.Right
        TextBoxIva.BackColor = Color.FromArgb(70, 75, 80) : TextBoxIva.BorderStyle = BorderStyle.None : TextBoxIva.Font = New Font("Segoe UI", 11.25F, FontStyle.Bold) : TextBoxIva.ForeColor = Color.WhiteSmoke : TextBoxIva.Location = New Point(1273, 486) : TextBoxIva.Name = "TextBoxIva" : TextBoxIva.ReadOnly = True : TextBoxIva.Size = New Size(256, 20) : TextBoxIva.TextAlign = HorizontalAlignment.Right
        TextBoxBase.BackColor = Color.FromArgb(70, 75, 80) : TextBoxBase.BorderStyle = BorderStyle.None : TextBoxBase.Font = New Font("Segoe UI", 11.25F, FontStyle.Bold) : TextBoxBase.ForeColor = Color.WhiteSmoke : TextBoxBase.Location = New Point(1273, 443) : TextBoxBase.Name = "TextBoxBase" : TextBoxBase.ReadOnly = True : TextBoxBase.Size = New Size(256, 20) : TextBoxBase.TextAlign = HorizontalAlignment.Right
        ' 
        ' Botones
        ' 
        ButtonNuevaLinea.BackColor = Color.FromArgb(128, 255, 255) : ButtonNuevaLinea.FlatStyle = FlatStyle.Flat : ButtonNuevaLinea.Font = New Font("Segoe UI", 9.75F, FontStyle.Bold) : ButtonNuevaLinea.ForeColor = Color.Black : ButtonNuevaLinea.Location = New Point(701, 627) : ButtonNuevaLinea.Name = "ButtonNuevaLinea" : ButtonNuevaLinea.Size = New Size(113, 29) : ButtonNuevaLinea.Text = "Nueva linea"
        ButtonBorrarLineas.BackColor = Color.FromArgb(255, 128, 128) : ButtonBorrarLineas.FlatStyle = FlatStyle.Flat : ButtonBorrarLineas.Font = New Font("Segoe UI", 9.75F, FontStyle.Bold) : ButtonBorrarLineas.ForeColor = Color.Black : ButtonBorrarLineas.Location = New Point(581, 627) : ButtonBorrarLineas.Name = "ButtonBorrarLineas" : ButtonBorrarLineas.Size = New Size(113, 29) : ButtonBorrarLineas.Text = "Borrar lineas"
        ButtonSiguiente.BackColor = Color.Silver : ButtonSiguiente.FlatStyle = FlatStyle.Flat : ButtonSiguiente.Font = New Font("Segoe UI", 9.75F, FontStyle.Bold) : ButtonSiguiente.ForeColor = Color.Black : ButtonSiguiente.Location = New Point(1134, 627) : ButtonSiguiente.Name = "ButtonSiguiente" : ButtonSiguiente.Size = New Size(113, 29) : ButtonSiguiente.Text = "Siguiente"
        ButtonAnterior.BackColor = Color.Silver : ButtonAnterior.FlatStyle = FlatStyle.Flat : ButtonAnterior.Font = New Font("Segoe UI", 9.75F, FontStyle.Bold) : ButtonAnterior.ForeColor = Color.Black : ButtonAnterior.Location = New Point(1013, 627) : ButtonAnterior.Name = "ButtonAnterior" : ButtonAnterior.Size = New Size(113, 29) : ButtonAnterior.Text = "Anterior"
        ButtonNuevoPed.BackColor = Color.FromArgb(192, 255, 255) : ButtonNuevoPed.FlatStyle = FlatStyle.Flat : ButtonNuevoPed.Font = New Font("Segoe UI", 9.75F, FontStyle.Bold) : ButtonNuevoPed.ForeColor = Color.Black : ButtonNuevoPed.Location = New Point(288, 627) : ButtonNuevoPed.Name = "ButtonNuevoPed" : ButtonNuevoPed.Size = New Size(113, 29) : ButtonNuevoPed.Text = "Nuevo"
        ButtonBorrar.BackColor = Color.FromArgb(255, 192, 192) : ButtonBorrar.FlatStyle = FlatStyle.Flat : ButtonBorrar.Font = New Font("Segoe UI", 9.75F, FontStyle.Bold) : ButtonBorrar.ForeColor = Color.Black : ButtonBorrar.Location = New Point(168, 627) : ButtonBorrar.Name = "ButtonBorrar" : ButtonBorrar.Size = New Size(113, 29) : ButtonBorrar.Text = "Borrar"
        ButtonGuardar.BackColor = Color.FromArgb(192, 255, 192) : ButtonGuardar.FlatStyle = FlatStyle.Flat : ButtonGuardar.Font = New Font("Segoe UI", 9.75F, FontStyle.Bold) : ButtonGuardar.ForeColor = Color.Black : ButtonGuardar.Location = New Point(48, 627) : ButtonGuardar.Name = "ButtonGuardar" : ButtonGuardar.Size = New Size(113, 29) : ButtonGuardar.Text = "Guardar"
        ' 
        ' Button1 (buscar albarán)
        ' 
        Button1.Text = "..." : Button1.Location = New Point(168, 67) : Button1.Name = "Button1" : Button1.Size = New Size(29, 25) : Button1.UseVisualStyleBackColor = True
        ' 
        ' DataGridView1
        ' 
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridView1.Location = New Point(47, 163)
        DataGridView1.Name = "DataGridView1"
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView1.Size = New Size(1195, 417)
        ' 
        ' Campos cabecera
        ' 
        TextBoxComprador.Location = New Point(110, 118) : TextBoxComprador.Name = "TextBoxComprador" : TextBoxComprador.Size = New Size(305, 25)
        TextBoxIdComprador.Location = New Point(48, 118) : TextBoxIdComprador.Name = "TextBoxIdComprador" : TextBoxIdComprador.Size = New Size(54, 25)
        Label4.AutoSize = True : Label4.Font = New Font("Segoe UI", 11.25F, FontStyle.Bold) : Label4.Location = New Point(48, 95) : Label4.Name = "Label4" : Label4.Text = "Comprador"
        TextBoxObservaciones.Location = New Point(422, 118) : TextBoxObservaciones.Name = "TextBoxObservaciones" : TextBoxObservaciones.Size = New Size(282, 25)
        Label6.AutoSize = True : Label6.Font = New Font("Segoe UI", 11.25F, FontStyle.Bold) : Label6.Location = New Point(711, 95) : Label6.Name = "Label6" : Label6.Text = "Estado"
        Label5.AutoSize = True : Label5.Font = New Font("Segoe UI", 11.25F, FontStyle.Bold) : Label5.Location = New Point(422, 95) : Label5.Name = "Label5" : Label5.Text = "Observaciones"
        TextBoxFecha.Location = New Point(584, 67) : TextBoxFecha.Name = "TextBoxFecha" : TextBoxFecha.Size = New Size(114, 25)
        TextBoxProveedor.Location = New Point(272, 67) : TextBoxProveedor.Name = "TextBoxProveedor" : TextBoxProveedor.Size = New Size(305, 25)
        TextBoxIdProveedor.Location = New Point(210, 67) : TextBoxIdProveedor.Name = "TextBoxIdProveedor" : TextBoxIdProveedor.Size = New Size(54, 25)
        TextBoxAlbaran.Location = New Point(48, 67) : TextBoxAlbaran.Name = "TextBoxAlbaran" : TextBoxAlbaran.Size = New Size(114, 25)
        Label3.AutoSize = True : Label3.Font = New Font("Segoe UI", 11.25F, FontStyle.Bold) : Label3.Location = New Point(584, 44) : Label3.Name = "Label3" : Label3.Text = "Fecha"
        Label2.AutoSize = True : Label2.Font = New Font("Segoe UI", 11.25F, FontStyle.Bold) : Label2.Location = New Point(210, 44) : Label2.Name = "Label2" : Label2.Text = "Proveedor"
        Label1.AutoSize = True : Label1.Font = New Font("Segoe UI", 11.25F, FontStyle.Bold) : Label1.Location = New Point(48, 44) : Label1.Name = "Label1" : Label1.Text = "Albaran"
        TextBoxPedidoOrigen.Location = New Point(865, 118) : TextBoxPedidoOrigen.Name = "TextBoxPedidoOrigen" : TextBoxPedidoOrigen.Size = New Size(114, 25)
        Label8.AutoSize = True : Label8.Font = New Font("Segoe UI", 11.25F, FontStyle.Bold) : Label8.Location = New Point(865, 95) : Label8.Name = "Label8" : Label8.Text = "Pedido"
        DateTimePickerFecha.Format = DateTimePickerFormat.Short : DateTimePickerFecha.Location = New Point(704, 67) : DateTimePickerFecha.Name = "DateTimePickerFecha" : DateTimePickerFecha.Size = New Size(125, 25)
        Label9.AutoSize = True : Label9.Font = New Font("Segoe UI", 11.25F, FontStyle.Bold) : Label9.Location = New Point(701, 44) : Label9.Name = "Label9" : Label9.Text = "Fecha entrega"
        btnImportarPedido.Text = "..." : btnImportarPedido.Location = New Point(985, 118) : btnImportarPedido.Name = "btnImportarPedido" : btnImportarPedido.Size = New Size(29, 25) : btnImportarPedido.UseVisualStyleBackColor = True
        ' 
        ' Datos logísticos
        ' 
        TextBoxDireccion.Location = New Point(48, 200) : TextBoxDireccion.Name = "TextBoxDireccion" : TextBoxDireccion.Size = New Size(350, 25)
        TextBoxPoblacion.Location = New Point(410, 200) : TextBoxPoblacion.Name = "TextBoxPoblacion" : TextBoxPoblacion.Size = New Size(170, 25)
        TextBoxCP.Location = New Point(590, 200) : TextBoxCP.Name = "TextBoxCP" : TextBoxCP.Size = New Size(80, 25)
        TextBoxBultos.Location = New Point(690, 200) : TextBoxBultos.Name = "TextBoxBultos" : TextBoxBultos.Size = New Size(60, 25) : TextBoxBultos.Text = "1"
        TextBoxPeso.Location = New Point(760, 200) : TextBoxPeso.Name = "TextBoxPeso" : TextBoxPeso.Size = New Size(80, 25) : TextBoxPeso.Text = "0"
        ComboBoxPortes.Location = New Point(850, 200) : ComboBoxPortes.Name = "ComboBoxPortes" : ComboBoxPortes.Size = New Size(100, 25) : ComboBoxPortes.DropDownStyle = ComboBoxStyle.DropDownList
        LabelDireccion.AutoSize = True : LabelDireccion.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold) : LabelDireccion.Location = New Point(48, 180) : LabelDireccion.Name = "LabelDireccion" : LabelDireccion.Text = "Dirección recepción"
        LabelPoblacion.AutoSize = True : LabelPoblacion.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold) : LabelPoblacion.Location = New Point(410, 180) : LabelPoblacion.Name = "LabelPoblacion" : LabelPoblacion.Text = "Población"
        LabelCP.AutoSize = True : LabelCP.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold) : LabelCP.Location = New Point(590, 180) : LabelCP.Name = "LabelCP" : LabelCP.Text = "C.P."
        LabelBultos.AutoSize = True : LabelBultos.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold) : LabelBultos.Location = New Point(690, 180) : LabelBultos.Name = "LabelBultos" : LabelBultos.Text = "Bultos"
        LabelPeso.AutoSize = True : LabelPeso.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold) : LabelPeso.Location = New Point(760, 180) : LabelPeso.Name = "LabelPeso" : LabelPeso.Text = "Peso (kg)"
        LabelPortes.AutoSize = True : LabelPortes.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold) : LabelPortes.Location = New Point(850, 180) : LabelPortes.Name = "LabelPortes" : LabelPortes.Text = "Portes"
        ' 
        ' FrmAlbaranesCompra
        ' 
        AutoScaleDimensions = New SizeF(8.0F, 17.0F)
        AutoScaleMode = AutoScaleMode.Font
        BackColor = Color.FromArgb(70, 75, 80)
        ClientSize = New Size(1696, 761)
        ControlBox = False
        Controls.Add(LabelDireccion) : Controls.Add(LabelPoblacion) : Controls.Add(LabelCP) : Controls.Add(LabelBultos) : Controls.Add(LabelPeso) : Controls.Add(LabelPortes)
        Controls.Add(TextBoxDireccion) : Controls.Add(TextBoxPoblacion) : Controls.Add(TextBoxCP) : Controls.Add(TextBoxBultos) : Controls.Add(TextBoxPeso) : Controls.Add(ComboBoxPortes)
        Controls.Add(btnImportarPedido)
        Controls.Add(Label9)
        Controls.Add(DateTimePickerFecha)
        Controls.Add(Label8)
        Controls.Add(TextBoxPedidoOrigen)
        Controls.Add(LabelStock)
        Controls.Add(Label7)
        Controls.Add(LabelIva)
        Controls.Add(LabelBase)
        Controls.Add(TextBoxTotalAlb)
        Controls.Add(TextBoxIva)
        Controls.Add(TextBoxBase)
        Controls.Add(ButtonNuevaLinea)
        Controls.Add(ButtonBorrarLineas)
        Controls.Add(ButtonSiguiente)
        Controls.Add(ButtonAnterior)
        Controls.Add(ButtonNuevoPed)
        Controls.Add(ButtonBorrar)
        Controls.Add(ButtonGuardar)
        Controls.Add(Button1)
        Controls.Add(DataGridView1)
        Controls.Add(TextBoxComprador)
        Controls.Add(TextBoxIdComprador)
        Controls.Add(Label4)
        Controls.Add(TextBoxObservaciones)
        Controls.Add(Label6)
        Controls.Add(Label5)
        Controls.Add(TextBoxFecha)
        Controls.Add(TextBoxProveedor)
        Controls.Add(TextBoxIdProveedor)
        Controls.Add(TextBoxAlbaran)
        Controls.Add(Label3)
        Controls.Add(Label2)
        Controls.Add(Label1)
        Font = New Font("Segoe UI Semibold", 9.75F, FontStyle.Bold)
        ForeColor = Color.WhiteSmoke
        MaximizeBox = False
        MinimizeBox = False
        Name = "FrmAlbaranesCompra"
        ShowIcon = False
        SizeGripStyle = SizeGripStyle.Show
        Text = "Optima - Albaranes de Compra"
        CType(DataGridView1, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents LabelStock As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents LabelIva As Label
    Friend WithEvents LabelBase As Label
    Friend WithEvents TextBoxTotalAlb As TextBox
    Friend WithEvents TextBoxIva As TextBox
    Friend WithEvents TextBoxBase As TextBox
    Friend WithEvents ButtonNuevaLinea As Button
    Friend WithEvents ButtonBorrarLineas As Button
    Friend WithEvents ButtonSiguiente As Button
    Friend WithEvents ButtonAnterior As Button
    Friend WithEvents ButtonNuevoPed As Button
    Friend WithEvents ButtonBorrar As Button
    Friend WithEvents ButtonGuardar As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents TextBoxComprador As TextBox
    Friend WithEvents TextBoxIdComprador As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents TextBoxObservaciones As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents TextBoxFecha As TextBox
    Friend WithEvents TextBoxProveedor As TextBox
    Friend WithEvents TextBoxIdProveedor As TextBox
    Friend WithEvents TextBoxAlbaran As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents TextBoxPedidoOrigen As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents DateTimePickerFecha As DateTimePicker
    Friend WithEvents Label9 As Label
    Friend WithEvents btnImportarPedido As Button
    Friend WithEvents TextBoxDireccion As TextBox
    Friend WithEvents TextBoxPoblacion As TextBox
    Friend WithEvents TextBoxCP As TextBox
    Friend WithEvents TextBoxBultos As TextBox
    Friend WithEvents TextBoxPeso As TextBox
    Friend WithEvents ComboBoxPortes As ComboBox
    Friend WithEvents LabelDireccion As Label
    Friend WithEvents LabelPoblacion As Label
    Friend WithEvents LabelCP As Label
    Friend WithEvents LabelBultos As Label
    Friend WithEvents LabelPeso As Label
    Friend WithEvents LabelPortes As Label
End Class