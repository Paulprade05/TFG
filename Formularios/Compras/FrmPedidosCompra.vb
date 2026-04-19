Imports System.Data.SQLite
Imports System.Windows.Forms
Public Class FrmPedidosCompra
    Inherits Form
    ' =========================================================
    ' CONTROLES DE LA INTERFAZ
    ' =========================================================
    ' Cabecera
    Private txtNumeroPedido As New TextBox()
    Private dtpFecha As New DateTimePicker()
    Private cboProveedor As New ComboBox()
    Private cboEstado As New ComboBox()
    Private txtObservaciones As New TextBox()

    ' Líneas (Rejilla)
    Private WithEvents dgvLineas As New DataGridView()

    ' Totales y Botones
    Private lblTotalPedido As New Label()
    Private WithEvents btnGuardar As New Button()
    Private WithEvents btnNuevo As New Button()
    Private WithEvents btnEliminarFila As New Button()

    Private Sub FrmPedidosCompra_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Nuevo Pedido a Proveedor"
        Me.BackColor = Color.FromArgb(240, 245, 250) ' Fondo gris muy clarito

        ConstruirInterfaz()
        ConfigurarGrid()

        CargarProveedores()
        PrepararNuevoPedido()
    End Sub

    ' =========================================================
    ' 1. CONSTRUCCIÓN DE LA INTERFAZ (DISEÑO RESPONSIVE)
    ' =========================================================
    Private Sub ConstruirInterfaz()
        Me.Controls.Clear()

        ' --- PANEL SUPERIOR (CABECERA OSCURA) ---
        Dim pnlHeader As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 180,
            .BackColor = Color.FromArgb(55, 65, 85)
        }
        Me.Controls.Add(pnlHeader)

        Dim lblTitulo As New Label With {.Text = "PEDIDO DE COMPRA A PROVEEDOR", .Font = New Font("Segoe UI", 16, FontStyle.Bold), .ForeColor = Color.WhiteSmoke, .AutoSize = True, .Location = New Point(30, 20)}
        pnlHeader.Controls.Add(lblTitulo)

        ' Función ayudante para crear campos en la cabecera
        Dim CrearCampoHeader = Sub(texto As String, ctrl As Control, x As Integer, y As Integer, w As Integer)
                                   Dim lbl As New Label With {.Text = texto, .Location = New Point(x, y), .AutoSize = True, .ForeColor = Color.FromArgb(200, 210, 220), .Font = New Font("Segoe UI", 9)}
                                   ctrl.Bounds = New Rectangle(x, y + 20, w, 28) : ctrl.Font = New Font("Segoe UI", 10)
                                   pnlHeader.Controls.Add(lbl) : pnlHeader.Controls.Add(ctrl)
                               End Sub

        CrearCampoHeader("Nº de Pedido", txtNumeroPedido, 30, 70, 150)
        txtNumeroPedido.BackColor = Color.FromArgb(40, 50, 70)
        txtNumeroPedido.ForeColor = Color.White
        txtNumeroPedido.ReadOnly = True ' Se autogenera

        CrearCampoHeader("Fecha", dtpFecha, 200, 70, 130)
        dtpFecha.Format = DateTimePickerFormat.Short

        CrearCampoHeader("Proveedor", cboProveedor, 350, 70, 350)
        cboProveedor.DropDownStyle = ComboBoxStyle.DropDownList

        CrearCampoHeader("Estado del Pedido", cboEstado, 30, 125, 150)
        cboEstado.DropDownStyle = ComboBoxStyle.DropDownList
        cboEstado.Items.AddRange({"Pendiente", "Recibido", "Cancelado"})
        cboEstado.SelectedIndex = 0

        CrearCampoHeader("Observaciones / Notas para el almacén", txtObservaciones, 200, 125, 500)

        ' --- PANEL CENTRAL (GRID DE LÍNEAS) ---
        Dim pnlCentral As New Panel With {
            .Dock = DockStyle.Fill,
            .Padding = New Padding(30, 20, 30, 20)
        }
        Me.Controls.Add(pnlCentral)
        pnlCentral.BringToFront()

        dgvLineas.Dock = DockStyle.Fill
        dgvLineas.BackgroundColor = Color.White
        dgvLineas.BorderStyle = BorderStyle.None
        Try : FrmPresupuestos.EstilizarGrid(dgvLineas) : Catch : End Try ' Usa tu estilo si existe
        pnlCentral.Controls.Add(dgvLineas)

        ' --- PANEL INFERIOR (TOTALES Y BOTONES) ---
        Dim pnlFooter As New Panel With {
            .Dock = DockStyle.Bottom,
            .Height = 80,
            .BackColor = Color.White,
            .Padding = New Padding(30, 0, 30, 0)
        }
        Me.Controls.Add(pnlFooter)

        ' Línea separadora
        Dim linea As New Label With {.Dock = DockStyle.Top, .Height = 1, .BackColor = Color.LightGray}
        pnlFooter.Controls.Add(linea)

        btnNuevo.Text = "+ Nuevo Pedido" : btnNuevo.Bounds = New Rectangle(30, 20, 150, 40)
        btnNuevo.BackColor = Color.FromArgb(70, 75, 80) : btnNuevo.ForeColor = Color.White : btnNuevo.FlatStyle = FlatStyle.Flat : btnNuevo.FlatAppearance.BorderSize = 0 : btnNuevo.Cursor = Cursors.Hand
        pnlFooter.Controls.Add(btnNuevo)

        btnGuardar.Text = "💾 Guardar Pedido" : btnGuardar.Bounds = New Rectangle(190, 20, 160, 40)
        btnGuardar.BackColor = Color.FromArgb(0, 120, 215) : btnGuardar.ForeColor = Color.White : btnGuardar.Font = New Font("Segoe UI", 10, FontStyle.Bold) : btnGuardar.FlatStyle = FlatStyle.Flat : btnGuardar.FlatAppearance.BorderSize = 0 : btnGuardar.Cursor = Cursors.Hand
        pnlFooter.Controls.Add(btnGuardar)

        btnEliminarFila.Text = "🗑️ Quitar Línea" : btnEliminarFila.Bounds = New Rectangle(360, 20, 130, 40)
        btnEliminarFila.BackColor = Color.FromArgb(209, 52, 56) : btnEliminarFila.ForeColor = Color.White : btnEliminarFila.FlatStyle = FlatStyle.Flat : btnEliminarFila.FlatAppearance.BorderSize = 0 : btnEliminarFila.Cursor = Cursors.Hand
        pnlFooter.Controls.Add(btnEliminarFila)

        Dim lblTextoTotal As New Label With {.Text = "TOTAL PEDIDO:", .AutoSize = True, .Location = New Point(650, 30), .Font = New Font("Segoe UI", 12, FontStyle.Bold), .ForeColor = Color.DimGray, .Anchor = AnchorStyles.Right Or AnchorStyles.Top}
        pnlFooter.Controls.Add(lblTextoTotal)

        lblTotalPedido.Text = "0,00 €" : lblTotalPedido.AutoSize = True : lblTotalPedido.Location = New Point(790, 25) : lblTotalPedido.Font = New Font("Segoe UI", 16, FontStyle.Bold) : lblTotalPedido.ForeColor = Color.FromArgb(40, 140, 90) : lblTotalPedido.Anchor = AnchorStyles.Right Or AnchorStyles.Top
        pnlFooter.Controls.Add(lblTotalPedido)
    End Sub

    ' =========================================================
    ' 2. CONFIGURACIÓN DEL GRID DE LÍNEAS
    ' =========================================================
    Private Sub ConfigurarGrid()
        dgvLineas.Columns.Clear()

        dgvLineas.Columns.Add("ID_Articulo", "Cód. Art.")
        dgvLineas.Columns("ID_Articulo").Width = 80

        dgvLineas.Columns.Add("Descripcion", "Descripción del Artículo")
        dgvLineas.Columns("Descripcion").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        dgvLineas.Columns.Add("Cantidad", "Cantidad")
        dgvLineas.Columns("Cantidad").Width = 100
        dgvLineas.Columns("Cantidad").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvLineas.Columns("Cantidad").DefaultCellStyle.Format = "N2"

        dgvLineas.Columns.Add("Precio", "Precio Coste")
        dgvLineas.Columns("Precio").Width = 120
        dgvLineas.Columns("Precio").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvLineas.Columns("Precio").DefaultCellStyle.Format = "C2"

        dgvLineas.Columns.Add("Total", "Total Línea")
        dgvLineas.Columns("Total").Width = 130
        dgvLineas.Columns("Total").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvLineas.Columns("Total").DefaultCellStyle.Format = "C2"
        dgvLineas.Columns("Total").ReadOnly = True ' El total se calcula solo
        dgvLineas.Columns("Total").DefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245)
        dgvLineas.Columns("Total").DefaultCellStyle.Font = New Font("Segoe UI", 10, FontStyle.Bold)

        dgvLineas.AllowUserToAddRows = True
    End Sub

    ' Calcula el total de la fila cuando cambias Cantidad o Precio
    Private Sub dgvLineas_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgvLineas.CellEndEdit
        If e.RowIndex >= 0 Then
            Dim row = dgvLineas.Rows(e.RowIndex)
            Dim cant As Decimal = 0
            Dim precio As Decimal = 0

            Decimal.TryParse(If(row.Cells("Cantidad").Value Is Nothing, "0", row.Cells("Cantidad").Value.ToString()), cant)
            Decimal.TryParse(If(row.Cells("Precio").Value Is Nothing, "0", row.Cells("Precio").Value.ToString()), precio)

            row.Cells("Total").Value = cant * precio
            RecalcularTotalPedido()
        End If
    End Sub

    Private Sub dgvLineas_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles dgvLineas.RowsRemoved
        RecalcularTotalPedido()
    End Sub

    Private Sub btnEliminarFila_Click(sender As Object, e As EventArgs) Handles btnEliminarFila.Click
        If dgvLineas.CurrentRow IsNot Nothing AndAlso Not dgvLineas.CurrentRow.IsNewRow Then
            dgvLineas.Rows.Remove(dgvLineas.CurrentRow)
        End If
    End Sub

    Private Sub RecalcularTotalPedido()
        Dim totalBase As Decimal = 0
        For Each row As DataGridViewRow In dgvLineas.Rows
            If Not row.IsNewRow Then
                Dim valLinea As Decimal = 0
                Decimal.TryParse(If(row.Cells("Total").Value Is Nothing, "0", row.Cells("Total").Value.ToString()), valLinea)
                totalBase += valLinea
            End If
        Next
        lblTotalPedido.Text = totalBase.ToString("C2")
        lblTotalPedido.Tag = totalBase ' Guardamos el valor numérico oculto
    End Sub

    ' =========================================================
    ' 3. LÓGICA DE BASE DE DATOS
    ' =========================================================
    Private Sub CargarProveedores()
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()
            ' Ajusta "NombreFiscal" y "ID_Proveedor" a como se llamen en tu tabla Proveedores
            Dim cmd As New SQLiteCommand("SELECT ID_Proveedor, NombreFiscal FROM Proveedores ORDER BY NombreFiscal", c)
            Dim da As New SQLiteDataAdapter(cmd)
            Dim dt As New DataTable()
            da.Fill(dt)

            cboProveedor.DisplayMember = "NombreFiscal"
            cboProveedor.ValueMember = "ID_Proveedor"
            cboProveedor.DataSource = dt
            cboProveedor.SelectedIndex = -1
        Catch ex As Exception
            ' Si no tienes la tabla Proveedores aún, mostramos un error amigable
            MessageBox.Show("Aviso: No se ha podido cargar la lista de proveedores. Asegúrate de tener proveedores creados.")
        End Try
    End Sub

    Private Sub PrepararNuevoPedido()
        txtNumeroPedido.Text = "PED-" & DateTime.Now.ToString("yyMMddHHmmss") ' Autogeneramos uno temporal único
        dtpFecha.Value = DateTime.Now
        cboProveedor.SelectedIndex = -1
        cboEstado.SelectedIndex = 0
        txtObservaciones.Text = ""
        dgvLineas.Rows.Clear()
        RecalcularTotalPedido()
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        If MessageBox.Show("¿Limpiar todo para crear un nuevo pedido?", "Nuevo", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            PrepararNuevoPedido()
        End If
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        If cboProveedor.SelectedIndex = -1 Then
            MessageBox.Show("Debes seleccionar un proveedor.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning) : Return
        End If

        If dgvLineas.Rows.Count <= 1 Then ' <= 1 porque siempre hay una fila vacía nueva
            MessageBox.Show("El pedido debe tener al menos una línea de artículo.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning) : Return
        End If

        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()

            ' Usamos una transacción por si falla algo en las líneas, que no se guarde el pedido a medias
            Using tr As SQLiteTransaction = c.BeginTransaction()
                Try
                    ' 1. Guardar la Cabecera (PedidosCompra)
                    Dim sqlCabecera As String = "INSERT OR REPLACE INTO PedidosCompra (NumeroPedido, ID_Proveedor, Fecha, Estado, Total, Observaciones) " &
                                                "VALUES (@num, @prov, @fecha, @estado, @total, @obs)"
                    Using cmd As New SQLiteCommand(sqlCabecera, c, tr)
                        cmd.Parameters.AddWithValue("@num", txtNumeroPedido.Text)
                        cmd.Parameters.AddWithValue("@prov", cboProveedor.SelectedValue)
                        cmd.Parameters.AddWithValue("@fecha", dtpFecha.Value.ToString("yyyy-MM-dd"))
                        cmd.Parameters.AddWithValue("@estado", cboEstado.Text)
                        cmd.Parameters.AddWithValue("@total", CDec(lblTotalPedido.Tag))
                        cmd.Parameters.AddWithValue("@obs", txtObservaciones.Text.Trim())
                        cmd.ExecuteNonQuery()
                    End Using

                    ' 2. Borrar líneas anteriores si estamos sobreescribiendo el mismo pedido
                    Using cmdDel As New SQLiteCommand("DELETE FROM LineasPedidoCompra WHERE NumeroPedido=@num", c, tr)
                        cmdDel.Parameters.AddWithValue("@num", txtNumeroPedido.Text)
                        cmdDel.ExecuteNonQuery()
                    End Using

                    ' 3. Guardar las Líneas (LineasPedidoCompra)
                    Dim sqlLineas As String = "INSERT INTO LineasPedidoCompra (NumeroPedido, ID_Articulo, Descripcion, Cantidad, PrecioCosto, Total) " &
                                              "VALUES (@num, @art, @desc, @cant, @precio, @total)"

                    Using cmdLin As New SQLiteCommand(sqlLineas, c, tr)
                        For Each row As DataGridViewRow In dgvLineas.Rows
                            If Not row.IsNewRow Then
                                cmdLin.Parameters.Clear()
                                cmdLin.Parameters.AddWithValue("@num", txtNumeroPedido.Text)
                                cmdLin.Parameters.AddWithValue("@art", If(row.Cells("ID_Articulo").Value, 0))
                                cmdLin.Parameters.AddWithValue("@desc", If(row.Cells("Descripcion").Value, ""))
                                cmdLin.Parameters.AddWithValue("@cant", CDec(If(row.Cells("Cantidad").Value, 0)))
                                cmdLin.Parameters.AddWithValue("@precio", CDec(If(row.Cells("Precio").Value, 0)))
                                cmdLin.Parameters.AddWithValue("@total", CDec(If(row.Cells("Total").Value, 0)))
                                cmdLin.ExecuteNonQuery()
                            End If
                        Next
                    End Using

                    tr.Commit()
                    MessageBox.Show("Pedido guardado con éxito.", "Guardado", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    ' Bloqueamos el número para que si le da a guardar otra vez, se actualice en vez de crear otro
                Catch ex As Exception
                    tr.Rollback()
                    Throw ex
                End Try
            End Using

        Catch ex As Exception
            MessageBox.Show("Error al guardar el pedido: " & ex.Message, "Error SQL", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class