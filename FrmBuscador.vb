Imports System.Data.SQLite

Public Class FrmBuscador
    ' CAMBIO 1: Ahora devolvemos un STRING para soportar "PRE-001" y "1"
    Public Property Resultado As String = ""

    ' Esta propiedad define qué vamos a buscar
    Public Property TablaABuscar As String = "Presupuestos"

    Private _dtDatos As DataTable

    Private Sub FrmBuscador_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ConfigurarGrid()
        CargarDatos()
        TextBoxBuscar.Focus()
    End Sub

    Private Sub ConfigurarGrid()
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.AllowUserToDeleteRows = False
        DataGridView1.ReadOnly = True
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView1.MultiSelect = False
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        DataGridView1.RowHeadersVisible = False
        DataGridView1.BackgroundColor = Color.White
        ' Estilo básico
        DataGridView1.DefaultCellStyle.Font = New Font("Segoe UI", 9)
        DataGridView1.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 10, FontStyle.Bold)
    End Sub

    Private Sub CargarDatos()
        Try
            Dim conexion = ConexionBD.GetConnection()
            Dim sql As String = ""

            Select Case TablaABuscar
                Case "Presupuestos"
                    Me.Text = "Buscar Presupuesto"
                    ' El campo clave es NumeroPresupuesto
                    sql = "SELECT NumeroPresupuesto AS [Nº Doc], " &
                          "C.NombreFiscal AS Cliente, P.Fecha, P.Estado " &
                          "FROM Presupuestos P " &
                          "LEFT JOIN Clientes C ON P.CodigoCliente = C.CodigoCliente " &
                          "ORDER BY P.NumeroPresupuesto DESC"
                Case "Clientes"
                    Me.Text = "Buscar Cliente"
                    ' El campo clave es ID_Cliente (Alias ID)
                    sql = "SELECT ID_Cliente AS ID, NombreFiscal AS Nombre, CIF, Poblacion FROM Clientes WHERE Activo = 1"

                Case "Articulos"
                    Me.Text = "Buscar Artículo"
                    ' El campo clave es ID_Articulo (Alias ID)
                    sql = "SELECT ID_Articulo AS ID, CodigoReferencia AS Codigo, Descripcion, PrecioVenta, StockActual AS Stock FROM Articulos WHERE Activo = 1"

                Case "Pedidos"
                    Me.Text = "Buscar Pedido"
                    ' El campo clave es NumeroPedido
                    sql = "SELECT NumeroPedido AS [Nº Doc], " &
                          "C.NombreFiscal AS Cliente, P.Fecha, P.Estado " &
                          "FROM Pedidos P " &
                          "LEFT JOIN Clientes C ON P.CodigoCliente = C.CodigoCliente " &
                          "ORDER BY P.NumeroPedido DESC"
                Case "Albaranes"
                    Me.Text = "Buscar Albaran"
                    ' El campo clave es NumeroPedido
                    sql = "SELECT NumeroAlbaran AS [Nº Doc], " &
                          "C.NombreFiscal AS Cliente, A.Fecha, A.Estado " &
                          "FROM Albaranes A " &
                          "LEFT JOIN Clientes C ON A.CodigoCliente = C.CodigoCliente " &
                          "ORDER BY A.NumeroAlbaran DESC"
            End Select

            If String.IsNullOrEmpty(sql) Then Return

            If conexion.State <> ConnectionState.Open Then conexion.Open()
            Dim da As New SQLiteDataAdapter(sql, conexion)
            _dtDatos = New DataTable()
            da.Fill(_dtDatos)

            DataGridView1.DataSource = _dtDatos

            ' Ocultamos la columna ID solo si es numérica interna (Clientes/Artículos)
            ' Para Presupuestos queremos ver el Nº Doc
            If DataGridView1.Columns.Contains("ID") AndAlso (TablaABuscar = "Clientes" Or TablaABuscar = "Articulos") Then
                DataGridView1.Columns("ID").Visible = False
            End If

        Catch ex As Exception
            MessageBox.Show("Error al cargar datos: " & ex.Message)
        End Try
    End Sub

    Private Sub TextBoxBuscar_TextChanged(sender As Object, e As EventArgs) Handles TextBoxBuscar.TextChanged
        If _dtDatos Is Nothing Then Return
        Try
            Dim filtro As String = TextBoxBuscar.Text.Replace("'", "''")
            Select Case TablaABuscar
                Case "Presupuestos", "Pedidos", "Albaranes"
                    _dtDatos.DefaultView.RowFilter = $"[Nº Doc] LIKE '%{filtro}%' OR Cliente LIKE '%{filtro}%'"
                Case "Clientes"
                    _dtDatos.DefaultView.RowFilter = $"Nombre LIKE '%{filtro}%' OR CIF LIKE '%{filtro}%'"
                Case "Articulos"
                    _dtDatos.DefaultView.RowFilter = $"Descripcion LIKE '%{filtro}%' OR Codigo LIKE '%{filtro}%'"

            End Select
        Catch
        End Try
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        SeleccionarYSalir()
    End Sub

    Private Sub ButtonAceptar_Click(sender As Object, e As EventArgs) Handles ButtonAceptar.Click
        SeleccionarYSalir()
    End Sub

    ' CAMBIO 2: Lógica de selección agnóstica (número o texto)
    Private Sub SeleccionarYSalir()
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim fila = DataGridView1.SelectedRows(0)
            Dim valorRecuperado As String = ""

            ' Dependiendo de la tabla, la columna clave se llama distinto
            If TablaABuscar = "Presupuestos" Or TablaABuscar = "Pedidos" Or TablaABuscar = "Albaranes" Then
                ' Buscamos "Nº Doc" (que es PRE-xxx o PED-xxx)
                valorRecuperado = fila.Cells("Nº Doc").Value.ToString()
            Else
                ' Buscamos "ID" (que es numérico 1, 2, 3...)
                valorRecuperado = fila.Cells("ID").Value.ToString()
            End If

            If Not String.IsNullOrEmpty(valorRecuperado) Then
                Resultado = valorRecuperado
                Me.DialogResult = DialogResult.OK
                Me.Close()
            End If
        Else
            MessageBox.Show("Por favor, selecciona un registro.")
        End If
    End Sub

    Private Sub ButtonCancelar_Click(sender As Object, e As EventArgs) Handles ButtonCancelar.Click
        Resultado = ""
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class