Imports System.Data.SQLite
Imports System.IO

Public Class FrmInformeFacturas
    ' Controles generados por código para no tener que arrastrarlos a mano
    Private dtpDesde As New DateTimePicker()
    Private dtpHasta As New DateTimePicker()
    Private btnGenerar As New Button()
    Private btnExportar As New Button()
    Private dgvResultados As New DataGridView()
    Private lblTotalBase As New Label()
    Private lblTotalIva As New Label()
    Private lblTotalFacturado As New Label()

    Private Sub FrmInformeFacturas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 1. CONFIGURACIÓN DEL FONDO (Modo Oscuro)
        Me.BackColor = Color.FromArgb(40, 50, 70)
        Me.ForeColor = Color.WhiteSmoke
        Me.Text = "Informe de Facturas Emitidas"

        ' 2. PANEL SUPERIOR (Filtros y Botones)
        Dim panelFiltros As New Panel With {.Dock = DockStyle.Top, .Height = 90}
        Me.Controls.Add(panelFiltros)

        Dim lblTitulo As New Label With {.Text = "LISTADO DE FACTURAS EMITIDAS (IVA)", .Location = New Point(30, 15), .AutoSize = True, .Font = New Font("Segoe UI", 12, FontStyle.Bold), .ForeColor = Color.FromArgb(0, 150, 255)}

        Dim lblDesde As New Label With {.Text = "Desde:", .Location = New Point(30, 50), .AutoSize = True, .Font = New Font("Segoe UI", 10)}
        dtpDesde.Bounds = New Rectangle(85, 48, 110, 25)
        dtpDesde.Format = DateTimePickerFormat.Short
        ' Por defecto, ponemos el primer día del mes actual
        dtpDesde.Value = New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1)

        Dim lblHasta As New Label With {.Text = "Hasta:", .Location = New Point(210, 50), .AutoSize = True, .Font = New Font("Segoe UI", 10)}
        dtpHasta.Bounds = New Rectangle(260, 48, 110, 25)
        dtpHasta.Format = DateTimePickerFormat.Short

        ' Botón Generar
        btnGenerar.Text = "Buscar Facturas"
        btnGenerar.Bounds = New Rectangle(390, 45, 140, 30)
        EstilizarBoton(btnGenerar, Color.FromArgb(0, 120, 215))
        AddHandler btnGenerar.Click, AddressOf GenerarInforme

        ' Botón Exportar (Verde Excel)
        btnExportar.Text = "Exportar a Excel"
        btnExportar.Bounds = New Rectangle(540, 45, 140, 30)
        EstilizarBoton(btnExportar, Color.FromArgb(40, 140, 90))
        AddHandler btnExportar.Click, AddressOf ExportarCSV

        panelFiltros.Controls.AddRange({lblTitulo, lblDesde, dtpDesde, lblHasta, dtpHasta, btnGenerar, btnExportar})

        ' 3. PANEL INFERIOR (Totales)
        Dim panelTotales As New Panel With {.Dock = DockStyle.Bottom, .Height = 60, .BackColor = Color.FromArgb(25, 30, 40)}
        lblTotalBase.Bounds = New Rectangle(30, 15, 200, 30)
        lblTotalBase.Font = New Font("Segoe UI", 11, FontStyle.Bold)

        lblTotalIva.Bounds = New Rectangle(250, 15, 200, 30)
        lblTotalIva.Font = New Font("Segoe UI", 11, FontStyle.Bold)

        lblTotalFacturado.Bounds = New Rectangle(470, 15, 250, 30)
        lblTotalFacturado.Font = New Font("Segoe UI", 13, FontStyle.Bold)
        lblTotalFacturado.ForeColor = Color.FromArgb(0, 150, 255)

        panelTotales.Controls.AddRange({lblTotalBase, lblTotalIva, lblTotalFacturado})
        Me.Controls.Add(panelTotales)

        ' 4. TABLA DE RESULTADOS (Se estira en el centro)
        dgvResultados.Dock = DockStyle.Fill
        dgvResultados.AllowUserToAddRows = False
        dgvResultados.ReadOnly = True
        dgvResultados.BackgroundColor = Me.BackColor
        dgvResultados.BorderStyle = BorderStyle.None
        dgvResultados.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        ' Intentamos usar tu estilo de tablas, si da error se queda el estándar
        Try : FrmPresupuestos.EstilizarGrid(dgvResultados) : Catch : End Try

        Me.Controls.Add(dgvResultados)
        dgvResultados.BringToFront()

        ' Nada más abrir, generamos el informe del mes actual
        GenerarInforme(Nothing, Nothing)
    End Sub

    Private Sub EstilizarBoton(btn As Button, colorFondo As Color)
        btn.BackColor = colorFondo
        btn.ForeColor = Color.White
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderSize = 0
        btn.Font = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        btn.Cursor = Cursors.Hand
    End Sub

    Private Sub GenerarInforme(sender As Object, e As EventArgs)
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()

            ' Buscamos las facturas en ese rango de fechas
            ' (Asegúrate de que tu tabla se llama 'Facturas', si no, cámbialo aquí abajo)
            Dim sql As String = "SELECT NumeroFactura AS [Nº Factura], Fecha, CodigoCliente AS [Cliente], " &
                                "BaseImponible AS [Base], (Total - BaseImponible) AS [Cuota IVA], Total AS [Total Factura], Estado " &
                                "FROM Facturas WHERE Fecha >= @desde AND Fecha <= @hasta ORDER BY Fecha ASC"

            Using cmd As New SQLiteCommand(sql, c)
                ' Formateamos las fechas para que SQLite las entienda (Año-Mes-Día)
                cmd.Parameters.AddWithValue("@desde", dtpDesde.Value.ToString("yyyy-MM-dd") & " 00:00:00")
                cmd.Parameters.AddWithValue("@hasta", dtpHasta.Value.ToString("yyyy-MM-dd") & " 23:59:59")

                Dim da As New SQLiteDataAdapter(cmd)
                Dim dt As New DataTable()
                da.Fill(dt)
                dgvResultados.DataSource = dt

                ' Ponemos símbolo de Euro a las columnas de dinero
                If dgvResultados.Columns.Contains("Base") Then dgvResultados.Columns("Base").DefaultCellStyle.Format = "C2"
                If dgvResultados.Columns.Contains("Cuota IVA") Then dgvResultados.Columns("Cuota IVA").DefaultCellStyle.Format = "C2"
                If dgvResultados.Columns.Contains("Total Factura") Then dgvResultados.Columns("Total Factura").DefaultCellStyle.Format = "C2"

                ' Calculamos las Sumas Totales abajo
                Dim sumBase As Decimal = 0
                Dim sumIva As Decimal = 0
                Dim sumTotal As Decimal = 0

                For Each row As DataRow In dt.Rows
                    sumBase += If(IsDBNull(row("Base")), 0, Convert.ToDecimal(row("Base")))
                    sumIva += If(IsDBNull(row("Cuota IVA")), 0, Convert.ToDecimal(row("Cuota IVA")))
                    sumTotal += If(IsDBNull(row("Total Factura")), 0, Convert.ToDecimal(row("Total Factura")))
                Next

                lblTotalBase.Text = "Base Total: " & sumBase.ToString("C2")
                lblTotalIva.Text = "IVA Total: " & sumIva.ToString("C2")
                lblTotalFacturado.Text = "TOTAL FACTURADO: " & sumTotal.ToString("C2")
            End Using
        Catch ex As Exception
            MsgBox("Error al generar informe (Comprueba que tengas facturas guardadas): " & ex.Message)
        End Try
    End Sub

    ' Exportación "inteligente" a Excel sin necesidad de instalar librerías raras
    Private Sub ExportarCSV(sender As Object, e As EventArgs)
        If dgvResultados.Rows.Count = 0 Then
            MsgBox("No hay facturas en esas fechas para exportar.", MsgBoxStyle.Exclamation)
            Return
        End If

        Dim sfd As New SaveFileDialog()
        sfd.Filter = "Archivo Excel (*.csv)|*.csv"
        sfd.FileName = "Informe_Facturas_" & DateTime.Now.ToString("yyyyMMdd") & ".csv"

        If sfd.ShowDialog() = DialogResult.OK Then
            Try
                Dim lineas As New List(Of String)

                ' Cabeceras (Las separamos por punto y coma para que el Excel de España las ponga en columnas)
                Dim cabeceras As String = ""
                For Each col As DataGridViewColumn In dgvResultados.Columns
                    cabeceras &= col.HeaderText & ";"
                Next
                lineas.Add(cabeceras.TrimEnd(";"c))

                ' Filas
                For Each row As DataGridViewRow In dgvResultados.Rows
                    Dim linea As String = ""
                    For Each cell As DataGridViewCell In row.Cells
                        linea &= If(cell.Value IsNot Nothing, cell.Value.ToString(), "") & ";"
                    Next
                    lineas.Add(linea.TrimEnd(";"c))
                Next

                ' Añadimos los Totales al final del Excel
                lineas.Add("")
                lineas.Add(";;TOTALES;" & lblTotalBase.Text.Replace("Base Total: ", "") & ";" & lblTotalIva.Text.Replace("IVA Total: ", "") & ";" & lblTotalFacturado.Text.Replace("TOTAL FACTURADO: ", "") & ";")

                File.WriteAllLines(sfd.FileName, lineas, System.Text.Encoding.UTF8)
                MsgBox("¡Facturas exportadas con éxito a Excel!", MsgBoxStyle.Information)
            Catch ex As Exception
                MsgBox("Error al exportar a Excel: " & ex.Message, MsgBoxStyle.Critical)
            End Try
        End If
    End Sub
End Class