Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Data.SQLite
Public Class FrmDashboard
    Private chartVentas As New Chart()

    Private Sub FrmDashboard_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Color.FromArgb(70, 75, 80)
        Me.Controls.Clear()

        ' 1. TÍTULO
        Dim lblTitulo As New Label With {
            .Text = "ANÁLISIS DE VENTAS " & DateTime.Now.Year,
            .Font = New Font("Segoe UI", 18, FontStyle.Bold),
            .ForeColor = Color.WhiteSmoke,
            .AutoSize = True,
            .Location = New Point(40, 25)
        }
        Me.Controls.Add(lblTitulo)

        ' 2. CONTENEDOR DE TARJETAS (Arriba)
        Dim panelTarjetas As New FlowLayoutPanel With {
            .Location = New Point(40, 75),
            .Size = New Size(Me.Width - 80, 150),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right,
            .FlowDirection = FlowDirection.LeftToRight,
            .WrapContents = False
        }
        Me.Controls.Add(panelTarjetas)

        ' 3. CONTENEDOR PARA LA GRÁFICA (Abajo)
        Dim panelGrafica As New Panel With {
            .Location = New Point(40, 240),
            .Size = New Size(Me.Width - 80, Me.Height - 280),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right,
            .BackColor = Color.FromArgb(40, 50, 70)
        }
        Me.Controls.Add(panelGrafica)

        ' --- CARGAR DATOS ---
        CargarTarjetas(panelTarjetas)
        ConfigurarGraficaInteractiva(panelGrafica)
    End Sub

    Private Sub CargarTarjetas(contenedor As FlowLayoutPanel)
        ' (Aquí va la lógica de las consultas SQL que ya teníamos para los 0,00€)
        ' He simplificado para centrarnos en la gráfica, pero mantén tus cmd1, cmd2, etc.
        ' Por ejemplo:
        contenedor.Controls.Add(CrearTarjeta("Ventas Año", GetTotalAño().ToString("C2"), Color.FromArgb(0, 150, 255)))
        contenedor.Controls.Add(CrearTarjeta("Pendiente", GetTotalEstado("Pendiente").ToString("C2"), Color.FromArgb(230, 120, 20)))
        contenedor.Controls.Add(CrearTarjeta("Vencido", GetTotalEstado("Vencida").ToString("C2"), Color.FromArgb(209, 52, 56)))
    End Sub

    Private Sub ConfigurarGraficaInteractiva(contenedor As Panel)
        chartVentas.Dock = DockStyle.Fill
        chartVentas.BackColor = Color.Transparent

        ' Área del gráfico
        Dim area As New ChartArea("Principal")
        area.BackColor = Color.Transparent
        area.AxisX.LabelStyle.ForeColor = Color.WhiteSmoke
        area.AxisX.LineColor = Color.Gray
        area.AxisX.MajorGrid.LineColor = Color.FromArgb(60, 60, 60)

        area.AxisY.LabelStyle.ForeColor = Color.WhiteSmoke
        area.AxisY.LineColor = Color.Gray
        area.AxisY.MajorGrid.LineColor = Color.FromArgb(60, 60, 60)
        area.AxisY.LabelStyle.Format = "{0} €"

        chartVentas.ChartAreas.Clear()
        chartVentas.ChartAreas.Add(area)

        ' Serie de Datos (Barras modernas)
        Dim serie As New Series("Ventas")
        serie.ChartType = SeriesChartType.Column
        serie.Color = Color.FromArgb(0, 150, 255)
        serie.BorderRadius = 5
        ' INTERACTIVIDAD: Al pasar el ratón muestra el valor
        serie.ToolTip = "Total: #VALY{C2}"

        chartVentas.Series.Clear()
        chartVentas.Series.Add(serie)

        ' --- SQL PARA SACAR VENTAS POR MES ---
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()

            ' Consulta que agrupa por mes (strftime saca 01, 02, etc)
            Dim sql = "SELECT strftime('%m', Fecha) as Mes, SUM(TotalFactura) as Total " &
                      "FROM Facturas WHERE Fecha LIKE @año AND Estado <> 'Cancelada' " &
                      "GROUP BY Mes ORDER BY Mes ASC"

            Using cmd As New SQLiteCommand(sql, c)
                cmd.Parameters.AddWithValue("@año", DateTime.Now.Year.ToString() & "-%")
                Using r = cmd.ExecuteReader()
                    Dim meses() As String = {"", "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"}
                    While r.Read()
                        Dim nMes As Integer = Convert.ToInt32(r("Mes"))
                        Dim total As Decimal = Convert.ToDecimal(r("Total"))
                        serie.Points.AddXY(meses(nMes), total)
                    End While
                End Using
            End Using
        Catch
        End Try

        contenedor.Controls.Add(chartVentas)
    End Sub

    ' --- HELPERS DE APOYO ---
    Private Function GetTotalAño() As Decimal
        Try
            Dim c = ConexionBD.GetConnection()
            Dim cmd As New SQLiteCommand("SELECT SUM(TotalFactura) FROM Facturas WHERE Fecha LIKE @anio AND Estado <> 'Cancelada'", c)
            cmd.Parameters.AddWithValue("@anio", DateTime.Now.Year.ToString() & "%")
            Return Convert.ToDecimal(If(cmd.ExecuteScalar(), 0))
        Catch : Return 0 : End Try
    End Function

    Private Function GetTotalEstado(estado As String) As Decimal
        Try
            Dim c = ConexionBD.GetConnection()
            Dim cmd As New SQLiteCommand("SELECT SUM(TotalFactura) FROM Facturas WHERE Estado = @est", c)
            cmd.Parameters.AddWithValue("@est", estado)
            Return Convert.ToDecimal(If(cmd.ExecuteScalar(), 0))
        Catch : Return 0 : End Try
    End Function

    Private Function CrearTarjeta(titulo As String, valor As String, colorAcento As Color) As Panel
        Dim pnl As New Panel With {.Size = New Size(220, 100), .BackColor = Color.FromArgb(40, 50, 70), .Margin = New Padding(0, 0, 20, 0)}
        Dim pBorde As New Panel With {.Height = 4, .Dock = DockStyle.Top, .BackColor = colorAcento}
        Dim lTit As New Label With {.Text = titulo, .Location = New Point(10, 20), .ForeColor = Color.Gray, .Font = New Font("Segoe UI", 8, FontStyle.Bold), .AutoSize = True}
        Dim lVal As New Label With {.Text = valor, .Location = New Point(10, 45), .ForeColor = colorAcento, .Font = New Font("Segoe UI", 16, FontStyle.Bold), .AutoSize = True}
        pnl.Controls.AddRange({pBorde, lTit, lVal})
        Return pnl
    End Function
End Class