Imports System.Data.SQLite
Imports System.IO
Imports System.Drawing.Imaging

Public Class FrmConfiguracion

    ' Controles UI principales
    Private tcConfig As New TabControl()
    Private tabEmpresa As New TabPage("1. Datos de Empresa")
    Private tabUsuarios As New TabPage("2. Usuarios y Permisos")
    Private tabBackups As New TabPage("3. Copias de Seguridad")
    Private WithEvents btnGuardar As New Button()

    ' --- CONSTANTE DE TU BASE DE DATOS (¡CÁMBIALO POR EL TUYO!) ---
    ' --- CONSTANTES DE TU BASE DE DATOS ---
    Private Const NOMBRE_ARCHIVO_BD As String = "BD\Optima.db"
    Private ReadOnly RUTA_ACTUAL_BD As String = Path.Combine(Application.StartupPath, NOMBRE_ARCHIVO_BD)
    Private ReadOnly RUTA_PAPELERA As String = Path.Combine(Application.StartupPath, "PapeleraDB")

    ' NUEVO: La ruta estándar donde irán las copias si el usuario no elige otra
    Private ReadOnly RUTA_BACKUP_DEFECTO As String = Path.Combine(Application.StartupPath, "BackUp")

    ' --- Controles Pestaña 1 (Empresa) ---
    Private txtNombreFiscal, txtNombreComercial, txtCIF, txtDireccion As New TextBox()
    Private txtPoblacion, txtCP, txtProvincia, txtTelefono, txtEmail, txtWeb, txtRegistro As New TextBox()
    Private picLogo As New PictureBox()
    Private WithEvents btnCargarLogo As New Button()

    ' --- Controles Pestaña 2 (Usuarios) ---
    Private WithEvents dgvUsuarios As New DataGridView()
    Private txtIdUsuario, txtUsername, txtPassword, txtNombreCompleto, txtEmailUsuario As New TextBox()
    Private cboRol As New ComboBox()
    Private chkActivo As New CheckBox()
    Private WithEvents btnNuevoUsuario, btnGuardarUsuario, btnEliminarUsuario As New Button()
    Private _dtUsuarios As New DataTable()

    ' --- Controles Pestaña 3 (Backups) ---
    Private txtRutaBackup As New TextBox()
    Private WithEvents btnExaminarRuta As New Button()
    Private chkAutoBackup As New CheckBox()
    Private WithEvents cboFrecuencia As New ComboBox()
    Private dtpHoraBackup As New DateTimePicker()
    Private chkDias As New CheckedListBox()
    Private WithEvents btnHacerCopia As New Button()
    Private WithEvents btnRestaurarCopia As New Button()
    Private WithEvents btnVaciarPapelera As New Button()

    Private Sub FrmConfiguracion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Configuración del Sistema"
        Me.Size = New Size(880, 700)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.FromArgb(70, 75, 80)

        ' Creamos las carpetas base si no existen
        If Not Directory.Exists(RUTA_PAPELERA) Then Directory.CreateDirectory(RUTA_PAPELERA)
        If Not Directory.Exists(RUTA_BACKUP_DEFECTO) Then Directory.CreateDirectory(RUTA_BACKUP_DEFECTO) ' <--- LÍNEA NUEVA

        ConstruirInterfaz()
        CargarDatosGenerales()
        CargarListaUsuarios()
    End Sub

    Public Sub IrAPestana(indice As Integer)
        If indice < tcConfig.TabCount Then tcConfig.SelectedIndex = indice
    End Sub

    ' =========================================================
    ' 1. DISEÑO DE LA INTERFAZ VISUAL (LAS 3 PESTAÑAS)
    ' =========================================================
    Private Sub ConstruirInterfaz()
        Me.Controls.Clear()

        Dim lblTitulo As New Label With {.Text = "AJUSTES DEL SISTEMA", .Font = New Font("Segoe UI", 18, FontStyle.Bold), .ForeColor = Color.WhiteSmoke, .AutoSize = True, .Location = New Point(30, 15)}
        Me.Controls.Add(lblTitulo)

        btnGuardar.Text = "Guardar Ajustes Globales" : btnGuardar.Bounds = New Rectangle(630, 15, 200, 35)
        btnGuardar.BackColor = Color.FromArgb(40, 140, 90) : btnGuardar.ForeColor = Color.White : btnGuardar.FlatStyle = FlatStyle.Flat : btnGuardar.FlatAppearance.BorderSize = 0 : btnGuardar.Font = New Font("Segoe UI", 10, FontStyle.Bold) : btnGuardar.Cursor = Cursors.Hand
        btnGuardar.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        Me.Controls.Add(btnGuardar)

        tcConfig.Bounds = New Rectangle(30, 65, 800, 560)
        tcConfig.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        tcConfig.Font = New Font("Segoe UI", 10)
        tcConfig.ItemSize = New Size(220, 32)
        Me.Controls.Add(tcConfig)

        tabEmpresa.BackColor = Color.White : tabUsuarios.BackColor = Color.White : tabBackups.BackColor = Color.White
        tcConfig.TabPages.Add(tabEmpresa) : tcConfig.TabPages.Add(tabUsuarios) : tcConfig.TabPages.Add(tabBackups)

        Dim CrearTituloSeccion = Sub(tab As TabPage, texto As String, x As Integer, y As Integer, ancho As Integer)
                                     Dim lbl As New Label With {.Text = texto, .Location = New Point(x, y), .AutoSize = True, .ForeColor = Color.FromArgb(0, 120, 215), .Font = New Font("Segoe UI", 10, FontStyle.Bold)}
                                     Dim lin As New Label With {.Bounds = New Rectangle(x, y + 22, ancho, 2), .BackColor = Color.FromArgb(230, 230, 230)}
                                     tab.Controls.Add(lbl) : tab.Controls.Add(lin)
                                 End Sub

        Dim CrearCampo = Sub(tab As TabPage, texto As String, txt As TextBox, x As Integer, y As Integer, w As Integer)
                             Dim lbl As New Label With {.Text = texto, .Location = New Point(x, y), .AutoSize = True, .ForeColor = Color.DimGray, .Font = New Font("Segoe UI", 9, FontStyle.Regular)}
                             txt.Bounds = New Rectangle(x, y + 20, w, 26) : txt.Font = New Font("Segoe UI", 10)
                             tab.Controls.Add(lbl) : tab.Controls.Add(txt)
                         End Sub

        Dim xIzq As Integer = 30
        Dim anchoTotal As Integer = 730

        ' -----------------------------------------------------------
        ' PESTAÑA 1: EMPRESA
        ' -----------------------------------------------------------
        Dim yPos As Integer = 20
        CrearTituloSeccion(tabEmpresa, "DATOS FISCALES Y COMERCIALES", xIzq, yPos, anchoTotal)
        yPos += 40
        CrearCampo(tabEmpresa, "Nombre Fiscal", txtNombreFiscal, xIzq, yPos, 320)
        CrearCampo(tabEmpresa, "Nombre Comercial", txtNombreComercial, xIzq + 340, yPos, 240)
        CrearCampo(tabEmpresa, "CIF / NIF", txtCIF, xIzq + 600, yPos, 130)
        yPos += 60
        CrearCampo(tabEmpresa, "Dirección Completa", txtDireccion, xIzq, yPos, 400)
        CrearCampo(tabEmpresa, "Población", txtPoblacion, xIzq + 420, yPos, 200)
        CrearCampo(tabEmpresa, "C.P.", txtCP, xIzq + 640, yPos, 90)
        yPos += 60
        CrearCampo(tabEmpresa, "Provincia", txtProvincia, xIzq, yPos, 180)
        CrearCampo(tabEmpresa, "Teléfono", txtTelefono, xIzq + 200, yPos, 180)
        CrearCampo(tabEmpresa, "Email", txtEmail, xIzq + 400, yPos, 330)
        yPos += 60
        CrearCampo(tabEmpresa, "Sitio Web", txtWeb, xIzq, yPos, 280)
        CrearCampo(tabEmpresa, "Datos Registrales (Para pie de facturas)", txtRegistro, xIzq + 300, yPos, 430)

        yPos += 75
        CrearTituloSeccion(tabEmpresa, "IDENTIDAD CORPORATIVA", xIzq, yPos, anchoTotal)
        yPos += 40
        picLogo.Bounds = New Rectangle(xIzq, yPos, 220, 100) : picLogo.BorderStyle = BorderStyle.FixedSingle : picLogo.SizeMode = PictureBoxSizeMode.Zoom : picLogo.BackColor = Color.FromArgb(245, 245, 245)
        tabEmpresa.Controls.Add(picLogo)
        btnCargarLogo.Text = "Examinar Imagen..." : btnCargarLogo.Bounds = New Rectangle(xIzq + 240, yPos + 35, 150, 32) : btnCargarLogo.BackColor = Color.FromArgb(220, 220, 220) : btnCargarLogo.ForeColor = Color.Black : btnCargarLogo.FlatStyle = FlatStyle.Flat : btnCargarLogo.FlatAppearance.BorderSize = 0 : btnCargarLogo.Cursor = Cursors.Hand
        tabEmpresa.Controls.Add(btnCargarLogo)

        ' -----------------------------------------------------------
        ' PESTAÑA 2: USUARIOS
        ' -----------------------------------------------------------
        Dim yUsr As Integer = 20
        CrearTituloSeccion(tabUsuarios, "LISTADO DE USUARIOS DEL SISTEMA", xIzq, yUsr, anchoTotal)
        yUsr += 35

        dgvUsuarios.Bounds = New Rectangle(xIzq, yUsr, anchoTotal, 180)
        dgvUsuarios.AllowUserToAddRows = False
        dgvUsuarios.ReadOnly = True
        dgvUsuarios.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvUsuarios.RowHeadersVisible = False
        dgvUsuarios.BackgroundColor = Color.White
        Try : FrmPresupuestos.EstilizarGrid(dgvUsuarios) : Catch : End Try
        tabUsuarios.Controls.Add(dgvUsuarios)

        yUsr += 200
        CrearTituloSeccion(tabUsuarios, "FICHA DEL USUARIO", xIzq, yUsr, anchoTotal)
        yUsr += 40

        txtIdUsuario.Visible = False
        tabUsuarios.Controls.Add(txtIdUsuario)
        CrearCampo(tabUsuarios, "Usuario (Login)", txtUsername, xIzq, yUsr, 180)
        CrearCampo(tabUsuarios, "Contraseña", txtPassword, xIzq + 200, yUsr, 180)
        txtPassword.UseSystemPasswordChar = True

        Dim lblRol As New Label With {.Text = "Rol del Sistema", .Location = New Point(xIzq + 400, yUsr), .AutoSize = True, .ForeColor = Color.DimGray, .Font = New Font("Segoe UI", 9)}
        cboRol.Bounds = New Rectangle(xIzq + 400, yUsr + 20, 160, 26) : cboRol.Font = New Font("Segoe UI", 10) : cboRol.DropDownStyle = ComboBoxStyle.DropDownList : cboRol.Items.AddRange({"Administrador", "Usuario"})
        tabUsuarios.Controls.Add(lblRol) : tabUsuarios.Controls.Add(cboRol)

        chkActivo.Text = "Cuenta Activa" : chkActivo.Bounds = New Rectangle(xIzq + 580, yUsr + 22, 130, 24) : chkActivo.Font = New Font("Segoe UI", 10, FontStyle.Bold) : chkActivo.Checked = True
        tabUsuarios.Controls.Add(chkActivo)

        yUsr += 60
        CrearCampo(tabUsuarios, "Nombre Completo", txtNombreCompleto, xIzq, yUsr, 350)
        CrearCampo(tabUsuarios, "Correo Electrónico", txtEmailUsuario, xIzq + 370, yUsr, 360)

        yUsr += 70
        btnNuevoUsuario.Text = "+ Nuevo" : btnNuevoUsuario.Bounds = New Rectangle(xIzq, yUsr, 100, 32) : btnNuevoUsuario.BackColor = Color.FromArgb(0, 120, 215) : btnNuevoUsuario.ForeColor = Color.White : btnNuevoUsuario.FlatStyle = FlatStyle.Flat : btnNuevoUsuario.FlatAppearance.BorderSize = 0 : btnNuevoUsuario.Cursor = Cursors.Hand
        tabUsuarios.Controls.Add(btnNuevoUsuario)
        btnGuardarUsuario.Text = "Guardar Usuario" : btnGuardarUsuario.Bounds = New Rectangle(xIzq + 110, yUsr, 140, 32) : btnGuardarUsuario.BackColor = Color.FromArgb(40, 140, 90) : btnGuardarUsuario.ForeColor = Color.White : btnGuardarUsuario.FlatStyle = FlatStyle.Flat : btnGuardarUsuario.FlatAppearance.BorderSize = 0 : btnGuardarUsuario.Cursor = Cursors.Hand
        tabUsuarios.Controls.Add(btnGuardarUsuario)
        btnEliminarUsuario.Text = "Eliminar" : btnEliminarUsuario.Bounds = New Rectangle(xIzq + 260, yUsr, 100, 32) : btnEliminarUsuario.BackColor = Color.FromArgb(209, 52, 56) : btnEliminarUsuario.ForeColor = Color.White : btnEliminarUsuario.FlatStyle = FlatStyle.Flat : btnEliminarUsuario.FlatAppearance.BorderSize = 0 : btnEliminarUsuario.Cursor = Cursors.Hand
        tabUsuarios.Controls.Add(btnEliminarUsuario)

        ' -----------------------------------------------------------
        ' PESTAÑA 3: COPIAS DE SEGURIDAD
        ' -----------------------------------------------------------
        Dim yBak As Integer = 20
        CrearTituloSeccion(tabBackups, "COPIA MANUAL Y RESTAURACIÓN", xIzq, yBak, anchoTotal)
        yBak += 40
        CrearCampo(tabBackups, "Carpeta destino de las copias", txtRutaBackup, xIzq, yBak, 400)

        btnExaminarRuta.Text = "Examinar..." : btnExaminarRuta.Bounds = New Rectangle(xIzq + 410, yBak + 19, 100, 28)
        btnExaminarRuta.BackColor = Color.FromArgb(220, 220, 220) : btnExaminarRuta.FlatStyle = FlatStyle.Flat : btnExaminarRuta.FlatAppearance.BorderSize = 0
        tabBackups.Controls.Add(btnExaminarRuta)

        btnHacerCopia.Text = "💾 HACER COPIA AHORA" : btnHacerCopia.Bounds = New Rectangle(xIzq, yBak + 60, 220, 40)
        btnHacerCopia.BackColor = Color.FromArgb(0, 120, 215) : btnHacerCopia.ForeColor = Color.White : btnHacerCopia.Font = New Font("Segoe UI", 10, FontStyle.Bold) : btnHacerCopia.FlatStyle = FlatStyle.Flat : btnHacerCopia.FlatAppearance.BorderSize = 0 : btnHacerCopia.Cursor = Cursors.Hand
        tabBackups.Controls.Add(btnHacerCopia)

        btnRestaurarCopia.Text = "🔄 RESTAURAR COPIA (PELIGRO)" : btnRestaurarCopia.Bounds = New Rectangle(xIzq + 240, yBak + 60, 260, 40)
        btnRestaurarCopia.BackColor = Color.FromArgb(209, 52, 56) : btnRestaurarCopia.ForeColor = Color.White : btnRestaurarCopia.Font = New Font("Segoe UI", 10, FontStyle.Bold) : btnRestaurarCopia.FlatStyle = FlatStyle.Flat : btnRestaurarCopia.FlatAppearance.BorderSize = 0 : btnRestaurarCopia.Cursor = Cursors.Hand
        tabBackups.Controls.Add(btnRestaurarCopia)

        yBak += 130
        CrearTituloSeccion(tabBackups, "COPIAS AUTOMÁTICAS (Requiere programa abierto)", xIzq, yBak, anchoTotal)
        yBak += 40

        chkAutoBackup.Text = "Activar copias de seguridad automáticas" : chkAutoBackup.Bounds = New Rectangle(xIzq, yBak, 350, 24) : chkAutoBackup.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        tabBackups.Controls.Add(chkAutoBackup)

        yBak += 35
        Dim lblFrec As New Label With {.Text = "Frecuencia:", .Location = New Point(xIzq, yBak + 5), .AutoSize = True}
        cboFrecuencia.Bounds = New Rectangle(xIzq + 80, yBak, 120, 26) : cboFrecuencia.DropDownStyle = ComboBoxStyle.DropDownList : cboFrecuencia.Items.AddRange({"Diaria", "Semanal"}) : cboFrecuencia.SelectedIndex = 0
        tabBackups.Controls.Add(lblFrec) : tabBackups.Controls.Add(cboFrecuencia)

        Dim lblHora As New Label With {.Text = "Hora:", .Location = New Point(xIzq + 220, yBak + 5), .AutoSize = True}
        dtpHoraBackup.Bounds = New Rectangle(xIzq + 270, yBak, 100, 26) : dtpHoraBackup.Format = DateTimePickerFormat.Time : dtpHoraBackup.ShowUpDown = True
        tabBackups.Controls.Add(lblHora) : tabBackups.Controls.Add(dtpHoraBackup)

        Dim lblDias As New Label With {.Text = "Días de la semana (Si es Semanal):", .Location = New Point(xIzq + 400, yBak - 20), .AutoSize = True, .ForeColor = Color.DimGray}
        chkDias.Bounds = New Rectangle(xIzq + 400, yBak, 150, 110) : chkDias.Items.AddRange({"Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"}) : chkDias.CheckOnClick = True : chkDias.Enabled = False
        tabBackups.Controls.Add(lblDias) : tabBackups.Controls.Add(chkDias)

        yBak += 140
        CrearTituloSeccion(tabBackups, "MANTENIMIENTO", xIzq, yBak, anchoTotal)
        yBak += 40

        Dim lblPapelera As New Label With {.Text = "Al restaurar una copia, la base de datos antigua se guarda por seguridad.", .Location = New Point(xIzq, yBak), .AutoSize = True, .ForeColor = Color.DimGray}
        tabBackups.Controls.Add(lblPapelera)

        btnVaciarPapelera.Text = "🗑️ Vaciar Papelera de Bases de Datos (" & ObtenerTamanoPapelera() & ")" : btnVaciarPapelera.Bounds = New Rectangle(xIzq, yBak + 30, 350, 32) : btnVaciarPapelera.BackColor = Color.FromArgb(70, 75, 80) : btnVaciarPapelera.ForeColor = Color.White : btnVaciarPapelera.FlatStyle = FlatStyle.Flat : btnVaciarPapelera.FlatAppearance.BorderSize = 0 : btnVaciarPapelera.Cursor = Cursors.Hand
        tabBackups.Controls.Add(btnVaciarPapelera)
    End Sub

    ' =========================================================
    ' EVENTOS UI SECUNDARIOS
    ' =========================================================
    Private Sub cboFrecuencia_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboFrecuencia.SelectedIndexChanged
        chkDias.Enabled = (cboFrecuencia.Text = "Semanal")
    End Sub

    Private Sub btnCargarLogo_Click(sender As Object, e As EventArgs) Handles btnCargarLogo.Click
        Dim opf As New OpenFileDialog() With {.Filter = "Imágenes (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg"}
        If opf.ShowDialog() = DialogResult.OK Then
            Try : picLogo.Image = Image.FromFile(opf.FileName) : Catch : MessageBox.Show("Error al cargar.") : End Try
        End If
    End Sub

    Private Function ImageToBytes(img As Image) As Byte()
        If img Is Nothing Then Return Nothing
        Using ms As New MemoryStream()
            img.Save(ms, ImageFormat.Png)
            Return ms.ToArray()
        End Using
    End Function

    Private Function BytesToImage(b As Byte()) As Image
        If b Is Nothing OrElse b.Length = 0 Then Return Nothing
        Using ms As New MemoryStream(b)
            Return Image.FromStream(ms)
        End Using
    End Function

    ' =========================================================
    ' LÓGICA DE USUARIOS (CRUD)
    ' =========================================================
    Private Sub CargarListaUsuarios()
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()

            Dim sql As String = "SELECT ID_Usuario, NombreUsuario, Password, Rol, NombreCompleto, Email, Activo, FechaRegistro FROM Usuarios"
            Using cmd As New SQLiteCommand(sql, c)
                Dim da As New SQLiteDataAdapter(cmd)
                _dtUsuarios.Clear()
                da.Fill(_dtUsuarios)
            End Using

            dgvUsuarios.DataSource = _dtUsuarios

            If dgvUsuarios.Columns.Count > 0 Then
                dgvUsuarios.Columns("ID_Usuario").Visible = False
                dgvUsuarios.Columns("Password").Visible = False
                dgvUsuarios.Columns("NombreUsuario").HeaderText = "Login"
                dgvUsuarios.Columns("NombreUsuario").Width = 100
                dgvUsuarios.Columns("NombreCompleto").HeaderText = "Nombre Completo"
                dgvUsuarios.Columns("NombreCompleto").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                dgvUsuarios.Columns("Rol").Width = 120
                dgvUsuarios.Columns("Email").Width = 200
                dgvUsuarios.Columns("Activo").Width = 60
                dgvUsuarios.Columns("FechaRegistro").HeaderText = "F. Registro"
                dgvUsuarios.Columns("FechaRegistro").Width = 120
            End If
        Catch ex As Exception
            MessageBox.Show("Error al cargar usuarios: " & ex.Message)
        End Try
    End Sub

    Private Sub dgvUsuarios_SelectionChanged(sender As Object, e As EventArgs) Handles dgvUsuarios.SelectionChanged
        If dgvUsuarios.CurrentRow IsNot Nothing Then
            Dim row As DataGridViewRow = dgvUsuarios.CurrentRow
            txtIdUsuario.Text = row.Cells("ID_Usuario").Value.ToString()
            txtUsername.Text = If(IsDBNull(row.Cells("NombreUsuario").Value), "", row.Cells("NombreUsuario").Value.ToString())
            txtPassword.Text = If(IsDBNull(row.Cells("Password").Value), "", row.Cells("Password").Value.ToString())
            cboRol.Text = If(IsDBNull(row.Cells("Rol").Value), "Usuario", row.Cells("Rol").Value.ToString())
            txtNombreCompleto.Text = If(IsDBNull(row.Cells("NombreCompleto").Value), "", row.Cells("NombreCompleto").Value.ToString())
            txtEmailUsuario.Text = If(IsDBNull(row.Cells("Email").Value), "", row.Cells("Email").Value.ToString())
            chkActivo.Checked = If(IsDBNull(row.Cells("Activo").Value), False, Convert.ToBoolean(row.Cells("Activo").Value))
        End If
    End Sub

    Private Sub btnNuevoUsuario_Click(sender As Object, e As EventArgs) Handles btnNuevoUsuario.Click
        dgvUsuarios.ClearSelection()
        txtIdUsuario.Text = "" : txtUsername.Text = "" : txtPassword.Text = "" : txtNombreCompleto.Text = "" : txtEmailUsuario.Text = ""
        cboRol.SelectedIndex = 1 : chkActivo.Checked = True : txtUsername.Focus()
    End Sub

    Private Sub btnGuardarUsuario_Click(sender As Object, e As EventArgs) Handles btnGuardarUsuario.Click
        If String.IsNullOrWhiteSpace(txtUsername.Text) OrElse String.IsNullOrWhiteSpace(txtPassword.Text) Then
            MessageBox.Show("Login y Contraseña son obligatorios.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning) : Return
        End If
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()
            Dim cmd As New SQLiteCommand(c)

            If String.IsNullOrEmpty(txtIdUsuario.Text) Then
                cmd.CommandText = "INSERT INTO Usuarios (NombreUsuario, Password, Rol, NombreCompleto, Email, Activo, FechaRegistro) VALUES (@user, @pass, @rol, @nom, @email, @act, @fecha)"
                cmd.Parameters.AddWithValue("@fecha", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
            Else
                cmd.CommandText = "UPDATE Usuarios SET NombreUsuario=@user, Password=@pass, Rol=@rol, NombreCompleto=@nom, Email=@email, Activo=@act WHERE ID_Usuario=@id"
                cmd.Parameters.AddWithValue("@id", txtIdUsuario.Text)
            End If

            cmd.Parameters.AddWithValue("@user", txtUsername.Text.Trim())
            cmd.Parameters.AddWithValue("@pass", txtPassword.Text.Trim())
            cmd.Parameters.AddWithValue("@rol", cboRol.Text)
            cmd.Parameters.AddWithValue("@nom", txtNombreCompleto.Text.Trim())
            cmd.Parameters.AddWithValue("@email", txtEmailUsuario.Text.Trim())
            cmd.Parameters.AddWithValue("@act", If(chkActivo.Checked, 1, 0))

            cmd.ExecuteNonQuery()
            CargarListaUsuarios()
            MessageBox.Show("Usuario guardado.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error al guardar el usuario: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnEliminarUsuario_Click(sender As Object, e As EventArgs) Handles btnEliminarUsuario.Click
        If String.IsNullOrEmpty(txtIdUsuario.Text) Then Return
        If MessageBox.Show("¿Eliminar usuario '" & txtUsername.Text & "'?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Try
                Dim c = ConexionBD.GetConnection()
                If c.State <> ConnectionState.Open Then c.Open()
                Dim cmd As New SQLiteCommand("DELETE FROM Usuarios WHERE ID_Usuario = @id", c)
                cmd.Parameters.AddWithValue("@id", txtIdUsuario.Text)
                cmd.ExecuteNonQuery()
                btnNuevoUsuario.PerformClick()
                CargarListaUsuarios()
            Catch ex As Exception
                MessageBox.Show("Error al eliminar: " & ex.Message)
            End Try
        End If
    End Sub

    ' =========================================================
    ' LÓGICA DE COPIAS DE SEGURIDAD
    ' =========================================================
    Private Sub btnExaminarRuta_Click(sender As Object, e As EventArgs) Handles btnExaminarRuta.Click
        Dim fbd As New FolderBrowserDialog() With {.Description = "Selecciona la carpeta para copias de seguridad"}
        If fbd.ShowDialog() = DialogResult.OK Then txtRutaBackup.Text = fbd.SelectedPath
    End Sub

    Private Sub btnHacerCopia_Click(sender As Object, e As EventArgs) Handles btnHacerCopia.Click
        If String.IsNullOrWhiteSpace(txtRutaBackup.Text) OrElse Not Directory.Exists(txtRutaBackup.Text) Then
            MessageBox.Show("Selecciona una carpeta válida.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning) : Return
        End If
        Try
            Dim nombreCopia As String = "OptimaDB_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".sqlite"
            Dim rutaDestino As String = Path.Combine(txtRutaBackup.Text, nombreCopia)
            File.Copy(RUTA_ACTUAL_BD, rutaDestino, True)
            MessageBox.Show("Copia realizada con éxito en:" & vbCrLf & rutaDestino, "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error al hacer copia: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnRestaurarCopia_Click(sender As Object, e As EventArgs) Handles btnRestaurarCopia.Click
        Dim opf As New OpenFileDialog() With {.Filter = "BD SQLite (*.db;*.sqlite)|*.db;*.sqlite|Todos (*.*)|*.*"}
        If opf.ShowDialog() = DialogResult.OK Then
            If MessageBox.Show("¡ATENCIÓN! Vas a reemplazar la BD actual." & vbCrLf & "Se creará una copia en papelera." & vbCrLf & "¿Seguro?", "Peligro", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes Then
                Try
                    ' 1. MATAMOS LA CONEXIÓN ABIERTA (La clave del éxito)
                    Dim c = ConexionBD.GetConnection()
                    If c.State = ConnectionState.Open Then c.Close()

                    ' 2. Limpiamos las "piscinas" de caché y forzamos a Windows a vaciar la memoria
                    SQLiteConnection.ClearAllPools()
                    GC.Collect()
                    GC.WaitForPendingFinalizers()

                    ' 3. Le damos al sistema operativo medio segundo para soltar el archivo del todo
                    System.Threading.Thread.Sleep(500)

                    ' 4. Ahora sí, movemos y copiamos
                    Dim rutaPapeleraDestino As String = Path.Combine(RUTA_PAPELERA, "BD_Eliminada_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".sqlite")

                    If File.Exists(RUTA_ACTUAL_BD) Then
                        File.Move(RUTA_ACTUAL_BD, rutaPapeleraDestino)
                    End If

                    File.Copy(opf.FileName, RUTA_ACTUAL_BD)

                    MessageBox.Show("Restauración completada." & vbCrLf & "El programa se cerrará para aplicar los cambios. Vuelve a abrirlo.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Application.Exit()

                Catch ex As Exception
                    MessageBox.Show("Error Crítico al restaurar: " & ex.Message & vbCrLf & vbCrLf &
                                    "CONSEJO: Comprueba que NO tienes la base de datos abierta en el programa 'DB Browser for SQLite'.", "Error de Bloqueo", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End If
    End Sub

    Private Function ObtenerTamanoPapelera() As String
        If Not Directory.Exists(RUTA_PAPELERA) Then Return "0 MB"
        Dim size As Long = 0
        For Each f As FileInfo In New DirectoryInfo(RUTA_PAPELERA).GetFiles()
            size += f.Length
        Next
        Return (size / 1048576.0).ToString("0.00") & " MB"
    End Function

    Private Sub btnVaciarPapelera_Click(sender As Object, e As EventArgs) Handles btnVaciarPapelera.Click
        If MessageBox.Show("¿Borrar permanentemente la papelera?", "Confirmar", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            Try
                For Each f As FileInfo In New DirectoryInfo(RUTA_PAPELERA).GetFiles()
                    f.Delete()
                Next
                btnVaciarPapelera.Text = "🗑️ Vaciar Papelera de Bases de Datos (0.00 MB)"
                MessageBox.Show("Papelera vaciada.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End If
    End Sub

    ' =========================================================
    ' CARGAR Y GUARDAR TODA LA CONFIGURACIÓN (EMPRESA + BACKUP)
    ' =========================================================
    Private Sub CargarDatosGenerales()
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()
            Dim sql As String = "SELECT * FROM Empresa WHERE ID = 1"
            Using cmd As New SQLiteCommand(sql, c)
                Using r As SQLiteDataReader = cmd.ExecuteReader()
                    If r.Read() Then
                        txtNombreFiscal.Text = If(IsDBNull(r("NombreFiscal")), "", r("NombreFiscal").ToString())
                        txtNombreComercial.Text = If(IsDBNull(r("NombreComercial")), "", r("NombreComercial").ToString())
                        txtCIF.Text = If(IsDBNull(r("CIF")), "", r("CIF").ToString())
                        txtDireccion.Text = If(IsDBNull(r("Direccion")), "", r("Direccion").ToString())
                        txtPoblacion.Text = If(IsDBNull(r("Poblacion")), "", r("Poblacion").ToString())
                        txtCP.Text = If(IsDBNull(r("CodigoPostal")), "", r("CodigoPostal").ToString())
                        txtProvincia.Text = If(IsDBNull(r("Provincia")), "", r("Provincia").ToString())
                        txtTelefono.Text = If(IsDBNull(r("Telefono")), "", r("Telefono").ToString())
                        txtEmail.Text = If(IsDBNull(r("Email")), "", r("Email").ToString())
                        txtWeb.Text = If(IsDBNull(r("SitioWeb")), "", r("SitioWeb").ToString())
                        txtRegistro.Text = If(IsDBNull(r("RegistroMercantil")), "", r("RegistroMercantil").ToString())
                        If Not IsDBNull(r("Logo")) Then picLogo.Image = BytesToImage(CType(r("Logo"), Byte()))

                        ' --- Datos Backup ---

                        ' Leemos la ruta que hay guardada en la base de datos
                        Dim rutaGuardada As String = If(IsDBNull(r("RutaBackup")), "", r("RutaBackup").ToString().Trim())

                        ' Si está vacía (es la primera vez que abre el programa), le ponemos la de defecto
                        If String.IsNullOrEmpty(rutaGuardada) Then
                            txtRutaBackup.Text = RUTA_BACKUP_DEFECTO
                        Else
                            txtRutaBackup.Text = rutaGuardada
                        End If
                        chkAutoBackup.Checked = If(IsDBNull(r("AutoBackup")), False, Convert.ToBoolean(r("AutoBackup")))
                        cboFrecuencia.Text = If(IsDBNull(r("FrecuenciaBackup")), "Diaria", r("FrecuenciaBackup").ToString())
                        If Not IsDBNull(r("HoraBackup")) Then
                            Dim horaG As DateTime
                            If DateTime.TryParse(r("HoraBackup").ToString(), horaG) Then dtpHoraBackup.Value = horaG
                        End If
                        If Not IsDBNull(r("DiasBackup")) Then
                            Dim diasG As String() = r("DiasBackup").ToString().Split(","c)
                            For i As Integer = 0 To chkDias.Items.Count - 1
                                chkDias.SetItemChecked(i, diasG.Contains(i.ToString()))
                            Next
                        End If
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error al cargar configuración: " & ex.Message)
        End Try
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Try
            Dim c = ConexionBD.GetConnection()
            If c.State <> ConnectionState.Open Then c.Open()

            Dim diasSeleccionados As New List(Of String)
            For i As Integer = 0 To chkDias.Items.Count - 1
                If chkDias.GetItemChecked(i) Then diasSeleccionados.Add(i.ToString())
            Next
            Dim diasString As String = String.Join(",", diasSeleccionados)

            Dim sql As String = "UPDATE Empresa SET NombreFiscal=@nomF, NombreComercial=@nomC, CIF=@cif, Direccion=@dir, Poblacion=@pob, CodigoPostal=@cp, Provincia=@prov, Telefono=@tel, Email=@ema, SitioWeb=@web, RegistroMercantil=@reg, Logo=@logo, RutaBackup=@rutaB, AutoBackup=@autoB, FrecuenciaBackup=@frecB, HoraBackup=@horaB, DiasBackup=@diasB WHERE ID = 1"

            Using cmd As New SQLiteCommand(sql, c)
                cmd.Parameters.AddWithValue("@nomF", txtNombreFiscal.Text.Trim())
                cmd.Parameters.AddWithValue("@nomC", txtNombreComercial.Text.Trim())
                cmd.Parameters.AddWithValue("@cif", txtCIF.Text.Trim())
                cmd.Parameters.AddWithValue("@dir", txtDireccion.Text.Trim())
                cmd.Parameters.AddWithValue("@pob", txtPoblacion.Text.Trim())
                cmd.Parameters.AddWithValue("@cp", txtCP.Text.Trim())
                cmd.Parameters.AddWithValue("@prov", txtProvincia.Text.Trim())
                cmd.Parameters.AddWithValue("@tel", txtTelefono.Text.Trim())
                cmd.Parameters.AddWithValue("@ema", txtEmail.Text.Trim())
                cmd.Parameters.AddWithValue("@web", txtWeb.Text.Trim())
                cmd.Parameters.AddWithValue("@reg", txtRegistro.Text.Trim())
                If picLogo.Image IsNot Nothing Then cmd.Parameters.AddWithValue("@logo", ImageToBytes(picLogo.Image)) Else cmd.Parameters.AddWithValue("@logo", DBNull.Value)

                cmd.Parameters.AddWithValue("@rutaB", txtRutaBackup.Text.Trim())
                cmd.Parameters.AddWithValue("@autoB", If(chkAutoBackup.Checked, 1, 0))
                cmd.Parameters.AddWithValue("@frecB", cboFrecuencia.Text)
                cmd.Parameters.AddWithValue("@horaB", dtpHoraBackup.Value.ToString("HH:mm"))
                cmd.Parameters.AddWithValue("@diasB", diasString)

                cmd.ExecuteNonQuery()
            End Using

            MessageBox.Show("Ajustes globales guardados correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error al guardar: " & ex.Message)
        End Try
    End Sub
End Class