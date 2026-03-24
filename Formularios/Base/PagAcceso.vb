Imports System.Drawing.Drawing2D

Public Class PagAcceso

    ' Creamos el panel central por código (La "Tarjeta" de Login)
    Private PanelLogin As New Panel()
    ' Importamos la función nativa para congelar el dibujo de la ventana
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Private Const WM_SETREDRAW As Integer = &HB
    Public Sub New()
        ' 1. Esta llamada es obligatoria para el diseñador
        InitializeComponent()

        ' 2. TRUCO MAESTRO: Movemos el formulario a "Cuenca" (fuera del monitor) 
        ' para que se dibuje en el vacío absoluto.
        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New Point(-10000, -10000)

        ' 3. Construimos todo ya mismo
        ConstruirInterfazPremium()
    End Sub
    Private Sub PagAcceso_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' La interfaz ya se construyó en el New, así que aquí solo centramos y mostramos
        CentrarPanelLogin()

        ' Traemos el formulario al centro de la pantalla del usuario
        Me.Location = New Point((Screen.PrimaryScreen.WorkingArea.Width - Me.Width) \ 2,
                            (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) \ 2)

        ' Forzamos un repintado final limpio
        Me.Refresh()
    End Sub

    ' =========================================================
    ' 1. CONSTRUCTOR DE INTERFAZ (TODO POR CÓDIGO)
    ' =========================================================
    Private Sub ConstruirInterfazPremium()
        ' --- A) CREAR LA TARJETA CENTRAL ---
        PanelLogin.Size = New Size(400, 520)
        PanelLogin.BackColor = Color.FromArgb(220, 25, 35, 45)

        ' NUEVO: Activamos DoubleBuffer por reflexión para evitar parpadeos en el panel transparente
        Dim meth = GetType(Control).GetMethod("SetStyle", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)
        meth.Invoke(PanelLogin, New Object() {ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint Or ControlStyles.DoubleBuffer, True})

        Me.Controls.Add(PanelLogin)

        ' --- B) ATRAPAR EL LOGO DEL DISEÑADOR ---
        ' Busca el PictureBox que tienes en el diseño visual y lo mete en la tarjeta
        Dim picLogo As PictureBox = Me.Controls.OfType(Of PictureBox)().FirstOrDefault()

        If picLogo IsNot Nothing Then
            PanelLogin.Controls.Add(picLogo)
            picLogo.Bounds = New Rectangle(50, 30, 300, 100) ' Lo coloca exacto arriba del centro
            picLogo.BackColor = Color.Transparent ' Para que se funda con la tarjeta oscura
            picLogo.SizeMode = PictureBoxSizeMode.Zoom ' Para que no se deforme
            picLogo.BringToFront()
        End If

        ' --- C) ORDENAR CAJAS DE TEXTO Y LABELS ---
        ConfigurarCajaTexto(TextBoxUsuario, "USUARIO", 50, 160)
        ConfigurarCajaTexto(TextBoxPassword, "CONTRASEÑA", 50, 240)

        ' --- D) CONFIGURAR EL BOTÓN DEL OJITO ---
        ' Lo colocamos justo encima del final de la caja de contraseña
        PanelLogin.Controls.Add(ButtonVerPasswd)
        ButtonVerPasswd.Bounds = New Rectangle(315, 267, 33, 26)
        ButtonVerPasswd.FlatStyle = FlatStyle.Flat
        ButtonVerPasswd.FlatAppearance.BorderSize = 0
        ButtonVerPasswd.BackColor = Color.White
        ButtonVerPasswd.Cursor = Cursors.Hand
        ButtonVerPasswd.BringToFront()

        ' --- E) CONFIGURAR EL COMBOBOX EMPRESA (Opcional) ---
        ' Busca el ComboBox en tu formulario y lo ordena
        Dim cmbEmpresa As ComboBox = Me.Controls.OfType(Of ComboBox).FirstOrDefault()
        If cmbEmpresa IsNot Nothing Then
            ConfigurarLabel("EMPRESA", 50, 320)
            PanelLogin.Controls.Add(cmbEmpresa)
            cmbEmpresa.Bounds = New Rectangle(50, 345, 300, 28)
            cmbEmpresa.Font = New Font("Segoe UI", 12)
            cmbEmpresa.DropDownStyle = ComboBoxStyle.DropDownList
        End If

        ' --- F) ORDENAR LOS BOTONES DE ACCIÓN ---
        PanelLogin.Controls.Add(ButtonCrearUsuario)
        PanelLogin.Controls.Add(ButtonAcceso)

        ' Matemáticas: Ancho total 300. Botón 1 = 140, Hueco = 20, Botón 2 = 140.
        EstilizarBoton(ButtonCrearUsuario, 50, 420, 140, 40, Color.FromArgb(80, 90, 100), "Crear")
        EstilizarBoton(ButtonAcceso, 210, 420, 140, 40, Color.FromArgb(0, 120, 215), "Entrar")

        ' Centrar por primera vez
        CentrarPanelLogin()
    End Sub

    ' =========================================================
    ' 2. FUNCIONES DE DISEÑO AUTOMÁTICO
    ' =========================================================
    Private Sub ConfigurarCajaTexto(txt As TextBox, titulo As String, x As Integer, y As Integer)
        ' 1. Crea la etiqueta del título
        ConfigurarLabel(titulo, x, y)

        ' 2. Mete la caja de texto en la tarjeta y le da estilo
        If txt IsNot Nothing Then
            PanelLogin.Controls.Add(txt)
            txt.Bounds = New Rectangle(x, y + 25, 300, 30)
            txt.Font = New Font("Segoe UI", 12)
            txt.BorderStyle = BorderStyle.FixedSingle
        End If
    End Sub

    Private Sub ConfigurarLabel(texto As String, x As Integer, y As Integer)
        Dim lbl As New Label()
        lbl.Text = texto
        lbl.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        lbl.ForeColor = Color.FromArgb(200, 200, 200) ' Gris clarito elegante
        lbl.AutoSize = True
        lbl.Location = New Point(x, y)
        PanelLogin.Controls.Add(lbl)
    End Sub

    Private Sub EstilizarBoton(btn As Button, x As Integer, y As Integer, w As Integer, h As Integer, bgColor As Color, texto As String)
        If btn IsNot Nothing Then
            btn.Bounds = New Rectangle(x, y, w, h)
            btn.FlatStyle = FlatStyle.Flat
            btn.FlatAppearance.BorderSize = 0
            btn.BackColor = bgColor
            btn.ForeColor = Color.White
            btn.Font = New Font("Segoe UI", 11, FontStyle.Bold)
            btn.Cursor = Cursors.Hand
            btn.Text = texto
        End If
    End Sub

    Private Sub CentrarPanelLogin()
        PanelLogin.Location = New Point((Me.ClientSize.Width - PanelLogin.Width) \ 2, (Me.ClientSize.Height - PanelLogin.Height) \ 2)
    End Sub

    ' Si la ventana cambia de tamaño, volvemos a centrar automáticamente
    Private Sub PagAcceso_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        CentrarPanelLogin()
    End Sub

    ' =========================================================
    ' 3. LÓGICA DE INTERACCIÓN (ACCESOS Y OJITO)
    ' =========================================================
    Private Sub ButtonVerPass_MouseDown(sender As Object, e As MouseEventArgs) Handles ButtonVerPasswd.MouseDown
        TextBoxPassword.UseSystemPasswordChar = False ' Muestra la contraseña
    End Sub

    Private Sub ButtonVerPass_MouseUp(sender As Object, e As MouseEventArgs) Handles ButtonVerPasswd.MouseUp
        TextBoxPassword.UseSystemPasswordChar = True  ' Oculta la contraseña
    End Sub

    Private Sub ButtonAcceso_Click(sender As Object, e As EventArgs) Handles ButtonAcceso.Click
        If String.IsNullOrWhiteSpace(TextBoxUsuario.Text) OrElse String.IsNullOrWhiteSpace(TextBoxPassword.Text) Then
            MessageBox.Show("Por favor, introduce usuario y contraseña.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ComunSesionActual.Usuario = TextBoxUsuario.Text
        ComunSesionActual.Contrasena = TextBoxPassword.Text

        Dim PagHome As New PagHome()
        AddHandler PagHome.FormClosed, Sub(s, args) Me.Close()
        PagHome.Show()
        Me.Hide()
    End Sub

    ' Pulsar ENTER en la contraseña para entrar rápido
    Private Sub TextBoxPassword_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBoxPassword.KeyDown
        If e.KeyCode = Keys.Enter Then
            ButtonAcceso.PerformClick()
            e.SuppressKeyPress = True
        End If
    End Sub

End Class

' (MANTÉN AQUÍ DEBAJO TU CLASE PanelTransparente)
' =========================================================
' PANEL SEMITRANSPARENTE CON LOGO AJUSTADO (ZOOM)
' =========================================================
Public Class PanelTransparente
    Inherits Panel

    ' Opacidad (0-255)
    Private _opacidad As Integer = 150
    Public Property Opacidad As Integer
        Get
            Return _opacidad
        End Get
        Set(value As Integer)
            _opacidad = value
            Me.Invalidate()
        End Set
    End Property

    ' Imagen del Logo
    Private _logo As Image
    Public Property Logo As Image
        Get
            Return _logo
        End Get
        Set(value As Image)
            _logo = value
            Me.Invalidate() ' Si cambia la foto, repintamos
        End Set
    End Property



    Protected Overrides Sub OnPaintBackground(e As PaintEventArgs)
        ' No pintamos fondo estándar
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        ' 1. Pintar el fondo blanco semitransparente (si la opacidad es > 0)
        If _opacidad > 0 Then
            Using pincel As New SolidBrush(Color.FromArgb(_opacidad, 255, 255, 255))
                e.Graphics.FillRectangle(pincel, Me.ClientRectangle)
            End Using
        End If

        ' 2. Pintar el Logo con modo ZOOM
        If _logo IsNot Nothing Then
            e.Graphics.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
            e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
            e.Graphics.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality

            Dim ratioX As Double = Me.Width / _logo.Width
            Dim ratioY As Double = Me.Height / _logo.Height
            Dim ratio As Double = Math.Min(ratioX, ratioY)

            Dim nuevoAncho As Integer = CInt(_logo.Width * ratio)
            Dim nuevoAlto As Integer = CInt(_logo.Height * ratio)

            Dim posX As Integer = (Me.Width - nuevoAncho) \ 2
            Dim posY As Integer = (Me.Height - nuevoAlto) \ 2

            e.Graphics.DrawImage(_logo, posX, posY, nuevoAncho, nuevoAlto)
        End If
    End Sub
    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            ' WS_EX_COMPOSITED (0x02000000) - Pinta todos los controles de una vez 
            ' de abajo hacia arriba en una sola pasada de memoria.
            cp.ExStyle = cp.ExStyle Or &H2000000
            Return cp
        End Get
    End Property
End Class
