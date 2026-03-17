Imports System.Drawing
Imports System.Windows.Forms

Public Class TabControlModerno
    Inherits TabControl

    Public Sub New()
        ' Activamos UserPaint para tener control total y evitar bordes blancos feos
        Me.SetStyle(ControlStyles.UserPaint Or ControlStyles.ResizeRedraw Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer, True)
        Me.ItemSize = New Size(120, 30)
        Me.SizeMode = TabSizeMode.Fixed
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim g As Graphics = e.Graphics

        ' 1. PINTAR EL FONDO (El hueco vacío a la derecha)
        ' Usamos el mismo gris oscuro que tu formulario para que se fusione
        Dim colorFondoGeneral As Color = Color.FromArgb(45, 45, 48)
        g.Clear(colorFondoGeneral)

        ' 2. PINTAR LAS PESTAÑAS (Iteramos sobre todas)
        For i As Integer = 0 To Me.TabCount - 1
            Dim tabRect As Rectangle = Me.GetTabRect(i)
            Dim isSelected As Boolean = (Me.SelectedIndex = i)
            Dim texto As String = Me.TabPages(i).Text

            ' Colores
            Dim colorFondoPestaña As Color
            Dim colorTexto As Color

            If isSelected Then
                colorFondoPestaña = Color.FromArgb(70, 75, 80)  ' Azul Activo
                colorTexto = Color.White
            Else
                colorFondoPestaña = colorFondoGeneral ' Gris Inactivo (Se funde con el fondo)
                colorTexto = Color.Gray
            End If

            ' A) Rellenar el rectángulo de la pestaña
            Using b As New SolidBrush(colorFondoPestaña)
                g.FillRectangle(b, tabRect)
            End Using

            ' B) Dibujar el texto centrado
            TextRenderer.DrawText(g, texto, Me.Font, tabRect, colorTexto, TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
        Next
    End Sub
End Class