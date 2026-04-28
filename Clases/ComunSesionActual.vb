Public Module ComunSesionActual
    Public Property Usuario As String
    Public Property Contrasena As String
    Public Property Empresa As String

    ' === AMPLIACIÓN: identidad real del usuario logueado ===
    ' Antes el sistema solo guardaba el nombre (texto) del usuario y nada más, así que los movimientos
    ' de almacén se atribuían siempre al usuario con ID=1 (hardcodeado). Ahora guardamos también el
    ' ID numérico (para auditoría) y el Rol (para los chequeos de permisos).
    Public Property IdUsuario As Integer = 0
    Public Property Rol As String = ""
    Public Property NombreCompleto As String = ""

    ' Helpers de comprobación de rol. Un poco más legibles que escribir If Rol = "Admin" por ahí.
    Public Function EsAdministrador() As Boolean
        Return Rol IsNot Nothing AndAlso Rol.Trim().Equals("Admin", StringComparison.OrdinalIgnoreCase)
    End Function

    Public Sub CerrarSesion()
        Usuario = ""
        Contrasena = ""
        Empresa = ""
        IdUsuario = 0
        Rol = ""
        NombreCompleto = ""
    End Sub
End Module
