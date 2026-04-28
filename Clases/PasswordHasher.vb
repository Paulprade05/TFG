Imports System.Security.Cryptography
Imports System.Text

' =====================================================================
'  PasswordHasher
'  Utilidad para almacenar contraseñas de forma segura.
'  ---------------------------------------------------------------------
'  Antes de este parche las contraseñas se guardaban en texto plano en
'  la columna Usuarios.Password. Ahora usamos PBKDF2-SHA256 con salt
'  aleatorio por usuario.
'
'  Formato almacenado en BD:    "PBKDF2$<iter>$<saltB64>$<hashB64>"
'  - iter: número de iteraciones (50.000 por defecto)
'  - saltB64: salt aleatorio (16 bytes) en Base64
'  - hashB64: hash de 32 bytes en Base64
'
'  RETROCOMPATIBILIDAD:
'  Si encontramos en BD una contraseña que NO empieza por "PBKDF2$"
'  asumimos que es texto plano (de la versión anterior) y la
'  comparamos directamente. En cuanto el usuario haga login bien,
'  la migramos al formato hasheado de forma transparente.
' =====================================================================
Public Class PasswordHasher

    Private Const PREFIJO As String = "PBKDF2"
    Private Const ITERACIONES As Integer = 50000
    Private Const SALT_BYTES As Integer = 16
    Private Const HASH_BYTES As Integer = 32

    ''' <summary>
    ''' Genera el string a guardar en BD para una contraseña.
    ''' </summary>
    Public Shared Function Hashear(password As String) As String
        If password Is Nothing Then password = ""
        Dim salt(SALT_BYTES - 1) As Byte
        Using rng As New RNGCryptoServiceProvider()
            rng.GetBytes(salt)
        End Using
        Dim hash As Byte() = DerivarHash(password, salt, ITERACIONES, HASH_BYTES)
        Return $"{PREFIJO}${ITERACIONES}${Convert.ToBase64String(salt)}${Convert.ToBase64String(hash)}"
    End Function

    ''' <summary>
    ''' Verifica una contraseña contra lo guardado en BD.
    ''' Acepta tanto el formato nuevo (PBKDF2$...) como texto plano (legacy).
    ''' </summary>
    Public Shared Function Verificar(passwordIntroducida As String, passwordEnBD As String) As Boolean
        If passwordEnBD Is Nothing Then Return False
        If passwordIntroducida Is Nothing Then passwordIntroducida = ""

        ' Caso legacy: contraseña antigua en texto plano
        If Not passwordEnBD.StartsWith(PREFIJO & "$") Then
            Return passwordIntroducida = passwordEnBD
        End If

        ' Formato nuevo
        Dim partes = passwordEnBD.Split("$"c)
        If partes.Length <> 4 Then Return False

        Try
            Dim iter As Integer = Integer.Parse(partes(1))
            Dim salt As Byte() = Convert.FromBase64String(partes(2))
            Dim hashEsperado As Byte() = Convert.FromBase64String(partes(3))
            Dim hashCalculado As Byte() = DerivarHash(passwordIntroducida, salt, iter, hashEsperado.Length)

            ' Comparación constante (evita timing attacks)
            Return CompararConstante(hashEsperado, hashCalculado)
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' True si la cadena guardada está en formato antiguo (texto plano)
    ''' y conviene migrarla.
    ''' </summary>
    Public Shared Function NecesitaMigracion(passwordEnBD As String) As Boolean
        If String.IsNullOrEmpty(passwordEnBD) Then Return False
        Return Not passwordEnBD.StartsWith(PREFIJO & "$")
    End Function

    Private Shared Function DerivarHash(password As String, salt As Byte(), iter As Integer, longitudBytes As Integer) As Byte()
        Using pbkdf2 As New Rfc2898DeriveBytes(password, salt, iter, HashAlgorithmName.SHA256)
            Return pbkdf2.GetBytes(longitudBytes)
        End Using
    End Function

    Private Shared Function CompararConstante(a As Byte(), b As Byte()) As Boolean
        If a.Length <> b.Length Then Return False
        Dim diff As Integer = 0
        For i As Integer = 0 To a.Length - 1
            diff = diff Or (a(i) Xor b(i))
        Next
        Return diff = 0
    End Function

End Class
