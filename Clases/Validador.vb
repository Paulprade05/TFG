Imports System.Text.RegularExpressions

' =====================================================================
'  Validador
'  Validaciones de datos de entrada para los formularios.
'  Cada función devuelve (Boolean, String):
'   - Boolean: True si es válido o vacío (vacío se considera válido a
'     propósito; la obligatoriedad se controla aparte).
'   - String: mensaje explicativo cuando es False.
'  ---------------------------------------------------------------------
'  Cubre: NIF (DNI), NIE, CIF, Email, IBAN, Código Postal (España),
'  Teléfono (España).
' =====================================================================
Public Class Validador

    ' --- Tabla de letras del DNI (mod 23) ---
    Private Const LETRAS_DNI As String = "TRWAGMYFPDXBNJZSQVHLCKE"

    ' =================================================================
    '  NIF (DNI español): 8 dígitos + letra de control
    ' =================================================================
    Public Shared Function ValidarNIF(nif As String) As (Ok As Boolean, Mensaje As String)
        If String.IsNullOrWhiteSpace(nif) Then Return (True, "")
        Dim s As String = nif.Trim().ToUpperInvariant()

        If Not Regex.IsMatch(s, "^[0-9]{8}[A-Z]$") Then
            Return (False, "El NIF debe tener 8 dígitos seguidos de una letra (ej. 12345678Z).")
        End If

        Dim numero As Integer = Integer.Parse(s.Substring(0, 8))
        Dim letraEsperada As Char = LETRAS_DNI(numero Mod 23)
        If s(8) <> letraEsperada Then
            Return (False, $"La letra del NIF no es correcta. Debería ser '{letraEsperada}'.")
        End If
        Return (True, "")
    End Function

    ' =================================================================
    '  NIE (extranjeros): X|Y|Z + 7 dígitos + letra
    ' =================================================================
    Public Shared Function ValidarNIE(nie As String) As (Ok As Boolean, Mensaje As String)
        If String.IsNullOrWhiteSpace(nie) Then Return (True, "")
        Dim s As String = nie.Trim().ToUpperInvariant()

        If Not Regex.IsMatch(s, "^[XYZ][0-9]{7}[A-Z]$") Then
            Return (False, "El NIE debe empezar por X, Y o Z, seguido de 7 dígitos y una letra (ej. X1234567L).")
        End If

        ' La primera letra se sustituye por un dígito: X=0, Y=1, Z=2
        Dim sustitucion As String = ""
        Select Case s(0)
            Case "X"c : sustitucion = "0"
            Case "Y"c : sustitucion = "1"
            Case "Z"c : sustitucion = "2"
        End Select

        Dim numero As Integer = Integer.Parse(sustitucion & s.Substring(1, 7))
        Dim letraEsperada As Char = LETRAS_DNI(numero Mod 23)
        If s(8) <> letraEsperada Then
            Return (False, $"La letra del NIE no es correcta. Debería ser '{letraEsperada}'.")
        End If
        Return (True, "")
    End Function

    ' =================================================================
    '  CIF (empresas): 1 letra + 7 dígitos + dígito/letra de control
    '  Algoritmo oficial AEAT.
    ' =================================================================
    Public Shared Function ValidarCIF(cif As String) As (Ok As Boolean, Mensaje As String)
        If String.IsNullOrWhiteSpace(cif) Then Return (True, "")
        Dim s As String = cif.Trim().ToUpperInvariant()

        If Not Regex.IsMatch(s, "^[A-HJNPQRSUVW][0-9]{7}[0-9A-J]$") Then
            Return (False, "El CIF debe empezar por una letra de organización válida (A-H, J, N, P-S, U-W), 7 dígitos y un carácter de control.")
        End If

        ' Cálculo del dígito de control:
        '   - Sumar las posiciones pares (2, 4, 6) tal cual.
        '   - Multiplicar por 2 las impares (1, 3, 5, 7), sumar dígitos del resultado.
        '   - Total = sumaPares + sumaImpares
        '   - Dígito control = (10 - (Total mod 10)) mod 10
        '   - Si el último carácter es letra, debe coincidir con la posición correspondiente; si es número, debe ser igual al dígito.
        Dim sumaPares As Integer = 0
        Dim sumaImpares As Integer = 0
        For i As Integer = 1 To 7
            Dim d As Integer = Integer.Parse(s(i).ToString())
            If (i Mod 2) = 0 Then
                sumaPares += d
            Else
                Dim x As Integer = d * 2
                sumaImpares += (x \ 10) + (x Mod 10)
            End If
        Next
        Dim total As Integer = sumaPares + sumaImpares
        Dim digitoControl As Integer = (10 - (total Mod 10)) Mod 10

        Dim letraInicial As Char = s(0)
        Dim ultimo As Char = s(8)

        ' Algunas letras iniciales OBLIGAN a que el control sea letra (P, Q, R, S, W, N).
        ' Otras OBLIGAN a que sea dígito (A, B, E, H). Las demás aceptan ambos.
        Dim letrasControl As String = "JABCDEFGHI"
        Dim letraEsperada As Char = letrasControl(digitoControl)

        Dim soloLetra As String = "PQRSWN"
        Dim soloNumero As String = "ABEH"

        If soloLetra.Contains(letraInicial) Then
            If ultimo <> letraEsperada Then
                Return (False, $"El carácter de control del CIF no es correcto. Debería ser '{letraEsperada}'.")
            End If
        ElseIf soloNumero.Contains(letraInicial) Then
            If ultimo <> Char.Parse(digitoControl.ToString()) Then
                Return (False, $"El dígito de control del CIF no es correcto. Debería ser '{digitoControl}'.")
            End If
        Else
            ' Ambos válidos
            If ultimo <> letraEsperada AndAlso ultimo <> Char.Parse(digitoControl.ToString()) Then
                Return (False, $"El carácter de control del CIF no es correcto. Debería ser '{letraEsperada}' o '{digitoControl}'.")
            End If
        End If

        Return (True, "")
    End Function

    ' =================================================================
    '  Documento de identificación: prueba CIF, luego NIF, luego NIE.
    '  Útil cuando un campo acepta cualquiera de los tres (clientes
    '  particulares, autónomos, empresas).
    ' =================================================================
    Public Shared Function ValidarDocumentoIdentidad(doc As String) As (Ok As Boolean, Mensaje As String)
        If String.IsNullOrWhiteSpace(doc) Then Return (True, "")
        Dim s As String = doc.Trim().ToUpperInvariant()

        ' Determinar tipo por el primer carácter
        Dim primerChar As Char = s(0)
        If Char.IsDigit(primerChar) Then
            Return ValidarNIF(s)
        ElseIf "XYZ".Contains(primerChar) Then
            Return ValidarNIE(s)
        Else
            Return ValidarCIF(s)
        End If
    End Function

    ' =================================================================
    '  Email
    ' =================================================================
    Public Shared Function ValidarEmail(email As String) As (Ok As Boolean, Mensaje As String)
        If String.IsNullOrWhiteSpace(email) Then Return (True, "")
        Dim s As String = email.Trim()

        ' Patrón razonable. No es la RFC al 100% pero atrapa el 99% de errores reales.
        Dim patron As String = "^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$"
        If Not Regex.IsMatch(s, patron) Then
            Return (False, "El email no tiene un formato válido (ejemplo: nombre@dominio.com).")
        End If
        Return (True, "")
    End Function

    ' =================================================================
    '  IBAN: validación con MOD-97 (estándar internacional).
    '  Acepta cualquier país, no solo España.
    ' =================================================================
    Public Shared Function ValidarIBAN(iban As String) As (Ok As Boolean, Mensaje As String)
        If String.IsNullOrWhiteSpace(iban) Then Return (True, "")

        ' Eliminamos espacios y pasamos a mayúsculas
        Dim s As String = Regex.Replace(iban.ToUpperInvariant(), "\s", "")

        If s.Length < 15 OrElse s.Length > 34 Then
            Return (False, "La longitud del IBAN no es válida (entre 15 y 34 caracteres).")
        End If
        If Not Regex.IsMatch(s, "^[A-Z]{2}[0-9]{2}[A-Z0-9]+$") Then
            Return (False, "El IBAN debe empezar por dos letras (país), dos dígitos (control) y luego cuenta.")
        End If

        ' Mover los 4 primeros al final
        Dim reordenado As String = s.Substring(4) & s.Substring(0, 4)

        ' Convertir letras a números (A=10, B=11, ..., Z=35)
        Dim numerico As New System.Text.StringBuilder()
        For Each ch As Char In reordenado
            If Char.IsDigit(ch) Then
                numerico.Append(ch)
            ElseIf Char.IsLetter(ch) Then
                numerico.Append((Asc(ch) - Asc("A"c) + 10).ToString())
            Else
                Return (False, "El IBAN contiene caracteres no válidos.")
            End If
        Next

        ' Calcular MOD 97 (en trozos para no desbordar Long)
        Dim resto As Integer = 0
        For Each ch As Char In numerico.ToString()
            resto = (resto * 10 + Integer.Parse(ch.ToString())) Mod 97
        Next

        If resto <> 1 Then
            Return (False, "El IBAN no es válido (los dígitos de control no coinciden).")
        End If
        Return (True, "")
    End Function

    ' =================================================================
    '  Código postal español: 5 dígitos, primeros 2 entre 01 y 52.
    ' =================================================================
    Public Shared Function ValidarCodigoPostal(cp As String) As (Ok As Boolean, Mensaje As String)
        If String.IsNullOrWhiteSpace(cp) Then Return (True, "")
        Dim s As String = cp.Trim()

        If Not Regex.IsMatch(s, "^[0-9]{5}$") Then
            Return (False, "El código postal debe tener exactamente 5 dígitos.")
        End If
        Dim provincia As Integer = Integer.Parse(s.Substring(0, 2))
        If provincia < 1 OrElse provincia > 52 Then
            Return (False, "El código postal no corresponde a ninguna provincia española (01-52).")
        End If
        Return (True, "")
    End Function

    ' =================================================================
    '  Teléfono español: 9 dígitos, empieza por 6, 7, 8 o 9.
    '  Acepta también prefijo +34 opcional. Ignora espacios y guiones.
    ' =================================================================
    Public Shared Function ValidarTelefono(tel As String) As (Ok As Boolean, Mensaje As String)
        If String.IsNullOrWhiteSpace(tel) Then Return (True, "")

        ' Limpiar espacios, guiones y paréntesis
        Dim s As String = Regex.Replace(tel.Trim(), "[\s\-()]", "")
        ' Quitar prefijo +34 o 0034 si lo tiene
        If s.StartsWith("+34") Then s = s.Substring(3)
        If s.StartsWith("0034") Then s = s.Substring(4)

        If Not Regex.IsMatch(s, "^[6789][0-9]{8}$") Then
            Return (False, "El teléfono debe tener 9 dígitos y empezar por 6, 7, 8 o 9 (móvil o fijo en España).")
        End If
        Return (True, "")
    End Function

    ' =================================================================
    '  HELPER DE FORMULARIO
    '  Recibe una lista de validaciones (con etiqueta) y, si alguna falla,
    '  muestra un único MessageBox listando todos los problemas y
    '  preguntando al usuario si desea guardar de todos modos.
    '
    '  Esto permite que las validaciones sean estrictas (avisan al usuario
    '  cuando algo no cuadra) sin ser bloqueantes (no impiden guardar
    '  datos existentes que ya estaban mal en BD antes de aplicar el
    '  parche). Para datos NUEVOS el usuario verá el aviso y normalmente
    '  los corregirá; para datos VIEJOS puede confirmar y continuar.
    '
    '  Devuelve True si se puede continuar guardando, False si no.
    ' =================================================================
    Public Shared Function ConfirmarSiHayProblemas(ParamArray validaciones As (Etiqueta As String, Resultado As (Ok As Boolean, Mensaje As String))()) As Boolean
        Dim problemas As New List(Of String)
        For Each v In validaciones
            If Not v.Resultado.Ok Then
                problemas.Add($"• {v.Etiqueta}: {v.Resultado.Mensaje}")
            End If
        Next
        If problemas.Count = 0 Then Return True

        Dim msg As String = "Se han detectado los siguientes problemas en los datos introducidos:" & vbCrLf & vbCrLf &
                            String.Join(vbCrLf, problemas) & vbCrLf & vbCrLf &
                            "¿Quieres guardar de todos modos?"

        Dim resp = MessageBox.Show(msg, "Datos con avisos", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)
        Return (resp = DialogResult.Yes)
    End Function

End Class
