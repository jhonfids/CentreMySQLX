Imports System.Net
Imports System.Globalization

Namespace Herramientas

    Public NotInheritable Class ValidacionFormato
        Const Clase As String = "Validación formato"

        Public Shared Function FormatoIP(ByVal IP As String) As Boolean
            Const fn As String = "FormatoIP"

            'Determina si tiene el formato de IP
            Try
                Dim valor As New IPAddress(New Byte() {0, 0, 0, 0})
                Return IPAddress.TryParse(IP, valor)

            Catch ex As Exception
                Throw New ExcepcionInfo("El formato de IP puede tener inconsistencias de contenido", Clase, fn, ex)
            End Try

        End Function

        Public Shared Sub SupresionCaracteresReservados(ByRef Valor As String)
            Valor = Valor.Replace("'", "´")

        End Sub

        Public Shared Sub FormatoNumerico(ByVal Valor As String)
            Const Funcion As String = "Numerico"
            If Valor.Length = 0 Then Exit Sub

            If IsNumeric(Valor) = False Then
                Throw New ExcepcionInfo("Este item debe ser solo numérico", Clase, Funcion)
            End If
        End Sub

    End Class

    Public NotInheritable Class ValidacionConexion
        Public Shared Function PingStatus(ByVal IP As String) As Boolean
            Try
                My.Computer.Network.Ping(IP, 500)
            Catch ex As Exception
                Return False
            End Try
            Return True
        End Function

    End Class

    Public NotInheritable Class Conversiones
        Const Clase As String = "Conversiones"
        Public Shared Function Fecha(ByVal InputDB As Object) As Date
            Const fn As String = "Fecha_DBinput"

            Dim f As Date
            Try
                If IsDBNull(InputDB) Then
                    Throw New ExcepcionInfo("Información nula", Clase, fn, 0)

                Else
                    f = CDate(CStr(InputDB))

                End If

            Catch ex As Exception
                Throw New ExcepcionInfo("Falla en conversión", Clase, fn, ex)
            End Try

            Return f

        End Function

        Public Shared Function Fecha(ByVal InputValue As Date) As String
            Const fn As String = "Fecha_ProgramInput"
            If IsDBNull(InputValue) Then
                Dim f As Date = New Date(1900, 1, 1)

            End If

            Try
                Return Format(InputValue, "yyyy-MM-dd HH:mm:ss")
            Catch ex As Exception
                Throw New ExcepcionInfo("Falla en conversión", Clase, fn, ex)
            End Try

        End Function

        Public Shared Function Booleano(ByVal ValorDB As Object) As Boolean
            Const fn As String = "Boolean_DBinput"

            If IsDBNull(ValorDB) Then
                Throw New ExcepcionInfo("Información nula", Clase, fn, 0)
            End If

            If CInt(ValorDB) = 1 Then
                Return True

            ElseIf CInt(ValorDB) = 0 Then
                Return False

            Else
                Throw New ExcepcionInfo("Valor de entrada no esperado", Clase, fn)
            End If

        End Function

        Public Shared Function Booleano(ByVal InputValue As Boolean) As Integer
            'Const fn As String = "Boolean_ProgramInput"

            If InputValue = True Then
                Return 1
            Else
                Return 0
            End If

        End Function

        Public Shared Function NumeracionDecimal(ByVal ValorDB As Object) As Decimal
            Const fn As String = "NumeraciónDecimal_DBinupt"

            Dim culture As New CultureInfo("es-VE")

            If IsDBNull(ValorDB) Then
                Throw New ExcepcionInfo("Información nula", Clase, fn, 0)
            End If

            Try
                Dim dec = CStr(ValorDB)
                Return Decimal.Parse(dec, culture)

            Catch ex As Exception
                Throw New ExcepcionInfo("Error de conversión", Clase, fn, ex)

            End Try

        End Function

        Public Shared Function NumeracionDecimal(ByVal ValorProgram As Decimal) As String
            Const fn As String = "NumeraciónDecimal_ProgramInput"

            Dim nfi As NumberFormatInfo = New CultureInfo("es-VE", True).NumberFormat
            nfi.NegativeSign = "-"
            nfi.CurrencyDecimalDigits = 2
            nfi.NumberDecimalDigits = 3

            Try
                Return ValorProgram.ToString(nfi)

            Catch ex As Exception
                Throw New ExcepcionInfo("Error de conversión", Clase, fn, ex)

            End Try

        End Function

    End Class

    Public NotInheritable Class LibreriaComandos

        Public Shared Function getfechaHoraString() As String
            Return "SELECT NOW()"
        End Function

    End Class

End Namespace

