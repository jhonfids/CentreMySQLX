Imports CentreMySQLX.Herramientas
Namespace Configuracion

    Public Structure DataConexion
        Private Const Clase As String = "DataConexión"
        Private _ip As String
        Private _puerto As String
        Private _nombreDatabase As String
        Private _usuario As String
        Private _contrasena As String

        Public Property IP As String
            Get
                Return _ip

            End Get
            Set(value As String)
                Const val As String = "IP v4"

                Try
                    If ValidacionFormato.FormatoIP(value) = False Then
                        Throw New ExcepcionPropiedades(val, "El formato no es válido", Clase)
                        Exit Try
                    End If

                Catch ex As Exception
                    Throw New ExcepcionPropiedades(val, "El formato es inconsistente", Clase, ex)
                End Try

                _ip = value

            End Set
        End Property

        Public Property Puerto As String
            Get
                Return _puerto
            End Get
            Set(value As String)
                Const val As String = "Puerto"

                If value.Length < 2 And value.Length > 6 Then
                    Throw New ExcepcionPropiedades(val, "La longitud de este valor debe estar entre 2 y 6 caracteres", Clase)
                End If

                Try
                    ValidacionFormato.FormatoNumerico(value)

                Catch ex As Exception
                    Throw New ExcepcionPropiedades(val, "Falla de validación", Clase, ex)
                End Try

                _puerto = value

            End Set
        End Property

        Public Property NombreDatabase As String
            Get
                Return _nombreDatabase
            End Get
            Set(value As String)
                Const val As String = "Database"

                If value.Length < 2 And value.Length > 50 Then
                    Throw New ExcepcionPropiedades(val, "La longitud de este valor debe estar entre 2 y 50 caracteres", Clase)
                End If

                ValidacionFormato.SupresionCaracteresReservados(value)

                _nombreDatabase = value

            End Set
        End Property

        Public Property Usuario As String
            Get
                Return _usuario
            End Get
            Set(value As String)
                Const val As String = "Usuario"

                If value.Length < 2 And value.Length > 50 Then
                    Throw New ExcepcionPropiedades(val, "La longitud de este valor debe estar entre 2 y 50 caracteres", Clase)
                End If

                ValidacionFormato.SupresionCaracteresReservados(value)

                _usuario = value
            End Set
        End Property

        Public Property Contrasena As String
            Get
                Return _contrasena
            End Get
            Set(value As String)
                Const val As String = "Usuario"

                If value.Length < 8 And value.Length > 100 Then
                    Throw New ExcepcionPropiedades(val, "La longitud de este valor debe estar entre 8 y 100 caracteres", Clase)
                End If

                _contrasena = value
            End Set
        End Property

    End Structure

    Public NotInheritable Class PerfilConexion
        Public Shared DataPerfil As New List(Of DataConexion)

    End Class

End Namespace


