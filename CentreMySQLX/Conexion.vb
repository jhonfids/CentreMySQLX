Imports CentreMySQLX.Configuracion
Imports CentreMySQLX.Herramientas
Imports MySql.Data.MySqlClient

Namespace Conexion

    Module Resources

        Public Function GetStringConnection(Data As DataConexion) As String
            Dim ln As String = "SERVER=" & Data.IP & ";" &
                                         "PORT=" & Data.Puerto & ";" &
                                         "DATABASE=" & Data.NombreDatabase & ";" &
                                         "USER=" & Data.Usuario & ";" &
                                         "PASSWORD=" & Data.Contrasena & ""

            Return ln
        End Function


    End Module


    Public NotInheritable Class Servicios
        Const Clase As String = "Servicios"

        Public Enum TipoInstruccion
            ExecuteScalar
            NonQuery
        End Enum

        Public Enum TipoFuncion
            Consulta = 1
            Instruccion = 2
        End Enum

        Public Enum TipoResultado
            Incorrecto
            Correcto
        End Enum

        Public Structure RegistroData
            Public IP As String
            Public Puerto As String
            Public DatabaseNombre As String
            Public Tipo As TipoFuncion
            Public Sintaxis As String
            Public Resultado As TipoResultado

        End Structure

        'Componentes base
        Private Shared mysqlConnect As MySqlConnection

        'Componentes servicio
        Private mysqlAdaptador As MySqlDataAdapter
        Private mysqlComando As MySqlCommand

        Private _dataconexion As DataConexion
        Private _testPing As Boolean

        'Registro
        Public Shared Transacciones As List(Of RegistroData)

        Sub New(ByVal DataConexion As DataConexion,
                Optional ByVal enable_PingTest As Boolean = False)
            _dataconexion = DataConexion
            _testPing = enable_PingTest
        End Sub

        Public Function TestConexion() As Boolean
            Const fn As String = "Prueba de conexión"

            Dim machConn As New Machine
            Try
                machConn.Conectar(_dataconexion, _testPing)
                machConn.Desconectar()

            Catch ex As Exception
                Throw New ExcepcionInfo("Falla en etapa de conexión", Clase, fn, ex)
            End Try

            Return True

        End Function

        Public Function Consulta(ByVal InstruccionesMySQL As List(Of String)) As DataTable()
            Const fn As String = "ConsultaMultiple"

            Dim dts() As DataTable = Nothing
            Dim indice As Integer
            Try
                For index As Integer = 0 To InstruccionesMySQL.Count - 1
                    indice = index

                    ReDim Preserve dts(index)

                    Dim syn As String = InstruccionesMySQL.Item(index)
                    dts(index) = Consulta(syn, False)

                Next

            Catch ex As Exception
                Throw New ExcepcionInfo("Proceso por lotes fallido en índice '" & indice & "'", Clase, fn, ex)
            End Try

            'Cerrar la conexion
            If mysqlConnect.State = ConnectionState.Open Then
                mysqlConnect.Close()
            End If

            Return dts

        End Function

        Public Function Instruccion(ByVal InstruccionesMySQL As List(Of String)) As String()
            Const fn As String = "InstrucciónMultiple"

            Dim strs() As String = Nothing
            Dim indice As Integer
            Try
                For index As Integer = 0 To InstruccionesMySQL.Count - 1
                    indice = index

                    ReDim Preserve strs(index)

                    Dim sql As String = InstruccionesMySQL.Item(index)
                    strs(index) = Instruccion(sql, False)

                Next

            Catch ex As Exception
                Throw New ExcepcionInfo("Proceso por lotes fallido en índice '" & indice & "'", Clase, fn, ex)
            End Try

            'Cerrar la conexion
            If mysqlConnect.State = ConnectionState.Open Then
                mysqlConnect.Close()
            End If

            Return strs

        End Function

        Public Function Consulta(ByVal InstruccionMySQL As String,
                                 Optional ByVal CerrarConexion As Boolean = True) As DataTable
            Const fn As String = "Consulta"
            Const ty As TipoFuncion = TipoFuncion.Consulta

            Dim dt As New DataTable

            'Abrir la conexión
            Dim machConn As New Machine
            Try
                machConn.Conectar(_dataconexion, _testPing)

            Catch ex As Exception
                Throw New ExcepcionInfo("Falla en etapa de conexión", Clase, fn, ex)
            End Try

            'Ejecutar
            Try
                mysqlAdaptador = New MySqlDataAdapter(InstruccionMySQL, mysqlConnect)
                mysqlAdaptador.Fill(dt)

            Catch ex As Exception
                RegistroTransaccion(ty, InstruccionMySQL, TipoResultado.Incorrecto)
                Throw New ExcepcionInfo("Falla en etapa de ejecución", Clase, fn, ex)

            End Try

            RegistroTransaccion(ty, InstruccionMySQL, TipoResultado.Correcto)

            'Desconexión
            If CerrarConexion = True Then
                Try
                    machConn.Desconectar()
                Catch ex As Exception
                    Throw New ExcepcionInfo("Falla en etapa de post conexión", Clase, fn, ex)
                End Try
            End If

            Return dt

        End Function

        Public Sub InstruccionConTransaccion(InstruccionesSQL As List(Of String))
            Const fn As String = "InstrucciónWithTransaccion"
            Const ty As TipoFuncion = TipoFuncion.Instruccion

            'Este nuevo metodo usa bloque transcaccion con commit and rollback
            Using connection As New MySqlConnection(GetStringConnection(_dataconexion))
                connection.Open()

                Dim command As MySqlCommand = connection.CreateCommand()
                Dim transaction As MySqlTransaction

                'transaction = connection.BeginTransaction("Instruccion")
                transaction = connection.BeginTransaction()
                command.Connection = connection
                command.Transaction = transaction

                Dim lntr As String = String.Empty
                Try
                    'Ejecutando cada instruccion y almacenando el resultado para devolver
                    For Each ln In InstruccionesSQL
                        command.CommandText = ln
                        command.ExecuteNonQuery()
                        lntr &= ln & vbNewLine

                    Next

                    'Realizando un commit de todas las intrucciones y cerrando conexion
                    transaction.Commit()
                    connection.Close()

                    RegistroTransaccion(ty, lntr, TipoResultado.Correcto)

                Catch ex As Exception
                    Dim err As String = "Falla durante el proceso de de ejecucion"

                    'Realizando rollback para deshacer cambios
                    Try
                        transaction.Rollback()

                    Catch ex1 As Exception
                        RegistroTransaccion(ty, lntr, TipoResultado.Incorrecto)
                        Throw New ExcepcionInfo(err & " y proceso rollback", Clase, fn, ex)

                    End Try

                    RegistroTransaccion(ty, lntr, TipoResultado.Incorrecto)
                    Throw New ExcepcionInfo(err & " y proceso rollback", Clase, fn, ex)

                End Try

            End Using

        End Sub

        Public Function Instruccion(ByVal InstruccionMySQL As String,
                                    Optional ByVal Tipo As TipoInstruccion = TipoInstruccion.NonQuery,
                                    Optional ByVal CerrarConexion As Boolean = True) As String
            Const fn As String = "Instrucción"
            Const ty As TipoFuncion = TipoFuncion.Instruccion

            'Abrir la conexión
            Dim machConn As New Machine
            Try
                machConn.Conectar(_dataconexion, _testPing)

            Catch ex As Exception
                Throw New ExcepcionInfo("Falla en etapa de conexión", Clase, fn, ex)
            End Try

            'Ejecutar
            Dim str As String
            Try
                mysqlComando = New MySqlCommand(InstruccionMySQL, mysqlConnect)
                If Tipo = TipoInstruccion.NonQuery Then
                    str = CStr(mysqlComando.ExecuteNonQuery())
                Else
                    str = CStr(mysqlComando.ExecuteScalar())
                End If

            Catch ex As Exception
                RegistroTransaccion(ty, InstruccionMySQL, TipoResultado.Incorrecto)
                Throw New ExcepcionInfo("Falla en etapa de ejecución", Clase, fn, ex)

            End Try

            RegistroTransaccion(ty, InstruccionMySQL, TipoResultado.Correcto)

            'Desconexión
            If CerrarConexion = True Then
                Try
                    machConn.Desconectar()
                Catch ex As Exception
                    Throw New ExcepcionInfo("Falla en etapa de post conexión", Clase, fn, ex)
                End Try
            End If

            Return str

        End Function


        Private Sub RegistroTransaccion(ByVal Funcion As TipoFuncion,
                                        ByVal InstruccionMySQL As String,
                                        ByVal Resultado As TipoResultado)

            If IsNothing(Transacciones) Then
                Transacciones = New List(Of RegistroData)
            End If

            Transacciones.Add(New RegistroData With {.IP = _dataconexion.IP,
                                                      .Puerto = _dataconexion.Puerto,
                                                      .DatabaseNombre = _dataconexion.NombreDatabase,
                                                      .Tipo = Funcion,
                                                      .Sintaxis = InstruccionMySQL,
                                                      .Resultado = Resultado})

        End Sub


        Friend Class Machine

            Public Sub Conectar(ByVal Data As DataConexion,
                                Optional ByVal enable_TestPing As Boolean = True)
                Const fn As String = "Conectar"

                If Not IsNothing(mysqlConnect) Then
                    If mysqlConnect.State = ConnectionState.Open Then
                        Exit Sub
                    End If
                End If

                Dim sql As String
                Try
                    sql = getStringConn(Data)
                Catch ex As Exception
                    Throw New ExcepcionInfo("Adquisición de cadena de conexión incorrecta", Clase, fn)
                End Try

                'Verificar la conexión con ping
                If enable_TestPing = True Then
                    If Not ValidacionConexion.PingStatus(Data.IP) Then
                        Throw New ExcepcionInfo("No superó la prueba de ping al destino de servicio", Clase, fn)
                    End If
                End If

                Try
                    mysqlConnect = New MySqlConnection(sql)
                    mysqlConnect.Open()

                Catch ex As Exception
                    Throw New ExcepcionInfo("El procesador de conexión reportó un fallo", Clase, fn, ex)
                End Try

            End Sub

            Public Sub Desconectar()
                Const fn As String = "Desconectar"

                Try
                    If Not IsNothing(mysqlConnect) Then
                        mysqlConnect.Close()
                    End If
                Catch ex As Exception
                    Throw New ExcepcionInfo("El procesador de conexión reportó un fallo", Clase, fn, ex)
                End Try

            End Sub


            Private Function getStringConn(ByVal DataConn As DataConexion)

                Dim ln As String = String.Empty
                ln = "SERVER=" & DataConn.IP & ";" &
                     "PORT=" & DataConn.Puerto & ";" &
                     "DATABASE=" & DataConn.NombreDatabase & ";" &
                     "USER=" & DataConn.Usuario & ";" &
                     "PASSWORD=" & DataConn.Contrasena & ""

                Return ln

            End Function

        End Class

    End Class

End Namespace

