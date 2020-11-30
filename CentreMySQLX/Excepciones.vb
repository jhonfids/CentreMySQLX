Public Class ExcepcionInfo
    Inherits Exception

    Public Sub New(ByVal MensajeError As String, ByVal Clase As String, ByVal Funcion As String)
        MyBase.New(MensajeError)

        HelpLink = Empresa_Email
        Source = Clase & "." & Funcion
        HResult = -10
    End Sub

    Public Sub New(ByVal MensajeError As String, ByVal Clase As String, ByVal Funcion As String, ByVal CodigoErrorInterno As Integer)
        MyBase.New(MensajeError)

        HelpLink = Empresa_Email
        Source = Clase & "." & Funcion
        HResult = CodigoErrorInterno
    End Sub

    Public Sub New(ByVal MensajeError As String, ByVal Clase As String, ByVal Funcion As String, ByVal InnerData As Exception)
        MyBase.New(MensajeError & ": (" & InnerData.Message & ")")

        HelpLink = Empresa_Email
        Source = Clase & "." & Funcion & "(" & InnerData.Source & ")"
        HResult = InnerData.HResult
    End Sub

End Class

Public Class ExcepcionPropiedades
    Inherits Exception
    Public Sub New(ByVal Item As String, ByVal MensajeError As String, ByVal Clase As String)
        MyBase.New("'" & Item & "': " & MensajeError)

        HelpLink = empresa_email
        Source = Clase
        HResult = -10
    End Sub

    Public Sub New(ByVal Item As String, ByVal MensajeError As String, ByVal Clase As String, ByVal InnerData As Exception)
        MyBase.New("'" & Item & "': " & MensajeError & " " & InnerData.Message)

        HelpLink = empresa_email
        Source = Clase
        HResult = InnerData.HResult
    End Sub

End Class
