Imports System.Data.SqlClient
Public Class clsConfiguracion

    'SQL
    Public Property sqlConexion As New SqlConnection(My.Settings.ConexionIntermedias)
    Public Property sqlComando As SqlCommand = New SqlCommand
    Public Property sqlAdaptador As SqlDataAdapter = New SqlDataAdapter

    'Propiedades de correo
    Public Property EnviarNotificaciones As Boolean
    Public Property ServidorDeCorreo As String
    Public Property Puerto As String
    Public Property RequiereAutenticacion As Boolean
    Public Property SSL As Boolean
    Public Property CorreoRemitente As String
    Public Property UsuarioMail As String
    Public Property ClaveMail As String
    Public Property CorreosNotificaciones As String
    Public Property CorreoNotificacionesError As String

    Public Property AdjuntarArchivoCorreo As String


    'Propiedades Multiproceso Hijo
    Public Property ProcesosParalelos As Integer
    Public Property NumFilasMultiProcesos As Integer
    Public Property RutaLog As String
    Public Property RutaPlanos As String

    'Propiedades Conexion
    Public Property ConexionWsSiesa As String
    Public Property CompaniaUnoEE As String
    Public Property UsuarioUnoEE As String
    Public Property ClaveUnoEE As String


    Public Sub New()

        Dim ds As New DataSet

        Try
            sqlComando.Connection = sqlConexion
            sqlComando.CommandType = CommandType.StoredProcedure
            sqlComando.CommandText = "sp_Propiedades_Select"
            sqlAdaptador.SelectCommand = sqlComando
            sqlAdaptador.Fill(ds)


            For Each Parametro As DataRow In ds.Tables(0).Rows
                If Parametro.Item("nombrePropiedad").ToString = "ProcesosParalelos" Then
                    ProcesosParalelos = Parametro.Item("valorEntero")
                ElseIf Parametro.Item("nombrePropiedad").ToString = "NumFilasMultiProcesos" Then
                    NumFilasMultiProcesos = Parametro.Item("valorEntero")
                ElseIf Parametro.Item("nombrePropiedad").ToString = "RutaLog" Then
                    RutaLog = Parametro.Item("valorTexto1")
                ElseIf Parametro.Item("nombrePropiedad").ToString = "RutaPlanos" Then
                    RutaPlanos = Parametro.Item("valorTexto1")
                ElseIf Parametro.Item("nombrePropiedad").ToString = "EnviarNotificaciones" Then
                    EnviarNotificaciones = Parametro.Item("valorBooleano")
                ElseIf Parametro.Item("nombrePropiedad").ToString = "ServidorDeCorreo" Then
                    ServidorDeCorreo = Parametro.Item("valorTexto1").ToString
                ElseIf Parametro.Item("nombrePropiedad").ToString = "Puerto" Then
                    Puerto = Parametro.Item("valorTexto1").ToString
                ElseIf Parametro.Item("nombrePropiedad").ToString = "RequiereAutenticacion" Then
                    RequiereAutenticacion = Parametro.Item("valorBooleano")
                ElseIf Parametro.Item("nombrePropiedad").ToString = "SSL" Then
                    SSL = Parametro.Item("valorBooleano")
                ElseIf Parametro.Item("nombrePropiedad").ToString = "CorreoRemitente" Then
                    CorreoRemitente = Parametro.Item("valorTexto1").ToString
                ElseIf Parametro.Item("nombrePropiedad").ToString = "UsuarioMail" Then
                    UsuarioMail = Parametro.Item("valorTexto1").ToString
                ElseIf Parametro.Item("nombrePropiedad").ToString = "ClaveMail" Then
                    ClaveMail = Parametro.Item("valorTexto1").ToString
                ElseIf Parametro.Item("nombrePropiedad").ToString = "CorreosNotificaciones" Then
                    CorreosNotificaciones = Parametro.Item("valorTexto1").ToString
                ElseIf Parametro.Item("nombrePropiedad").ToString = "AdjuntarArchivoCorreo" Then
                    AdjuntarArchivoCorreo = Parametro.Item("valorTexto1").ToString
                ElseIf Parametro.Item("nombrePropiedad").ToString = "ConexionWsSiesa" Then
                    ConexionWsSiesa = Parametro.Item("valorTexto1").ToString
                ElseIf Parametro.Item("nombrePropiedad").ToString = "CompaniaUnoEE" Then
                    CompaniaUnoEE = Parametro.Item("valorTexto1").ToString
                ElseIf Parametro.Item("nombrePropiedad").ToString = "UsuarioUnoEE" Then
                    UsuarioUnoEE = Parametro.Item("valorTexto1").ToString
                ElseIf Parametro.Item("nombrePropiedad").ToString = "ClaveUnoEE" Then
                    ClaveUnoEE = Parametro.Item("valorTexto1").ToString
                ElseIf Parametro.Item("nombrePropiedad").ToString = "CorreoNotificacionesError" Then
                    CorreoNotificacionesError = Parametro.Item("valorTexto1").ToString
                End If
            Next

        Catch ex As Exception
            MessageBox.Show("Error: " + ex.Message.ToString)
        Finally
            sqlComando.Parameters.Clear()
            sqlComando.Connection.Close()
        End Try

    End Sub

End Class