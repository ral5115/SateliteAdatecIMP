Imports System.Data.SqlClient

Public Class Form1

    Dim objConfiguraciones As New clsConfiguracion
    Dim objCorreo As New clsCorreo
    Dim AdatecDat As New AdatecDat
    Dim clsAdatectPed As New AdactecPedidos
    Dim clsAdatectRec As New AdactecRecibos
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try
            clsAdatectPed.AdatectPedidos()
            fnc_IntegrarPedido()
            'clsAdatectRec.AdactecRecibos()
            'fnc_IntegrarRecibo()
            Me.Close()
        Catch ex As Exception
            objCorreo.EnviarCorreoTareaExe("DisLaCosta: Error al cargar funciones de importación: ", ex.Message)
        End Try

    End Sub

    Private Function fnc_IntegrarRecibo(Optional XML As String = "") As String
        Try
            Dim dsDatos As New DataSet
            Dim strResultado As String
            Dim objGenericTransfer As New wsGenericTransfer.wsGenerarPlano
            Dim RutaPlano As String = objConfiguraciones.RutaPlanos
            Dim FechaLog = Date.Now.ToString("dd-MM-yyyy HH:mm:ss")

            MsgBox("aca Estoy consultando ")
            AdatecDat.ConsultarRecibo(dsDatos)
            MsgBox("ya consulte")
            objGenericTransfer.Timeout = 1800000000

            'Valida que el dataset este cargado 
            If dsDatos.Tables.Count >= 1 Then
                If dsDatos.Tables(0).Rows.Count >= 1 Then


                    MsgBox("Entrar a consumir")
                    strResultado = objGenericTransfer.ImportarDatosDS(72209, "RECIBO_CAJA", 2, 1, "gt", "gt", dsDatos, RutaPlano)
                    MsgBox("ya consumi" + strResultado)
                    'mensaje de error
                    'MsgBox(strResultado)


                    objGenericTransfer.Timeout = 1800000000

                If strResultado.Contains("Importacion exitosa") Then

                        '------------------------
                        Dim sqlComando As SqlCommand = New SqlCommand
                        Dim SqlCon = New SqlConnection("Data Source=192.168.0.241;Initial Catalog=GTIntegration;User ID=sa;Password=Decepticon2014;Integrated Security=False")
                        'Dim SqlCon = New SqlConnection("Data Source=siesa.generictransfer.com,1434;Initial Catalog=GTintegrationLACOSTA;User ID=admincali;Password=4217;Integrated Security=False")

                        sqlComando.CommandTimeout = 18000000
                        sqlComando.Connection = SqlCon
                        sqlComando.CommandType = CommandType.StoredProcedure
                        sqlComando.CommandText = "sp_CambioEstadoRecibo"
                        sqlComando.Parameters.AddWithValue("@estado", "2")

                        '------------------------
                        Try
                            SqlCon.Open()
                            sqlComando.ExecuteNonQuery()
                        Catch ex As Exception
                        Throw ex
                    Finally
                            SqlCon.Close()
                        End Try

                    AlmacenarLog(FechaLog & " Resultado: " & strResultado)

                End If
            End If




            ElseIf strResultado.Contains("consumir el Web Service") Then

            Dim sqlComando As SqlCommand = New SqlCommand
                Dim SqlCon = New SqlConnection("Data Source=192.168.0.241;Initial Catalog=GTintegration;User ID=sa;Password=Decepticon2014;Integrated Security=false")
                'Dim SqlCon = New SqlConnection("Data Source=siesa.generictransfer.com,1434;Initial Catalog=GTintegrationLACOSTA;User ID=admincali;Password=4217;Integrated Security=False")


                sqlComando.CommandTimeout = 18000000
            sqlComando.Connection = SqlCon
            sqlComando.CommandType = CommandType.StoredProcedure
            sqlComando.CommandText = "sp_CambioEstado"
            sqlComando.Parameters.AddWithValue("@estado", "0")

            Try
                SqlCon.Open()
                sqlComando.ExecuteNonQuery()
            Catch ex As Exception
                Throw ex
            Finally
                SqlCon.Close()
            End Try

            AlmacenarLog(FechaLog & " Resultado: " & strResultado)
            ElseIf strResultado = "99" Then
            Dim sqlComando As SqlCommand = New SqlCommand
                Dim SqlCon = New SqlConnection("Data Source=192.168.0.241;Initial Catalog=GTintegration;User ID=sa;Password=Decepticon2014;Integrated Security=false")
                'Dim SqlCon = New SqlConnection("Data Source=siesa.generictransfer.com,1434;Initial Catalog=GTintegrationLACOSTA;User ID=admincali;Password=4217;Integrated Security=False")


                sqlComando.CommandTimeout = 18000000
            sqlComando.Connection = SqlCon
            sqlComando.CommandType = CommandType.StoredProcedure
            sqlComando.CommandText = "sp_CambioEstado"
            sqlComando.Parameters.AddWithValue("@estado", "0")

            Try
                SqlCon.Open()
                sqlComando.ExecuteNonQuery()
            Catch ex As Exception
                Throw ex
            Finally
                SqlCon.Close()
            End Try
            AlmacenarLog(FechaLog & " Resultado: " & "El Web Service de SIESA no responde, por favor validar los permisos de lectura y escritura del usuario generic. En caso de contar con ambos permisos validar si la dirección del web service ha cambiado, para ello solicitar revisión por parte de soporte SIESA ERP.")
                objCorreo.EnviarCorreoTarea("Novedades de importación ", "El Web Service de SIESA no responde, por favor validar los permisos de lectura y escritura del usuario generic. En caso de contar con ambos permisos validar si la dirección del web service ha cambiado, para ello solicitar revisión por parte de soporte SIESA ERP.")

            Else
                objCorreo.EnviarCorreoTarea("Novedades de importación", "Error al cargar a SIESA. Revisar Log de Importaciones.")
                AlmacenarLog(FechaLog & " Resultado: " & strResultado.ToString)
            End If
            Return strResultado

        Catch ex As Exception
            objCorreo.EnviarCorreoTareaExe("DisLaCosta: Error al realizar la integración de datos: ", ex.ToString)
        Finally
            Me.Close()
        End Try
    End Function


    Public Function fnc_IntegrarPedido(Optional XML As String = "") As String
        Try
            Dim count As Integer
            Dim dsDatos As New DataSet
            Dim strResultado As String
            Dim objGenericTransfer As New wsGenericTransfer.wsGenerarPlano
            Dim RutaPlano As String = objConfiguraciones.RutaPlanos
            Dim FechaLog = Date.Now.ToString("dd-MM-yyyy HH:mm:ss")
            MsgBox("aca Estoy consultando ")
            AdatecDat.ConsultarPedido(dsDatos)
            MsgBox("aca ya consulte ")
            objGenericTransfer.Timeout = 1800000000
            MsgBox("Entre a consumir ")

            'Valida que el dataset este cargado 
            If dsDatos.Tables.Count >= 1 Then
                If dsDatos.Tables(0).Rows.Count >= 1 Then
                    'strResultado = objGenericTransfer.ImportarDatosDS(70675, "Pedidos", 2, 1, "gt", "gt", dsDatos, RutaPlano)

                    strResultado = objGenericTransfer.ImportarDatosDS(76880, "PEDIDOS_DESCUENTOS", 2, 1, "gt", "gt", dsDatos, RutaPlano)
                    MsgBox("Sali de  consumir /n" + strResultado)
                    objGenericTransfer.Timeout = 1800000000
                    If strResultado.Contains("Importacion exitosa") Then
                        'Inicio de llamado de SP para cambio de estado YSK
                        'For Each campo As DataRow In dsDatos.Tables(0).Rows()
                        'Next
                        Dim sqlComando As SqlCommand = New SqlCommand
                        Dim SqlCon = New SqlConnection("Data Source=192.168.0.241;Initial Catalog=GTintegration;User ID=sa;Password=Decepticon2014;Integrated Security=false")

                        sqlComando.CommandTimeout = 18000000
                        sqlComando.Connection = SqlCon
                        sqlComando.CommandType = CommandType.StoredProcedure
                        sqlComando.CommandText = "sp_CambioEstado"
                        sqlComando.Parameters.AddWithValue("@estado", "2")

                        Try
                            SqlCon.Open()
                            sqlComando.ExecuteNonQuery()
                        Catch ex As Exception
                            Throw ex
                        Finally
                            SqlCon.Close()
                        End Try

                        AlmacenarLog(FechaLog & " Resultado: " & strResultado)

                    End If
                End If


            ElseIf strResultado.Contains("consumir el Web Service") Then

                Dim sqlComando As SqlCommand = New SqlCommand
                Dim SqlCon = New SqlConnection("Data Source=192.168.0.241;Initial Catalog=GTintegration;User ID=sa;Password=Decepticon2014;Integrated Security=false")

                sqlComando.CommandTimeout = 18000000
                sqlComando.Connection = SqlCon
                sqlComando.CommandType = CommandType.StoredProcedure
                sqlComando.CommandText = "sp_CambioEstado"
                sqlComando.Parameters.AddWithValue("@estado", "0")

                Try
                    SqlCon.Open()
                    sqlComando.ExecuteNonQuery()
                Catch ex As Exception
                    Throw ex
                Finally
                    SqlCon.Close()
                End Try

                AlmacenarLog(FechaLog & " Resultado: " & strResultado)
            ElseIf strResultado = "99" Then
                Dim sqlComando As SqlCommand = New SqlCommand
                Dim SqlCon = New SqlConnection("Data Source=192.168.0.241;Initial Catalog=GTintegration;User ID=sa;Password=Decepticon2014;Integrated Security=false")

                sqlComando.CommandTimeout = 18000000
                sqlComando.Connection = SqlCon
                sqlComando.CommandType = CommandType.StoredProcedure
                sqlComando.CommandText = "sp_CambioEstado"
                sqlComando.Parameters.AddWithValue("@estado", "0")

                Try
                    SqlCon.Open()
                    sqlComando.ExecuteNonQuery()
                Catch ex As Exception
                    Throw ex
                Finally
                    SqlCon.Close()
                End Try
                AlmacenarLog(FechaLog & " Resultado: " & "El Web Service de SIESA no responde, por favor validar los permisos de lectura y escritura del usuario generic. En caso de contar con ambos permisos validar si la dirección del web service ha cambiado, para ello solicitar revisión por parte de soporte SIESA ERP.")
                objCorreo.EnviarCorreoTarea("Novedades de importación ", "El Web Service de SIESA no responde, por favor validar los permisos de lectura y escritura del usuario generic. En caso de contar con ambos permisos validar si la dirección del web service ha cambiado, para ello solicitar revisión por parte de soporte SIESA ERP.")

            Else
                Dim sqlComando As SqlCommand = New SqlCommand
                Dim SqlCon = New SqlConnection("Data Source=192.168.0.241;Initial Catalog=GTintegration;User ID=sa;Password=Decepticon2014;Integrated Security=false")

                sqlComando.CommandTimeout = 18000000
                sqlComando.Connection = SqlCon
                sqlComando.CommandType = CommandType.StoredProcedure
                sqlComando.CommandText = "sp_CambioEstado"
                sqlComando.Parameters.AddWithValue("@estado", "0")

                Try
                    SqlCon.Open()
                    sqlComando.ExecuteNonQuery()
                Catch ex As Exception
                    Throw ex
                Finally
                    SqlCon.Close()
                End Try
                objCorreo.EnviarCorreoTarea("Novedades de importación", "Error al cargar a SIESA. Revisar Log de Importaciones.")
                AlmacenarLog(FechaLog & " Resultado: " & strResultado.ToString)
            End If
            Return strResultado

        Catch ex As Exception
            objCorreo.EnviarCorreoTareaExe("DisLaCosta: Error al realizar la integración de datos: ", ex.ToString)
        Finally
            Me.Close()
        End Try
    End Function

    Public Sub AlmacenarLog(ByVal Mensaje As String)

        Dim FileLogImpXml As System.IO.StreamWriter
        FileLogImpXml = My.Computer.FileSystem.OpenTextFileWriter(objConfiguraciones.RutaLog & "Log_Importaciones.txt", True)
        FileLogImpXml.WriteLine(Mensaje)
        FileLogImpXml.Close()
    End Sub

End Class