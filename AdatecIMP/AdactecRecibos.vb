Imports System.Data.SqlClient

Public Class AdactecRecibos


    '************************

    Dim objConfiguraciones As New clsConfiguracion
    Dim objCorreo As New clsCorreo
    Public Sub AdactecRecibos()
        Try

            '1. Consumir el ws de siesa para consultas
            Dim objSiesa As New wsUnoEE.WSUNOEE
            objSiesa.Timeout = 1800000000
            Dim ParametroWS As String = "<?xml version=""1.0"" encoding=""utf-8""?>
                <Consulta>
                <NombreConexion>" & objConfiguraciones.ConexionWsSiesa & "</NombreConexion>
                <IdCia>" & objConfiguraciones.CompaniaUnoEE & "</IdCia>
                <IdProveedor>INTERFACES_Y_SOLUCIONES</IdProveedor>
                <IdConsulta>CONSULTA_WS_RECIBOS</IdConsulta>
                <Usuario>" & objConfiguraciones.UsuarioUnoEE & "</Usuario>
                <Clave>" & objConfiguraciones.ClaveUnoEE & "</Clave>
                <Parametros></Parametros>
                </Consulta>"

            Dim dsResultadoSiesa As DataSet
            dsResultadoSiesa = objSiesa.EjecutarConsultaXML(ParametroWS)


            Dim datos As DataTable = New DataTable("datos")

            '    Dim Columna1 As DataColumn = New DataColumn()
            '    Columna1.DataType = System.Type.GetType("System.String")
            '    Columna1.ColumnName = "Columna1"
            '    Dim Columna2 As DataColumn = New DataColumn()
            '    Columna2.DataType = System.Type.GetType("System.String")
            '    Columna2.ColumnName = "Columna2"
            '    Dim Columna3 As DataColumn = New DataColumn()
            '    Columna3.DataType = System.Type.GetType("System.String")
            '    Columna3.ColumnName = "Columna3"
            '    Dim Columna4 As DataColumn = New DataColumn()
            '    Columna4.DataType = System.Type.GetType("System.String")
            '    Columna4.ColumnName = "Columna4"

            '    datos.Columns.Add(Columna1)
            '    datos.Columns.Add(Columna2)
            '    datos.Columns.Add(Columna3)
            '    datos.Columns.Add(Columna4)

            '    'LLENAR DESDE LA BASE DE DATOS el DATATABLE
            '    Dim fila As DataRow
            '    For Each FilaSiesa As DataRow In dsResultadoSiesa.Tables(0).Rows
            '        fila = datos.NewRow()

            '        fila("Columna1") = FilaSiesa.Item("f430_id_co")
            '        fila("Columna2") = FilaSiesa.Item("f430_id_tipo_docto")
            '        fila("Columna3") = FilaSiesa.Item("f430_consec_docto")
            '        fila("Columna4") = FilaSiesa.Item("f430_notas")
            '        datos.Rows.Add(fila)
            '    Next

            '    CargarTablaBulkCopy(datos, "ADATEC_SIESA_PEDIDOS")

            '    'ejecutar SP actualizar estado INTEGRADO
            Dim ObjConexion As New SqlConnection(My.MySettings.Default.ConexionIntermedias)
            Dim objComando As New SqlCommand
            Dim dsCentroOp As New DataSet
            Dim objDA As New SqlDataAdapter
            objComando.Connection = ObjConexion
            '    objComando.CommandType = CommandType.StoredProcedure
            '    objComando.CommandText = "SAVICOLVerificarDctoPedidos"
            Try
                ObjConexion.Open()
                objComando.ExecuteNonQuery()
            Catch ex As Exception
                'objCorreo.EnviarCorreoTareaExe(" Error al ejecutar procedimiento almacenado: ", ex.ToString)
            Finally
            objComando.Parameters.Clear()
            objComando.Connection.Close()
            ObjConexion.Close()
        End Try
        '''''''''''''''''''''''''''''''''''''''''''''
        Catch ex As Exception
        objCorreo.EnviarCorreoTareaExe(" Error al sincronziar pedidos desde SIESA: ", ex.ToString)
        End Try
    End Sub

    Private Sub CargarTablaBulkCopy(ByVal Datos As DataTable, ByVal Tabla As String)
        Dim ObjConexion As New SqlConnection(My.MySettings.Default.ConexionIntermedias)

        Try

            Dim bulkCopy As New SqlBulkCopy(ObjConexion)
            ObjConexion.Open()
            bulkCopy.BulkCopyTimeout = 3600000

            limpiar_Tablas(Tabla)
            bulkCopy.DestinationTableName = "dbo." & Tabla
            bulkCopy.WriteToServer(Datos)

        Catch ex As Exception
            objCorreo.EnviarCorreoTareaExe(" Error al cargar datos a las tablas auxiliares: ", ex.ToString)
        Finally
            ObjConexion.Close()
        End Try
    End Sub

    Public Sub limpiar_Tablas(Tabla As String)
        Dim ObjConexion As New SqlConnection(My.MySettings.Default.ConexionIntermedias)
        Dim objComando As New SqlCommand
        Dim dsCentroOp As New DataSet
        Dim objDA As New SqlDataAdapter
        objComando.Connection = ObjConexion
        objComando.CommandType = CommandType.StoredProcedure
        objComando.CommandText = "LimpiarTablas"

        Try
            ObjConexion.Open()
            objComando.Parameters.AddWithValue("@NombreTabla", Tabla)
            objComando.ExecuteNonQuery()
        Catch ex As Exception
            objCorreo.EnviarCorreoTareaExe("Error al limpiar datos de las tablas auxiliares: ", ex.ToString)
        Finally
            objComando.Parameters.Clear()
            objComando.Connection.Close()
            ObjConexion.Close()
        End Try
    End Sub
    '************************


End Class
