Imports System.Data.SqlClient
Public Class AdatecDat

    Public Sub ConsultarPedido(ByRef dsDatos As DataSet)

        Dim objCorreo As New clsCorreo
        Dim sqlConexion As New SqlConnection(My.Settings.ConexionIntermedias)
        Dim sqlComando As SqlCommand = New SqlCommand
        Dim daDatos As New SqlDataAdapter
        sqlComando.Connection = sqlConexion
        sqlComando.CommandType = CommandType.StoredProcedure
        sqlComando.CommandText = "sp_DatosMigrar_Pedidos"
        sqlComando.CommandTimeout = 18000000
        Try
            daDatos.SelectCommand = sqlComando
            daDatos.Fill(dsDatos)
            dsDatos.Tables(0).TableName = "Pedidos"
            dsDatos.Tables(1).TableName = "Descuentos"
            dsDatos.Tables(2).TableName = "MovtoPedidosComercial"




        Catch ex As Exception
            objCorreo.EnviarCorreoTareaExe("DisLaCosta: Error al vincular estructura: ", ex.ToString)
        Finally
            sqlComando.Parameters.Clear()
            sqlComando.Connection.Close()
        End Try
    End Sub


    Public Sub ConsultarRecibo(ByRef dsDatos As DataSet)

        Dim objCorreo As New clsCorreo
        Dim sqlConexion As New SqlConnection(My.Settings.ConexionIntermedias)
        Dim sqlComando As SqlCommand = New SqlCommand
        Dim daDatos As New SqlDataAdapter
        sqlComando.Connection = sqlConexion
        sqlComando.CommandType = CommandType.StoredProcedure
        sqlComando.CommandText = "sp_DatosMigrar_Recibos"
        sqlComando.CommandTimeout = 18000000
        Try
            daDatos.SelectCommand = sqlComando
            daDatos.Fill(dsDatos)
            dsDatos.Tables(0).TableName = "Recibos"
            dsDatos.Tables(1).TableName = "Caja"
            dsDatos.Tables(2).TableName = "CxC"
        Catch ex As Exception
            objCorreo.EnviarCorreoTareaExe("DisLaCosta: Error al vincular estructura: ", ex.ToString)
        Finally
            sqlComando.Parameters.Clear()
            sqlComando.Connection.Close()
        End Try
    End Sub
End Class
