Imports System.Data.SqlClient

Public Class clsBaseDatos


    Public Function AgenciaCerrada(ByVal nCodAge As Integer, ByVal pdFecha As Date) As Integer


        Dim objCon As New clsConecta
        Dim oComm As SqlClient.SqlCommand

        oComm = New SqlClient.SqlCommand("PA_AgenciaCerrada", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@pnCodage", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@pdFecha", SqlDbType.Date).Value = pdFecha
        oComm.Parameters.Add("@pnResp", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()
        objCon.CierraConexion()

        AgenciaCerrada = oComm.Parameters("@pnResp").Value


    End Function


    Public Function ListaIngresos(ByVal psCodUsu As String, ByVal pdFecha As Date) As DataTable


        Dim objCon As New clsConecta
        Dim oComm As SqlClient.SqlCommand

        oComm = New SqlClient.SqlCommand("PA_RecuperaIngresosxFecha", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@psCodUsu", SqlDbType.VarChar, 50).Value = psCodUsu
        oComm.Parameters.Add("@dFecha", SqlDbType.Date).Value = pdFecha

        Dim Table As DataTable
        Table = New DataTable
        Dim da As SqlDataAdapter
        da = New SqlDataAdapter(oComm)
        da.Fill(Table)

        objCon.CierraConexion()

        ListaIngresos = Table


    End Function

    Public Function ListaEgresos(ByVal psCodUsu As String, ByVal pdFecha As Date) As DataTable


        Dim objCon As New clsConecta
        Dim oComm As SqlClient.SqlCommand

        oComm = New SqlClient.SqlCommand("PA_RecuperaEgresosxFecha", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@psCodUsu", SqlDbType.VarChar, 50).Value = psCodUsu
        oComm.Parameters.Add("@pdFecCuadre", SqlDbType.Date).Value = pdFecha

        Dim Table As DataTable
        Table = New DataTable
        Dim da As SqlDataAdapter
        da = New SqlDataAdapter(oComm)
        da.Fill(Table)

        objCon.CierraConexion()

        ListaEgresos = Table


    End Function

    Public Function DevuelveSaldoCajaAnterior(ByVal psCodUSu As String, ByVal pdFecha As Date) As DataTable


        Dim objCon As New clsConecta
        Dim oComm As SqlClient.SqlCommand

        oComm = New SqlClient.SqlCommand("PA_RecuperaSaldoAnterior", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@psCodUsu", SqlDbType.VarChar, 50).Value = psCodUSu
        oComm.Parameters.Add("@pdFecha", SqlDbType.Date).Value = pdFecha

        Dim Table As DataTable
        Table = New DataTable
        Dim da As SqlDataAdapter
        da = New SqlDataAdapter(oComm)
        da.Fill(Table)

        objCon.CierraConexion()

        DevuelveSaldoCajaAnterior = Table


    End Function

    Public Function DevuelveSaldoCajaHoy(ByVal psCodUSu As String, ByVal pdFecha As Date) As DataTable


        Dim objCon As New clsConecta
        Dim oComm As SqlClient.SqlCommand

        oComm = New SqlClient.SqlCommand("PA_RecuperaSaldoEfectivoHOY", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@psCodUsu", SqlDbType.VarChar, 50).Value = psCodUSu
        oComm.Parameters.Add("@pdFecha", SqlDbType.Date).Value = pdFecha

        Dim Table As DataTable
        Table = New DataTable
        Dim da As SqlDataAdapter
        da = New SqlDataAdapter(oComm)
        da.Fill(Table)

        objCon.CierraConexion()

        DevuelveSaldoCajaHoy = Table


    End Function

    Public Function ListaUsuariosCuadre(ByVal psCodUsu As String) As DataTable


        Dim objCon As New clsConecta
        Dim oComm As SqlClient.SqlCommand

        oComm = New SqlClient.SqlCommand("PA_RecuperaIngresosxFecha", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@psCodUsu", SqlDbType.VarChar, 50).Value = psCodUsu

        Dim Table As DataTable
        Table = New DataTable
        Dim da As SqlDataAdapter
        da = New SqlDataAdapter(oComm)
        da.Fill(Table)

        objCon.CierraConexion()

        ListaUsuariosCuadre = Table


    End Function

End Class
