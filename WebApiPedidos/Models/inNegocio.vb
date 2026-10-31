Imports System.Data.SqlClient

Public Class inNegocio

    Public Property sToken As String
    Public Property nSolicitud As Integer
    Public Property sCodUsu As String
    Public Property oComercio As clsComercio
    Public Property cValorBusqCom As String
    Public Property nTipoBusqCom As Integer
    Public Property oOrden As clsOrden
    Public Property oPedido As clsPedidos
    Public Property nIdCom As Integer
    Public Property nLat As Double
    Public Property nLong As Double
    Public Property sClave As String
    Public Property oUbicacion As clsUbicacion
    Public Property oNoGestion As clsNogestion
    Public Property nCodCatalogo As Integer

    Public Property nNumTran As Long
    Public Property nPagoEntrega As Integer


    Public Property nCodRol As Integer
    Public Function ValidaRol() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i, j As Integer
        Dim nCorrecto As Integer
        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_ValidaRol", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("psCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@nCodRol", SqlDbType.Int).Value = nCodRol
            oComm.Parameters.Add("@pnResp", SqlDbType.Int).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            nCorrecto = oComm.Parameters("@pnResp").Value

            oConex.Close()

            If nCorrecto = 1 Then
                D(0) = "1"
                D(1) = "Acceso Correcto"
            Else
                D(0) = "0"
                D(1) = "Clave o Usuario Invalidos"
            End If


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function

    Public Function RegistraAnulaPago() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i, j As Integer
        Dim nCorrecto As Integer
        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_AnulaPago", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nNumTran", SqlDbType.Int).Value = nNumTran

            oComm.ExecuteNonQuery()


            oConex.Close()


            D(0) = "1"
            D(1) = "Registro Correcto"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function


    Public Function ValidaZona() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i, j As Integer
        Dim nCorrecto As Integer
        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_ValidaZona", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@nLat", SqlDbType.Float).Value = nLat
            oComm.Parameters.Add("@nlong", SqlDbType.Float).Value = nLong
            oComm.Parameters.Add("@nResultado", SqlDbType.Int).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()


            oConex.Close()

            If oComm.Parameters("@nResultado").Value = 1 Then
                D(0) = "1"
                D(1) = "Zona Correcta"
            Else
                D(0) = "0"
                D(1) = "Usted NO PERTENECE a esta ZONA"

            End If


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function

    Public Function RegistraUbicacion() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i, j As Integer
        Dim nCorrecto As Integer
        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_RegistraUbicacion", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = oUbicacion.cCodUsu
            oComm.Parameters.Add("@nLat", SqlDbType.Float).Value = oUbicacion.nLat
            oComm.Parameters.Add("@nlong", SqlDbType.Float).Value = oUbicacion.nLong

            oComm.ExecuteNonQuery()


            oConex.Close()


            D(0) = "1"
            D(1) = "Registro Correcto"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function

    Public Function ValidaClave() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i, j As Integer
        Dim nCorrecto As Integer
        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_ValidaClave", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@cClave", SqlDbType.VarChar, 50).Value = sClave
            oComm.Parameters.Add("@nCorrecto", SqlDbType.Int).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            nCorrecto = oComm.Parameters("@nCorrecto").Value

            oConex.Close()

            If nCorrecto = 1 Then
                D(0) = "1"
                D(1) = "Acceso Correcto"
            Else
                D(0) = "0"
                D(1) = "Clave o Usuario Invalidos"
            End If


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function


    Public Function RegistraEntregaPedido() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i, j As Integer
        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_RegistraEntregaPedido_2", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nIdOrden", SqlDbType.Int).Value = oPedido.nIdOrden
            oComm.Parameters.Add("@nIdProd", SqlDbType.Int).Value = oPedido.nIdProd
            oComm.Parameters.Add("@nCantidadEntregada", SqlDbType.Int).Value = oPedido.nCantidad
            oComm.Parameters.Add("@nLat", SqlDbType.Float).Value = nLat
            oComm.Parameters.Add("@nLong", SqlDbType.Float).Value = nLong
            oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@nPagoEntrega", SqlDbType.Int).Value = nPagoEntrega

            oComm.ExecuteNonQuery()

            oConex.Close()

            D(0) = "1"
            D(1) = "Registro Correcto"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function


    Public Function RecuperaCobroComercio() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i, j, k As Integer
        Dim oCobros As New clsCobro
        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try

            oComm = New SqlClient.SqlCommand("PA_RecuperaComercio_CobroPendiente", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nIDCom", SqlDbType.Int).Value = nIdCom
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read

                oCobros.nIDCom = dr.Item("nIDCom")
                oCobros.cRazonSocial = dr.Item("cRazonSocial")

                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            oComm = New SqlClient.SqlCommand("PA_RecuperaOrdenCobroComercio", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nIdCom", SqlDbType.Int).Value = nIdCom
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read

                oCobros.oOrdenCobro = New clsOrden
                oCobros.oOrdenCobro.cRazonSocial = dr.Item("cRazonSocial")
                oCobros.oOrdenCobro.cUser = ""
                oCobros.oOrdenCobro.nEstado = dr.Item("nEstado")
                oCobros.oOrdenCobro.nIdOrden = dr.Item("nIdOrden")
                oCobros.oOrdenCobro.nLat = dr.Item("nLat")
                oCobros.oOrdenCobro.nLong = dr.Item("nLong")

                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            If Not IsNothing(oCobros.oOrdenCobro) Then
                oComm = New SqlClient.SqlCommand("PA_RecuperaOrdenesEntregadasParaCobro", oConex)
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.Parameters.Add("@nIdOrden", SqlDbType.Int).Value = oCobros.oOrdenCobro.nIdOrden
                oComm.ExecuteNonQuery()

                dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                i = 0
                While dr.Read
                    ReDim Preserve oCobros.oOrdenCobro.oPedidos(i)
                    oCobros.oOrdenCobro.oPedidos(i) = New clsPedidos
                    oCobros.oOrdenCobro.oPedidos(i).nIdOrden = dr.Item("nIdOrden")
                    oCobros.oOrdenCobro.oPedidos(i).nIdProd = dr.Item("nIdProd")
                    oCobros.oOrdenCobro.oPedidos(i).nCantidad = dr.Item("nCantidad")
                    oCobros.oOrdenCobro.oPedidos(i).nMonto = dr.Item("nMonto")
                    oCobros.oOrdenCobro.oPedidos(i).cNomProd = dr.Item("cNomProd")
                    oCobros.oOrdenCobro.oPedidos(i).nPrecioUnitario = dr.Item("nPreciounitario")
                    oCobros.oOrdenCobro.oPedidos(i).nEstado = dr.Item("nCobrado")
                    oCobros.oOrdenCobro.oPedidos(i).nEntregado = dr.Item("nEntregado")

                    i = i + 1
                End While
                dr.Close()
                dr = Nothing
            End If

            '***** PARA ENTREGAS
            oComm = New SqlClient.SqlCommand("PA_RecuperaOrdenCobroEntregaPedido", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nIdCom", SqlDbType.Int).Value = oCobros.nIDCom
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read

                oCobros.oOrdenEntrega = New clsOrden
                oCobros.oOrdenEntrega.cRazonSocial = dr.Item("cRazonSocial")
                oCobros.oOrdenEntrega.cUser = ""
                oCobros.oOrdenEntrega.dFecha = dr.Item("dFecha")
                oCobros.oOrdenEntrega.nEstado = dr.Item("nEstado")
                oCobros.oOrdenEntrega.nIdOrden = dr.Item("nIdOrden")
                oCobros.oOrdenEntrega.nLat = dr.Item("nLat")
                oCobros.oOrdenEntrega.nLong = dr.Item("nLong")

                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            If Not IsNothing(oCobros.oOrdenEntrega) Then


                oComm = New SqlClient.SqlCommand("PA_RecuperaOrdenesPedidos", oConex)
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.Parameters.Add("@nIdOrden", SqlDbType.Int).Value = oCobros.oOrdenEntrega.nIdOrden
                oComm.ExecuteNonQuery()

                dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                i = 0
                While dr.Read
                    ReDim Preserve oCobros.oOrdenEntrega.oPedidos(i)
                    oCobros.oOrdenEntrega.oPedidos(i) = New clsPedidos
                    oCobros.oOrdenEntrega.oPedidos(i).nIdOrden = dr.Item("nIdOrden")
                    oCobros.oOrdenEntrega.oPedidos(i).nIdProd = dr.Item("nIdProd")
                    oCobros.oOrdenEntrega.oPedidos(i).nCantidad = dr.Item("nCantidad")
                    oCobros.oOrdenEntrega.oPedidos(i).nMonto = dr.Item("nMonto")
                    oCobros.oOrdenEntrega.oPedidos(i).cNomProd = dr.Item("cNomProd")
                    oCobros.oOrdenEntrega.oPedidos(i).nPrecioUnitario = dr.Item("nPreciounitario")
                    oCobros.oOrdenEntrega.oPedidos(i).nEstado = dr.Item("nCobrado")
                    oCobros.oOrdenEntrega.oPedidos(i).nEntregado = dr.Item("nEntregado")


                    i = i + 1
                End While
                dr.Close()
                dr = Nothing
            End If


            oConex.Close()

            R.oCobroCom = oCobros

            D(0) = "1"
            D(1) = "Registro Correcto"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function

    Public Function RecuperaCobros() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i, j As Integer
        Dim oCobros() As clsCobro
        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_RecuperaComerciosCobroPendiente", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read

                ReDim Preserve oCobros(i)
                oCobros(i) = New clsCobro
                oCobros(i).nIDCom = dr.Item("nIDCom")
                oCobros(i).cRazonSocial = dr.Item("cRazonSocial")
                oCobros(i).cDoc = dr.Item("cDoc")

                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            If IsNothing(oCobros) Then
                D(0) = "1"
                D(1) = "NO EXISTEN DATOS"
                R.nResultado = CInt(D(0))
                R.sRespuesta = D(1)
                Return R
            End If


            If oCobros.Count = 0 Then
                D(0) = "1"
                D(1) = "NO EXISTEN DATOS"
                R.nResultado = CInt(D(0))
                R.sRespuesta = D(1)
                Return R
            End If

            For j = 0 To oCobros.Count - 1


                oComm = New SqlClient.SqlCommand("PA_RecuperaOrdenCobroComercio", oConex)
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.Parameters.Add("@nIdCom", SqlDbType.Int).Value = oCobros(j).nIDCom
                oComm.ExecuteNonQuery()

                dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                i = 0
                While dr.Read

                    oCobros(j).oOrdenCobro = New clsOrden
                    oCobros(j).oOrdenCobro.cRazonSocial = dr.Item("cRazonSocial")
                    oCobros(j).oOrdenCobro.cUser = ""
                    oCobros(j).oOrdenCobro.nEstado = dr.Item("nEstado")
                    oCobros(j).oOrdenCobro.nIdOrden = dr.Item("nIdOrden")
                    oCobros(j).oOrdenCobro.nLat = dr.Item("nLat")
                    oCobros(j).oOrdenCobro.nLong = dr.Item("nLong")

                    i = i + 1
                End While
                dr.Close()
                dr = Nothing

                If Not IsNothing(oCobros(j).oOrdenCobro) Then
                    oComm = New SqlClient.SqlCommand("PA_RecuperaOrdenesEntregadasParaCobro", oConex)
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.Parameters.Add("@nIdOrden", SqlDbType.Int).Value = oCobros(j).oOrdenCobro.nIdOrden
                    oComm.ExecuteNonQuery()

                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0
                    While dr.Read
                        ReDim Preserve oCobros(j).oOrdenCobro.oPedidos(i)
                        oCobros(j).oOrdenCobro.oPedidos(i) = New clsPedidos
                        oCobros(j).oOrdenCobro.oPedidos(i).nIdOrden = dr.Item("nIdOrden")
                        oCobros(j).oOrdenCobro.oPedidos(i).nIdProd = dr.Item("nIdProd")
                        oCobros(j).oOrdenCobro.oPedidos(i).nCantidad = dr.Item("nCantidad")
                        oCobros(j).oOrdenCobro.oPedidos(i).nMonto = dr.Item("nMonto")
                        oCobros(j).oOrdenCobro.oPedidos(i).cNomProd = dr.Item("cNomProd")
                        oCobros(j).oOrdenCobro.oPedidos(i).nPrecioUnitario = dr.Item("nPreciounitario")
                        oCobros(j).oOrdenCobro.oPedidos(i).nEstado = dr.Item("nCobrado")
                        oCobros(j).oOrdenCobro.oPedidos(i).nEntregado = dr.Item("nEntregado")

                        i = i + 1
                    End While
                    dr.Close()
                    dr = Nothing
                End If

                '***** PARA ENTREGAS
                oComm = New SqlClient.SqlCommand("PA_RecuperaOrdenCobroEntregaPedido", oConex)
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.Parameters.Add("@nIdCom", SqlDbType.Int).Value = oCobros(j).nIDCom
                oComm.ExecuteNonQuery()

                dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                i = 0
                While dr.Read

                    oCobros(j).oOrdenEntrega = New clsOrden
                    oCobros(j).oOrdenEntrega.cRazonSocial = dr.Item("cRazonSocial")
                    oCobros(j).oOrdenEntrega.cUser = ""
                    oCobros(j).oOrdenEntrega.dFecha = dr.Item("dFecha")
                    oCobros(j).oOrdenEntrega.nEstado = dr.Item("nEstado")
                    oCobros(j).oOrdenEntrega.nIdOrden = dr.Item("nIdOrden")
                    oCobros(j).oOrdenEntrega.nLat = dr.Item("nLat")
                    oCobros(j).oOrdenEntrega.nLong = dr.Item("nLong")

                    i = i + 1
                End While
                dr.Close()
                dr = Nothing

                If Not IsNothing(oCobros(j).oOrdenEntrega) Then
                    oComm = New SqlClient.SqlCommand("PA_RecuperaOrdenesPedidos", oConex)
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.Parameters.Add("@nIdOrden", SqlDbType.Int).Value = oCobros(j).oOrdenEntrega.nIdOrden
                    oComm.ExecuteNonQuery()

                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0
                    While dr.Read
                        ReDim Preserve oCobros(j).oOrdenEntrega.oPedidos(i)
                        oCobros(j).oOrdenEntrega.oPedidos(i) = New clsPedidos
                        oCobros(j).oOrdenEntrega.oPedidos(i).nIdOrden = dr.Item("nIdOrden")
                        oCobros(j).oOrdenEntrega.oPedidos(i).nIdProd = dr.Item("nIdProd")
                        oCobros(j).oOrdenEntrega.oPedidos(i).nCantidad = dr.Item("nCantidad")
                        oCobros(j).oOrdenEntrega.oPedidos(i).nMonto = dr.Item("nMonto")
                        oCobros(j).oOrdenEntrega.oPedidos(i).cNomProd = dr.Item("cNomProd")
                        oCobros(j).oOrdenEntrega.oPedidos(i).nPrecioUnitario = dr.Item("nPreciounitario")
                        oCobros(j).oOrdenEntrega.oPedidos(i).nEstado = dr.Item("nCobrado")
                        oCobros(j).oOrdenEntrega.oPedidos(i).nEntregado = dr.Item("nEntregado")
                        i = i + 1
                    End While
                    dr.Close()
                    dr = Nothing
                End If
            Next
            oConex.Close()

            R.oCobros = oCobros

            D(0) = "1"
            D(1) = "Registro Correcto"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function

    Public Function RegistraCobroPedido() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i, j As Integer
        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_RegistraCobroPedido_2", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nIdOrden", SqlDbType.Int).Value = oPedido.nIdOrden
            oComm.Parameters.Add("@nIdProd", SqlDbType.Int).Value = oPedido.nIdProd
            oComm.Parameters.Add("@nCobrado", SqlDbType.Int).Value = 1
            oComm.Parameters.Add("@nMontoxCobrar", SqlDbType.Money).Value = oPedido.nMontoxCobrar
            oComm.Parameters.Add("@nMontoCobrado", SqlDbType.Money).Value = oPedido.nMontoCobrado
            oComm.Parameters.Add("@nCantidadVendida", SqlDbType.Int).Value = oPedido.nCantidadVendida
            oComm.Parameters.Add("@nCantidadPorDevolver", SqlDbType.Int).Value = oPedido.nCantidadPorDevolver
            oComm.Parameters.Add("@nCantidadEntregada", SqlDbType.Int).Value = oPedido.nCantidadEntregada
            oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = sCodUsu

            oComm.Parameters.Add("@pnResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@psRespuesta", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            oConex.Close()

            D(0) = Str(oComm.Parameters("@pnResp").Value)
            D(1) = oComm.Parameters("@psRespuesta").Value


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function


    Public Function RecuperaOrdenes() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim oComm2 As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim dr2 As SqlClient.SqlDataReader
        Dim i, j As Integer
        Dim nOrden As Long
        Dim oOrdenes() As clsOrden
        Dim oPedidos() As clsPedidos
        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_RecuperaOrdenes", oConex)
            oComm.CommandTimeout = 0
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.CommandType = CommandType.StoredProcedure

            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read

                ReDim Preserve oOrdenes(i)
                oOrdenes(i) = New clsOrden
                oOrdenes(i).nIdOrden = dr.Item("nIdOrden")
                oOrdenes(i).dFecha = dr.Item("dFecha")
                oOrdenes(i).nIdCom = dr.Item("nIdCom")
                oOrdenes(i).cUser = dr.Item("cUser")
                oOrdenes(i).nEstado = dr.Item("nEstado")
                oOrdenes(i).nLat = dr.Item("nLat")
                oOrdenes(i).nLong = dr.Item("nLong")
                oOrdenes(i).cRazonSocial = dr.Item("cRazonSocial")


                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            If IsNothing(oOrdenes) Then
                D(0) = "1"
                D(1) = "NO Se Encontraron Ordenes"
                R.nResultado = CInt(D(0))
                R.sRespuesta = D(1)
                Return R

            End If

            i = 0
            For i = 0 To oOrdenes.Count - 1

                oComm2 = New SqlClient.SqlCommand("PA_RecuperaOrdenesPedidos", oConex)
                oComm2.CommandTimeout = 0
                oComm2.CommandType = CommandType.StoredProcedure
                oComm2.Parameters.Add("@nIdOrden", SqlDbType.Int).Value = oOrdenes(i).nIdOrden

                oComm2.ExecuteNonQuery()

                dr2 = oComm2.ExecuteReader(CommandBehavior.SingleResult)
                j = 0

                While dr2.Read
                    ReDim Preserve oPedidos(j)
                    oPedidos(j) = New clsPedidos
                    oPedidos(j).nIdProd = dr2.Item("nIdProd")
                    oPedidos(j).nCantidad = dr2.Item("nCantidad")
                    oPedidos(j).nMonto = dr2.Item("nMonto")
                    oPedidos(j).cNomProd = dr2.Item("cNomProd")

                    j = j + 1
                End While
                oOrdenes(i).oPedidos = oPedidos

                dr2.Close()
                dr2 = Nothing

            Next

            oConex.Close()

            R.oOrdenes = oOrdenes

            D(0) = "1"
            D(1) = "Registro Correcto"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function

    Public Function EliminaOrden() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i As Integer
        Dim nOrden As Long

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_EliminaOrden_2", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nIdOrden", SqlDbType.Int).Value = oOrden.nIdOrden
            oComm.Parameters.Add("@psUser", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            nOrden = oComm.Parameters("@nResp").Value


            oConex.Close()

            D(0) = "1"
            D(1) = "Registro Correcto"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function


    Public Function ValidaSiExisteOrden() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i As Integer
        Dim nOrden As Long

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_ValidaExisteOrden", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nIdCom", SqlDbType.Int).Value = oOrden.nIdCom
            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            nOrden = oComm.Parameters("@nResp").Value


            oConex.Close()

            D(0) = Str(nOrden).Trim
            D(1) = "Registro Correcto"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function


    Public Function RegistraOrden() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i As Integer
        Dim nOrden As Long

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_RegistraOrden_2", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nIdCom", SqlDbType.Int).Value = oOrden.nIdCom
            oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = oOrden.cUser
            oComm.Parameters.Add("@nEstado", SqlDbType.Int).Value = oOrden.nEstado
            oComm.Parameters.Add("@nLat", SqlDbType.Float).Value = oOrden.nLat
            oComm.Parameters.Add("@nLong", SqlDbType.Float).Value = oOrden.nLong
            oComm.Parameters.Add("@dFechaEntrega", SqlDbType.Date).Value = oOrden.oPedidos(0).dFechaEntrega
            oComm.Parameters.Add("@nIDOrden", SqlDbType.Int).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            nOrden = oComm.Parameters("@nIDOrden").Value

            For i = 0 To oOrden.oPedidos.Count - 1

                oComm = New SqlClient.SqlCommand("PA_RegistraOrdenPedido_2", oConex)
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.Parameters.Add("@nIdOrden", SqlDbType.Int).Value = nOrden
                oComm.Parameters.Add("@nIdProd", SqlDbType.Int).Value = oOrden.oPedidos(i).nIdProd
                oComm.Parameters.Add("@nCantidad", SqlDbType.Int).Value = oOrden.oPedidos(i).nCantidad
                oComm.Parameters.Add("@nMonto", SqlDbType.Int).Value = oOrden.oPedidos(i).nMonto
                oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = oOrden.cUser
                oComm.Parameters.Add("@nPedidoHoy", SqlDbType.Int).Value = oOrden.oPedidos(i).nPedidoHoy

                oComm.ExecuteNonQuery()

            Next



            oConex.Close()

            D(0) = "1"
            D(1) = "Registro Correcto"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function


    Public Function RecuperaCatalogo() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim Com() As clsCatalogo
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i As Integer

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_RecuperaCatalogo", oConex)
            oComm.Parameters.Add("@nCodCat", SqlDbType.Int).Value = nCodCatalogo
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read

                ReDim Preserve Com(i)
                Com(i) = New clsCatalogo
                Com(i).nIDCod = dr.Item("nIdCod")
                Com(i).cNomCod = dr.Item("cNomCod")
                Com(i).cValorCod = dr.Item("cValorcod")

                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            R.oCatalogo = Com

            oConex.Close()

            D(0) = "1"
            D(1) = "Registro Correcto"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function

    Public Function RecuperaProductos() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim Com() As clsProducto
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i As Integer

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_RecuperaProductos", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read

                ReDim Preserve Com(i)
                Com(i) = New clsProducto
                Com(i).nIdProd = dr.Item("nIdProd")
                Com(i).cNomProd = dr.Item("cNomProd")
                Com(i).nPrecio = dr.Item("nPrecio")

                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            R.oProductos = Com

            oConex.Close()

            D(0) = "1"
            D(1) = "Registro Correcto"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function

    Public Function RecuperaPagoParaAnulacion() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i As Integer
        Dim oPagoAnulacion As clsAnulacionPago

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_RecuperaPagoParaAnular", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nIdCom", SqlDbType.Int).Value = nIdCom
            oComm.Parameters.Add("@dFecha", SqlDbType.Date).Value = Now()

            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            oPagoAnulacion = New clsAnulacionPago
            While dr.Read

                oPagoAnulacion.cCodUsu = dr.Item("cUser")
                oPagoAnulacion.cNomComercio = dr.Item("cRazonSocial")
                oPagoAnulacion.dFecha = dr.Item("dFecha")
                oPagoAnulacion.nCantidad = dr.Item("nMermaEntregada")
                oPagoAnulacion.nIDCom = dr.Item("nIDCom")
                oPagoAnulacion.nIdOrden = dr.Item("nIDOrden")
                oPagoAnulacion.nMonto = dr.Item("nMontoCobrado")
                oPagoAnulacion.nNumTran = dr.Item("nNroTrans")

                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            R.oPagoAnulacion = oPagoAnulacion

            oConex.Close()

            D(0) = "1"
            D(1) = "Registro Correcto"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function

    Public Function RecuperaBusquedaSoloComercio() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim Com() As clsComercio
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i As Integer

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_RecuperaBusqSoloComercio", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nIDCom", SqlDbType.Int).Value = nIdCom

            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read

                ReDim Preserve Com(i)
                Com(i) = New clsComercio
                Com(i).nIDCom = dr.Item("nIDCom")
                Com(i).cRazonSocial = dr.Item("cRazonSocial")
                Com(i).cDoc = dr.Item("cDoc")
                Com(i).cApePat = dr.Item("cApePat")
                Com(i).cApeMat = dr.Item("cApeMat")
                Com(i).cNombres = dr.Item("cNombres")
                Com(i).cTelefFijo = dr.Item("cTelefFijo")
                Com(i).cMovil = dr.Item("cMovil")
                Com(i).cDireccion = dr.Item("cDireccion")
                Com(i).cReferencia = dr.Item("cReferencia")
                Com(i).nLat = dr.Item("nLat")
                Com(i).nLong = dr.Item("nLong")
                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            R.oComercios = Com

            oConex.Close()

            D(0) = "1"
            D(1) = "Registro Correcto"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function

    Public Function RecuperaBusquedaComercios() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim Com() As clsComercio
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i As Integer

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_RecuperaBusqComercio", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nTipoBusq", SqlDbType.Int).Value = nTipoBusqCom
            oComm.Parameters.Add("@psValor", SqlDbType.VarChar, 150).Value = cValorBusqCom

            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read

                ReDim Preserve Com(i)
                Com(i) = New clsComercio
                Com(i).nIDCom = dr.Item("nIDCom")
                Com(i).cRazonSocial = dr.Item("cRazonSocial")
                Com(i).cDoc = dr.Item("cDoc")
                Com(i).cApePat = dr.Item("cApePat")
                Com(i).cApeMat = dr.Item("cApeMat")
                Com(i).cNombres = dr.Item("cNombres")
                Com(i).cTelefFijo = dr.Item("cTelefFijo")
                Com(i).cMovil = dr.Item("cMovil")
                Com(i).cDireccion = dr.Item("cDireccion")
                Com(i).cReferencia = dr.Item("cReferencia")
                Com(i).nLat = dr.Item("nLat")
                Com(i).nLong = dr.Item("nLong")
                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            R.oComercios = Com

            oConex.Close()

            D(0) = "1"
            D(1) = "Registro Correcto"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function

    Public Function RegistraComercio() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim i As Integer
        Dim nIdCom As Long
        Dim nResp As Integer
        Dim sRespuesta As String

        nResp = 0
        sRespuesta = ""


        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try

            If oComercio.nNoInteresado = 0 Then

                R = ValidaZona()
                If R.nResultado = 0 Then
                    Return R
                End If

                oComm = New SqlClient.SqlCommand("PA_RegistraComercio_3", oConex)
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.Parameters.Add("@cRazonSocial", SqlDbType.VarChar, 150).Value = oComercio.cRazonSocial
                oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = oComercio.cDoc
                oComm.Parameters.Add("@cApePat", SqlDbType.VarChar, 50).Value = oComercio.cApePat
                oComm.Parameters.Add("@cApeMat", SqlDbType.VarChar, 50).Value = oComercio.cApeMat
                oComm.Parameters.Add("@cNombres", SqlDbType.VarChar, 150).Value = oComercio.cNombres
                oComm.Parameters.Add("@cTelefFijo", SqlDbType.VarChar, 50).Value = oComercio.cTelefFijo
                oComm.Parameters.Add("@cMovil", SqlDbType.VarChar, 50).Value = oComercio.cMovil
                oComm.Parameters.Add("@cDireccion", SqlDbType.VarChar, 150).Value = oComercio.cDireccion
                oComm.Parameters.Add("@cReferencia", SqlDbType.VarChar, 250).Value = oComercio.cReferencia
                oComm.Parameters.Add("@nLat", SqlDbType.Float).Value = oComercio.nLat
                oComm.Parameters.Add("@nLong", SqlDbType.Float).Value = oComercio.nLong
                oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
                oComm.Parameters.Add("@nIdCom", SqlDbType.Int).Direction = ParameterDirection.Output
                oComm.Parameters.Add("@pnResp", SqlDbType.Int).Direction = ParameterDirection.Output
                oComm.Parameters.Add("@psRespuesta", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
                oComm.Parameters.Add("@nIDComAct", SqlDbType.Int).Value = oComercio.nIDCom

                oComm.ExecuteNonQuery()

                nResp = oComm.Parameters("@pnResp").Value
                sRespuesta = oComm.Parameters("@psRespuesta").Value

                If nResp = 0 Then
                    D(0) = Str(nResp)
                    D(1) = sRespuesta

                    R.nResultado = CInt(D(0))
                    R.sRespuesta = D(1)

                    Return R

                End If

                nIdCom = oComm.Parameters("@nIdCom").Value

                If (Not IsNothing(oComercio.oImagenes)) Then
                    For i = 0 To oComercio.oImagenes.Count - 1

                        oComm = New SqlClient.SqlCommand("PA_RegistraComercioFoto", oConex)
                        oComm.CommandTimeout = 0
                        oComm.CommandType = CommandType.StoredProcedure
                        oComm.Parameters.Add("@nIDCom", SqlDbType.Int).Value = nIdCom
                        oComm.Parameters.Add("@nFoto", SqlDbType.Image).Value = oComercio.oImagenes(i).nImg

                        oComm.ExecuteNonQuery()
                    Next
                End If


            Else

                oComm = New SqlClient.SqlCommand("PA_RegistraNoInteresado", oConex)
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.Parameters.Add("@cNomComercio", SqlDbType.VarChar, 150).Value = oComercio.cRazonSocial
                oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = sCodUsu
                oComm.Parameters.Add("@nIDNointeresado", SqlDbType.Int).Direction = ParameterDirection.Output

                oComm.ExecuteNonQuery()

                nIdCom = oComm.Parameters("@nIDNointeresado").Value

                For i = 0 To oComercio.oImagenes.Count - 1

                    oComm = New SqlClient.SqlCommand("PA_RegistraFotoNoInteresado", oConex)
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.Parameters.Add("@nIDNointeresado", SqlDbType.Int).Value = nIdCom
                    oComm.Parameters.Add("@nFoto", SqlDbType.Image).Value = oComercio.oImagenes(i).nImg

                    oComm.ExecuteNonQuery()


                Next

                nResp = 1
                sRespuesta = ""

            End If

            oConex.Close()

            D(0) = Str(nResp)
            D(1) = sRespuesta

            'PA_RegistraComercioFoto

        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try

        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R


    End Function


    Public Function RegistraNogestion() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim i As Integer
        Dim nIdCom As Long

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_RegistraNoGestion", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nIdOrden", SqlDbType.Int).Value = oNoGestion.nIdOrden
            oComm.Parameters.Add("@nCodMotivo", SqlDbType.Int).Value = oNoGestion.nCodMotivo
            oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = oNoGestion.cUser
            oComm.Parameters.Add("@nID", SqlDbType.Int).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            nIdCom = oComm.Parameters("@nID").Value


            oComm = New SqlClient.SqlCommand("PA_RegistraNoGestionFoto", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nID", SqlDbType.Int).Value = nIdCom
            oComm.Parameters.Add("@nFoto", SqlDbType.Image).Value = oNoGestion.nImg

            oComm.ExecuteNonQuery()




            oConex.Close()

            D(0) = "1"
            D(1) = "Registro Correcto"



        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try

        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R


    End Function



End Class
