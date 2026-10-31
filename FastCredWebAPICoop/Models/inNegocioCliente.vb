Imports System.Data.SqlClient

Public Class inNegocioCliente

    Public Property sToken As String
    Public Property nSolicitud As Integer
    Public Property sClave As String
    Public Property sCelular As String
    Public Property sCodUsu As String
    Public Property nCodigoCatalogo As Integer
    Public Property oTraza As clsTraza
    Public Property oPersona As clsPersona
    Public Property sCodSMS As String
    Public Property nLatitud As Double
    Public Property nLongitud As Double
    Public Property nCodCred As Integer
    Public Property sChequeo As String
    Public Property oSolicitudApp As clsSolicitudApp

    Public Function PuntoDeChequeo() As outNegocioCliente
        Dim R As New outNegocioCliente
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


            oComm = New SqlClient.SqlCommand("PA_RegistraVitAppCliente", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@ccodusu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@cComent", SqlDbType.VarChar, 500).Value = sChequeo

            oComm.ExecuteNonQuery()


            oConex.Close()

            If oComm.Parameters("@nResultado").Value = 1 Then
                D(0) = "1"
                D(1) = "OK"
            Else
                D(0) = "0"
                D(1) = ""

            End If


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function


    Public Function ValidaZona() As outNegocioCliente
        Dim R As New outNegocioCliente
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
            oComm.Parameters.Add("@nLat", SqlDbType.Float).Value = nLatitud
            oComm.Parameters.Add("@nlong", SqlDbType.Float).Value = nLongitud
            oComm.Parameters.Add("@nResultado", SqlDbType.Int).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()


            oConex.Close()

            If oComm.Parameters("@nResultado").Value = 1 Then
                D(0) = "1"
                D(1) = "Zona Correcta"
            Else
                D(0) = "0"
                D(1) = "Por el momento no atendemos en la Zona donde usted se ubica o Debe aceptar el uso de su localización para poder validar su Zona"

            End If


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function

    Public Function RegistraSolicitudNuevaRapida() As outNegocioCliente
        Dim R As New outNegocioCliente
        Dim oConex As SqlConnection
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim D() As String
        Dim sCadConexion As String
        Dim oConect As New clsConecta
        Dim nCodEvalMax As Long
        Dim nCodEval As Integer

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing



        'Siempre es Nuevo
        nCodEval = 0


        'Registra Solicitud Pyme
        oComm = New SqlClient.SqlCommand("PA_RegistraCreditoPorProducto_AppCliente", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = oSolicitudApp.cDoc
        oComm.Parameters.Add("@nPrestamo", SqlDbType.Money).Value = oSolicitudApp.nPrestamo
        oComm.Parameters.Add("@nNumCuotas", SqlDbType.Int).Value = oSolicitudApp.nNumCuotas
        oComm.Parameters.Add("@cCodUsuAna", SqlDbType.VarChar, 50).Value = "APP"
        oComm.Parameters.Add("@nMontoCuota", SqlDbType.Money).Value = oSolicitudApp.nMontoCuota
        oComm.Parameters.Add("@nTipoPlazo", SqlDbType.VarChar, 50).Value = "1"
        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = "ADMIN"
        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@psProd", SqlDbType.VarChar, 50).Value = Str(oSolicitudApp.nProd)
        oComm.Parameters.Add("@pdFecgaPago", SqlDbType.Date).Value = Now()
        oComm.Parameters.Add("@nIdSolicitud", SqlDbType.Int).Value = 0
        oComm.Parameters.Add("@nCampCuotasSinInt", SqlDbType.Int).Value = 0

        oComm.ExecuteNonQuery()

        ReDim Preserve D(2)

        D(0) = oComm.Parameters("@nResp").Value
        D(1) = oComm.Parameters("@Resultado").Value



        If D(0) = "0" Then
            oConex.Close()
            R.nResultado = 0
            R.sRespuesta = D(1)
            Return R

        End If


        oConex.Close()

        R.sRespuesta = "Registro Correcto"
        R.nResultado = 1

        Return R


    End Function

    Public Function RegistraSolicitudNueva() As outNegocioCliente
        Dim R As New outNegocioCliente
        Dim oConex As SqlConnection
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim D() As String
        Dim sCadConexion As String
        Dim oConect As New clsConecta
        Dim nCodEvalMax As Long
        Dim nCodEval As Integer

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing



        'Siempre es Nuevo
        nCodEval = 0

        oComm = New SqlClient.SqlCommand("PA_Movil2_RegistraNegocio_2", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = oSolicitudApp.cDoc
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@psNuevo", SqlDbType.VarChar, 50).Value = "1"
        oComm.Parameters.Add("@cRazonSocial", SqlDbType.VarChar, 150).Value = oSolicitudApp.cRazonSocial
        oComm.Parameters.Add("@cDirecNeg", SqlDbType.VarChar, 150).Value = oSolicitudApp.cDirecNeg
        oComm.Parameters.Add("@cCodZonaNeg", SqlDbType.VarChar, 50).Value = oSolicitudApp.cCodZona
        oComm.Parameters.Add("@cCodActEcon", SqlDbType.VarChar, 50).Value = oSolicitudApp.cCodActEcon
        oComm.Parameters.Add("@dFechaIniNeg", SqlDbType.DateTime).Value = oSolicitudApp.dFechaNeg
        oComm.Parameters.Add("@nNumTrab", SqlDbType.Int).Value = 0
        oComm.Parameters.Add("@nGastosFam", SqlDbType.Money).Value = 0
        oComm.Parameters.Add("@nGastosMen", SqlDbType.Money).Value = oSolicitudApp.nMontoVentasDiarias
        oComm.Parameters.Add("@cReferNeg", SqlDbType.VarChar, 150).Value = oSolicitudApp.cRefDirecNeg
        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = "APP"
        oComm.Parameters.Add("@nLatitud", SqlDbType.Float).Value = oSolicitudApp.nLatitud
        oComm.Parameters.Add("@nLong", SqlDbType.Float).Value = oSolicitudApp.nLongitud

        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nCodEvalMax", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()

        ReDim Preserve D(2)

        D(0) = oComm.Parameters("@nResp").Value
        D(1) = oComm.Parameters("@Resultado").Value
        nCodEvalMax = oComm.Parameters("@nCodEvalMax").Value



        'Registra Fotos
        If D(0) <> "1" Then
            oConex.Close()
            R.nResultado = 0
            R.sRespuesta = D(1)
            Return R
        End If

        Dim i As Integer

        For i = 0 To oSolicitudApp.oFotos.Count - 1
            D = GuardarFotoNegocio(nCodEvalMax, i + 1, oSolicitudApp.oFotos(i).nFoto)
            If D(0) = "0" Then
                R.nResultado = 0
                R.sRespuesta = D(1)
                Return R

            End If
        Next

        'Registra Solicitud Pyme
        oComm = New SqlClient.SqlCommand("PA_RegistraCreditoPorProducto_AppCliente", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = oSolicitudApp.cDoc
        oComm.Parameters.Add("@nPrestamo", SqlDbType.Money).Value = oSolicitudApp.nPrestamo
        oComm.Parameters.Add("@nNumCuotas", SqlDbType.Int).Value = oSolicitudApp.nNumCuotas
        oComm.Parameters.Add("@cCodUsuAna", SqlDbType.VarChar, 50).Value = "APP"
        oComm.Parameters.Add("@nMontoCuota", SqlDbType.Money).Value = oSolicitudApp.nMontoCuota
        oComm.Parameters.Add("@nTipoPlazo", SqlDbType.VarChar, 50).Value = "1"
        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = "ADMIN"
        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@psProd", SqlDbType.VarChar, 50).Value = Str(oSolicitudApp.nProd)
        oComm.Parameters.Add("@pdFecgaPago", SqlDbType.Date).Value = Now()
        oComm.Parameters.Add("@nIdSolicitud", SqlDbType.Int).Value = 0
        oComm.Parameters.Add("@nCampCuotasSinInt", SqlDbType.Int).Value = 0

        oComm.ExecuteNonQuery()

        ReDim Preserve D(2)

        D(0) = oComm.Parameters("@nResp").Value
        D(1) = oComm.Parameters("@Resultado").Value



        If D(0) = "0" Then
            oConex.Close()
            R.nResultado = 0
            R.sRespuesta = D(1)
            Return R

        End If

        'Actualizar los Datos de la Persona

        oComm = New SqlClient.SqlCommand("PA_ActualizaDatosPersona_AppCliente", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = oSolicitudApp.cDoc
        oComm.Parameters.Add("@cDireccionDom", SqlDbType.VarChar, 150).Value = oSolicitudApp.cDirecDom
        oComm.Parameters.Add("@cCodZonaDom", SqlDbType.VarChar, 50).Value = oSolicitudApp.cZonaDom
        oComm.Parameters.Add("@cReferenciaDom", SqlDbType.VarChar, 150).Value = oSolicitudApp.cRefDirDom

        oComm.ExecuteNonQuery()

        oConex.Close()

        R.sRespuesta = "Registro Correcto"
        R.nResultado = 1

        Return R


    End Function

    Public Function GuardarFotoNegocio(ByVal nCodEvalMax As Long, ByVal nTipoFoto As Integer, ByVal nImg As Byte()) As String()
        'PA_RegistraImagen
        Dim oConex As SqlConnection
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim D() As String
        Dim sCadConexion As String
        Dim oConect As New clsConecta

        sCadConexion = oConect.gsCadConex


        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        oComm = New SqlClient.SqlCommand("PA_RegistraFotoNegocio_2", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodEval", SqlDbType.Int).Value = nCodEvalMax
        oComm.Parameters.Add("@iDoc", SqlDbType.Image).Value = nImg
        oComm.Parameters.Add("@nTipoFoto", SqlDbType.Int).Value = nTipoFoto
        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output


        oComm.ExecuteNonQuery()


        ReDim Preserve D(2)

        D(0) = oComm.Parameters("@nResp").Value
        D(1) = oComm.Parameters("@Resultado").Value

        oConex.Close()

        Return D

    End Function

    Public Function RecuperaPrecalifiNew() As outNegocioCliente

        Dim R As New outNegocioCliente
        Dim i, k As Integer
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Try

            oComm = New SqlClient.SqlCommand("PA_RecuperaPrecalifiNew_2", oConex)
            oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@psCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu

            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            R.oPreCalif = New clsPrecalif
            While dr.Read

                R.oPreCalif.cCliente = dr.Item("cCliente")
                R.oPreCalif.nCantEnt = dr.Item("nCantEnt")
                R.oPreCalif.cCalif = dr.Item("cCalif")
                R.oPreCalif.ResultadoSemanal = dr.Item("ResultadoSemanal")
                R.oPreCalif.nRespSemanal = dr.Item("nRespSemanal")
                R.oPreCalif.ResultadoJunior = dr.Item("ResultadoJunior")
                R.oPreCalif.nRespJunior = dr.Item("nRespJunior")
                R.oPreCalif.ResultadoMicro = Replace(dr.Item("ResultadoMicro"), "RECHAZADO", "LO SENTIMOS NO PODEMOS ATENDERTE")
                R.oPreCalif.nRespMicro = dr.Item("nRespMicro")
                R.oPreCalif.ResultadoCredParalelo = dr.Item("ResultadoParalelo")
                R.oPreCalif.nRespParalelo = dr.Item("nRespParalelo")

                i = i + 1

            End While
            dr.Close()
            dr = Nothing

            oConex.Close()

            D(0) = "1"
            D(1) = "Registro Correcto"

        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message

        End Try


        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))

        Return R

    End Function


    Public Function CliRecuperaDatosPersona() As outNegocioCliente

        Dim R As New outNegocioCliente
        Dim i, k As Integer
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Try

            oComm = New SqlClient.SqlCommand("PA_CliRecuperaDatosPersona", oConex)
            oComm.Parameters.Add("@sDoc", SqlDbType.VarChar, 50).Value = sCodUsu

            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                R.oclsCliente = New clsCliente

                R.oclsCliente.cDoc = dr.Item("cDoc")
                R.oclsCliente.dFechaNac = dr.Item("dFechaNac")
                R.oclsCliente.dFechareg = dr.Item("dFechareg")
                R.oclsCliente.cApellidos = dr.Item("cApellidos")
                R.oclsCliente.cNombres = dr.Item("cNombres")
                R.oclsCliente.cCodGenero = dr.Item("cCodGenero")
                R.oclsCliente.cMovil1 = dr.Item("cMovil1")
                R.oclsCliente.cDirecDom = dr.Item("cDirecDom")
                R.oclsCliente.cDistritoDom = dr.Item("cDistritoDom")
                R.oclsCliente.cZonaDom = dr.Item("cZonaDom")
                R.oclsCliente.cRefDirDom = dr.Item("cRefDirDom")
                R.oclsCliente.cNomBanco = dr.Item("cNomBanco")
                R.oclsCliente.cNroCtaAho = dr.Item("cNroCtaAho")
                R.oclsCliente.cCategoria = dr.Item("cCategoria")
                R.oclsCliente.nCodAge = dr.Item("nCodAge")
                R.oclsCliente.nNumCiclos = dr.Item("nNumCiclos")

                i = i + 1

            End While
            dr.Close()
            dr = Nothing

            oComm = New SqlClient.SqlCommand("PA_CliRecuperaCreditos", oConex)
            oComm.Parameters.Add("@sDoc", SqlDbType.VarChar, 50).Value = sCodUsu

            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oclsCliente.oCreditos(i)
                R.oclsCliente.oCreditos(i) = New clsCreditoData

                R.oclsCliente.oCreditos(i).nCodAge = dr.Item("nCodAge")
                R.oclsCliente.oCreditos(i).nCodCred = dr.Item("nCodCred")
                R.oclsCliente.oCreditos(i).nCodEval = dr.Item("nCodEval")
                R.oclsCliente.oCreditos(i).nEstado = dr.Item("nEstado")
                R.oclsCliente.oCreditos(i).nSaldok = dr.Item("nSaldok")
                R.oclsCliente.oCreditos(i).dFecVig = dr.Item("dFecVig")
                R.oclsCliente.oCreditos(i).nPrestamo = dr.Item("nPrestamo")
                R.oclsCliente.oCreditos(i).nNumCuotas = dr.Item("nNumCuotas")
                R.oclsCliente.oCreditos(i).nMontoCuota = dr.Item("nMontoCuota")
                R.oclsCliente.oCreditos(i).nTasa = dr.Item("nTasa")
                R.oclsCliente.oCreditos(i).nProd = dr.Item("nProd")
                R.oclsCliente.oCreditos(i).nAmpliado = dr.Item("nAmpliado")
                R.oclsCliente.oCreditos(i).nDiasAtraso = dr.Item("nDiasAtraso")
                R.oclsCliente.oCreditos(i).cAmpliacion = dr.Item("cAmpliacion")
                R.oclsCliente.oCreditos(i).nDeudaAFecha = dr.Item("nDeudaAFecha")
                R.oclsCliente.oCreditos(i).nEfectivoDisponible = dr.Item("nEfectivoDisponible")

                i = i + 1

            End While
            dr.Close()
            dr = Nothing


            oComm = New SqlClient.SqlCommand("PA_CliRecuperaNegocio", oConex)
            oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = sCodUsu

            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                R.oclsCliente.oNegocio = New clsNegocio

                R.oclsCliente.oNegocio.cDoc = dr.Item("cDoc")
                R.oclsCliente.oNegocio.nCodEval = dr.Item("nCodEval")
                R.oclsCliente.oNegocio.dFechareg = dr.Item("dFechareg")
                R.oclsCliente.oNegocio.cRazonSocial = dr.Item("cRazonSocial")
                R.oclsCliente.oNegocio.cDirecNeg = dr.Item("cDirecNeg")
                R.oclsCliente.oNegocio.cRefDirecNeg = dr.Item("cRefDirecNeg")
                R.oclsCliente.oNegocio.cCodDist = dr.Item("cCodDist")
                R.oclsCliente.oNegocio.cCodZona = dr.Item("cCodZona")
                R.oclsCliente.oNegocio.cCodActEcon = dr.Item("cCodActEcon")
                R.oclsCliente.oNegocio.nMontoVentasDiarias = dr.Item("nMontoVentasDiarias")
                R.oclsCliente.oNegocio.nLatitud = dr.Item("nLatitud")
                R.oclsCliente.oNegocio.nLongitud = dr.Item("nLongitud")

                i = i + 1

            End While
            dr.Close()
            dr = Nothing

            If Not IsNothing(R.oclsCliente.oCreditos) Then
                For k = 0 To R.oclsCliente.oCreditos.Count - 1

                    '*************************************
                    'Carga kardex
                    '*************************************
                    oComm = New SqlClient.SqlCommand("PA_CliRecuperaKardex", oConex)
                    oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = R.oclsCliente.oCreditos(k).nCodCred

                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()

                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0

                    While dr.Read
                        ReDim Preserve R.oclsCliente.oCreditos(k).oKardexPago(i)
                        R.oclsCliente.oCreditos(k).oKardexPago(i) = New clsKardexPago

                        R.oclsCliente.oCreditos(k).oKardexPago(i).dFecha = dr.Item("dFecha")
                        R.oclsCliente.oCreditos(k).oKardexPago(i).nCodAge = dr.Item("nCodAge")
                        R.oclsCliente.oCreditos(k).oKardexPago(i).nCodCred = dr.Item("nCodCred")
                        R.oclsCliente.oCreditos(k).oKardexPago(i).cNomOpe = dr.Item("cNomOpe")
                        R.oclsCliente.oCreditos(k).oKardexPago(i).nMonto = dr.Item("nMonto")
                        R.oclsCliente.oCreditos(k).oKardexPago(i).nCapital = dr.Item("nCapital")
                        R.oclsCliente.oCreditos(k).oKardexPago(i).nInteres = dr.Item("nInteres")
                        R.oclsCliente.oCreditos(k).oKardexPago(i).nIgv = dr.Item("nIgv")
                        R.oclsCliente.oCreditos(k).oKardexPago(i).cUser = dr.Item("cUser")
                        R.oclsCliente.oCreditos(k).oKardexPago(i).nCodOpe = dr.Item("nCodOpe")
                        R.oclsCliente.oCreditos(k).oKardexPago(i).nDiasAtraso = dr.Item("nDiasAtraso")
                        R.oclsCliente.oCreditos(k).oKardexPago(i).nNroTrans = dr.Item("nNroTrans")

                        i = i + 1

                    End While
                    dr.Close()
                    dr = Nothing


                    '*************************************
                    'Carga Credcalendario
                    '*************************************
                    oComm = New SqlClient.SqlCommand("PA_CliRecuperaCalendario", oConex)
                    oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = R.oclsCliente.oCreditos(k).nCodCred

                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()

                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0

                    While dr.Read
                        ReDim Preserve R.oclsCliente.oCreditos(k).oCalendario(i)
                        R.oclsCliente.oCreditos(k).oCalendario(i) = New clsCalendarioCli

                        R.oclsCliente.oCreditos(k).oCalendario(i).nCodAge = dr.Item("nCodAge")
                        R.oclsCliente.oCreditos(k).oCalendario(i).nCodCred = dr.Item("nCodCred")
                        R.oclsCliente.oCreditos(k).oCalendario(i).nNroCuota = dr.Item("nNroCuota")
                        R.oclsCliente.oCreditos(k).oCalendario(i).dFechaVenc = dr.Item("dFechaVenc")
                        If IsDBNull(dr.Item("dFechaPago")) Then
                            R.oclsCliente.oCreditos(k).oCalendario(i).dFechaPago = CDate("01/01/1900")
                        Else
                            R.oclsCliente.oCreditos(k).oCalendario(i).dFechaPago = dr.Item("dFechaPago")
                        End If

                        R.oclsCliente.oCreditos(k).oCalendario(i).nEstado = dr.Item("nEstado")
                        R.oclsCliente.oCreditos(k).oCalendario(i).nCapital = dr.Item("nCapital")
                        R.oclsCliente.oCreditos(k).oCalendario(i).nCapPag = dr.Item("nCapPag")
                        R.oclsCliente.oCreditos(k).oCalendario(i).nInteres = dr.Item("nInteres")
                        R.oclsCliente.oCreditos(k).oCalendario(i).nIntPag = dr.Item("nIntPag")
                        R.oclsCliente.oCreditos(k).oCalendario(i).nIgv = dr.Item("nIgv")
                        R.oclsCliente.oCreditos(k).oCalendario(i).nIgvPag = dr.Item("nIgvPag")

                        i = i + 1

                    End While
                    dr.Close()
                    dr = Nothing

                Next
            End If

            oConex.Close()

            D(0) = "1"
            D(1) = "Registro Correcto"

        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message

        End Try


        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))

        Return R

    End Function

    Public Function RecuperaCuotasCliente() As outNegocioCliente
        Dim R As New outNegocioCliente
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nResp As Integer
        Dim sRespuesta As String
        Dim dr As SqlClient.SqlDataReader

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

            oComm = New SqlClient.SqlCommand("PA_RecuperaCuotasCliente", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = nCodCred

            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

            Dim i As Integer = 0

            While dr.Read

                ReDim Preserve R.oCuotas(i)
                R.oCuotas(i) = New clsCuotasCliente()
                R.oCuotas(i).nNroCuota = dr.Item("nNroCuota")
                R.oCuotas(i).dFecVenc = dr.Item("dFecVenc")
                R.oCuotas(i).nMonto = dr.Item("nMonto")

                i = i + 1
            End While

            dr.Close()
            dr = Nothing

            nResp = 1
            sRespuesta = "PROCESO CORRECTO"

            oConex.Close()

            D(0) = Str(nResp)
            D(1) = sRespuesta

        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try

        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R


    End Function


    Public Function RecuperaPagosCliente() As outNegocioCliente
        Dim R As New outNegocioCliente
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nResp As Integer
        Dim sRespuesta As String
        Dim dr As SqlClient.SqlDataReader

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

            oComm = New SqlClient.SqlCommand("PA_RecuperaPagosCliente", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = nCodCred

            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

            Dim i As Integer = 0

            While dr.Read

                ReDim Preserve R.oPagos(i)
                R.oPagos(i) = New clsPagosCliente()
                R.oPagos(i).cNomOpe = dr.Item("cCodOpe")
                R.oPagos(i).nNrotrans = dr.Item("nNroTrans")
                R.oPagos(i).nCodcred = dr.Item("nCodCred")
                R.oPagos(i).dFecPago = dr.Item("dFecha")
                R.oPagos(i).nMontoPago = dr.Item("nMonto")
                R.oPagos(i).nDiasAtraso = dr.Item("nDiasAtraso")

                i = i + 1
            End While

            dr.Close()
            dr = Nothing

            nResp = 1
            sRespuesta = "PROCESO CORRECTO"

            oConex.Close()

            D(0) = Str(nResp)
            D(1) = sRespuesta

        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try

        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R


    End Function

    Public Function RecuperaCreditos() As outNegocioCliente
        Dim R As New outNegocioCliente
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nResp As Integer
        Dim sRespuesta As String
        Dim dr As SqlClient.SqlDataReader

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

            oComm = New SqlClient.SqlCommand("PA_AppCli_RecuperaCreditoVigentes", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@psDNI", SqlDbType.VarChar, 50).Value = sCodUsu

            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

            Dim i As Integer = 0

            While dr.Read
                ReDim Preserve R.oCreditos(i)
                R.oCreditos(i) = New clsCredito()
                R.oCreditos(i).nCodAge = dr.Item("nCodAge")
                R.oCreditos(i).nCodCred = dr.Item("nCodCred")
                R.oCreditos(i).dFecVig = dr.Item("dFecVig")
                R.oCreditos(i).nPrestamo = dr.Item("nPrestamo")
                R.oCreditos(i).nSaldoK = dr.Item("nSaldoK")
                R.oCreditos(i).nCuotas = dr.Item("nCuotas")
                R.oCreditos(i).cPlazo = dr.Item("cPlazo")
                R.oCreditos(i).nCuotasPagadas = dr.Item("nCuotasPagadas")
                R.oCreditos(i).nCuotasPendientes = dr.Item("nCuotasPendientes")
                R.oCreditos(i).dFecVenc = dr.Item("dFecVenc")
                R.oCreditos(i).nMontoPago = dr.Item("nMontoPago")
                R.oCreditos(i).nDiasAtraso = dr.Item("nDiasAtraso")
                R.oCreditos(i).nDeudaFecha = dr.Item("nDeudaFecha")
                R.oCreditos(i).nEfectivoDisponible = dr.Item("nEfectivoDisponible")
                R.oCreditos(i).nEstado = dr.Item("nEstado")


                i = i + 1
            End While

            dr.Close()
            dr = Nothing

            nResp = 1
            sRespuesta = "PROCESO CORRECTO"

            oConex.Close()

            D(0) = Str(nResp)
            D(1) = sRespuesta

        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try

        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R


    End Function

    Public Function RegistraUsuarioNuevo() As outNegocioCliente
        Dim R As New outNegocioCliente
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nResp As Integer
        Dim sRespuesta As String

        Dim sMensaje As String
        Dim sCodigo As String

        nResp = 0
        sRespuesta = ""
        sCodigo = GeneraCodigo()
        sMensaje = sCodigo

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try

            oComm = New SqlClient.SqlCommand("pa_usuarionuevo", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@psDNI", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@psCodigoSMS", SqlDbType.VarChar, 50).Value = sCodSMS
            oComm.Parameters.Add("@psClave", SqlDbType.VarChar, 50).Value = sClave
            oComm.Parameters.Add("@pnResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@psRespuesta", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@cNombres", SqlDbType.VarChar, 100).Value = oPersona.cNombres
            oComm.Parameters.Add("@cApePat", SqlDbType.VarChar, 100).Value = oPersona.cApePat
            oComm.Parameters.Add("@cApeMat", SqlDbType.VarChar, 100).Value = oPersona.cApeMat
            oComm.Parameters.Add("@dFechaNac", SqlDbType.DateTime).Value = oPersona.dFechaNac
            oComm.Parameters.Add("@cGenero", SqlDbType.VarChar, 50).Value = oPersona.cCodGenero
            oComm.Parameters.Add("@cMovil1", SqlDbType.VarChar, 50).Value = oPersona.cMovil1
            oComm.Parameters.Add("@nLatitud", SqlDbType.Float).Value = nLatitud
            oComm.Parameters.Add("@nLongitud", SqlDbType.Float).Value = nLongitud

            oComm.ExecuteNonQuery()

            oConex.Close()

            D(0) = CInt(oComm.Parameters("@pnResp").Value)
            D(1) = oComm.Parameters("@psRespuesta").Value



        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try

        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R


    End Function

    Public Function RegistraUsuarioExistente() As outNegocioCliente
        Dim R As New outNegocioCliente
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nResp As Integer
        Dim sRespuesta As String

        Dim sMensaje As String
        Dim sCodigo As String

        nResp = 0
        sRespuesta = ""
        sCodigo = GeneraCodigo()
        sMensaje = sCodigo

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try

            oComm = New SqlClient.SqlCommand("usuarioexistente", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@psDNI", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@psCodigoSMS", SqlDbType.VarChar, 50).Value = sCodSMS
            oComm.Parameters.Add("@psClave", SqlDbType.VarChar, 50).Value = sClave
            oComm.Parameters.Add("@pnResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@psRespuesta", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            oConex.Close()

            D(0) = CInt(oComm.Parameters("@pnResp").Value)
            D(1) = oComm.Parameters("@psRespuesta").Value



        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try

        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R


    End Function

    Public Function RecuperaZonas() As outNegocioCliente
        Dim R As New outNegocioCliente
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nResp As Integer
        Dim sRespuesta As String
        Dim dr As SqlClient.SqlDataReader

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

            oComm = New SqlClient.SqlCommand("PA_RecuperaTodasZonas", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

            Dim i As Integer = 0

            While dr.Read
                ReDim Preserve R.oZonas(i)
                R.oZonas(i) = New clsZonas()
                R.oZonas(i).cCodZona = dr.Item("cCodZona")
                R.oZonas(i).cNomZona = dr.Item("cNomZona")

                i = i + 1
            End While

            dr.Close()
            dr = Nothing

            nResp = 1
            sRespuesta = "PROCESO CORRECTO"

            oConex.Close()

            D(0) = Str(nResp)
            D(1) = sRespuesta

        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try

        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R


    End Function

    Public Function RecuperaCatalogo() As outNegocioCliente
        Dim R As New outNegocioCliente
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nResp As Integer
        Dim sRespuesta As String
        Dim dr As SqlClient.SqlDataReader

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

            oComm = New SqlClient.SqlCommand("PA_RecuperaTodoCatalogo", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

            Dim i As Integer = 0

            While dr.Read
                ReDim Preserve R.oCatalogo(i)
                R.oCatalogo(i) = New clsCatalogocodigos()
                R.oCatalogo(i).nCodigo = dr.Item("nCodigo")
                R.oCatalogo(i).cDescripcion = dr.Item("cDescripcion")
                R.oCatalogo(i).cValor = dr.Item("cValor")

                i = i + 1
            End While

            dr.Close()
            dr = Nothing



            oComm = New SqlClient.SqlCommand("PA_RecuperaTodasZonasCliente", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read
                ReDim Preserve R.oZonas(i)
                R.oZonas(i) = New clsZonas()
                R.oZonas(i).cCodZona = dr.Item("cCodZona")
                R.oZonas(i).cNomZona = dr.Item("cNomZona")

                i = i + 1
            End While

            dr.Close()
            dr = Nothing


            oComm = New SqlClient.SqlCommand("PA_CliRecuperaTarifarioMicro", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read
                ReDim Preserve R.oCliTarifarioMicro(i)
                R.oCliTarifarioMicro(i) = New clsCliTarifarioMicro()
                R.oCliTarifarioMicro(i).cCodCat = dr.Item("cCodCat")
                R.oCliTarifarioMicro(i).cDescripcion = dr.Item("cDescripcion")
                R.oCliTarifarioMicro(i).nCodAge = dr.Item("nCodAge")
                R.oCliTarifarioMicro(i).nCuotasMax = dr.Item("nCuotasMax")
                R.oCliTarifarioMicro(i).nCuotasMin = dr.Item("nCuotasMin")
                R.oCliTarifarioMicro(i).nPrestamoFin = dr.Item("nPrestamoFin")
                R.oCliTarifarioMicro(i).nPretamoIni = dr.Item("nPretamoIni")
                R.oCliTarifarioMicro(i).nTasaComp = dr.Item("nTasaComp")
                R.oCliTarifarioMicro(i).nTasaCompNeto = dr.Item("nTasaCompNeto")


                i = i + 1
            End While

            dr.Close()
            dr = Nothing



            oComm = New SqlClient.SqlCommand("PA_CliRecuperaTarifarioSemanal", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read
                ReDim Preserve R.oCliTarifarioSemanal(i)
                R.oCliTarifarioSemanal(i) = New clsCliTarifarioSemanal()
                R.oCliTarifarioSemanal(i).cCodCat = dr.Item("cCodCat")
                R.oCliTarifarioSemanal(i).cDescripcion = dr.Item("cDescripcion")
                R.oCliTarifarioSemanal(i).nCodAge = dr.Item("nCodAge")
                R.oCliTarifarioSemanal(i).nCuotasMax = dr.Item("nCuotasMax")
                R.oCliTarifarioSemanal(i).nCuotasMin = dr.Item("nCuotasMin")
                R.oCliTarifarioSemanal(i).nPrestamoFin = dr.Item("nPrestamoFin")
                R.oCliTarifarioSemanal(i).nPretamoIni = dr.Item("nPretamoIni")
                R.oCliTarifarioSemanal(i).nTasaComp = dr.Item("nTasaComp")
                R.oCliTarifarioSemanal(i).nTasaCompNeto = dr.Item("nTasaCompNeto")


                i = i + 1
            End While

            dr.Close()
            dr = Nothing



            oComm = New SqlClient.SqlCommand("PA_CliRecuperaTarifarioSemanalJunior", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read
                ReDim Preserve R.oCliTarifarioSemanalJunior(i)
                R.oCliTarifarioSemanalJunior(i) = New clsCliTarifarioSemanalJunior()
                R.oCliTarifarioSemanalJunior(i).cCodCat = dr.Item("cCodCat")
                R.oCliTarifarioSemanalJunior(i).cDescripcion = dr.Item("cDescripcion")
                R.oCliTarifarioSemanalJunior(i).nCodAge = dr.Item("nCodAge")
                R.oCliTarifarioSemanalJunior(i).nCuotasMax = dr.Item("nCuotasMax")
                R.oCliTarifarioSemanalJunior(i).nCuotasMin = dr.Item("nCuotasMin")
                R.oCliTarifarioSemanalJunior(i).nPrestamoFin = dr.Item("nPrestamoFin")
                R.oCliTarifarioSemanalJunior(i).nPretamoIni = dr.Item("nPretamoIni")
                R.oCliTarifarioSemanalJunior(i).nTasaComp = dr.Item("nTasaComp")
                R.oCliTarifarioSemanalJunior(i).nTasaCompNeto = dr.Item("nTasaCompNeto")


                i = i + 1
            End While

            dr.Close()
            dr = Nothing



            nResp = 1
            sRespuesta = "PROCESO CORRECTO"

            oConex.Close()

            D(0) = Str(nResp)
            D(1) = sRespuesta

        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try

        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R


    End Function

    Public Function RegistraUsuario() As outNegocioCliente
        Dim R As New outNegocioCliente
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nResp As Integer
        Dim sRespuesta As String

        Dim sMensaje As String
        Dim sCodigo As String

        nResp = 0
        sRespuesta = ""
        sCodigo = GeneraCodigo()
        sMensaje = sCodigo

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try

            oComm = New SqlClient.SqlCommand("PA_EnviaSMS", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@psCelular", SqlDbType.VarChar, 50).Value = sCelular
            oComm.Parameters.Add("@psMensaje", SqlDbType.VarChar, 150).Value = sMensaje

            oComm.ExecuteNonQuery()


            oComm = New SqlClient.SqlCommand("PA_RegistraEnvioSMS", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@cMovil", SqlDbType.VarChar, 50).Value = sCelular
            oComm.Parameters.Add("@cSMS", SqlDbType.VarChar, 50).Value = sCodigo


            oComm.ExecuteNonQuery()

            oConex.Close()

            D(0) = "1"
            D(1) = "OK"



        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try

        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R


    End Function

    Public Function ValidaClaveCliente() As outNegocioCliente
        Dim R As New outNegocioCliente
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nResp As Integer
        Dim sRespuesta As String
        Dim dr As SqlClient.SqlDataReader

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


            oComm = New SqlClient.SqlCommand("PA_RegistraVitAppCliente", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@ccodusu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@cComent", SqlDbType.VarChar, 500).Value = "LOGIN"

            oComm.ExecuteNonQuery()


            oComm = New SqlClient.SqlCommand("PA_ValidaClaveClienteApp", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@psCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@psClave", SqlDbType.VarChar, 50).Value = sClave
            oComm.Parameters.Add("@pnResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@psRespuesta", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            nResp = oComm.Parameters("@pnResp").Value
            sRespuesta = oComm.Parameters("@psRespuesta").Value

            oConex.Close()

            D(0) = Str(nResp)
            D(1) = sRespuesta

            R.nResultado = CInt(D(0))
            R.sRespuesta = D(1)

            If nResp = 1 Then

                R = CliRecuperaDatosPersona()

            End If

        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message
            R.nResultado = CInt(D(0))
            R.sRespuesta = D(1)

        End Try


        Return R


    End Function

    Public Function GeneraCodigo() As String
        Randomize()
        Dim value As Integer = CInt(Int((8999 * Rnd()) + 1000))
        GeneraCodigo = Str(value)

    End Function

    Public Function RegistraSMS() As outNegocioCliente
        Dim R As New outNegocioCliente
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nResp As Integer
        Dim sRespuesta As String

        Dim sMensaje As String
        Dim sCodigo As String

        nResp = 0
        sRespuesta = ""
        sCodigo = GeneraCodigo()
        sMensaje = sCodigo

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try

            oComm = New SqlClient.SqlCommand("PA_EnviaSMS", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@psCelular", SqlDbType.VarChar, 50).Value = sCelular
            oComm.Parameters.Add("@psMensaje", SqlDbType.VarChar, 150).Value = sMensaje

            oComm.ExecuteNonQuery()


            oComm = New SqlClient.SqlCommand("RegistraCodigoSMS", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@cCodVerif", SqlDbType.VarChar, 50).Value = sCodigo


            oComm.ExecuteNonQuery()

            oConex.Close()

            D(0) = "1"
            D(1) = "TE HEMOS ENVIADO TU CODIGO DE VERIFICACION A TU CELULAR"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try

        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R


    End Function

    Public Function RegistraSMSXDNI() As outNegocioCliente
        Dim R As New outNegocioCliente
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nResp As Integer
        Dim sRespuesta As String

        Dim sMensaje As String
        Dim sCodigo As String

        nResp = 0
        sRespuesta = ""
        sCodigo = GeneraCodigo()
        sMensaje = sCodigo

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try

            oComm = New SqlClient.SqlCommand("PA_EnviaSMSXDNI", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@psDNI", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@psMensaje", SqlDbType.VarChar, 150).Value = sMensaje
            oComm.Parameters.Add("@pnResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@psMensajeResp", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@psCelularNew", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            R.nResultado = oComm.Parameters("@pnResp").Value
            R.sRespuesta = oComm.Parameters("@psMensajeResp").Value

            If R.nResultado = 0 Then
                oConex.Close()
                Return R
            End If


            sCelular = oComm.Parameters("@psCelularNew").Value


            oComm = New SqlClient.SqlCommand("RegistraCodigoSMS", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@cCodVerif", SqlDbType.VarChar, 50).Value = sCodigo


            oComm.ExecuteNonQuery()

            oConex.Close()

            D(0) = "1"
            D(1) = "SE HA ENVIADO UN SMS CON EL CODIGO DE VERIFICACION"



        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try

        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R


    End Function

End Class
