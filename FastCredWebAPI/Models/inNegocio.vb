Imports System.Data
Imports System.Data.Entity
Imports System.Data.Entity.Core.Objects
Imports System.Data.Entity.Infrastructure

Imports System
Imports System.Collections.Generic
Imports FastCredWebAPI.FastCredWebAPI
Imports System.Data.SqlClient
Imports Newtonsoft.Json
Imports System.IO
Imports System.Net
Imports FastCredWebAPI.dDesembolso
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text.Json
Imports System.Web.Http.Results

Public Class inNegocio

    Private db As New DBNEGOCIOEntities
    Public Property sToken As String

    Public Property nSolicitud As Integer

    Public Property nCodCatalogo As Integer

    Public Property cDocPersona As String

    Public ValoresPantalla() As SistFormularios
    Public DatosPantalla As SistFormularios

    Public Property nCodAge As Integer
    Public Property sCodUsu As String
    Public Property nIdGrupo As Long
    Public Property oGrupoComunal As clsGrupoComunal
    Public Property oSocio As clsGrupoComunalSocio

    Public Property cDoc As String

    Public Property nLatitud As Decimal
    Public Property nLongitud As Decimal
    Public Property nCodcred As Integer

    Public Property TranPago As dTransacPago

    Public Property oSolicitudPyme As dSolPyme

    Public Property oVerificacion As dVerificacion

    Public Property oDesembolso As dDesembolso

    Public Property oProspeccionEmp As dProspeccionEmp

    Public Property AsignaAsesor As dAsignaAsesor

    Public Property oVoucher As clsvoucher

    Public Property ListaAbonos As dAbonoDesemb()

    Public Property nPrestamo As Double
    Public Property nCuotas As Integer

    Public Property nProd As Integer
    Public Property cCodCat As String

    Public Property oFirmaContrato As clsFirmaContrato
    Public Property oAsistencia As clsAsistencia
    Public Property oActLinea As clsActLinea

    Public Property nTasa As Double
    Public Property oAprobacion As clsAprobacionRegistro

    Public Property oMotivoNoPago As clsMotivoNoPago

    Public Property nCodRol As Integer
    Public Property oVerifCliente As clsVerificacionCliente()

    Public Property nCodMotivoNoDesembolso As Integer

    Public Property oBandejaAprobacion As clsAprobacionNueva()

    Public Property oBandejaDesembolsoNuevo As clsDesembolsoNuevo()

    Public Property cDireccionDomicilio As String
    Public Property cDireccionNegocio As String

    Public Property nCambiaDireccionDomicilio As Integer
    Public Property nCambiaDireccionNegocio As Integer

    Public Property nResetIMEI As Integer
    Public Property nAnulaCierre As Integer

    Public Property nCierraOperacionesUsu As Integer
    Public Property cMovil1 As String

    Public Property sClave As String
    Public Property sClaveNueva As String

    Public Property oScore As clsScore
    Public Property oListaPagoLoteCreditos As clsPagoLoteItem()

    Public Property oVerifReferencia As clsVerificacionReferencia()



    Public Property oTasacionRegistro As clsTasacionCliente
    Public Property oDataExperian As ResponseHistorial
    Public Property nImg As Byte()
    Public Property oDataEquifax As ResponseHistorialEquifax
    Public Property nCampSinInt As Integer

    Public Property nOperador As Integer
    Public Property cMovil As String
    Public Property sNroTran As String



    Public Function RecuperaEfectivoDisponible() As outNegocio

        Dim R As New outNegocio
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nResp As Integer
        Dim sRespuesta As String
        Dim dr As SqlClient.SqlDataReader
        Dim i, j As Integer

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


            oComm = New SqlClient.SqlCommand("PA_RecuperaEfectivoDisponible", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@nDisponible", SqlDbType.Money).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            R.nDisponible = oComm.Parameters("@nDisponible").Value

            nResp = 1
            sRespuesta = "OK"

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


    Public Function CargaUsuarios() As outNegocio

        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_RecuperaUsuariosAgencia_2", oConex)
            oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = nCodAge

            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oUsuarios(i)
                R.oUsuarios(i) = dr.Item("cCodUsu")

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

    Public Function RegistraEntregaEfectivo() As outNegocio

        Dim R As New outNegocio
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nResp As Integer
        Dim sRespuesta As String
        Dim dr As SqlClient.SqlDataReader
        Dim i, j As Integer

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


            oComm = New SqlClient.SqlCommand("PA_RegistraEntregaEfectivo", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@pnMonto", SqlDbType.Money).Value = nPrestamo
            oComm.Parameters.Add("@ncodage", SqlDbType.Int).Value = nCodAge
            oComm.Parameters.Add("@psUser", SqlDbType.VarChar, 50).Value = cDocPersona
            oComm.Parameters.Add("@psResponsable", SqlDbType.VarChar, 50).Value = sCodUsu

            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            nResp = oComm.Parameters("@nResp").Value
            sRespuesta = oComm.Parameters("@Resultado").Value

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



    Public Function RecuperaScore() As outNegocio

        Dim R As New outNegocio
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


            '*********************************************
            'RecuperaDatosCreditoPorDNI
            '*********************************************
            oComm = New SqlClient.SqlCommand("PA_RecuperaDatosUltCredXDNI", oConex)
            oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = cDoc
            oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Direction = ParameterDirection.Output

            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            nCodAge = oComm.Parameters("@pnCodAge").Value
            nCodcred = oComm.Parameters("@pnCodCred").Value

            If nCodcred = 0 Then
                R.sRespuesta = "No se Encontró Creditos"
                R.nResultado = 0
                Return R
            End If

            '*********************************************
            'Genera el Score
            '*********************************************
            oComm = New SqlClient.SqlCommand("PA_GeneraScore", oConex)
            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
            oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodcred

            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            '*********************************************
            'Recupera Detalle del Score
            '*********************************************
            oComm = New SqlClient.SqlCommand("PA_RecuperaDataScore", oConex)
            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
            oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodcred

            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            R.oScore = New clsScore
            ReDim Preserve R.oScore.oScoreDetalle(0)

            While dr.Read
                ReDim Preserve R.oScore.oScoreDetalle(i)
                R.oScore.oScoreDetalle(i) = New clsScoreDetalle

                R.oScore.oScoreDetalle(i).cPerfil = dr.Item("PERFIL")
                R.oScore.oScoreDetalle(i).cIndicador = dr.Item("INDICADOR")
                R.oScore.oScoreDetalle(i).cValor = dr.Item("VALOR")
                R.oScore.oScoreDetalle(i).cRango = dr.Item("RANGO")
                R.oScore.oScoreDetalle(i).cPuntaje = dr.Item("PUNTAJE")
                R.oScore.oScoreDetalle(i).cPeso = dr.Item("PESO")
                R.oScore.oScoreDetalle(i).cPuntajePonderado = Format(CDbl(dr.Item("PUNTAJE_PONDERADO")), "#0.00")

                i = i + 1

            End While
            dr.Close()
            dr = Nothing

            '*********************************************
            'Recupera Cabecera del Score
            '*********************************************
            oComm = New SqlClient.SqlCommand("PA_RecuperaRecomendacionScore", oConex)
            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
            oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodcred

            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            R.oScore.cPuntaje = "0"
            R.oScore.cRecomendacion = "NO EXISTE SCORE"

            While dr.Read
                R.oScore.nCodCred = nCodcred
                R.oScore.nCodAge = nCodAge
                R.oScore.cPuntaje = Format(CDbl(dr.Item("nPuntaje")), "#0.00")
                R.oScore.cRecomendacion = dr.Item("cRecomendacion")

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


    Public Function CambioClave() As outNegocio

        Dim R As New outNegocio
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nResp As Integer
        Dim sRespuesta As String
        Dim dr As SqlClient.SqlDataReader
        Dim i, j As Integer

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


            oComm = New SqlClient.SqlCommand("PA_CambiaClave", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@psCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@psClaveActual", SqlDbType.VarChar, 50).Value = sClave
            oComm.Parameters.Add("@psClaveNueva", SqlDbType.VarChar, 50).Value = sClaveNueva
            oComm.Parameters.Add("@pnResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@psRespuesta", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            nResp = oComm.Parameters("@pnResp").Value
            sRespuesta = oComm.Parameters("@psRespuesta").Value

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


    Public Function RegistraaDatosCambioTasa() As outNegocio

        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_RegsitraCambioTasa", oConex)
            oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodcred
            oComm.Parameters.Add("@nTasaComp", SqlDbType.Money).Value = nTasa
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu

            oComm.Parameters.Add("@pnResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@psRespuesta", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output

            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            R.nResultado = oComm.Parameters("@pnResp").Value
            R.sRespuesta = oComm.Parameters("@psRespuesta").Value

            D(0) = Str(R.nResultado).Trim
            D(1) = R.sRespuesta


            oConex.Close()

        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message

        End Try


        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))

        Return R

    End Function



    Public Function RecuperaDatosCambioTasa() As outNegocio

        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_RecuperaDatosCambioTasa", oConex)
            oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = cDoc
            oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@cCliente", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nTasaComp", SqlDbType.Money).Direction = ParameterDirection.Output

            oComm.Parameters.Add("@pnResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@cRespuesta", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output



            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            R.nResultado = oComm.Parameters("@pnResp").Value
            R.sRespuesta = oComm.Parameters("@cRespuesta").Value

            D(0) = Str(R.nResultado).Trim
            D(1) = R.sRespuesta


            If R.nResultado = 1 Then
                R.nCodCred = oComm.Parameters("@nCodCred").Value
                R.cCliente = oComm.Parameters("@cCliente").Value
                R.nTasaComp = oComm.Parameters("@nTasaComp").Value
            End If



            oConex.Close()

        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message

        End Try


        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))

        Return R

    End Function


    Public Function ActualzaCelularCliente() As outNegocio

        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_RegistraCelularCliente", oConex)
            oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = cDoc
            oComm.Parameters.Add("@cCelular", SqlDbType.VarChar, 50).Value = cMovil1
            oComm.Parameters.Add("@pnResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@psRespuesta", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output


            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            R.nResultado = oComm.Parameters("@pnResp").Value
            R.sRespuesta = oComm.Parameters("@psRespuesta").Value


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



    Public Function RecuperaCelularesActuales() As outNegocio

        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_RecuperaCelularCliente", oConex)
            oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = cDoc
            oComm.Parameters.Add("@cCelular", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@pnResp", SqlDbType.Int).Direction = ParameterDirection.Output


            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            R.cMovil1 = oComm.Parameters("@cCelular").Value
            R.nResultado = oComm.Parameters("@pnResp").Value

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



    Public Function ActualzaImeiAnulaCierre() As outNegocio

        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_ActualizaImeiAnulaCierre", oConex)
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@nResetIMEI", SqlDbType.Int).Value = nResetIMEI
            oComm.Parameters.Add("@nAnulaCierre", SqlDbType.Int).Value = nAnulaCierre
            oComm.Parameters.Add("@nCierraOperacionesUsu", SqlDbType.Int).Value = nCierraOperacionesUsu
            oComm.Parameters.Add("@nResultado", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@sRespuesta", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output


            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            R.nResultado = oComm.Parameters("@nResultado").Value
            R.sRespuesta = oComm.Parameters("@sRespuesta").Value


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



    Public Function ActualzaDireccionesActuales() As outNegocio

        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_ActualizaDireccionesActuales", oConex)
            oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = cDoc
            oComm.Parameters.Add("@cDireccionDom", SqlDbType.VarChar, 150).Value = cDireccionDomicilio
            oComm.Parameters.Add("@cDireccionNeg", SqlDbType.VarChar, 150).Value = cDireccionNegocio
            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@sRespuesta", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nCambiaDirecDom", SqlDbType.Int).Value = nCambiaDireccionDomicilio
            oComm.Parameters.Add("@nCambiaDirecNeg", SqlDbType.Int).Value = nCambiaDireccionNegocio


            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            R.nResultado = oComm.Parameters("@nResp").Value
            R.sRespuesta = oComm.Parameters("@sRespuesta").Value


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



    Public Function RecuperaDireccionesActuales() As outNegocio

        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_RecuperaDireccionesActuales", oConex)
            oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = cDoc
            oComm.Parameters.Add("@cDireccionDom", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@cDireccionNeg", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output


            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            R.cDireccionDomicilio = oComm.Parameters("@cDireccionDom").Value
            R.cDireccionNegocio = oComm.Parameters("@cDireccionNeg").Value
            R.nResultado = oComm.Parameters("@nResp").Value

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



    Public Function RecuperaInactivos() As outNegocio
        Dim R As New outNegocio
        Dim Rtemp As New outNegocio

        Dim oListaCliInactivos = From T In db.ClientesInactivos
                                 Where T.nCodAge = nCodAge
                                 Select T

        R.oListaClientesInactivos = oListaCliInactivos

        Return R

    End Function

    Public Function RegistraEnvioAVerificacion() As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_EnviarAVerificacion", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@ncodcred", SqlDbType.Int).Value = nCodcred

            oComm.ExecuteNonQuery()

            nResp = 1
            sRespuesta = "CREDTIO SE ENVIÓ A VERIFICACION CORRECTAMENTE"

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

    Public Function RecuperaPrecalifiNew() As outNegocio

        Dim R As New outNegocio
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
            oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = cDoc
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
                R.oPreCalif.ResultadoMicro = dr.Item("ResultadoMicro")
                R.oPreCalif.nRespMicro = dr.Item("nRespMicro")
                R.oPreCalif.ResultadoCredParalelo = dr.Item("ResultadoParalelo")
                R.oPreCalif.nRespParalelo = dr.Item("nRespParalelo")

                R.oPreCalif.ResultadoCrediPlus = dr.Item("ResultadoCrediplus")
                R.oPreCalif.nRespCrediPlus = dr.Item("nRespCrediplus")

                i = i + 1

            End While
            dr.Close()
            dr = Nothing


            oComm = New SqlClient.SqlCommand("PA_RecuperaProductosCliente", oConex)
            oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = cDoc

            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

            i = 0

            While dr.Read
                ReDim Preserve R.oProductos(i)
                R.oProductos(i) = New clsProductos
                R.oProductos(i).sproducto = dr.Item("sproducto")
                R.oProductos(i).nprestamo = dr.Item("nPrestamo")
                R.oProductos(i).nmontocuota = dr.Item("nmontocuota")
                R.oProductos(i).ncuotastotales = dr.Item("ncuotastotales")
                R.oProductos(i).ncuotaspagadas = dr.Item("ncuotaspagadas")
                R.oProductos(i).nsaldo = dr.Item("nsaldo")
                R.oProductos(i).ndiasatraso = dr.Item("ndiasatraso")
                R.oProductos(i).dfecvig = dr.Item("dfecvig")
                R.oProductos(i).dfecvenc = dr.Item("dfecvenc")


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

    Public Function RecuperaDatosBonoSupvervisor() As outNegocio

        Dim R As New outNegocio
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



            oComm = New SqlClient.SqlCommand("PA_RecuperaDatosBonoSupvervisor", oConex)
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu

            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                R.oDatosBonoSupvervisor = New clsDatosBonoSupvervisor

                R.oDatosBonoSupvervisor.nMoraFinMes = dr.Item("nMoraFinMes")
                R.oDatosBonoSupvervisor.nMoraHoy = dr.Item("nMoraHoy")
                R.oDatosBonoSupvervisor.nSaldoFinMes = dr.Item("nSaldoFinMes")
                R.oDatosBonoSupvervisor.nSaldoHoy = dr.Item("nSaldoHoy")
                R.oDatosBonoSupvervisor.nSaldoVencidoFinMes = dr.Item("nSaldoVencidoFinMes")
                R.oDatosBonoSupvervisor.nSaldoVencidoHoy = dr.Item("nSaldoVencidoHoy")
                R.oDatosBonoSupvervisor.nBono = dr.Item("nBono")
                R.oDatosBonoSupvervisor.nDiferencia = dr.Item("nDiferencia")


                i = i + 1

            End While
            dr.Close()
            dr = Nothing


            'Bono para cartera menor a 31/03
            oComm = New SqlClient.SqlCommand("PA_RecuoeraBonoSupervisorCarMarzo", oConex)
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@Meta", SqlDbType.Money, 50).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nBono", SqlDbType.Money, 50).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nRecup", SqlDbType.Money, 50).Direction = ParameterDirection.Output

            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            R.oDatosBonoSupvervisor.nMetaCarMarzo = oComm.Parameters("@Meta").Value
            R.oDatosBonoSupvervisor.nBonoCarMarzo = oComm.Parameters("@nBono").Value
            R.oDatosBonoSupvervisor.nRecupMarzo = oComm.Parameters("@nRecup").Value

            oComm = New SqlClient.SqlCommand("PA_RecuperaCreditosVencidosNuevos", oConex)
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu

            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oDatosBonoSupvervisor.oCreditosVencidosNuevos(i)
                R.oDatosBonoSupvervisor.oCreditosVencidosNuevos(i) = New clsCreditosVencidosNuevos

                R.oDatosBonoSupvervisor.oCreditosVencidosNuevos(i).cCliente = dr.Item("cCliente")
                R.oDatosBonoSupvervisor.oCreditosVencidosNuevos(i).cCodUsuCob = dr.Item("cCodUsuCob")
                R.oDatosBonoSupvervisor.oCreditosVencidosNuevos(i).dFecVig = dr.Item("dFecVig")
                R.oDatosBonoSupvervisor.oCreditosVencidosNuevos(i).nCodCred = dr.Item("nCodCred")
                R.oDatosBonoSupvervisor.oCreditosVencidosNuevos(i).nDiasAtraso = dr.Item("nDiasAtraso")
                R.oDatosBonoSupvervisor.oCreditosVencidosNuevos(i).nSaldok = dr.Item("nSaldok")

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

    Public Function RecuperaListaCreditosPorDesembolsar() As outNegocio

        Dim R As New outNegocio
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


            oComm = New SqlClient.SqlCommand("PA_LimpiaTempoDesembolsoCredito", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()


            For k = 0 To oBandejaDesembolsoNuevo.Count - 1

                oComm = New SqlClient.SqlCommand("PA_RegistraTempoDesembolsoCredito", oConex)
                oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
                oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = oBandejaDesembolsoNuevo(k).nCodCred
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.ExecuteNonQuery()

            Next


            oComm = New SqlClient.SqlCommand("PA_RecuperaListaCreditosPorDesembolsar", oConex)
            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu

            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oBandejaDesembolsoNuevo(i)
                R.oBandejaDesembolsoNuevo(i) = New clsDesembolsoNuevo

                R.oBandejaDesembolsoNuevo(i).nCodAge = dr.Item("nCodAge")
                R.oBandejaDesembolsoNuevo(i).nCodCred = dr.Item("nCodCred")
                R.oBandejaDesembolsoNuevo(i).Credito = dr.Item("Credito")
                R.oBandejaDesembolsoNuevo(i).cDoc = dr.Item("cDoc")
                R.oBandejaDesembolsoNuevo(i).Cliente = dr.Item("Cliente")
                R.oBandejaDesembolsoNuevo(i).nPrestamo = dr.Item("nPrestamo")
                R.oBandejaDesembolsoNuevo(i).cCodUsuCob = dr.Item("cCodUsuCob")
                R.oBandejaDesembolsoNuevo(i).nTasaComp = dr.Item("nTasaComp")
                R.oBandejaDesembolsoNuevo(i).nCuotas = dr.Item("nCuotas")
                R.oBandejaDesembolsoNuevo(i).nMontoCuota = dr.Item("nMontoCuota")
                R.oBandejaDesembolsoNuevo(i).bAmpliado = dr.Item("bAmpliado")
                R.oBandejaDesembolsoNuevo(i).cCodUsuDesem = dr.Item("cCodUsuDesem")

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

    Public Function RegistraMotivoNODesembolso() As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_RegistraMotivoNODesembolso", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = nCodAge
            oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = nCodcred
            oComm.Parameters.Add("@psCodMotivo", SqlDbType.Int).Value = nCodMotivoNoDesembolso
            oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@nLat", SqlDbType.Float).Value = nLatitud
            oComm.Parameters.Add("@nLong", SqlDbType.Float).Value = nLongitud
            oComm.Parameters.Add("@psResul", SqlDbType.VarChar, 2500).Direction = ParameterDirection.Output


            oComm.ExecuteNonQuery()

            nResp = 1
            sRespuesta = oComm.Parameters("@psResul").Value

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

    Public Function RecuperaPagoClientePorProducto() As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_Movil_RecuperaCreditoClientePorProducto", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@pnCoAge", SqlDbType.Int).Value = nCodAge
            oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = nCodcred

            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            R.oClientePorProducto = New clsClientePorProducto

            While dr.Read

                R.oClientePorProducto.cCliente = dr.Item("cCliente")
                R.oClientePorProducto.sFechaVencimiento = dr.Item("sFechaVencimiento")
                R.oClientePorProducto.nMontoCuota = dr.Item("nMontoCuota")
                R.oClientePorProducto.nCodCred = dr.Item("nCodCred")
                R.oClientePorProducto.nPrestamo = dr.Item("nPrestamo")
                R.oClientePorProducto.cMovil1 = dr.Item("cMovil1")
                R.oClientePorProducto.cZonaPersona = dr.Item("cZonaPersona")
                R.oClientePorProducto.cDireccionDom = dr.Item("cDireccionDom")
                R.oClientePorProducto.cZonaNegocio = dr.Item("cZonaNegocio")
                R.oClientePorProducto.cDirecNeg = dr.Item("cDirecNeg")
                R.oClientePorProducto.nNroProxCuota = dr.Item("nNroProxCuota")
                R.oClientePorProducto.nDeudaAtrasada = dr.Item("nDeudaAtrasada")
                R.oClientePorProducto.nDiasAtraso = dr.Item("nDiasAtraso")
                R.oClientePorProducto.UltFecCob = dr.Item("UltFecCob")
                R.oClientePorProducto.cReferenciaNeg = dr.Item("cReferenciaNeg")
                R.oClientePorProducto.cRazonSocial = dr.Item("cRazonSocial")
                R.oClientePorProducto.cMovil2 = dr.Item("cMovil2")
                R.oClientePorProducto.nCodAge = dr.Item("nCodAge")
                R.oClientePorProducto.MontoCancelaCuota = dr.Item("MontoCancelaCuota")
                R.oClientePorProducto.MontoCancelaCredito = dr.Item("MontoCancelaCredito")
                R.oClientePorProducto.Ampliacion = dr.Item("Ampliacion")
                R.oClientePorProducto.cDoc = dr.Item("cDoc")
                R.oClientePorProducto.nLatitud = dr.Item("nLatitud")
                R.oClientePorProducto.nLongitud = dr.Item("nLongitud")
                R.oClientePorProducto.nPagoMinimo = dr.Item("nPagoMinimo")

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

    Public Function RecuperaPagosListaItem() As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_Movil_RecuperaCreditosItem", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nEstado", SqlDbType.Int).Value = 4
            oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = nCodAge
            oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodcred
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            R.oDesembolsoItem = New clsDesembolsoItem
            While dr.Read


                R.oDesembolsoItem.cCliente = dr.Item("cCliente")
                R.oDesembolsoItem.sFechaVencimiento = dr.Item("sFechaVencimiento")
                R.oDesembolsoItem.nMontoCuota = dr.Item("nMontoCuota")
                R.oDesembolsoItem.nCodCred = dr.Item("nCodCred")
                R.oDesembolsoItem.nPrestamo = dr.Item("nPrestamo")
                R.oDesembolsoItem.cMovil1 = dr.Item("cMovil1")
                R.oDesembolsoItem.cZonaPersona = dr.Item("cZonaPersona")
                R.oDesembolsoItem.cDireccionDom = dr.Item("cDireccionDom")
                R.oDesembolsoItem.cZonaNegocio = dr.Item("cZonaNegocio")
                R.oDesembolsoItem.cDirecNeg = dr.Item("cDirecNeg")
                R.oDesembolsoItem.nNroProxCuota = dr.Item("nNroProxCuota")
                R.oDesembolsoItem.nDeudaAtrasada = dr.Item("nDeudaAtrasada")
                R.oDesembolsoItem.nDiasAtraso = dr.Item("nDiasAtraso")
                R.oDesembolsoItem.UltFecCob = dr.Item("UltFecCob")
                R.oDesembolsoItem.cReferenciaNeg = dr.Item("cReferenciaNeg")
                R.oDesembolsoItem.cRazonSocial = dr.Item("cRazonSocial")
                R.oDesembolsoItem.cMovil2 = dr.Item("cMovil2")
                R.oDesembolsoItem.nCodAge = dr.Item("nCodAge")
                R.oDesembolsoItem.MontoCancelaCuota = dr.Item("MontoCancelaCuota")
                R.oDesembolsoItem.MontoCancelaCredito = dr.Item("MontoCancelaCredito")
                R.oDesembolsoItem.Ampliacion = dr.Item("Ampliacion")
                R.oDesembolsoItem.cDoc = dr.Item("cDoc")
                R.oDesembolsoItem.nLatitud = dr.Item("nLatitud")
                R.oDesembolsoItem.nLongitud = dr.Item("nLongitud")

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


    Public Function RecuperaDesembolsosDiaListaXProducto() As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_Movil_RecuperaCreditosXProdDesemDelDia", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nEstado", SqlDbType.Int).Value = 4
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oDesembolsos(i)
                R.oDesembolsos(i) = New clsDesembolsos

                R.oDesembolsos(i).cCliente = dr.Item("cCliente")
                R.oDesembolsos(i).sFechaVencimiento = dr.Item("sFechaVencimiento")
                R.oDesembolsos(i).nMontoCuota = dr.Item("nMontoCuota")
                R.oDesembolsos(i).nCodCred = dr.Item("nCodCred")
                R.oDesembolsos(i).nPrestamo = dr.Item("nPrestamo")
                R.oDesembolsos(i).UltFecCob = dr.Item("UltFecCob")
                R.oDesembolsos(i).nCodAge = dr.Item("nCodAge")
                R.oDesembolsos(i).Ampliacion = dr.Item("Ampliacion")
                R.oDesembolsos(i).cDoc = dr.Item("cDoc")
                R.oDesembolsos(i).nLatNeg = dr.Item("nLatitud")
                R.oDesembolsos(i).nLongNeg = dr.Item("nLongitud")

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

    Public Function RecuperaBandejaAprobacion() As outNegocio
        Dim R As New outNegocio
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

            'Carga Gestores de Cobranza de Administrador
            oComm = New SqlClient.SqlCommand("PA_RecuperaGestoresDeCobranzaDeAdm", oConex)
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu

            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            R.oUsuariosEquipo = Nothing

            While dr.Read

                ReDim Preserve R.oUsuariosEquipo(i)
                R.oUsuariosEquipo(i) = New clsUsuariosEquipo

                R.oUsuariosEquipo(i).cCodUsu = dr.Item("cCodUsu")
                R.oUsuariosEquipo(i).cRol = dr.Item("cRol")

                i = i + 1
            End While
            dr.Close()
            dr = Nothing


            oComm = New SqlClient.SqlCommand("PA_LimpiaTempoAprobacionCredito", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()


            For k = 0 To oBandejaAprobacion.Count - 1

                oComm = New SqlClient.SqlCommand("PA_RegistraTempoAprobacionCredito", oConex)
                oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
                oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = oBandejaAprobacion(k).ncodcred
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.ExecuteNonQuery()

            Next

            oComm = New SqlClient.SqlCommand("PA_RecuperaBandejaAprobacion", oConex)
            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu

            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oAprobaciones(i)
                R.oAprobaciones(i) = New clsAprobacionNueva

                R.oAprobaciones(i).cDoc = dr.Item("cDoc")
                R.oAprobaciones(i).cCliente = dr.Item("cCliente")
                R.oAprobaciones(i).cMovil1 = dr.Item("cMovil1")
                R.oAprobaciones(i).cGenero = dr.Item("cGenero")
                R.oAprobaciones(i).cEstadoCiv = dr.Item("cEstadoCiv")
                R.oAprobaciones(i).cCategoria2 = dr.Item("cCategoria2")
                R.oAprobaciones(i).cCodUsuCob = dr.Item("cCodUsuCob")
                R.oAprobaciones(i).cDistritoDom = dr.Item("cDistritoDom")
                R.oAprobaciones(i).cZonaDom = dr.Item("cZonaDom")
                R.oAprobaciones(i).cDirecccionDom = dr.Item("cDirecccionDom")
                R.oAprobaciones(i).nLatitud = dr.Item("nLatitud")
                R.oAprobaciones(i).nLongitud = dr.Item("nLongitud")
                R.oAprobaciones(i).ccodusu = dr.Item("ccodusu")
                R.oAprobaciones(i).dFecha = dr.Item("dFecha")
                R.oAprobaciones(i).ncodcred = dr.Item("ncodcred")
                R.oAprobaciones(i).cRazonSocial = dr.Item("cRazonSocial")
                R.oAprobaciones(i).cActividadEcon = dr.Item("cActividadEcon")
                R.oAprobaciones(i).nLatNeg = dr.Item("nLatNeg")
                R.oAprobaciones(i).nLongNeg = dr.Item("nLongNeg")
                R.oAprobaciones(i).nVentasDiarias = dr.Item("nVentasDiarias")
                R.oAprobaciones(i).cDistritoNeg = dr.Item("cDistritoNeg")
                R.oAprobaciones(i).cZonaNeg = dr.Item("cZonaNeg")
                R.oAprobaciones(i).cDireccionNeg = dr.Item("cDireccionNeg")
                R.oAprobaciones(i).cResultadoCall = dr.Item("cResultadoCall")
                R.oAprobaciones(i).dFechaCall = dr.Item("dFechaCall")
                R.oAprobaciones(i).cResultadoVerif = dr.Item("cResultadoVerif")
                R.oAprobaciones(i).cUsuVerificador = dr.Item("cUsuVerificador")
                R.oAprobaciones(i).cComentarioSustento = dr.Item("cComentarioSustento")
                R.oAprobaciones(i).nPrestamoMaximo = dr.Item("nPrestamoMaximo")
                R.oAprobaciones(i).nTasaAsignada = dr.Item("nTasaAsignada")
                R.oAprobaciones(i).nPrestamo = dr.Item("nPrestamo")
                R.oAprobaciones(i).nSaldoAmpliado = dr.Item("nSaldoAmpliacion")
                R.oAprobaciones(i).nCuotas = dr.Item("nCuotas")
                R.oAprobaciones(i).nMontoCuota = dr.Item("nMontoCuota")
                R.oAprobaciones(i).cTipoAprobacion = dr.Item("cTipoAprobacion")
                R.oAprobaciones(i).nPagoPromedio = dr.Item("nPagoPromedio")
                R.oAprobaciones(i).nMontoCuotaAnt = dr.Item("nMontoCuotaAnt")
                R.oAprobaciones(i).nPrestamoAnt = dr.Item("nPrestamoAnt")
                R.oAprobaciones(i).nProd = dr.Item("nProd")
                R.oAprobaciones(i).nCodAge = dr.Item("nCodAge")
                R.oAprobaciones(i).cProducto = dr.Item("cProducto")
                R.oAprobaciones(i).nVerificacion = dr.Item("nVerificacion")
                R.oAprobaciones(i).nExcepcion = dr.Item("nExcepcion")
                R.oAprobaciones(i).cComentarioExcepcion = dr.Item("cComentarioExcepcion")
                R.oAprobaciones(i).cDesembolsoBanco = dr.Item("cDesembolsoBanco")
                R.oAprobaciones(i).nCampCuotasSinInt = dr.Item("nCampSinInt")

                i = i + 1

            End While
            dr.Close()
            dr = Nothing

            If i > 0 Then

                For j = 0 To R.oAprobaciones.Count - 1

                    oComm = New SqlClient.SqlCommand("PA_RecuperaRecordCliente", oConex)
                    oComm.Parameters.Add("@psDoc", SqlDbType.Int).Value = R.oAprobaciones(j).cDoc
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()

                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0

                    While dr.Read
                        ReDim Preserve R.oAprobaciones(j).oRecord(i)
                        R.oAprobaciones(j).oRecord(i) = New clsRecordNuevo
                        R.oAprobaciones(j).oRecord(i).ncodcred = dr.Item("ncodcred")
                        R.oAprobaciones(j).oRecord(i).cEstado = dr.Item("cEstado")
                        R.oAprobaciones(j).oRecord(i).nDiasAtraso = dr.Item("nDiasAtraso")
                        R.oAprobaciones(j).oRecord(i).dFecVig = dr.Item("dFecVig")
                        R.oAprobaciones(j).oRecord(i).dFecUltPago = dr.Item("dFecUltPago")
                        R.oAprobaciones(j).oRecord(i).nDiasAtrasoProm = dr.Item("nDiasAtrasoProm")
                        R.oAprobaciones(j).oRecord(i).nPrestamo = dr.Item("nPrestamo")
                        R.oAprobaciones(j).oRecord(i).nMontoCuota = dr.Item("nMontoCuota")
                        R.oAprobaciones(j).oRecord(i).nMontoPagoPromedio = dr.Item("nMontoPagoPromedio")
                        R.oAprobaciones(j).oRecord(i).nTasaComp = dr.Item("nTasaComp")
                        R.oAprobaciones(j).oRecord(i).nCuotasPag = dr.Item("nCuotasPagadas")
                        R.oAprobaciones(j).oRecord(i).nCuotasPend = dr.Item("nCuotasPendientes")

                        i = i + 1

                    End While
                    dr.Close()
                    dr = Nothing

                Next


                For j = 0 To R.oAprobaciones.Count - 1

                    oComm = New SqlClient.SqlCommand("PA_RecuperaHistoriaSBS", oConex)
                    oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = R.oAprobaciones(j).cDoc
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()

                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0

                    While dr.Read
                        ReDim Preserve R.oAprobaciones(j).oSBS(i)
                        R.oAprobaciones(j).oSBS(i) = New clsSBSHist
                        R.oAprobaciones(j).oSBS(i).Fec_rep = dr.Item("Fec_rep")
                        R.oAprobaciones(j).oSBS(i).Calif_0 = dr.Item("Calif_0")
                        R.oAprobaciones(j).oSBS(i).Calif_1 = dr.Item("Calif_1")
                        R.oAprobaciones(j).oSBS(i).Calif_2 = dr.Item("Calif_2")
                        R.oAprobaciones(j).oSBS(i).Calif_3 = dr.Item("Calif_3")
                        R.oAprobaciones(j).oSBS(i).Calif_4 = dr.Item("Calif_4")
                        R.oAprobaciones(j).oSBS(i).Can_Ents = dr.Item("Can_Ents")

                        i = i + 1

                    End While
                    dr.Close()
                    dr = Nothing

                Next



                For j = 0 To R.oAprobaciones.Count - 1

                    oComm = New SqlClient.SqlCommand("PA_DetalleSBS", oConex)
                    oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = R.oAprobaciones(j).cDoc
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()

                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0

                    While dr.Read
                        ReDim Preserve R.oAprobaciones(j).oSBSDetalle(i)
                        R.oAprobaciones(j).oSBSDetalle(i) = New clsSBSHistDet
                        R.oAprobaciones(j).oSBSDetalle(i).nombre = dr.Item("nombre")
                        R.oAprobaciones(j).oSBSDetalle(i).nDeuda = dr.Item("nDeuda")
                        R.oAprobaciones(j).oSBSDetalle(i).clasificacion = dr.Item("clasificacion")

                        i = i + 1

                    End While
                    dr.Close()
                    dr = Nothing

                Next


                For j = 0 To R.oAprobaciones.Count - 1

                    oComm = New SqlClient.SqlCommand("PA_ImagenesNegocioAprobacion", oConex)
                    oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = R.oAprobaciones(j).ncodcred
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()

                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0

                    While dr.Read
                        ReDim Preserve R.oAprobaciones(j).oImagenes(i)
                        R.oAprobaciones(j).oImagenes(i) = New clsVerificacionClienteImagenes
                        R.oAprobaciones(j).oImagenes(i).nCodCred = R.oAprobaciones(j).ncodcred
                        R.oAprobaciones(j).oImagenes(i).iImagen = UnicodeBytesToString(dr.Item("nFoto"))
                        R.oAprobaciones(j).oImagenes(i).cTipo = dr.Item("cTipo")
                        R.oAprobaciones(j).oImagenes(i).nID = i

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

    Public Function RecuperaDatosAprobacionIndividual() As outNegocio
        Dim R As New outNegocio
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



            oComm = New SqlClient.SqlCommand("PA_RecuperaDatosAprobacionIndividual", oConex)
            oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = cDoc

            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oAprobaciones(i)
                R.oAprobaciones(i) = New clsAprobacionNueva

                R.oAprobaciones(i).cDoc = dr.Item("cDoc")
                R.oAprobaciones(i).cCliente = dr.Item("cCliente")
                R.oAprobaciones(i).cMovil1 = dr.Item("cMovil1")
                R.oAprobaciones(i).cGenero = dr.Item("cGenero")
                R.oAprobaciones(i).cEstadoCiv = dr.Item("cEstadoCiv")
                R.oAprobaciones(i).cCategoria2 = dr.Item("cCategoria2")
                R.oAprobaciones(i).cCodUsuCob = dr.Item("cCodUsuCob")
                R.oAprobaciones(i).cDistritoDom = dr.Item("cDistritoDom")
                R.oAprobaciones(i).cZonaDom = dr.Item("cZonaDom")
                R.oAprobaciones(i).cDirecccionDom = dr.Item("cDirecccionDom")
                R.oAprobaciones(i).nLatitud = dr.Item("nLatitud")
                R.oAprobaciones(i).nLongitud = dr.Item("nLongitud")
                R.oAprobaciones(i).ccodusu = dr.Item("ccodusu")
                R.oAprobaciones(i).dFecha = dr.Item("dFecha")
                R.oAprobaciones(i).ncodcred = dr.Item("ncodcred")
                R.oAprobaciones(i).cRazonSocial = dr.Item("cRazonSocial")
                R.oAprobaciones(i).cActividadEcon = dr.Item("cActividadEcon")
                R.oAprobaciones(i).nLatNeg = dr.Item("nLatNeg")
                R.oAprobaciones(i).nLongNeg = dr.Item("nLongNeg")
                R.oAprobaciones(i).nVentasDiarias = dr.Item("nVentasDiarias")
                R.oAprobaciones(i).cDistritoNeg = dr.Item("cDistritoNeg")
                R.oAprobaciones(i).cZonaNeg = dr.Item("cZonaNeg")
                R.oAprobaciones(i).cDireccionNeg = dr.Item("cDireccionNeg")
                R.oAprobaciones(i).cResultadoCall = dr.Item("cResultadoCall")
                R.oAprobaciones(i).dFechaCall = dr.Item("dFechaCall")
                R.oAprobaciones(i).cResultadoVerif = dr.Item("cResultadoVerif")
                R.oAprobaciones(i).cUsuVerificador = dr.Item("cUsuVerificador")
                R.oAprobaciones(i).cComentarioSustento = dr.Item("cComentarioSustento")
                R.oAprobaciones(i).nPrestamoMaximo = dr.Item("nPrestamoMaximo")
                R.oAprobaciones(i).nTasaAsignada = dr.Item("nTasaAsignada")
                R.oAprobaciones(i).nPrestamo = dr.Item("nPrestamo")
                R.oAprobaciones(i).nSaldoAmpliado = dr.Item("nSaldoAmpliacion")
                R.oAprobaciones(i).nCuotas = dr.Item("nCuotas")
                R.oAprobaciones(i).nMontoCuota = dr.Item("nMontoCuota")
                R.oAprobaciones(i).cTipoAprobacion = dr.Item("cTipoAprobacion")
                R.oAprobaciones(i).nPagoPromedio = dr.Item("nPagoPromedio")
                R.oAprobaciones(i).nMontoCuotaAnt = dr.Item("nMontoCuotaAnt")
                R.oAprobaciones(i).nPrestamoAnt = dr.Item("nPrestamoAnt")
                R.oAprobaciones(i).nProd = dr.Item("nProd")
                R.oAprobaciones(i).nCodAge = dr.Item("nCodAge")
                R.oAprobaciones(i).cProducto = dr.Item("cProducto")
                R.oAprobaciones(i).nVerificacion = dr.Item("nVerificacion")
                R.oAprobaciones(i).nExcepcion = dr.Item("nExcepcion")
                R.oAprobaciones(i).cComentarioExcepcion = dr.Item("cComentarioExcepcion")
                R.oAprobaciones(i).cDesembolsoBanco = dr.Item("cDesembolsoBanco")
                R.oAprobaciones(i).nCampCuotasSinInt = dr.Item("nCampSinInt")

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

    Public Function RecuperaDataPrecalifNueva() As outNegocio
        Dim R As New outNegocio
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

            ReDim Preserve R.oAprobaciones(0)
            R.oAprobaciones(0) = New clsAprobacionNueva
            R.oAprobaciones(0).cDoc = cDoc

            oComm = New SqlClient.SqlCommand("PA_RecuperaRecordCliente", oConex)
            oComm.Parameters.Add("@psDoc", SqlDbType.Int).Value = cDoc
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oAprobaciones(0).oRecord(i)
                R.oAprobaciones(0).oRecord(i) = New clsRecordNuevo
                R.oAprobaciones(0).oRecord(i).ncodcred = dr.Item("ncodcred")
                R.oAprobaciones(0).oRecord(i).cEstado = dr.Item("cEstado")
                R.oAprobaciones(0).oRecord(i).nDiasAtraso = dr.Item("nDiasAtraso")
                R.oAprobaciones(0).oRecord(i).dFecVig = dr.Item("dFecVig")
                R.oAprobaciones(0).oRecord(i).dFecUltPago = dr.Item("dFecUltPago")
                R.oAprobaciones(0).oRecord(i).nDiasAtrasoProm = dr.Item("nDiasAtrasoProm")
                R.oAprobaciones(0).oRecord(i).nPrestamo = dr.Item("nPrestamo")
                R.oAprobaciones(0).oRecord(i).nMontoCuota = dr.Item("nMontoCuota")
                R.oAprobaciones(0).oRecord(i).nMontoPagoPromedio = dr.Item("nMontoPagoPromedio")

                i = i + 1

            End While
            dr.Close()
            dr = Nothing


            For j = 0 To R.oAprobaciones.Count - 1

                oComm = New SqlClient.SqlCommand("PA_RecuperaHistoriaSBS", oConex)
                oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = R.oAprobaciones(j).cDoc
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.ExecuteNonQuery()

                dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                i = 0

                While dr.Read
                    ReDim Preserve R.oAprobaciones(j).oSBS(i)
                    R.oAprobaciones(j).oSBS(i) = New clsSBSHist
                    R.oAprobaciones(j).oSBS(i).Fec_rep = dr.Item("Fec_rep")
                    R.oAprobaciones(j).oSBS(i).Calif_0 = dr.Item("Calif_0")
                    R.oAprobaciones(j).oSBS(i).Calif_1 = dr.Item("Calif_1")
                    R.oAprobaciones(j).oSBS(i).Calif_2 = dr.Item("Calif_2")
                    R.oAprobaciones(j).oSBS(i).Calif_3 = dr.Item("Calif_3")
                    R.oAprobaciones(j).oSBS(i).Calif_4 = dr.Item("Calif_4")
                    R.oAprobaciones(j).oSBS(i).Can_Ents = dr.Item("Can_Ents")

                    i = i + 1

                End While
                dr.Close()
                dr = Nothing

            Next



            For j = 0 To R.oAprobaciones.Count - 1

                oComm = New SqlClient.SqlCommand("PA_DetalleSBS", oConex)
                oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = R.oAprobaciones(j).cDoc
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.ExecuteNonQuery()

                dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                i = 0

                While dr.Read
                    ReDim Preserve R.oAprobaciones(j).oSBSDetalle(i)
                    R.oAprobaciones(j).oSBSDetalle(i) = New clsSBSHistDet
                    R.oAprobaciones(j).oSBSDetalle(i).nombre = dr.Item("nombre")
                    R.oAprobaciones(j).oSBSDetalle(i).nDeuda = dr.Item("nDeuda")
                    R.oAprobaciones(j).oSBSDetalle(i).clasificacion = dr.Item("clasificacion")

                    i = i + 1

                End While
                dr.Close()
                dr = Nothing

            Next




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

    Public Function RegistraVerificacionNueva() As outNegocio
        Dim R As New outNegocio
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

            If oVerifCliente(0).nResultado = 2 Then

                oComm = New SqlClient.SqlCommand("PA_RegistraVerificacionRechazo", oConex)
                oComm.CommandType = CommandType.StoredProcedure
                oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = oVerifCliente(0).nCodCred
                oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = sCodUsu
                oComm.Parameters.Add("@nResultado", SqlDbType.Int).Value = oVerifCliente(0).nResultado
                oComm.Parameters.Add("@cComentarioSustento", SqlDbType.VarChar, 1500).Value = oVerifCliente(0).cComentarioSustento
                oComm.Parameters.Add("@nLatNeg", SqlDbType.Float).Value = oVerifCliente(0).nLatNeg
                oComm.Parameters.Add("@nLongNeg", SqlDbType.Float).Value = oVerifCliente(0).nLongNeg
                oComm.Parameters.Add("@nLatDom", SqlDbType.Float).Value = oVerifCliente(0).nLatDom
                oComm.Parameters.Add("@nLongDom", SqlDbType.Float).Value = oVerifCliente(0).nLongDom
                oComm.Parameters.Add("@cDirecDom", SqlDbType.VarChar, 500).Value = oVerifCliente(0).cDirecDom
                oComm.Parameters.Add("@cRefDirecDom", SqlDbType.VarChar, 500).Value = oVerifCliente(0).cReferenciaDom
                oComm.Parameters.Add("@cDirecNeg", SqlDbType.VarChar, 500).Value = oVerifCliente(0).cDirecNeg
                oComm.Parameters.Add("@cRefDirecNeg", SqlDbType.VarChar, 500).Value = oVerifCliente(0).cReferenciaNeg
                oComm.Parameters.Add("@cRazonSocial", SqlDbType.VarChar, 500).Value = oVerifCliente(0).cRazonSocial
                oComm.Parameters.Add("@nVerifDom", SqlDbType.Int).Value = oVerifCliente(0).nVerifDom
                oComm.Parameters.Add("@nVerifNeg", SqlDbType.Int).Value = oVerifCliente(0).nVerifNeg
                oComm.Parameters.Add("@cCodAct", SqlDbType.VarChar, 50).Value = oVerifCliente(0).cCodActEcon


                oComm.ExecuteNonQuery()


            Else

                oComm = New SqlClient.SqlCommand("PA_RegistraVerificacionNueva_1", oConex)
                oComm.CommandType = CommandType.StoredProcedure
                oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = oVerifCliente(0).nCodCred
                oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = sCodUsu
                oComm.Parameters.Add("@nResultado", SqlDbType.Int).Value = oVerifCliente(0).nResultado
                oComm.Parameters.Add("@cComentarioSustento", SqlDbType.VarChar, 1500).Value = oVerifCliente(0).cComentarioSustento
                oComm.Parameters.Add("@nLatNeg", SqlDbType.Float).Value = oVerifCliente(0).nLatNeg
                oComm.Parameters.Add("@nLongNeg", SqlDbType.Float).Value = oVerifCliente(0).nLongNeg
                oComm.Parameters.Add("@nLatDom", SqlDbType.Float).Value = oVerifCliente(0).nLatDom
                oComm.Parameters.Add("@nLongDom", SqlDbType.Float).Value = oVerifCliente(0).nLongDom
                oComm.Parameters.Add("@cDirecDom", SqlDbType.VarChar, 500).Value = oVerifCliente(0).cDirecDom
                oComm.Parameters.Add("@cRefDirecDom", SqlDbType.VarChar, 500).Value = oVerifCliente(0).cReferenciaDom
                oComm.Parameters.Add("@cDirecNeg", SqlDbType.VarChar, 500).Value = oVerifCliente(0).cDirecNeg
                oComm.Parameters.Add("@cRefDirecNeg", SqlDbType.VarChar, 500).Value = oVerifCliente(0).cReferenciaNeg
                oComm.Parameters.Add("@cRazonSocial", SqlDbType.VarChar, 500).Value = oVerifCliente(0).cRazonSocial
                oComm.Parameters.Add("@nVerifDom", SqlDbType.Int).Value = oVerifCliente(0).nVerifDom
                oComm.Parameters.Add("@nVerifNeg", SqlDbType.Int).Value = oVerifCliente(0).nVerifNeg
                oComm.Parameters.Add("@cCodAct", SqlDbType.VarChar, 50).Value = oVerifCliente(0).cCodActEcon
                'mdipas 2024
                oComm.Parameters.Add("@nVerifCuali", SqlDbType.Int).Value = oVerifCliente(0).nVerifDom
                oComm.Parameters.Add("@nCumpleConAntNeg", SqlDbType.Int).Value = oVerifCliente(0).nCumpleConAntNeg
                oComm.Parameters.Add("@nEsPropietario", SqlDbType.Int).Value = oVerifCliente(0).nEsPropietario
                oComm.Parameters.Add("@nPoseeEmpleados", SqlDbType.Int).Value = oVerifCliente(0).nPoseeEmpleados
                oComm.Parameters.Add("@nNroEmpleados", SqlDbType.Int).Value = oVerifCliente(0).nNroEmpleados
                oComm.Parameters.Add("@nTieneRotuloVisible", SqlDbType.Int).Value = oVerifCliente(0).nTieneRotuloVisible
                oComm.Parameters.Add("@nNegoCuentaInfra", SqlDbType.Int).Value = oVerifCliente(0).nNegoCuentaInfra
                oComm.Parameters.Add("@nCuentaPermiOpe", SqlDbType.Int).Value = oVerifCliente(0).nCuentaPermiOpe
                oComm.Parameters.Add("@nObservaAfluCli ", SqlDbType.Int).Value = oVerifCliente(0).nObservaAfluCli
                oComm.Parameters.Add("@nUbicaFacilAcceso ", SqlDbType.Int).Value = oVerifCliente(0).nUbicaFacilAcceso
                oComm.Parameters.Add("@nTieneBuenaRefeZona ", SqlDbType.Int).Value = oVerifCliente(0).nTieneBuenaRefeZona
                oComm.Parameters.Add("@nTieneProspBuenaActi", SqlDbType.Int).Value = oVerifCliente(0).nTieneProspBuenaActi
                oComm.Parameters.Add("@nTieneHoraAtencion ", SqlDbType.Int).Value = oVerifCliente(0).nTieneHoraAtencion
                oComm.Parameters.Add("@nAquilaCasa ", SqlDbType.Int).Value = oVerifCliente(0).nAquilaCasa
                oComm.Parameters.Add("@nTiempoAquila ", SqlDbType.Int).Value = oVerifCliente(0).nTiempoAquila
                oComm.Parameters.Add("@nTieneDependientes ", SqlDbType.Int).Value = oVerifCliente(0).nTieneDependientes
                oComm.Parameters.Add("@nNroDependientes ", SqlDbType.Int).Value = oVerifCliente(0).nNroDependientes
                oComm.Parameters.Add("@cHoraInicio", SqlDbType.VarChar, 50).Value = oVerifCliente(0).cHoraInicio
                oComm.Parameters.Add("@cHoraFinal", SqlDbType.VarChar, 50).Value = oVerifCliente(0).cHoraFinal
                oComm.Parameters.Add("@nNivelDelincuencia", SqlDbType.Int).Value = oVerifCliente(0).nNivelDelincuencia


                oComm.ExecuteNonQuery()

                Dim i As Integer

                For i = 0 To oVerifCliente(0).oImagenes.Count - 1


                    oComm = New SqlClient.SqlCommand("PA_RegistraVerificacionImagenNuevo", oConex)
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = oVerifCliente(0).nCodCred
                    oComm.Parameters.Add("@nImagen", SqlDbType.Image).Value = UnicodeStringToBytes(oVerifCliente(0).oImagenes(i).iImagen)
                    oComm.Parameters.Add("@cTipoFoto", SqlDbType.VarChar, 50).Value = oVerifCliente(0).oImagenes(i).cTipo


                    oComm.ExecuteNonQuery()

                Next

                'mdipas 2024
                For i = 0 To oVerifReferencia.Count - 1

                    oComm = New SqlClient.SqlCommand("PA_RegistraVerificacionReferenciaNuevo", oConex)
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = oVerifReferencia(i).nCodCred
                    oComm.Parameters.Add("@cNombres", SqlDbType.VarChar, 150).Value = oVerifReferencia(i).cNombres
                    oComm.Parameters.Add("@cTelefono", SqlDbType.VarChar, 50).Value = oVerifReferencia(i).cTelefono
                    oComm.Parameters.Add("@nRecomendado", SqlDbType.Int).Value = oVerifReferencia(i).nRecomendado
                    oComm.Parameters.Add("@cRecomendado", SqlDbType.VarChar, 500).Value = oVerifReferencia(i).cRecomendado
                    oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = sCodUsu
                    oComm.ExecuteNonQuery()

                Next
                'end


            End If

            nResp = 1
            sRespuesta = "Registro Correcto"

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

    Public Function RecuperaEvaluacion() As outNegocio
        Dim R As New outNegocio
        Dim AbonoDesemb() As dAbonoDesemb
        Dim ListaDesembolsador() As String
        Dim i As Integer
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


            oComm = New SqlClient.SqlCommand("PA_RecuperaDatosNegocioVerif", oConex)
            oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = nCodcred
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            R.oEvalNegocio = New clsEvalNegocio()

            While dr.Read

                R.oEvalNegocio.cRazonSocial = dr.Item("cRazonSocial")
                R.oEvalNegocio.cDirecNeg = dr.Item("cDirecNeg")
                R.oEvalNegocio.cReferenciaNeg = dr.Item("cReferenciaNeg")
                R.oEvalNegocio.nLatNeg = dr.Item("nLatNeg")
                R.oEvalNegocio.nLongNeg = dr.Item("nLongNeg")
                R.oEvalNegocio.cCodActEcon = dr.Item("cCodActEcon")


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

    Public Function RecuperaImagenesNegocio() As outNegocio
        Dim R As New outNegocio
        Dim AbonoDesemb() As dAbonoDesemb
        Dim ListaDesembolsador() As String
        Dim i As Integer
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


            oComm = New SqlClient.SqlCommand("PA_ImagenesNegocio", oConex)
            oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodcred
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oFotosEvaluacion(i)
                R.oFotosEvaluacion(i) = UnicodeBytesToString(dr.Item("nFoto"))

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

    Public Function RecuperaBandejaVerificacion() As outNegocio
        Dim R As New outNegocio
        Dim AbonoDesemb() As dAbonoDesemb
        Dim ListaDesembolsador() As String
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

            oComm = New SqlClient.SqlCommand("PA_LimpiaTempoVerificCliente", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()


            For k = 0 To oVerifCliente.Count - 1

                oComm = New SqlClient.SqlCommand("PA_RegistraTempoVerificCliente", oConex)
                oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
                oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = oVerifCliente(k).nCodCred
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.ExecuteNonQuery()


            Next

            oComm = New SqlClient.SqlCommand("PA_RecuperaBandejaVerificacion_1", oConex)
            oComm.Parameters.Add("@sCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oBandejaVerificacion(i)
                R.oBandejaVerificacion(i) = New clsVerificacionCliente

                R.oBandejaVerificacion(i).cDoc = dr.Item("cDoc")
                R.oBandejaVerificacion(i).cCliente = dr.Item("cCliente")
                R.oBandejaVerificacion(i).nVerifDom = dr.Item("nVerifDom")
                R.oBandejaVerificacion(i).nVerifNeg = dr.Item("nVerifNeg")
                R.oBandejaVerificacion(i).nResultado = dr.Item("nResultado")
                R.oBandejaVerificacion(i).cUserSolic = dr.Item("ccodusu")
                R.oBandejaVerificacion(i).nCodCred = dr.Item("nCodcred")
                R.oBandejaVerificacion(i).cRazonSocial = dr.Item("cRazonSocial")
                R.oBandejaVerificacion(i).cDirecNeg = dr.Item("cDirecNeg")
                R.oBandejaVerificacion(i).cReferenciaNeg = dr.Item("cReferenciaNeg")
                R.oBandejaVerificacion(i).nLatNeg = dr.Item("nLatNeg")
                R.oBandejaVerificacion(i).nLongNeg = dr.Item("nLongNeg")
                R.oBandejaVerificacion(i).cCodActEcon = dr.Item("cCodActEcon")
                R.oBandejaVerificacion(i).cDirecDom = dr.Item("cDireccionDom")
                R.oBandejaVerificacion(i).cReferenciaDom = dr.Item("cReferenciaDom")
                R.oBandejaVerificacion(i).nPrestampo = dr.Item("nPrestamo")
                R.oBandejaVerificacion(i).nNumCuotas = dr.Item("nCuotas")
                R.oBandejaVerificacion(i).nMontoCuota = dr.Item("nMontoCuota")
                R.oBandejaVerificacion(i).cMovil = dr.Item("cMovil1")

                i = i + 1

            End While
            dr.Close()
            dr = Nothing

            If i > 0 Then

                For j = 0 To R.oBandejaVerificacion.Count - 1

                    oComm = New SqlClient.SqlCommand("PA_ImagenesNegocio", oConex)
                    oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = R.oBandejaVerificacion(j).nCodCred
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()

                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0

                    While dr.Read
                        ReDim Preserve R.oBandejaVerificacion(j).oImagenes(i)
                        R.oBandejaVerificacion(j).oImagenes(i) = New clsVerificacionClienteImagenes
                        R.oBandejaVerificacion(j).oImagenes(i).nCodCred = R.oBandejaVerificacion(j).nCodCred
                        R.oBandejaVerificacion(j).oImagenes(i).iImagen = UnicodeBytesToString(dr.Item("nFoto"))
                        R.oBandejaVerificacion(j).oImagenes(i).cTipo = "NE0"
                        R.oBandejaVerificacion(j).oImagenes(i).nID = i

                        i = i + 1

                    End While
                    dr.Close()
                    dr = Nothing


                Next



                For j = 0 To R.oBandejaVerificacion.Count - 1

                    oComm = New SqlClient.SqlCommand("PA_RecuperaRecordCliente", oConex)
                    oComm.Parameters.Add("@psDoc", SqlDbType.Int).Value = R.oBandejaVerificacion(j).cDoc
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()

                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0

                    While dr.Read
                        ReDim Preserve R.oBandejaVerificacion(j).oRecord(i)
                        R.oBandejaVerificacion(j).oRecord(i) = New clsRecordNuevo
                        R.oBandejaVerificacion(j).oRecord(i).ncodcred = dr.Item("ncodcred")
                        R.oBandejaVerificacion(j).oRecord(i).cEstado = dr.Item("cEstado")
                        R.oBandejaVerificacion(j).oRecord(i).nDiasAtraso = dr.Item("nDiasAtraso")
                        R.oBandejaVerificacion(j).oRecord(i).dFecVig = dr.Item("dFecVig")
                        R.oBandejaVerificacion(j).oRecord(i).dFecUltPago = dr.Item("dFecUltPago")
                        R.oBandejaVerificacion(j).oRecord(i).nDiasAtrasoProm = dr.Item("nDiasAtrasoProm")
                        R.oBandejaVerificacion(j).oRecord(i).nPrestamo = dr.Item("nPrestamo")
                        R.oBandejaVerificacion(j).oRecord(i).nMontoCuota = dr.Item("nMontoCuota")
                        R.oBandejaVerificacion(j).oRecord(i).nMontoPagoPromedio = dr.Item("nMontoPagoPromedio")
                        R.oBandejaVerificacion(j).oRecord(i).nTasaComp = dr.Item("nTasaComp")
                        R.oBandejaVerificacion(j).oRecord(i).nCuotasPag = dr.Item("nCuotasPagadas")
                        R.oBandejaVerificacion(j).oRecord(i).nCuotasPend = dr.Item("nCuotasPendientes")

                        i = i + 1

                    End While
                    dr.Close()
                    dr = Nothing

                Next


                For j = 0 To R.oBandejaVerificacion.Count - 1

                    oComm = New SqlClient.SqlCommand("PA_RecuperaHistoriaSBS", oConex)
                    oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = R.oBandejaVerificacion(j).cDoc
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()

                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0

                    While dr.Read
                        ReDim Preserve R.oBandejaVerificacion(j).oSBS(i)
                        R.oBandejaVerificacion(j).oSBS(i) = New clsSBSHist
                        R.oBandejaVerificacion(j).oSBS(i).Fec_rep = dr.Item("Fec_rep")
                        R.oBandejaVerificacion(j).oSBS(i).Calif_0 = dr.Item("Calif_0")
                        R.oBandejaVerificacion(j).oSBS(i).Calif_1 = dr.Item("Calif_1")
                        R.oBandejaVerificacion(j).oSBS(i).Calif_2 = dr.Item("Calif_2")
                        R.oBandejaVerificacion(j).oSBS(i).Calif_3 = dr.Item("Calif_3")
                        R.oBandejaVerificacion(j).oSBS(i).Calif_4 = dr.Item("Calif_4")
                        R.oBandejaVerificacion(j).oSBS(i).Can_Ents = dr.Item("Can_Ents")

                        i = i + 1

                    End While
                    dr.Close()
                    dr = Nothing

                Next



                For j = 0 To R.oBandejaVerificacion.Count - 1

                    oComm = New SqlClient.SqlCommand("PA_DetalleSBS", oConex)
                    oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = R.oBandejaVerificacion(j).cDoc
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()

                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0

                    While dr.Read
                        ReDim Preserve R.oBandejaVerificacion(j).oSBSDetalle(i)
                        R.oBandejaVerificacion(j).oSBSDetalle(i) = New clsSBSHistDet
                        R.oBandejaVerificacion(j).oSBSDetalle(i).nombre = dr.Item("nombre")
                        R.oBandejaVerificacion(j).oSBSDetalle(i).nDeuda = dr.Item("nDeuda")
                        R.oBandejaVerificacion(j).oSBSDetalle(i).clasificacion = dr.Item("clasificacion")

                        i = i + 1

                    End While
                    dr.Close()
                    dr = Nothing

                Next

                'mdipas 2024
                For j = 0 To R.oBandejaVerificacion.Count - 1

                    oComm = New SqlClient.SqlCommand("PA_ReferenciasVericacion", oConex)
                    oComm.Parameters.Add("@nCodCred", SqlDbType.VarChar, 50).Value = R.oBandejaVerificacion(j).nCodCred
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()

                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0

                    While dr.Read
                        ReDim Preserve R.oBandejaVerificacion(j).oReferencias(i)
                        R.oBandejaVerificacion(j).oReferencias(i) = New clsVerificacionReferencia
                        R.oBandejaVerificacion(j).oReferencias(i).nCodCred = dr.Item("nCodCred")
                        R.oBandejaVerificacion(j).oReferencias(i).cNombres = dr.Item("cNombres")
                        R.oBandejaVerificacion(j).oReferencias(i).cTelefono = dr.Item("cTelefono")
                        R.oBandejaVerificacion(j).oReferencias(i).nRecomendado = dr.Item("nRecomienda")
                        R.oBandejaVerificacion(j).oReferencias(i).cRecomendado = dr.Item("cRecomiendaSustento")
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

    Public Function RecuperaBandejaVerificacionParaActualizar() As outNegocio
        Dim R As New outNegocio
        Dim AbonoDesemb() As dAbonoDesemb
        Dim ListaDesembolsador() As String
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

            oComm = New SqlClient.SqlCommand("PA_RecuperaBandejaVerificacionActualiza_1", oConex)
            oComm.Parameters.Add("@sCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oBandejaVerificacion(i)
                R.oBandejaVerificacion(i) = New clsVerificacionCliente

                R.oBandejaVerificacion(i).cDoc = dr.Item("cDoc")
                R.oBandejaVerificacion(i).cCliente = dr.Item("cCliente")
                R.oBandejaVerificacion(i).nVerifDom = dr.Item("nVerifDom")
                R.oBandejaVerificacion(i).nVerifNeg = dr.Item("nVerifNeg")
                R.oBandejaVerificacion(i).nResultado = dr.Item("nResultado")
                R.oBandejaVerificacion(i).cUserSolic = dr.Item("ccodusu")
                R.oBandejaVerificacion(i).nCodCred = dr.Item("nCodcred")
                R.oBandejaVerificacion(i).cRazonSocial = dr.Item("cRazonSocial")
                R.oBandejaVerificacion(i).cDirecNeg = dr.Item("cDirecNeg")
                R.oBandejaVerificacion(i).cReferenciaNeg = dr.Item("cReferenciaNeg")
                R.oBandejaVerificacion(i).nLatNeg = dr.Item("nLatNeg")
                R.oBandejaVerificacion(i).nLongNeg = dr.Item("nLongNeg")
                R.oBandejaVerificacion(i).cCodActEcon = dr.Item("cCodActEcon")
                R.oBandejaVerificacion(i).cDirecDom = dr.Item("cDireccionDom")
                R.oBandejaVerificacion(i).cReferenciaDom = dr.Item("cReferenciaDom")
                R.oBandejaVerificacion(i).nPrestampo = dr.Item("nPrestamo")
                R.oBandejaVerificacion(i).nNumCuotas = dr.Item("nCuotas")
                R.oBandejaVerificacion(i).nMontoCuota = dr.Item("nMontoCuota")
                R.oBandejaVerificacion(i).cMovil = dr.Item("cMovil1")
                R.oBandejaVerificacion(i).nVerifCuali = dr.Item("nVerifCuali")
                R.oBandejaVerificacion(i).nVerifRefe = dr.Item("nVerifRefe")

                R.oBandejaVerificacion(i).nCumpleConAntNeg = dr.Item("nCumpleConAntNeg")
                R.oBandejaVerificacion(i).nEsPropietario = dr.Item("nEsPropietario")
                R.oBandejaVerificacion(i).nPoseeEmpleados = dr.Item("nPoseeEmpleados")
                R.oBandejaVerificacion(i).nNroEmpleados = dr.Item("nNroEmpleados")
                R.oBandejaVerificacion(i).nTieneRotuloVisible = dr.Item("nTieneRotuloVisible")
                R.oBandejaVerificacion(i).nNegoCuentaInfra = dr.Item("nNegoCuentaInfra")
                R.oBandejaVerificacion(i).nCuentaPermiOpe = dr.Item("nCuentaPermiOpe")
                R.oBandejaVerificacion(i).nObservaAfluCli = dr.Item("nObservaAfluCli")
                R.oBandejaVerificacion(i).nUbicaFacilAcceso = dr.Item("nUbicaFacilAcceso")
                R.oBandejaVerificacion(i).nTieneBuenaRefeZona = dr.Item("nTieneBuenaRefeZona")
                R.oBandejaVerificacion(i).nTieneProspBuenaActi = dr.Item("nTieneProspBuenaActi")
                R.oBandejaVerificacion(i).nTieneHoraAtencion = dr.Item("nTieneHoraAtencion")
                R.oBandejaVerificacion(i).nAquilaCasa = dr.Item("nAquilaCasa")
                R.oBandejaVerificacion(i).nTiempoAquila = dr.Item("nTiempoAquila")
                R.oBandejaVerificacion(i).nTieneDependientes = dr.Item("nTieneDependientes")
                R.oBandejaVerificacion(i).nNroDependientes = dr.Item("nNroDependientes")
                R.oBandejaVerificacion(i).cHoraInicio = dr.Item("cHoraInicio")
                R.oBandejaVerificacion(i).cHoraFinal = dr.Item("cHoraFinal")
                R.oBandejaVerificacion(i).nNivelDelincuencia = dr.Item("nNivelDelincuencia")
                R.oBandejaVerificacion(i).cComentarioCPD = dr.Item("cComentarioCPD")
                R.oBandejaVerificacion(i).cNombreGrupo = dr.Item("cNomGrupo")

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

    Public Function ValidaRol() As outNegocio
        Dim R As New outNegocio
        Dim AbonoDesemb() As dAbonoDesemb
        Dim ListaDesembolsador() As String
        Dim i As Integer
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


            oComm = New SqlClient.SqlCommand("PA_ValidaRolUsuario", oConex)
            oComm.Parameters.Add("@psCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@psRol", SqlDbType.VarChar, 50).Value = Str(nCodRol).Trim
            oComm.Parameters.Add("@psResult", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@pnResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()


            D(0) = Str(oComm.Parameters("@pnResp").Value).Trim
            D(1) = oComm.Parameters("@psResult").Value

            oConex.Close()

        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message

        End Try


        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))

        Return R

    End Function


    Public Function PrefiltroXProd() As outNegocio
        Dim R As New outNegocio
        Dim AbonoDesemb() As dAbonoDesemb
        Dim ListaDesembolsador() As String
        Dim i As Integer
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



            If nProd = 3 Then
                oComm = New SqlClient.SqlCommand("PA_PreCalificacion", oConex)
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = cDoc
                oComm.Parameters.Add("@cResultado", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                oComm.Parameters.Add("@bResp", SqlDbType.Int).Direction = ParameterDirection.Output
                oComm.Parameters.Add("@cuser", SqlDbType.VarChar, 50).Value = sCodUsu
                oComm.Parameters.Add("@nLatitud", SqlDbType.Float).Value = nLatitud
                oComm.Parameters.Add("@nLongitud", SqlDbType.Float).Value = nLongitud

                oComm.ExecuteNonQuery()

            Else

                oComm = New SqlClient.SqlCommand("PA_PreCalificacion_Consumo", oConex)
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = cDoc
                oComm.Parameters.Add("@cResultado", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                oComm.Parameters.Add("@bResp", SqlDbType.Int).Direction = ParameterDirection.Output
                oComm.Parameters.Add("@cuser", SqlDbType.VarChar, 50).Value = sCodUsu
                oComm.Parameters.Add("@nLatitud", SqlDbType.Float).Value = nLatitud
                oComm.Parameters.Add("@nLongitud", SqlDbType.Float).Value = nLongitud

                oComm.ExecuteNonQuery()

            End If


            D(0) = Str(oComm.Parameters("@bResp").Value)
            D(1) = oComm.Parameters("@cResultado").Value

        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message

        End Try


        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))

        Return R

    End Function

    Public Function RecuperaBonosXColocacion() As outNegocio
        Dim R As New outNegocio
        Dim AbonoDesemb() As dAbonoDesemb
        Dim ListaDesembolsador() As String
        Dim i As Integer
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


            oComm = New SqlClient.SqlCommand("PA_RecuperaBonoXColocacion_2", oConex)
            oComm.Parameters.Add("@sUser", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oBonosXColocacion(i)
                R.oBonosXColocacion(i) = New clsBonoXColocacion

                R.oBonosXColocacion(i).dfecha = dr.Item("dfecha")
                R.oBonosXColocacion(i).nAvance = dr.Item("nAvance")
                R.oBonosXColocacion(i).nComision = dr.Item("nComision")
                R.oBonosXColocacion(i).nMonto = dr.Item("nMonto")
                R.oBonosXColocacion(i).nMontoAcumulado = dr.Item("nMontoAcumulado")
                R.oBonosXColocacion(i).nMeta = dr.Item("nMeta")

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


    Public Function RegistraMotivoNOPago() As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_RegistraMotivoNOPago_2", oConex)
            oComm.CommandType = CommandType.StoredProcedure
            '
            oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = oMotivoNoPago.nCodAge
            oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = oMotivoNoPago.nCodCred
            oComm.Parameters.Add("@psCodMotivo", SqlDbType.Int).Value = oMotivoNoPago.nMotivo
            oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = oMotivoNoPago.cUser
            oComm.Parameters.Add("@nLat", SqlDbType.Float).Value = oMotivoNoPago.nLatitud
            oComm.Parameters.Add("@nLong", SqlDbType.Float).Value = oMotivoNoPago.nLongitud
            oComm.Parameters.Add("@sComentario", SqlDbType.VarChar, 1500).Value = oMotivoNoPago.sComentario
            oComm.Parameters.Add("@psResul", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output


            oComm.ExecuteNonQuery()

            nResp = 1
            sRespuesta = oComm.Parameters("@psResul").Value

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

    Public Function RecuperaAsesores() As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_RecupeaAmpliaciones", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu

            'oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

            Dim i As Integer = 0

            While dr.Read
                ReDim Preserve R.oAmpliaciones(i)
                R.oAmpliaciones(i) = New clsItemAmpliacion()
                R.oAmpliaciones(i).cDNI = dr.Item("cDoc")
                R.oAmpliaciones(i).cCliente = dr.Item("cCliente")
                R.oAmpliaciones(i).nLatNeg = dr.Item("nLatNeg")
                R.oAmpliaciones(i).nLongNeg = dr.Item("nLongNeg")
                R.oAmpliaciones(i).cDirecNeg = dr.Item("cDirecNeg")
                R.oAmpliaciones(i).cRazonSocial = dr.Item("cRazonSocial")
                R.oAmpliaciones(i).nPrestamo = dr.Item("nPrestamo")
                R.oAmpliaciones(i).nSaldoK = dr.Item("nSaldoK")
                R.oAmpliaciones(i).nExcedente = dr.Item("nExcedente")


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

    Public Function RecuperaProyeccionAmortizaciones() As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_RecuperaProyeccionAmortizacion", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu

            'oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

            Dim i As Integer = 0

            While dr.Read
                ReDim Preserve R.oProyecamortizacion(i)
                R.oProyecamortizacion(i) = New clsproyecamortizacion()
                R.oProyecamortizacion(i).nSaldoKAyer = dr.Item("nSaldoKAyer")
                R.oProyecamortizacion(i).AmortizacionPromedio7Dias = dr.Item("AmortizacionPromedio7Dias")
                R.oProyecamortizacion(i).AmortizacionEsperada = dr.Item("AmortizacionEsperada")
                R.oProyecamortizacion(i).AmortizacionActual = dr.Item("AmortizacionActual")
                R.oProyecamortizacion(i).ColocacionNeta = dr.Item("ColocacionNeta")

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



    Public Function RecuperaAmpliaciones() As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_RecupeaAmpliaciones", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu

            'oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

            Dim i As Integer = 0

            While dr.Read
                ReDim Preserve R.oAmpliaciones(i)
                R.oAmpliaciones(i) = New clsItemAmpliacion()
                R.oAmpliaciones(i).cDNI = dr.Item("cDoc")
                R.oAmpliaciones(i).cCliente = dr.Item("cCliente")
                R.oAmpliaciones(i).nLatNeg = dr.Item("nLatNeg")
                R.oAmpliaciones(i).nLongNeg = dr.Item("nLongNeg")
                R.oAmpliaciones(i).cDirecNeg = dr.Item("cDirecNeg")
                R.oAmpliaciones(i).cRazonSocial = dr.Item("cRazonSocial")
                R.oAmpliaciones(i).nPrestamo = dr.Item("nPrestamo")
                R.oAmpliaciones(i).nSaldoK = dr.Item("nSaldoK")
                R.oAmpliaciones(i).nCuotas = dr.Item("nCuotas")
                R.oAmpliaciones(i).nCuotaPendiente = dr.Item("nnroproxcuota")
                R.oAmpliaciones(i).nTasacomp = dr.Item("nTasacomp")
                R.oAmpliaciones(i).nExcedente = dr.Item("nExcedente")


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

    Private Function UnicodeStringToBytes(ByVal sImagen As String) As Byte()

        'Return System.Text.Encoding.Default.GetString(bytes)
        Return System.Convert.FromBase64String(sImagen)

    End Function

    Private Function UnicodeBytesToString(ByVal bytes() As Byte) As String

        'Return System.Text.Encoding.Default.GetString(bytes)
        Return System.Convert.ToBase64String(bytes)

    End Function

    Public Function RecuperaFotosVerificacion() As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_RecuperaFotoVerif", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = cDoc

            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

            Dim i As Integer = 0

            While dr.Read
                ReDim Preserve R.oFotosVerificacion(i)

                R.oFotosVerificacion(i) = UnicodeBytesToString(dr.Item("nFoto"))

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

    Public Function RecuperaTamañoNegocio() As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_Movil2_TamanoNegocios", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

            Dim i As Integer = 0

            While dr.Read
                ReDim Preserve R.oCatalogocodigos_2(i)
                R.oCatalogocodigos_2(i) = New clsCatalogocodigos()

                R.oCatalogocodigos_2(i).nCodigo = 0
                R.oCatalogocodigos_2(i).cDescripcion = dr.Item("cDescripcion").ToString + Space(100) + dr.Item("cValor").ToString
                R.oCatalogocodigos_2(i).cValor = dr.Item("cValor").ToString

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


    Public Function RegistraRechazoCredito() As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_Movil_RechazarCredito", oConex)
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
            oComm.Parameters.Add("@ncodcred", SqlDbType.Int).Value = nCodcred
            oComm.Parameters.Add("@cComentario", SqlDbType.VarChar, 2500).Value = ""
            oComm.Parameters.Add("@psCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu

            oComm.ExecuteNonQuery()

            nResp = 1
            sRespuesta = "REGISTRO CORRECTO"

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


    Public Function RegistraAprobacion() As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_Movil_AprobacionCredito_7", oConex)
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = CInt(oAprobacion.psCodage)
            oComm.Parameters.Add("@ncodcred", SqlDbType.Int).Value = CInt(oAprobacion.psCodCred)
            oComm.Parameters.Add("@cComentario", SqlDbType.VarChar, 2500).Value = oAprobacion.psComentario
            oComm.Parameters.Add("@pnPrestamo", SqlDbType.Money).Value = CDbl(oAprobacion.psPrestamo)
            oComm.Parameters.Add("@pnCuotas", SqlDbType.Int).Value = CInt(oAprobacion.psCuotas)
            oComm.Parameters.Add("@nMontoCuota", SqlDbType.Money).Value = CDbl(oAprobacion.psMontoCuota)
            oComm.Parameters.Add("@psCudUsu", SqlDbType.VarChar, 50).Value = oAprobacion.psCodUsu
            oComm.Parameters.Add("@psResult", SqlDbType.VarChar, 250).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@pnCodTamNeg", SqlDbType.Int).Value = oAprobacion.pnCodTamNeg
            oComm.Parameters.Add("@pcCodUsuCob", SqlDbType.VarChar, 100).Value = oAprobacion.cCodusuCob
            oComm.Parameters.Add("@nCampSinInt", SqlDbType.Int).Value = nCampSinInt

            oComm.ExecuteNonQuery()

            nResp = 1
            sRespuesta = oComm.Parameters("@psResult").Value

            If (Mid(sRespuesta, 1, 2) <> "01") Then
                nResp = 0
            End If


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


    Public Function MontoCuotaPorProd() As outNegocio
        Dim R As New outNegocio
        Dim D As New dDesembolso
        Dim sMontoCuota As String()

        sMontoCuota = D.MontoCuotaPorProd(Str(nPrestamo), Str(nCuotas), Str(nTasa), "1", "1", Str(nProd), "1")

        R.nResultado = 1
        R.sRespuesta = sMontoCuota(0)

        Return R

    End Function


    Public Function RecuperaDatosAprobacion() As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_Movil_RecuperaDatosAprobacion_3", oConex)
            oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = cDoc
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

            Dim i As Integer = 0

            While dr.Read

                R.oDatosAprobacion = New clsDatosAprobacion()

                R.oDatosAprobacion.cApeNom = dr.Item("cApeNom").ToString
                R.oDatosAprobacion.cCodUsuCob = dr.Item("cCodUsuCob").ToString
                R.oDatosAprobacion.psComenVerfi = dr.Item("psComenVerfi").ToString
                R.oDatosAprobacion.cCodUsuVerif = dr.Item("cCodUsuVerif").ToString
                R.oDatosAprobacion.dFechaVerif = Format(CDate(dr.Item("dFechaVerif").ToString), "dd/MM/yyyy HH:mm")
                R.oDatosAprobacion.nDiasAtraso = dr.Item("nDiasAtraso").ToString
                R.oDatosAprobacion.PrestamoMax = Format(CDbl(dr.Item("PrestamoMax").ToString), "#0.00")
                R.oDatosAprobacion.CuotasMax = dr.Item("CuotasMax").ToString
                R.oDatosAprobacion.Tasa = Format(CDbl(dr.Item("Tasa").ToString), "#0.00")
                R.oDatosAprobacion.nSaldoActual = Format(CDbl(dr.Item("nSaldoActual").ToString), "#0.00")
                R.oDatosAprobacion.nMontoPagoUltimo = Format(CDbl(dr.Item("nMontoPagoUltimo").ToString), "#0.00")
                R.oDatosAprobacion.nPrestamo = Format(CDbl(dr.Item("nPrestamo").ToString), "#0.00")
                R.oDatosAprobacion.nCuotas = Format(CDbl(dr.Item("nCuotas").ToString), "#0.00")
                R.oDatosAprobacion.nCodAge = dr.Item("nCodAge").ToString
                R.oDatosAprobacion.nCodCred = dr.Item("nCodCred").ToString
                R.oDatosAprobacion.Latitud = dr.Item("Latitud").ToString
                R.oDatosAprobacion.nLongitud = dr.Item("nLongitud").ToString
                R.oDatosAprobacion.cReferVecino1 = dr.Item("cReferVecino1").ToString
                R.oDatosAprobacion.cReferVecino2 = dr.Item("cReferVecino2").ToString
                R.oDatosAprobacion.cReferVecino3 = dr.Item("cReferVecino3").ToString
                R.oDatosAprobacion.pnTarifario = dr.Item("pnTarifario").ToString
                R.oDatosAprobacion.cResultadoCall = dr.Item("cResultadoCall").ToString
                R.oDatosAprobacion.cNomProd = dr.Item("cNomProd").ToString
                R.oDatosAprobacion.nProd = dr.Item("nProd").ToString
                R.oDatosAprobacion.bAmpliado = dr.Item("bAmpliado").ToString
                R.oDatosAprobacion.SaldoAnterior = dr.Item("SaldoAnterior").ToString
                R.oDatosAprobacion.cCategoria2 = dr.Item("cCategoria2").ToString
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

    Public Function RecuperaDatosBandejaAprobacion() As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_Movil_BandejaAprobacion", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

            Dim i As Integer = 0

            While dr.Read
                ReDim Preserve R.oCreditosAAprobar(i)
                R.oCreditosAAprobar(i) = New clsCreditosAAprobar()

                R.oCreditosAAprobar(i).credito = dr.Item("credito").ToString
                R.oCreditosAAprobar(i).cDoc = dr.Item("cDoc").ToString
                R.oCreditosAAprobar(i).Cliente = dr.Item("Cliente").ToString
                R.oCreditosAAprobar(i).nPrestamo = Format(CDbl(dr.Item("nPrestamo").ToString), "#0.00")
                R.oCreditosAAprobar(i).ESTADO = dr.Item("ESTADO").ToString
                R.oCreditosAAprobar(i).dFecVig = Format(CDate(dr.Item("dFecVig").ToString), "dd/MM/yyyy")
                R.oCreditosAAprobar(i).cUser = dr.Item("cUser").ToString

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


    Public Function RecuperaDatosCambioLinea() As outNegocio
        Dim R As New outNegocio
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


            oComm = New SqlClient.SqlCommand("PA_Movil_RecupelaLineaCredito", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@psDNI", SqlDbType.VarChar, 50).Value = cDoc

            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            R.oDatosActLinea = New clsDatosActLinea()


            While dr.Read
                R.oDatosActLinea.nLineaCredito = Format(CDbl(dr.Item("nLineaCredito").ToString), "#0.00")
                R.oDatosActLinea.nCuotasMax = Format(CInt(dr.Item("nCuotasMax").ToString), "#0")

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


    Public Function RegistraCambioLinea() As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_Movil_DatosLineaCredito", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@psDNI", SqlDbType.VarChar, 50).Value = oActLinea.psDNI
            oComm.Parameters.Add("@nMontoLinea", SqlDbType.Money).Value = CDbl(oActLinea.psMontoLinea)
            oComm.Parameters.Add("@nNumCuotas", SqlDbType.Int).Value = CInt(oActLinea.psNumCuotas)

            oComm.ExecuteNonQuery()

            nResp = 1
            sRespuesta = "REGISTRO CORRECTO"

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

    Public Function RegistraExcepcion() As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_Movil_RegistraExcepcion", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = cDoc

            oComm.ExecuteNonQuery()

            nResp = 1
            sRespuesta = "REGISTRO CORRECTO"

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

    Public Function RegistraMarcadoAsistencia() As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_RegistraFotoAsistencia_2", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@psCodUsu", SqlDbType.VarChar, 50).Value = oAsistencia.sCodUsu
            oComm.Parameters.Add("@iDoc", SqlDbType.Image).Value = oAsistencia.nImg
            oComm.Parameters.Add("@pnLatitud", SqlDbType.Float).Value = CDec(oAsistencia.sLatitud)
            oComm.Parameters.Add("@pnLongitud", SqlDbType.Float).Value = CDec(oAsistencia.sLongitud)
            oComm.Parameters.Add("@psTipoMarcado", SqlDbType.VarChar, 100).Value = oAsistencia.sTipoMarcado

            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            nResp = oComm.Parameters("@nResp").Value
            sRespuesta = oComm.Parameters("@Resultado").Value

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

    Public Function RegistraEnvioSMSCliente() As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_ActualizaEnvioSMSCliente", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@pnOperador", SqlDbType.Int).Value = nOperador
            oComm.Parameters.Add("@nNroTran", SqlDbType.Int).Value = CLng(sNroTran)

            oComm.ExecuteNonQuery()

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

    Public Function RecuperaCreditosOpeEspec() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nResp As Integer
        Dim sRespuesta As String
        Dim dr As SqlClient.SqlDataReader
        Dim i As Integer

        nResp = 0
        sRespuesta = ""
        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        oComm = New SqlClient.SqlCommand("PA_RecuperaCreditosOperEspec", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = cDoc

        oComm.ExecuteNonQuery()
        dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

        i = 0

        While dr.Read

            ReDim Preserve R.ListaDesembolsadoresAbonar(i)
            R.ListaDesembolsadoresAbonar(i) = dr.Item("nCodCred").ToString
            i = i + 1
        End While

        dr.Close()
        dr = Nothing

        nResp = 1
        sRespuesta = "PROCESO CORRECTO"

        oConex.Close()

        D(0) = Str(nResp)
        D(1) = sRespuesta
        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function



    Public Function RecuperaDatosEnvioSMS_Cliente() As outNegocio
        Dim R As New outNegocio
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


            oComm = New SqlClient.SqlCommand("PA_RecuperaParaEnvioSMSCliente", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@pnOperador", SqlDbType.Int).Value = nOperador

            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            R.oSMSCliente = New clsSMSCliente()

            ReDim Preserve D(3)
            While dr.Read
                R.oSMSCliente.cMovil1 = dr.Item("cMovil1").ToString
                R.oSMSCliente.sSerieGen = dr.Item("sSerieGen").ToString
                R.oSMSCliente.sNumeroGen = dr.Item("sNumeroGen").ToString
                R.oSMSCliente.nMonto = Format(CDbl(dr.Item("nMonto").ToString), "#,0.00")
                R.oSMSCliente.CuotasPagadas = dr.Item("CuotasPagadas").ToString
                R.oSMSCliente.CuotasTotales = dr.Item("CuotasTotales").ToString
                R.oSMSCliente.nNroTran = dr.Item("nNroTran").ToString
                R.oSMSCliente.Vecimiento = dr.Item("Vecimiento").ToString

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


    Public Function RegistraEnvioSMS() As outNegocio
        Dim R As New outNegocio
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


            oComm = New SqlClient.SqlCommand("PA_ActualizaEnvioSMSAvisoNuevo", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nOperador", SqlDbType.Int).Value = nOperador
            oComm.Parameters.Add("@cMovil", SqlDbType.VarChar, 50).Value = cMovil

            oComm.ExecuteNonQuery()

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



    Public Function RecuperaDatosEnvioSMS() As outNegocio
        Dim R As New outNegocio
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


            oComm = New SqlClient.SqlCommand("PA_RecuperaSMSAvisosNuevo", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nOperador", SqlDbType.Int).Value = nOperador

            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            R.oSMS = New clsSMS()

            ReDim Preserve D(3)
            While dr.Read
                R.oSMS.cMovil = dr.Item("cMovil").ToString
                R.oSMS.MSJ = dr.Item("MSJ").ToString
                R.oSMS.nNroTrans = dr.Item("nNroTrans")
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


    Public Function RecuperaFotoFirmaContrato() As outNegocio
        Dim R As New outNegocio
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

            Dim i As Integer

            oComm = New SqlClient.SqlCommand("PA_Movil2_RecuperaDatosFotoFirmaContrato", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = cDoc

            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            ReDim Preserve D(8)

            R.oFirmaContratoData = New clsfirmacontratodata()

            While dr.Read

                R.oFirmaContratoData.nCodPers = dr.Item("nCodPers")
                R.oFirmaContratoData.nNumFotos = dr.Item("nNumFotos").ToString
                R.oFirmaContratoData.Cliente = dr.Item("Cliente")
                R.oFirmaContratoData.nCodAge = dr.Item("nCodAge")
                R.oFirmaContratoData.nCodCred = dr.Item("nCodCred")
                R.oFirmaContratoData.nLatitud = dr.Item("nLatitud")
                R.oFirmaContratoData.nLongitud = dr.Item("nLongitud")

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

    Public Function GuardarFotoFirmaContrato() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
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

            oComm = New SqlClient.SqlCommand("PA_RegistraFotoFirmaContrato_2", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@pnCodPers", SqlDbType.Int).Value = CLng(oFirmaContrato.sCodPers)
            oComm.Parameters.Add("@iDoc", SqlDbType.Image).Value = oFirmaContrato.nImg
            oComm.Parameters.Add("@psCodUsu", SqlDbType.VarChar, 50).Value = oFirmaContrato.psCodUsu

            oComm.Parameters.Add("@nLat", SqlDbType.Float).Value = nLatitud
            oComm.Parameters.Add("@nLong", SqlDbType.Float).Value = nLongitud

            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output


            oComm.ExecuteNonQuery()


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



    Public Function RecuperaBonosXCobranza() As outNegocio
        Dim R As New outNegocio
        Dim AbonoDesemb() As dAbonoDesemb
        Dim ListaDesembolsador() As String
        Dim i As Integer
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


            oComm = New SqlClient.SqlCommand("PA_RecuperaBonoXCobranzaDelInteres", oConex)
            'oComm = New SqlClient.SqlCommand("PA_RecuperaBonoXCobranza", oConex)
            oComm.Parameters.Add("@sUser", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oBonosXCobranza(i)
                R.oBonosXCobranza(i) = New clsBonoXCobranza

                R.oBonosXCobranza(i).dFecha = dr.Item("dfecha")
                R.oBonosXCobranza(i).nMontoCobrado = dr.Item("nMonto")
                R.oBonosXCobranza(i).nMontoPagado = dr.Item("nPago")
                R.oBonosXCobranza(i).nBonoGenerado = dr.Item("nBono")

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

    Public Function RecuperaBonosXCobranzaNueva() As outNegocio
        Dim R As New outNegocio
        Dim AbonoDesemb() As dAbonoDesemb
        Dim ListaDesembolsador() As String
        Dim i As Integer
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


            oComm = New SqlClient.SqlCommand("PA_RecuperaBonoXCobranzaDelInteres_2", oConex)
            oComm.Parameters.Add("@sUser", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@nPorcComision", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nIncreMora", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nMoraAnt", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nMorahoy", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nSaldoFinMes", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nSaldoHoy", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nSaldoVencidoFinMes", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nSaldoVencidoHoy", SqlDbType.Money).Direction = ParameterDirection.Output

            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            R.nPorcComision = oComm.Parameters("@nPorcComision").Value
            R.nIncreMora = oComm.Parameters("@nIncreMora").Value
            R.nMoraAnt = oComm.Parameters("@nMoraAnt").Value
            R.nMorahoy = oComm.Parameters("@nMorahoy").Value

            R.nSaldoFinMes = oComm.Parameters("@nSaldoFinMes").Value
            R.nSaldoHoy = oComm.Parameters("@nSaldoHoy").Value
            R.nSaldoVencidoFinMes = oComm.Parameters("@nSaldoVencidoFinMes").Value
            R.nSaldoVencidoHoy = oComm.Parameters("@nSaldoVencidoHoy").Value


            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oBonosXCobranza(i)
                R.oBonosXCobranza(i) = New clsBonoXCobranza

                R.oBonosXCobranza(i).dFecha = dr.Item("dfecha")
                R.oBonosXCobranza(i).nMontoCobrado = dr.Item("nMonto")
                R.oBonosXCobranza(i).nMontoPagado = dr.Item("nPago")
                R.oBonosXCobranza(i).nBonoGenerado = dr.Item("nBono")

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


    Public Function RecuperaResumenCobranza() As outNegocio
        Dim R As New outNegocio
        Dim AbonoDesemb() As dAbonoDesemb
        Dim ListaDesembolsador() As String
        Dim i As Integer
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


            oComm = New SqlClient.SqlCommand("PA_ResumenCobranza", oConex)
            oComm.Parameters.Add("@psCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()


            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            R.oResumenCobranza = New clsResumenCobranza()
            While dr.Read

                R.oResumenCobranza.nPagos = dr.Item("nPagos")
                R.oResumenCobranza.nDepositos = dr.Item("nDepositos")
                R.oResumenCobranza.nDiferencia = dr.Item("nDiferencia")
                R.oResumenCobranza.nEntrega = dr.Item("nEntrega")

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




    Public Function RegistraVoucher() As outNegocio
        Dim R As New outNegocio

        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
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

            oComm = New SqlClient.SqlCommand("PA_RegistraFotoYDeposito_2", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@pnMonto", SqlDbType.Money).Value = oVoucher.nMonto
            oComm.Parameters.Add("@iDoc", SqlDbType.Image).Value = oVoucher.iFoto
            oComm.Parameters.Add("@ncodage", SqlDbType.Int).Value = oVoucher.nCodAge
            oComm.Parameters.Add("@psUser", SqlDbType.VarChar, 50).Value = oVoucher.cCodUsu

            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@cNumOperacion", SqlDbType.VarChar, 50).Value = oVoucher.cNumOperacion

            oComm.ExecuteNonQuery()


            oConex.Close()

            D(0) = Str(oComm.Parameters("@nResp").Value)
            D(1) = oComm.Parameters("@Resultado").Value

        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try

        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R



    End Function

    Public Function ValidaUsuario() As outNegocio
        Dim R As New outNegocio

        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
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

            oComm = New SqlClient.SqlCommand("PA_Movil2_ValidaUsuario", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@psCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@psClave", SqlDbType.VarChar, 50).Value = sToken
            oComm.Parameters.Add("@psRol", SqlDbType.VarChar, 50).Value = cDoc
            oComm.Parameters.Add("@psResult", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@pnResp", SqlDbType.Int).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            oConex.Close()

            D(0) = Str(oComm.Parameters("@pnResp").Value)
            D(1) = oComm.Parameters("@psResult").Value

        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try

        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R



    End Function

    Public Function ActualizaIMEI() As outNegocio
        Dim R As New outNegocio

        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
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

            oComm = New SqlClient.SqlCommand("PA_ActualizaIMEI", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@psUSER", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@psIMEI", SqlDbType.VarChar, 50).Value = sToken


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

    Public Function RegistraAbonosDesembolsos() As outNegocio
        Dim R As New outNegocio
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

            Dim i As Integer

            For i = 0 To ListaAbonos.Count - 1

                If ListaAbonos(i).nAbonar = 1 Then
                    oComm = New SqlClient.SqlCommand("PA_Movil_RegistraEntregaEfectivoParaDesembolso_2", oConex)
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = ListaAbonos(i).nMontoNeto
                    oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = sCodUsu
                    oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = ListaAbonos(i).nCodCred
                    oComm.ExecuteNonQuery()

                End If
            Next

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

    Public Function RecuperaListaDesembolsadores() As outNegocio
        Dim R As New outNegocio
        Dim AbonoDesemb() As dAbonoDesemb
        Dim ListaDesembolsador() As String
        Dim i As Integer
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


            oComm = New SqlClient.SqlCommand("PA_RecuperaDesembolsadores", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read

                ReDim Preserve ListaDesembolsador(i)
                ListaDesembolsador(i) = UCase(dr.Item("cValor"))
                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            R.ListaDesembolsadoresAbonar = ListaDesembolsador


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

    Public Function RecuperaListaXAbonar() As outNegocio
        Dim R As New outNegocio
        Dim AbonoDesemb() As dAbonoDesemb
        Dim ListaDesembolsador() As String
        Dim i As Integer
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


            oComm = New SqlClient.SqlCommand("PA_RecuperaCreditosXAbonarADesemb", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@psCodDesemb", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read
                ReDim Preserve AbonoDesemb(i)
                AbonoDesemb(i) = New dAbonoDesemb
                AbonoDesemb(i).nAbonar = 0
                AbonoDesemb(i).nCodAge = dr.Item("nCodAge")
                AbonoDesemb(i).nCodCred = dr.Item("nCodCred")
                AbonoDesemb(i).nDeudaVigente = dr.Item("nDeudaVigente")
                AbonoDesemb(i).nMontoNeto = dr.Item("nPrestamoNeto")
                AbonoDesemb(i).nPrestamo = dr.Item("nPrestamo")
                AbonoDesemb(i).cCiente = dr.Item("cCliente")

                i = i + 1

            End While
            dr.Close()
            dr = Nothing
            R.ListaAbonos = AbonoDesemb


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


    Public Function RegistraAsignacionAsesor() As outNegocio
        Dim R As New outNegocio
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


            oComm = New SqlClient.SqlCommand("PA_AsignaAsesorDesembACredito", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = AsignaAsesor.nCodCred
            oComm.Parameters.Add("@psCodAsesor", SqlDbType.VarChar, 150).Value = AsignaAsesor.cCodUsuAsesorNuevo
            oComm.Parameters.Add("@psCodDesembolsador", SqlDbType.VarChar, 150).Value = AsignaAsesor.cCodUsuDesemNuevo
            oComm.ExecuteNonQuery()

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


    Public Function RecuperaListaXAsignar() As outNegocio
        Dim R As New outNegocio
        Dim AsignaAsesor() As dAsignaAsesor
        Dim ListaAsesor() As String
        Dim ListaDesembolsador() As String
        Dim i As Integer
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


            oComm = New SqlClient.SqlCommand("PA_RecuperaDesembolsadores", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read

                ReDim Preserve ListaDesembolsador(i)
                ListaDesembolsador(i) = UCase(dr.Item("cValor"))
                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            oComm = New SqlClient.SqlCommand("PA_RecuperaAsesoresAsigna", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read

                ReDim Preserve ListaAsesor(i)
                ListaAsesor(i) = UCase(dr.Item("cValor"))
                i = i + 1
            End While
            dr.Close()
            dr = Nothing


            oComm = New SqlClient.SqlCommand("PA_RecuperaCreditosXAsignarAsesor_2", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read
                ReDim Preserve AsignaAsesor(i)
                AsignaAsesor(i) = New dAsignaAsesor
                AsignaAsesor(i).ListaAsesores = ListaAsesor
                AsignaAsesor(i).ListaDesembolsadores = ListaDesembolsador
                AsignaAsesor(i).nCodAge = dr.Item("nCodAge")
                AsignaAsesor(i).nCodCred = dr.Item("nCodCred")
                AsignaAsesor(i).nDeudaVigente = dr.Item("nDeudaVigente")
                AsignaAsesor(i).nMontoNeto = dr.Item("nPrestamoNeto")
                AsignaAsesor(i).nPrestamo = dr.Item("nPrestamo")
                AsignaAsesor(i).cAsesorActual = dr.Item("ccodusucob")
                AsignaAsesor(i).cCliente = dr.Item("cCliente")
                AsignaAsesor(i).cCodUsuAsesorNuevo = ""
                AsignaAsesor(i).cCodUsuDesemNuevo = ""
                AsignaAsesor(i).cDistritoDom = dr.Item("cDistritoDom")
                AsignaAsesor(i).cDistritoNeg = dr.Item("cDistritoNeg")
                AsignaAsesor(i).cDoc = dr.Item("cdoc")
                AsignaAsesor(i).cZonaDom = dr.Item("cZonaDom")
                AsignaAsesor(i).cZonaNeg = dr.Item("cZonaNeg")
                i = i + 1

            End While
            dr.Close()
            dr = Nothing
            R.ListaCreditosXAsignar = AsignaAsesor


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
    Public Function EliminaInactivo() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Try


            oComm = New SqlClient.SqlCommand("PA_EliminaClienteInactivo", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = cDoc

            oComm.ExecuteNonQuery()

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

    Public Function RegistraProspeccion() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Try


            oComm = New SqlClient.SqlCommand("PA_RegistraProspeccionEmpenio", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = oProspeccionEmp.cDoc
            oComm.Parameters.Add("@cNomApe", SqlDbType.VarChar, 150).Value = oProspeccionEmp.cNomApe
            oComm.Parameters.Add("@cMovil", SqlDbType.VarChar, 50).Value = oProspeccionEmp.cMovil
            oComm.Parameters.Add("@cDireccion", SqlDbType.VarChar, 150).Value = oProspeccionEmp.cDireccion
            oComm.Parameters.Add("@cReferencia", SqlDbType.VarChar, 150).Value = oProspeccionEmp.cReferencia
            oComm.Parameters.Add("@nPrestamo", SqlDbType.Money).Value = oProspeccionEmp.nPrestamo
            oComm.Parameters.Add("@cArticulo", SqlDbType.VarChar, 150).Value = oProspeccionEmp.cArticulo
            oComm.Parameters.Add("@nDelivery", SqlDbType.Int).Value = oProspeccionEmp.nDelivery
            oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@nLat", SqlDbType.Money).Value = oProspeccionEmp.nLatitud
            oComm.Parameters.Add("@nLon", SqlDbType.Money).Value = oProspeccionEmp.nLongitud
            oComm.Parameters.Add("@cHoraVisita", SqlDbType.VarChar, 150).Value = oProspeccionEmp.cHoraVisita

            oComm.ExecuteNonQuery()

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
    Public Function CierreOperaciones() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String

        Dim oUsu As New Usuarios

        db.Configuration.AutoDetectChangesEnabled = False

        oUsu.cCodUsu = sCodUsu
        oUsu = db.Usuarios.Find(oUsu.cCodUsu)
        db.Usuarios.Attach(oUsu)
        oUsu.nCierre = 1

        db.Entry(oUsu).Property("nCierre").IsModified = True

        Try

            db.SaveChanges()
            R.nResultado = 1
            R.sRespuesta = "Registro Correcto"


        Catch ex As Exception
            R.nResultado = 1
            R.sRespuesta = ex.Message

        End Try

        db.Configuration.AutoDetectChangesEnabled = True


        Return R

    End Function


    Public Function Movil2_RegistraDesembolsoPorProducto() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim nAmpliado As Boolean

        Dim oDatos = From C In db.Creditos
                     Where C.nCodCred = CInt(oDesembolso.psCodcred)
                     Select C

        If IsNothing(oDatos.First.bAmpliado) Then
            nAmpliado = False
        Else
            nAmpliado = oDatos.First.bAmpliado
        End If

        If nAmpliado = True Then
            D = oDesembolso.Movil2_RegistraDesembolsoAmpliacionXProd()
        Else
            D = oDesembolso.Movil2_RegistraDesembolsoPorProducto()
        End If

        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R
    End Function

    Public Function RegistraVerificacion() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String


        D = oVerificacion.RegistraVerificacion()

        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R
    End Function

    Public Function RegistraSolcitudEmpenio() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String

        D = oSolicitudPyme.RegistraSolcitudEmpenio()

        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R
    End Function

    Public Function RegistraSolcitudConsumo() As outNegocio
        Dim R As New outNegocio
        Dim D() As String

        D = oSolicitudPyme.RegistraSolcitudConsumo()

        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R
    End Function

    Public Function RegistraSolcitudPymeExcepcion() As outNegocio
        Dim R As New outNegocio
        Dim D() As String

        D = oSolicitudPyme.RegistraSolcitudPymeExcepcion()

        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R
    End Function

    Public Function RegistraSolcitudPyme() As outNegocio
        Dim R As New outNegocio
        Dim D() As String

        D = oSolicitudPyme.RegistraSolcitudPyme()

        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R
    End Function
    Public Function RecuperaPersona() As outNegocio

        Dim R As New outNegocio
        'Dim oPersonas = From P In db.vPersonasBase
        '                Where P.cDoc = cDoc.Trim
        '                Select New dPersona With {.cApellidos = P.cApellidos, .cCodGenero = P.cGenero, .cDirecDom = P.cDireccionDom, .cDistritoDom = "", .cDoc = P.cDoc, .cDocConyuge = "", .cMovil1 = "", .cMovil2 = "", .cNomBanco = "", .cNombres = P.cNombres, .cNroCtaAho = "", .cRefDirDom = P.cReferenciaDom, .cTelefFijo = "", .cZonaDom = P.cCodZonaDom, .dFechaNac = P.dFechaNac, .nCodCred = P.nCodCred, .nCodEstadoCiv = P.nEstadoCiv, .nNumDep = 0, .nSaldoCap = P.nSaldoK}

        'R.oPersonas = oPersonas

        Dim RpersonaBase As New outNegocio
        RpersonaBase = RecuperaPersonaBase(nCodAge)
        R.oPersonas = RpersonaBase.oPersonas



        Return R
    End Function
    Public Function RecuperaListaAgencias() As outNegocio
        Dim R As New outNegocio

        Dim oListaAgencias = From A In db.Agencias
                             Select A


        R.ListaAgencias = oListaAgencias

        R.nResultado = 1
        R.sRespuesta = "OK"

        Return R
    End Function

    Public Function RecuperaCuotaONLINE_Categorias() As outNegocio
        Dim R As New outNegocio
        Dim nTasa As Double
        Dim D As New dDesembolso
        Dim sMontoCuota As String()

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

        oComm = New SqlClient.SqlCommand("PA_RecuperaTasaSemAntigua", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@sDoc", SqlDbType.VarChar, 50).Value = cDoc
        oComm.Parameters.Add("@nTasaSemAnt", SqlDbType.Money).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()

        Dim nTasaSemAnt As Double = 0.0
        nTasaSemAnt = oComm.Parameters("@nTasaSemAnt").Value



        oComm = New SqlClient.SqlCommand("PA_ValidaClienteGrupalAntiguo", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = cDoc
        oComm.Parameters.Add("@pnClienteAntiguo", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()

        Dim nClienteAntiguo As Integer
        nClienteAntiguo = oComm.Parameters("@pnClienteAntiguo").Value


        If nProd = 6 Then

            If nTasaSemAnt <= 0 Then
                Dim oTasa = From T In db.TarifarioMicroSemanal
                            Where T.cCodCat = cCodCat And T.nCodAge = nCodAge And T.nPretamoIni <= nPrestamo And T.nPrestamoFin >= nPrestamo
                            Select T

                If oTasa.Count > 0 Then
                    nTasa = oTasa.First.nTasaComp
                Else
                    R.nResultado = 1
                    R.sRespuesta = "No se pudo encontrar la tasa para el rango de prestamo"
                    R.nTasaComp = 0

                    If oConex.State = ConnectionState.Open Then
                        oConex.Close()
                    End If

                    Return R
                End If

            Else
                nTasa = nTasaSemAnt
            End If


        ElseIf nProd = 7 Then

            Dim oTasa = From T In db.TarifarioMicroJuniorSemanal
                        Where T.cCodCat = cCodCat And T.nCodAge = nCodAge And T.nPretamoIni <= nPrestamo And T.nPrestamoFin >= nPrestamo
                        Select T

            If oTasa.Count > 0 Then
                nTasa = oTasa.First.nTasaComp
            Else
                R.nResultado = 1
                R.sRespuesta = "No se pudo encontrar la tasa para el rango de prestamo"
                R.nTasaComp = 0

                If oConex.State = ConnectionState.Open Then
                    oConex.Close()
                End If

                Return R
            End If

        ElseIf nProd = 8 Then
            Dim oTasa = From T In db.TarifarioMicro
                        Where T.nCodAge = nCodAge And T.nPretamoIni <= nPrestamo And T.nPrestamoFin >= nPrestamo And T.nProd = 8
                        Select T

            If oTasa.Count > 0 Then
                nTasa = oTasa.First.nTasaComp
            Else
                R.nResultado = 1
                R.sRespuesta = "No se pudo encontrar la tasa para el rango de prestamo"
                R.nTasaComp = 0

                If oConex.State = ConnectionState.Open Then
                    oConex.Close()
                End If

                Return R
            End If

        ElseIf nProd = 10 Or nProd = 9 Or nProd = 11 Or nProd = 12 Or nProd = 15 Then

            Dim oTasa = From T In db.TarifarioMicro
                        Where T.nCodAge = nCodAge And T.nPretamoIni <= nPrestamo And T.nPrestamoFin >= nPrestamo And T.nProd = nProd
                        Select T

            If oTasa.Count > 0 Then
                nTasa = oTasa.First.nTasaComp
            Else
                R.nResultado = 1
                R.sRespuesta = "No se pudo encontrar la tasa para el rango de prestamo"
                R.nTasaComp = 0

                If oConex.State = ConnectionState.Open Then
                    oConex.Close()
                End If

                Return R
            End If

            If nProd = 9 Then
                If nClienteAntiguo = 1 Then
                    nTasa = 6.4
                End If
            End If

        ElseIf nProd = 5 Then 'mdipas 2024
            Dim oTasa = From T In db.TarifarioMicro
                        Where T.nCodAge = nCodAge And T.nPretamoIni <= nPrestamo And T.nPrestamoFin >= nPrestamo And T.nProd = nProd
                        Select T

            If oTasa.Count > 0 Then
                nTasa = oTasa.First.nTasaComp
            Else
                R.nResultado = 1
                R.sRespuesta = "No se pudo encontrar la tasa para el rango de prestamo"
                R.nTasaComp = 0

                If oConex.State = ConnectionState.Open Then
                    oConex.Close()
                End If

                Return R
            End If
        Else


            Dim oTasa = From T In db.TarifarioMicro
                        Where T.cCodCat = cCodCat And T.nCodAge = nCodAge And T.nPretamoIni <= nPrestamo And T.nPrestamoFin >= nPrestamo And T.nProd = 3
                        Select T

            If oTasa.Count > 0 Then
                nTasa = oTasa.First.nTasaComp
            Else
                R.nResultado = 1
                R.sRespuesta = "No se pudo encontrar la tasa para el rango de prestamo"
                R.nTasaComp = 0

                If oConex.State = ConnectionState.Open Then
                    oConex.Close()
                End If

                Return R
            End If

        End If

        'If nCodAge = 1 Then
        '    sMontoCuota = D.MontoCuotaPorProd(Str(nPrestamo), Str(nCuotas), Str(nTasa), "1", "1", Str(nProd), "1")
        'Else
        '    sMontoCuota = D.MontoCuotaPorProd_flat(Str(nPrestamo), Str(nCuotas), Str(nTasa), "1", "1", Str(nProd), "1")
        'End If

        If nCodAge = 1 Then
            sMontoCuota = GeneaCalendarioPorProductoSimCuota(nTasa, nCuotas, nPrestamo, nProd, nCampSinInt)
        Else
            sMontoCuota = GeneaCalendarioPorProductoSimCuota_FLAT(nTasa, nCuotas, nPrestamo, nProd, nCampSinInt)
        End If


        If IsNumeric(sMontoCuota(0)) Then
            sMontoCuota(0) = Format(CDbl(sMontoCuota(0)), "#.##")
        End If

        R.nResultado = 1
        R.sRespuesta = sMontoCuota(0)
        R.nTasaComp = nTasa

        oConex.Close()

        Return R

    End Function

    Public Function RecuperaAbonosPorAprobarPorAgencia() As outNegocio
        Dim R As New outNegocio
        Dim AbonoDesemb() As dAbonoDesemb
        Dim ListaDesembolsador() As String
        Dim i As Integer
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


            oComm = New SqlClient.SqlCommand("PA_RecuperaAbonosPorAprobarPorAgencia", oConex)
            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oCredKeynua(i)
                R.oCredKeynua(i) = New clsCredKeynua

                R.oCredKeynua(i).sCodCred = Str(dr.Item("Credito"))
                R.oCredKeynua(i).cEstado = dr.Item("Estado")
                R.oCredKeynua(i).cDocumento = dr.Item("Documento")
                R.oCredKeynua(i).cCliente = dr.Item("cCliente")
                R.oCredKeynua(i).cMonto = Format(dr.Item("Monto"), "#,#0.00")
                R.oCredKeynua(i).cReconoc_Facial = dr.Item("Reconoc_Facial")
                R.oCredKeynua(i).cAmpliacion = dr.Item("Ampliacion")


                i = i + 1

            End While
            dr.Close()
            dr = Nothing


            oConex.Close()

            D(0) = "1"
            D(1) = Str(i) + " Registros Encontrados"

        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message

        End Try


        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))

        Return R

    End Function


    Public Function EjecutaOperacionesEspecialesAdmTasa() As outNegocio


        Dim R As New outNegocio
        Dim oConect As New clsConecta
        Dim oComm As SqlClient.SqlCommand
        Dim oConex As SqlConnection
        Dim sCadConexion As String
        Dim C As SqlCommand

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        oComm = New SqlClient.SqlCommand("PA_EjecutaOpeEspAdmTasasPreferenciales", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodcred
        oComm.Parameters.Add("@nTasaComp", SqlDbType.Money).Value = CDbl(oSolicitudPyme.nPrestamo)
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu


        oComm.ExecuteNonQuery()


        oConex.Close()
        R.nResultado = 1

        Return R

    End Function

    Public Function LinkKeynua() As outNegocio

        Dim R As New outNegocio
        Dim oConect As New clsConecta
        Dim oComm As SqlClient.SqlCommand
        Dim oConex As SqlConnection
        Dim sCadConexion As String
        Dim C As SqlCommand

        Dim sRutaContrato As String = ""
        Dim sRutaHojaResumen As String = ""
        Dim sRutaFichaDatos As String = ""
        Dim sRutaDDJJ As String = ""
        Dim sRutaPagare As String = ""
        Dim sRutaAnexoAmpliacion As String = ""


        Dim da As SqlDataAdapter
        Dim dtDatos As DataTable
        Dim oImagen1() As Byte
        Dim contractId As String
        Dim sToken As String

        dtDatos = New DataTable



        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        oComm = New SqlClient.SqlCommand("PA_RecuperaContratoKeynua", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodcred
        oComm.Parameters.Add("@contractId", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output


        oComm.ExecuteNonQuery()

        contractId = oComm.Parameters("@contractId").Value
        contractId = Trim(contractId)



        oConex.Close()


        Dim apiUrl As String = "https://api.keynua.com/contracts/v1/" + contractId



        ' Serializar a JSON
        Dim jsonPayload As String = ""

        Dim apikey As String = "gv4RqGRuPe1N2f9Zjr5Vr3XQKg2ntlshaUf3hTGI"
        Dim token As String = "MDg5Y2FkYmMtMzc5Mi00NGRlLWE0MjctZDU4MWZmOTE3MDBlOjhlNjUwNmRlM2JmODQ2YTlhN2ViMzQ2ZjdlOGRhNzZiOjhhYWE2NmE4NjQyOWM0NTU0NDBlZGIyMGQ3YzE3YTA2MTJlZjc1MjlkMDAxMzViNTg3MTYzY2FjY2E0NDZkNmE"

        ' Crear el cliente HTTP
        Using client As New HttpClient()

            ' Configurar encabezados
            client.DefaultRequestHeaders.Accept.Clear()
            client.DefaultRequestHeaders.Add("x-api-key", apikey)
            client.DefaultRequestHeaders.Add("authorization", token)
            'client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", token)
            client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))

            ' Crear contenido HTTP con JSON
            Dim content As New StringContent(jsonPayload, Encoding.UTF8, "application/json")


            'Enviar POST
            Dim response As HttpResponseMessage = client.GetAsync(apiUrl).GetAwaiter().GetResult()

            If response.IsSuccessStatusCode Then
                Dim result As String = response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                sToken = "https://sign.keynua.com/index.html?token="
                Using doc As JsonDocument = JsonDocument.Parse(result)
                    Dim Users = doc.RootElement.GetProperty("users")
                    For Each user In Users.EnumerateArray()
                        sToken = sToken + user.GetProperty("token").GetString()
                        Exit For
                    Next

                End Using


            Else


            End If
        End Using


        If Len(Trim(contractId)) = 0 Then

            R.sRespuesta = "No se encontró el Link"
            R.nResultado = 0
        Else
            R.sRespuesta = sToken
            R.nResultado = 1
        End If



        Return R

    End Function

    Public Function EjecutaOperacionesEspecialesAdmCambioProd() As outNegocio


        Dim R As New outNegocio
        Dim oConect As New clsConecta
        Dim oComm As SqlClient.SqlCommand
        Dim oConex As SqlConnection
        Dim sCadConexion As String
        Dim C As SqlCommand

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        oComm = New SqlClient.SqlCommand("PA_EjecutaOpeEspAdmCambioProd", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodcred
        oComm.Parameters.Add("@nProd", SqlDbType.Int).Value = CInt(oSolicitudPyme.psProd)
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu


        oComm.ExecuteNonQuery()


        oConex.Close()
        R.nResultado = 1

        Return R

    End Function


    Public Function EjecutaOperacionesEspecialesAdmAmpliacion() As outNegocio


        Dim R As New outNegocio
        Dim oConect As New clsConecta
        Dim oComm As SqlClient.SqlCommand
        Dim oConex As SqlConnection
        Dim sCadConexion As String
        Dim C As SqlCommand

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        oComm = New SqlClient.SqlCommand("PA_EjecutaOpeEspAdmAmpliacion", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodcred
        oComm.Parameters.Add("@nAmpliacion", SqlDbType.Int).Value = oSolicitudPyme.nAmpliado
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu


        oComm.ExecuteNonQuery()


        oConex.Close()
        R.nResultado = 1

        Return R

    End Function

    Public Function RecuperaDatosOperacionesEspecialesAdm() As outNegocio

        Dim R As New outNegocio
        Dim oConect As New clsConecta
        Dim oComm As SqlClient.SqlCommand
        Dim oConex As SqlConnection
        Dim sCadConexion As String
        Dim C As SqlCommand

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        oComm = New SqlClient.SqlCommand("PA_RecuperaDatosOpeEspAdm", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodcred
        oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cCliente", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nPrestamo", SqlDbType.Money).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cAnalista", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cEstado", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cAmpliado", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nEstado", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cProducto", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nTasaComp", SqlDbType.Money).Direction = ParameterDirection.Output



        oComm.ExecuteNonQuery()

        Dim oSolicitudPyme = New clsAprobacionNueva()

        oSolicitudPyme.cDoc = IIf(oComm.Parameters("@cDoc").Value = Nothing, 0, oComm.Parameters("@cDoc").Value)
        oSolicitudPyme.cCliente = IIf(oComm.Parameters("@cCliente").Value = Nothing, "", oComm.Parameters("@cCliente").Value)
        oSolicitudPyme.nPrestamo = IIf(oComm.Parameters("@nPrestamo").Value = Nothing, 0, oComm.Parameters("@nPrestamo").Value)
        oSolicitudPyme.cCodUsuCob = IIf(oComm.Parameters("@cAnalista").Value = Nothing, "", oComm.Parameters("@cAnalista").Value)
        oSolicitudPyme.cComentarioSustento = IIf(oComm.Parameters("@cEstado").Value = Nothing, "", oComm.Parameters("@cEstado").Value)
        oSolicitudPyme.cComentarioExcepcion = IIf(oComm.Parameters("@cAmpliado").Value = Nothing, "", oComm.Parameters("@cAmpliado").Value)
        oSolicitudPyme.nExcepcion = IIf(oComm.Parameters("@nEstado").Value = Nothing, 0, oComm.Parameters("@nEstado").Value)
        oSolicitudPyme.cProducto = IIf(oComm.Parameters("@cProducto").Value = Nothing, "", oComm.Parameters("@cProducto").Value)
        oSolicitudPyme.nTasaAsignada = IIf(oComm.Parameters("@nTasaComp").Value = Nothing, 0, oComm.Parameters("@nTasaComp").Value)


        oConex.Close()
        ReDim R.oAprobaciones(1)
        R.oAprobaciones(0) = oSolicitudPyme
        R.nResultado = 1

        Return R
    End Function

    Public Function RecuperaCuotaSimuladorONLINE_Categorias() As outNegocio
        Dim R As New outNegocio
        Dim D As New dDesembolso
        Dim sMontoCuota As String()

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

        oComm = New SqlClient.SqlCommand("PA_RecuperaTasaSemAntigua", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@sDoc", SqlDbType.VarChar, 50).Value = cDoc
        oComm.Parameters.Add("@nTasaSemAnt", SqlDbType.Money).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()

        Dim nTasaSemAnt As Double = 0.0
        nTasaSemAnt = oComm.Parameters("@nTasaSemAnt").Value



        oComm = New SqlClient.SqlCommand("PA_ValidaClienteGrupalAntiguo", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = cDoc
        oComm.Parameters.Add("@pnClienteAntiguo", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()

        Dim nClienteAntiguo As Integer
        nClienteAntiguo = oComm.Parameters("@pnClienteAntiguo").Value


        If nCodAge = 1 Then
            sMontoCuota = GeneaCalendarioPorProductoSimCuota(nTasa, nCuotas, nPrestamo, nProd, nCampSinInt)
        Else
            sMontoCuota = GeneaCalendarioPorProductoSimCuota_FLAT(nTasa, nCuotas, nPrestamo, nProd, nCampSinInt)
        End If


        If IsNumeric(sMontoCuota(0)) Then
            sMontoCuota(0) = Format(CDbl(sMontoCuota(0)), "#.##")
        End If

        R.nResultado = 1
        R.sRespuesta = sMontoCuota(0)
        R.nTasaComp = nTasa

        oConex.Close()

        Return R
    End Function

    Public Function CuotaPeriodoFijo(ByVal pTasaComp As Double, ByVal pTasaComision As Double,
            ByVal pnTasaSeguro As Double,
            ByVal pnCuotas As Double, ByVal vMonto As Double, ByVal pnPlazo As Integer) As Double

        Dim Pot1 As Double
        Dim nTasaTmp As Double
        Dim nTasaComiTmp As Double

        nTasaTmp = Format((((1 + (pTasaComp / 100)) ^ (pnPlazo / 30)) - 1), "#,0.0000000")

        If pnPlazo = 7 Then
            'Semanal debe venir sin igv luego de llevarlo al factor le sumo el igv
            nTasaTmp = nTasaTmp * 1.18
        End If


        nTasaComiTmp = Format((((1 + (pTasaComision / 100)) ^ (pnPlazo / 30)) - 1), "#,0.0000000")

        If pnPlazo = 14 Then
            nTasaTmp = Format((((1 + (pTasaComp / 100)) ^ (pnPlazo / 28)) - 1), "#,0.0000000")
            nTasaComiTmp = Format((((1 + (pTasaComision / 100)) ^ (pnPlazo / 28)) - 1), "#,0.0000000")
        End If

        nTasaTmp = nTasaTmp + nTasaComiTmp
        nTasaTmp = Format(nTasaTmp * (1 + (pnTasaSeguro / 100)), "#,0.0000000")

        'Obtengo la cuota de pago
        Pot1 = (1 + nTasaTmp) ^ pnCuotas
        CuotaPeriodoFijo = ((Pot1 * nTasaTmp) / (Pot1 - 1)) * vMonto
        CuotaPeriodoFijo = CDbl(Format(CuotaPeriodoFijo, "#0.00"))


    End Function

    Public Function MontoCuotaPorProd(ByVal pnPrestamo As String, ByVal pnCuotas As String, ByVal pnTasa As String, ByVal pnTipoPlazo As String, ByVal psTipoTarifario As String, ByVal psProd As String, ByVal pnPlazo As String) As String()
        Dim ValorCuota As Double
        Dim D(1) As String
        Dim pnTasaNeta As Double
        Dim nCuotaNeta As Double
        Dim nGastoFijo As Double
        Dim nTasaFactorDiario As Double
        Dim MontoCuota As Double


        If psProd.Trim = "3" Then
            pnPlazo = 1

            pnTasaNeta = 2.98

            'Diario
            If pnTipoPlazo = 1 Then

                ''llevar la tasa a factor diario
                'nTasaFactorDiario = (((1 + CDbl(pnTasa) / 100.0) ^ (1.0 / 20.0)) - 1) * 100.0
                'MontoCuota = CuotaPeriodoFijo((CDbl(nTasaFactorDiario) * 1.18), 0, 0, CDbl(pnCuotas), CDbl(pnPrestamo), 30)
                'ValorCuota = MontoCuota

                'Codigo Anterior
                'mdipas 2024
                'ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, CDbl(pnCuotas) / 30, CDbl(pnPrestamo), 30) / 30
                'nCuotaNeta = CuotaPeriodoFijo((pnTasaNeta * 1.18), 0, 0, CDbl(pnCuotas) / 30, CDbl(pnPrestamo), 30) / 30
                ValorCuota = CuotaPeriodoFijo((CDbl(Val(pnTasa)) * 1.18), 0, 0, CDbl(Val(pnCuotas)) / 30, CDbl(Val(pnPrestamo)), 30) / 30
                nCuotaNeta = CuotaPeriodoFijo((pnTasaNeta * 1.18), 0, 0, CDbl(Val(pnCuotas)) / 30, CDbl(Val(pnPrestamo)), 30) / 30
                'end
                nGastoFijo = ValorCuota - nCuotaNeta
                nGastoFijo = Math.Round((nGastoFijo / 1.18), 1, MidpointRounding.AwayFromZero)
                ValorCuota = nCuotaNeta + nGastoFijo + Format((nGastoFijo * 0.18), "#0.00")

            End If

            'Mensual
            If pnTipoPlazo = 4 Then
                ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, CDbl(pnCuotas), CDbl(pnPrestamo), CInt(pnPlazo))
            End If

        ElseIf psProd.Trim = "5" Then
            pnPlazo = 30
            ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, pnCuotas, CDbl(pnPrestamo), CInt(pnPlazo))

        ElseIf psProd.Trim = "9" Then
            pnPlazo = 28
            ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, pnCuotas, CDbl(pnPrestamo), CInt(pnPlazo))

        ElseIf psProd.Trim = "11" Or psProd.Trim = "12" Then
            pnPlazo = 14
            ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, pnCuotas, CDbl(pnPrestamo), CInt(pnPlazo))

        ElseIf psProd.Trim = "6" Or psProd.Trim = "7" Or psProd.Trim = "8" Then
            pnPlazo = 7
            'ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, CDbl(pnCuotas), CDbl(pnPrestamo), CInt(pnPlazo))
            ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa)), 0, 0, CDbl(pnCuotas), CDbl(pnPrestamo), CInt(pnPlazo))
            ValorCuota = ValorCuota
        ElseIf psProd.Trim = "10" Then
            pnPlazo = 30
            ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, pnCuotas, CDbl(pnPrestamo), CInt(pnPlazo))
            'ValorCuota = ValorCuota / 20
        Else
            'pnTasa = "16.00"
            'pnTasaNeta = 2.98
            pnPlazo = 30
            'ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, CDbl(pnCuotas), CDbl(pnPrestamo), 30)
            ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, 1, CDbl(pnPrestamo), CInt(pnPlazo))

        End If
        'mdipas 2024
        D(0) = Str(Format(ValorCuota, "##,##0.00")) 'Format(ValorCuota, "#0.00")
        MontoCuotaPorProd = D



    End Function

    Public Function MontoCuotaPorProdSinIGV(ByVal pnPrestamo As String, ByVal pnCuotas As String, ByVal pnTasa As String, ByVal pnTipoPlazo As String, ByVal psTipoTarifario As String, ByVal psProd As String, ByVal pnPlazo As String) As String()
        Dim ValorCuota As Double
        Dim D(1) As String
        Dim pnTasaNeta As Double
        Dim nCuotaNeta As Double
        Dim nGastoFijo As Double
        Dim nTasaFactorDiario As Double
        Dim MontoCuota As Double

        If psProd <> "4" Then

            pnTasaNeta = 2.9

            'If psTipoTarifario = "1" Then
            '    pnTasa = "16.30"
            '    pnTasaNeta = 2.98
            'ElseIf psTipoTarifario = "3" Then
            '    pnTasa = "17.50"
            '    pnTasaNeta = 2.98
            'Else
            '    pnTasa = "23.00"
            '    pnTasaNeta = 2.98
            'End If

            'Diario
            If pnTipoPlazo = 1 Then
                'llevar la tasa a factor diario
                nTasaFactorDiario = (((1 + CDbl(pnTasa) / 100.0) ^ (1.0 / 20.0)) - 1) * 100.0
                MontoCuota = CuotaPeriodoFijo(nTasaFactorDiario, 0, 0, CDbl(pnCuotas), CDbl(pnPrestamo), 30)
                ValorCuota = MontoCuota

            End If

            'Mensual
            If pnTipoPlazo = 4 Then
                ValorCuota = CuotaPeriodoFijo(CDbl(pnTasa), 0, 0, CDbl(pnCuotas), CDbl(pnPrestamo), CInt(pnPlazo))
            End If

        Else
            'pnTasa = "16.00"
            'pnTasaNeta = 2.98

            'ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, CDbl(pnCuotas), CDbl(pnPrestamo), 30)
            ValorCuota = CuotaPeriodoFijo(CDbl(pnTasa), 0, 0, 1, CDbl(pnPrestamo), CInt(pnPlazo))

        End If

        D(0) = Format(ValorCuota, "#0.00")
        MontoCuotaPorProdSinIGV = D


    End Function

    Public Function GeneaCalendarioPorProductoSimCuota(ByVal pnTasa As Double, ByVal pnCuotas As Integer, ByVal pnPrestamo As Double, ByVal pnProd As Integer, ByVal pnCamSinInt As Integer) As String()


        Dim nMontoCuota As Double
        Dim nMontoCuotaNeto As Double
        Dim nGastoFijo As Double
        Dim nGastoFijoIgv As Double
        Dim D(1) As String

        Dim nTasaFactorDiario As Double
        Dim nInteresEsperado As Double
        Dim nInteresNeto As Double

        Dim nTipoPlazo As Integer
        Dim nTasaCompNeto As Double
        Dim nInteresTempo As Double
        Dim nInteresNetoTempo As Double
        Dim nTasaCompNetoTempo As Double
        Dim nTasaCompTempo As Double

        Dim nSaldoTempo As Double
        Dim InteresNetoIgv As Double

        Dim nTipoTarifario As Integer



        Dim dr As SqlClient.SqlDataReader

        Dim oConex As SqlConnection
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer

        Dim MatCal() As Cuota
        Dim MatFeriados() As Date

        Dim dFechaTempo As Date
        Dim bFeriado As Boolean

        Dim nNumCuotasTempo As Integer
        Dim dFechaPrimPago As Date
        Dim nDiasTrans As Integer

        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim nMontoIntGraciaCuota As Double

        sCadConexion = oConect.gsCadConex


        Try


            oConex = New SqlConnection(sCadConexion)

            oConex.Open()
            C = oConex.CreateCommand()
            C.CommandText = "SET DATEFORMAT DMY"
            C.ExecuteNonQuery()
            C = Nothing


            nTipoPlazo = 1
            nTasaCompNeto = 2.98

            If nCodAge = 1 Then
                nTipoTarifario = 1
            Else
                nTipoTarifario = 2
            End If
            dFechaPrimPago = CDate("01/01/2019")
            nCampSinInt = 0



            oConex.Close()

            '*********************************************
            'CARGA FERIADOS
            '*********************************************
            oConex = New SqlConnection(sCadConexion)

            oConex.Open()
            C = oConex.CreateCommand()
            C.CommandText = "SET DATEFORMAT DMY"
            C.ExecuteNonQuery()
            C = Nothing

            oComm = New SqlClient.SqlCommand("PA_RecuperaFeriados", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                i = i + 1
                ReDim Preserve MatFeriados(i)
                MatFeriados(i - 1) = dr.Item("dFecha")

            End While

            dr.Close()
            dr = Nothing

            '****************************************************************************

            If pnProd = 10 Then

                nDiasTrans = 30
                D = MontoCuotaPorProd(Str(pnPrestamo).Trim, Str(nCuotas).Trim, Str(pnTasa).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(pnProd).Trim, Str(nDiasTrans).Trim)
                nMontoCuota = CDbl(D(0))
                D = MontoCuotaPorProd(Str(pnPrestamo).Trim, Str(nCuotas).Trim, Str(nTasaCompNeto).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(pnProd).Trim, Str(nDiasTrans).Trim)
                nMontoCuotaNeto = CDbl(D(0))

                nGastoFijo = nMontoCuota - nMontoCuotaNeto
                nGastoFijo = Math.Round((nGastoFijo / 1.18), 1, MidpointRounding.AwayFromZero)
                nGastoFijoIgv = Math.Round(nGastoFijo * 0.18, 1, MidpointRounding.AwayFromZero)

            ElseIf pnProd = 4 Then


                nDiasTrans = DateDiff(DateInterval.Day, CDate(Format(Now(), "dd/MM/yyyy")), dFechaPrimPago)
                D = MontoCuotaPorProd(Str(pnPrestamo).Trim, Str(pnCuotas).Trim, Str(pnTasa).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(pnProd).Trim, Str(nDiasTrans).Trim)
                nMontoCuota = CDbl(D(0))
                D = MontoCuotaPorProd(Str(pnPrestamo).Trim, Str(pnCuotas).Trim, Str(nTasaCompNeto).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(pnProd).Trim, Str(nDiasTrans).Trim)
                nMontoCuotaNeto = CDbl(D(0))

                nGastoFijo = nMontoCuota - nMontoCuotaNeto
                nGastoFijo = Math.Round((nGastoFijo / 1.18), 1, MidpointRounding.AwayFromZero)
                nGastoFijoIgv = Math.Round(nGastoFijo * 0.18, 1, MidpointRounding.AwayFromZero)

            ElseIf pnProd = 9 Then
                nDiasTrans = 28
                D = MontoCuotaPorProd(Str(pnPrestamo).Trim, Str(pnCuotas).Trim, Str((pnTasa)).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(pnProd).Trim, Str(nDiasTrans).Trim)
                nMontoCuota = CDbl(D(0))
                D = MontoCuotaPorProd(Str(pnPrestamo).Trim, Str(pnCuotas).Trim, Str((nTasaCompNeto)).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(pnProd).Trim, Str(nDiasTrans).Trim)
                nMontoCuotaNeto = CDbl(D(0))

                nGastoFijo = nMontoCuota - nMontoCuotaNeto
                nGastoFijo = Math.Round((nGastoFijo / 1.18), 1, MidpointRounding.AwayFromZero)
                nGastoFijoIgv = Math.Round(nGastoFijo * 0.18, 1, MidpointRounding.AwayFromZero)

            ElseIf pnProd = 11 Or pnProd = 12 Then
                nDiasTrans = 14
                D = MontoCuotaPorProd(Str(pnPrestamo).Trim, Str(pnCuotas).Trim, Str((pnTasa)).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(pnProd).Trim, Str(nDiasTrans).Trim)
                nMontoCuota = CDbl(D(0))
                D = MontoCuotaPorProd(Str(pnPrestamo).Trim, Str(pnCuotas).Trim, Str((nTasaCompNeto)).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(pnProd).Trim, Str(nDiasTrans).Trim)
                nMontoCuotaNeto = CDbl(D(0))

                nGastoFijo = nMontoCuota - nMontoCuotaNeto
                nGastoFijo = Math.Round((nGastoFijo / 1.18), 1, MidpointRounding.AwayFromZero)
                nGastoFijoIgv = Math.Round(nGastoFijo * 0.18, 1, MidpointRounding.AwayFromZero)


            ElseIf pnProd = 6 Or pnProd = 7 Or pnProd = 8 Or pnProd = 15 Then
                nDiasTrans = 7
                D = MontoCuotaPorProd(Str(pnPrestamo).Trim, Str(pnCuotas).Trim, Str((pnTasa)).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(pnProd).Trim, Str(nDiasTrans).Trim)
                nMontoCuota = CDbl(D(0))
                D = MontoCuotaPorProd(Str(pnPrestamo).Trim, Str(pnCuotas).Trim, Str((nTasaCompNeto)).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(pnProd).Trim, Str(nDiasTrans).Trim)
                nMontoCuotaNeto = CDbl(D(0))

                nGastoFijo = nMontoCuota - nMontoCuotaNeto
                nGastoFijo = Math.Round((nGastoFijo / 1.18), 1, MidpointRounding.AwayFromZero)
                nGastoFijoIgv = Math.Round(nGastoFijo * 0.18, 1, MidpointRounding.AwayFromZero)

            ElseIf pnProd = 5 Then

                nDiasTrans = DateDiff(DateInterval.Day, CDate(Format(Now(), "dd/MM/yyyy")), dFechaPrimPago)
                D = MontoCuotaPorProd(Str(pnPrestamo).Trim, Str(pnCuotas).Trim, Str(pnTasa).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(pnProd).Trim, "30")
                nMontoCuota = CDbl(D(0))
                D = MontoCuotaPorProd(Str(pnPrestamo).Trim, Str(pnCuotas).Trim, Str(nTasaCompNeto).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(pnProd).Trim, "30")
                nMontoCuotaNeto = CDbl(D(0))

                nGastoFijo = nMontoCuota - nMontoCuotaNeto
                nGastoFijo = Math.Round((nGastoFijo / 1.18), 1, MidpointRounding.AwayFromZero)
                nGastoFijoIgv = Math.Round(nGastoFijo * 0.18, 1, MidpointRounding.AwayFromZero)

            Else


                D = MontoCuotaPorProdSinIGV(Str(pnPrestamo).Trim, Str(pnCuotas).Trim, Str(pnTasa).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(pnProd).Trim, Str(nDiasTrans).Trim)
                nInteresEsperado = (CDbl(D(0)) * pnCuotas) - pnPrestamo

                D = MontoCuotaPorProdSinIGV(Str(pnPrestamo).Trim, Str(pnCuotas).Trim, Str(nTasaCompNeto).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(pnProd).Trim, Str(nDiasTrans).Trim)
                nInteresNeto = (CDbl(D(0)) * pnCuotas) - pnPrestamo

                'D = MontoCuota2(pnPrestamo, pnCuotas, pnTasa, nTipoPlazo)
                'nMontoCuota = CDbl(D(0))
                'D = MontoCuota2(pnPrestamo, pnCuotas, nTasaCompNeto, nTipoPlazo)
                'nMontoCuotaNeto = CDbl(D(0))


                D = MontoCuotaPorProd(Str(pnPrestamo), Str(pnCuotas), Str(pnTasa), "1", "1", Str(pnProd), "1")
                nMontoCuota = CDbl(D(0))
                D = MontoCuotaPorProd(Str(pnPrestamo), Str(pnCuotas), Str(nTasaCompNeto), "1", "1", Str(pnProd), "1")
                nMontoCuotaNeto = CDbl(D(0))


                nGastoFijo = nMontoCuota - nMontoCuotaNeto
                nGastoFijo = Math.Round((nGastoFijo / 1.18), 1, MidpointRounding.AwayFromZero)
                nGastoFijoIgv = Math.Round(nGastoFijo * 0.18, 1, MidpointRounding.AwayFromZero)

                'D = MontoCuotaPorProdSinIGV(Str(pnPrestamo).Trim, Str(pnCuotas).Trim, Str(Math.Round(nTasaCompNeto * 1.18, 2)).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(pnProd).Trim, Str(nDiasTrans).Trim)
                'nMontoCuotaNeto = CDbl(D(0))


            End If

            '****************************************************************************

            'Si producto Diario Pasar tasas de calculo a factor diario
            If pnProd = 3 Then
                nTasaCompNetoTempo = Math.Round((((1 + nTasaCompNeto / 100.0) ^ (1 / 20.0)) - 1) * 100.0, 4)
                nTasaCompTempo = Math.Round((((1 + pnTasa / 100.0) ^ (1 / 20.0)) - 1) * 100.0, 4)
            Else
                nTasaCompNetoTempo = nTasaCompNeto
                nTasaCompTempo = pnTasa
            End If

            Dim bCorrecto As Boolean
            Dim nDiaPago As Integer
            Dim nMesPago As Integer
            Dim bFechaEnc As Boolean
            Dim nDiasFeriado As Integer
            nDiasFeriado = 0

            bCorrecto = False
            nCampSinInt = pnCamSinInt

            nMontoIntGraciaCuota = 0
            If nCampSinInt = 1 Then
                nMontoIntGraciaCuota = Math.Round(((Math.Round((((1 + pnTasa / 100.0) ^ (7 / 30.0)) - 1), 4) * nPrestamo)) / pnCuotas, 2)
                nMontoCuota = nMontoCuota + nMontoIntGraciaCuota
            End If

            Do While Not bCorrecto


                ReDim MatCal(pnCuotas)
                dFechaTempo = CDate(Format(Now(), "dd/MM/yyyy"))

                If pnProd = 6 Or pnProd = 7 Or pnProd = 8 Or pnProd = 10 Or pnProd = 11 Or pnProd = 12 Or pnProd = 15 Or pnProd = 9 Then
                    dFechaTempo = CDate(Format(Now(), "dd/MM/yyyy"))
                Else
                    dFechaTempo = dFechaTempo.AddDays(1)

                End If

                '*******************************************************************
                'Si es campaña sin interes correr 1 dia si es diario
                ' o correr 7 dias si es semanal o mensual
                '*******************************************************************
                If nCampSinInt = 1 Then
                    dFechaTempo = CDate(Format(Now(), "dd/MM/yyyy"))
                    If pnProd = 6 Or pnProd = 7 Or pnProd = 8 Or pnProd = 10 Then
                        dFechaTempo = dFechaTempo.AddDays(7)
                    Else
                        dFechaTempo = dFechaTempo.AddDays(6)
                    End If

                End If

                nSaldoTempo = pnPrestamo
                nNumCuotasTempo = pnCuotas
                nDiaPago = dFechaPrimPago.Day

                For i = 0 To pnCuotas - 1

                    If pnProd = 4 Then
                        nDiasFeriado = 0
                        dFechaTempo = dFechaPrimPago
                    ElseIf pnProd = 6 Or pnProd = 7 Or pnProd = 8 Or pnProd = 15 Then

                        'Restar los feriados movidos para no cambiar de dia de fecha
                        If nDiasFeriado > 0 And nDiasFeriado < 7 Then
                            dFechaTempo = dFechaTempo.AddDays(7 - nDiasFeriado)
                        Else
                            dFechaTempo = dFechaTempo.AddDays(7)
                        End If
                        'Volver a cero variable que acumula feriados
                        nDiasFeriado = 0

                    ElseIf pnProd = 5 Then
                        nDiasFeriado = 0
                        If i = 0 Then
                            dFechaTempo = dFechaPrimPago
                        Else
                            dFechaTempo = DateAdd(DateInterval.Day, 1, MatCal(i - 1).Fecha)
                            'dFechaTempo = CDate("01/" + Microsoft.VisualBasic.Right("00" + Str(dFechaTempo.Month), 2) + "/" + Str(dFechaTempo.Year).Trim)
                            nMesPago = dFechaTempo.Month
                            bFechaEnc = False
                            Do While Not bFechaEnc
                                If dFechaTempo.Day = nDiaPago Or DateDiff(DateInterval.Month, MatCal(i - 1).Fecha, dFechaTempo) >= 2 Then
                                    bFechaEnc = True
                                    Exit Do
                                End If
                                dFechaTempo = dFechaTempo.AddDays(1)
                            Loop
                        End If
                    ElseIf pnProd = 10 Then
                        nDiasFeriado = 0
                        dFechaTempo = dFechaTempo.AddDays(30)
                    ElseIf pnProd = 9 Then
                        nDiasFeriado = 0
                        dFechaTempo = dFechaTempo.AddDays(28)
                    ElseIf pnProd = 11 Or pnProd = 12 Then
                        nDiasFeriado = 0
                        dFechaTempo = dFechaTempo.AddDays(14)
                    Else
                        nDiasFeriado = 0
                        dFechaTempo = dFechaTempo.AddDays(1)
                    End If


                    If dFechaTempo.DayOfWeek = DayOfWeek.Saturday Then
                        dFechaTempo = dFechaTempo.AddDays(1)
                        nDiasFeriado = nDiasFeriado + 1
                    End If


                    If dFechaTempo.DayOfWeek = DayOfWeek.Sunday Then
                        dFechaTempo = dFechaTempo.AddDays(1)
                        nDiasFeriado = nDiasFeriado + 1
                    End If


                    'SI es feriado incrementa en 1
                    bFeriado = False
                    For j = 0 To MatFeriados.Count - 1
                        If MatFeriados(j).Day = dFechaTempo.Day And MatFeriados(j).Month = dFechaTempo.Month And MatFeriados(j).Year = dFechaTempo.Year Then
                            bFeriado = True
                            Exit For
                        End If
                    Next
                    If bFeriado Then
                        dFechaTempo = dFechaTempo.AddDays(1)
                        nDiasFeriado = nDiasFeriado + 1
                        bFeriado = False
                    End If

                    'SI es feriado incrementa en 1
                    bFeriado = False
                    For j = 0 To MatFeriados.Count - 1
                        If MatFeriados(j).Day = dFechaTempo.Day And MatFeriados(j).Month = dFechaTempo.Month And MatFeriados(j).Year = dFechaTempo.Year Then
                            bFeriado = True
                            Exit For
                        End If
                    Next
                    If bFeriado Then
                        dFechaTempo = dFechaTempo.AddDays(1)
                        nDiasFeriado = nDiasFeriado + 1
                        bFeriado = False
                    End If

                    'SI es feriado incrementa en 1
                    bFeriado = False
                    For j = 0 To MatFeriados.Count - 1
                        If MatFeriados(j).Day = dFechaTempo.Day And MatFeriados(j).Month = dFechaTempo.Month And MatFeriados(j).Year = dFechaTempo.Year Then
                            bFeriado = True
                            Exit For
                        End If
                    Next
                    If bFeriado Then
                        dFechaTempo = dFechaTempo.AddDays(1)
                        nDiasFeriado = nDiasFeriado + 1
                        bFeriado = False
                    End If

                    If dFechaTempo.DayOfWeek = DayOfWeek.Saturday Then
                        dFechaTempo = dFechaTempo.AddDays(1)
                        nDiasFeriado = nDiasFeriado + 1
                    End If

                    If dFechaTempo.DayOfWeek = DayOfWeek.Sunday Then
                        dFechaTempo = dFechaTempo.AddDays(1)
                        nDiasFeriado = nDiasFeriado + 1
                    End If

                    MatCal(i).Fecha = CDate(Format(dFechaTempo, "dd/MM/yyyy"))
                    MatCal(i).NumCuota = i + 1



                    If pnProd = 4 Or pnProd = 9 Then
                        If i = 0 Then
                            nDiasTrans = DateDiff(DateInterval.Day, CDate(Format(Now(), "dd/MM/yyyy")), MatCal(i).Fecha)
                        Else
                            nDiasTrans = DateDiff(DateInterval.Day, MatCal(i - 1).Fecha, MatCal(i).Fecha)
                        End If

                        nInteresTempo = Format(CDbl(Format((((1 + (nTasaCompTempo / 100)) ^ (nDiasTrans / 30)) - 1), "#,0.0000000")) * nSaldoTempo, "#0.00")
                        nInteresNetoTempo = Format(CDbl(Format((((1 + (nTasaCompNetoTempo / 100)) ^ (nDiasTrans / 30)) - 1), "#,0.0000000")) * nSaldoTempo, "#0.00")

                    ElseIf pnProd = 6 Or pnProd = 7 Or pnProd = 8 Or pnProd = 15 Then
                        '**** Antes estaba asi y calculaba los dias transcurridos
                        'If i = 0 Then
                        '    nDiasTrans = DateDiff(DateInterval.Day, CDate(Format(Now(), "dd/MM/yyyy")), MatCal(i).Fecha)
                        'Else
                        '    nDiasTrans = DateDiff(DateInterval.Day, MatCal(i - 1).Fecha, MatCal(i).Fecha)
                        'End If

                        '**** Se cambio para que sea a plazo fijo
                        nDiasTrans = 7

                        nInteresTempo = Format(CDbl(Format((((1 + (nTasaCompTempo / 100)) ^ (nDiasTrans / 30)) - 1), "#,0.0000000")) * nSaldoTempo, "#0.00")
                        nInteresNetoTempo = Format(CDbl(Format((((1 + (nTasaCompNetoTempo / 100)) ^ (nDiasTrans / 30)) - 1), "#,0.0000000")) * nSaldoTempo, "#0.00")

                    ElseIf pnProd = 5 Or pnProd = 10 Then

                        If i = 0 Then
                            nDiasTrans = DateDiff(DateInterval.Day, CDate(Format(Now(), "dd/MM/yyyy")), MatCal(i).Fecha)
                        Else

                            nDiasTrans = DateDiff(DateInterval.Day, MatCal(i - 1).Fecha, MatCal(i).Fecha)
                        End If


                        nInteresTempo = Format(CDbl(Format((((1 + (nTasaCompTempo / 100)) ^ (nDiasTrans / 30)) - 1), "#,0.0000000")) * nSaldoTempo, "#0.00")
                        nInteresNetoTempo = Format(CDbl(Format((((1 + (nTasaCompNetoTempo / 100)) ^ (nDiasTrans / 30)) - 1), "#,0.0000000")) * nSaldoTempo, "#0.00")


                    ElseIf pnProd = 11 Or pnProd = 12 Then

                        If i = 0 Then
                            nDiasTrans = DateDiff(DateInterval.Day, CDate(Format(Now(), "dd/MM/yyyy")), MatCal(i).Fecha)
                        Else

                            nDiasTrans = DateDiff(DateInterval.Day, MatCal(i - 1).Fecha, MatCal(i).Fecha)
                        End If


                        nInteresTempo = Format(CDbl(Format((((1 + (nTasaCompTempo / 100)) ^ (nDiasTrans / 28)) - 1), "#,0.0000000")) * nSaldoTempo, "#0.00")
                        nInteresNetoTempo = Format(CDbl(Format((((1 + (nTasaCompNetoTempo / 100)) ^ (nDiasTrans / 28)) - 1), "#,0.0000000")) * nSaldoTempo, "#0.00")


                    Else
                        nInteresTempo = Format((nTasaCompTempo / 100.0) * nSaldoTempo, "#0.00")
                        nInteresNetoTempo = Format((nTasaCompNetoTempo / 100.0) * nSaldoTempo, "#0.00")
                    End If



                    MatCal(i).Interes = nInteresTempo + nMontoIntGraciaCuota
                    MatCal(i).InteresNeto = nInteresNetoTempo
                    nGastoFijo = MatCal(i).Interes - MatCal(i).InteresNeto
                    MatCal(i).Gasto = nGastoFijo
                    MatCal(i).Igv = CDbl(Format(MatCal(i).Interes * 0.18, "#0.00"))

                    If i = (pnCuotas - 1) Then
                        MatCal(i).Capital = Format(nSaldoTempo, "#0.00")
                    Else
                        MatCal(i).Capital = Format(nMontoCuota - MatCal(i).Interes - MatCal(i).Igv, "#0.00")
                    End If
                    If MatCal(i).Capital < 0 Then
                        MatCal(i).Capital = 0
                    End If

                    nSaldoTempo = nSaldoTempo - MatCal(i).Capital

                Next i


                If pnCuotas > 1 Then
                    If pnProd = 3 Then

                        i = pnCuotas - 1
                        If (CDbl(Format(MatCal(i).Capital + MatCal(i).Interes + MatCal(i).Igv, "#0.00")) - nMontoCuota) > 0 Then
                            bCorrecto = False
                            nTasaCompTempo = nTasaCompTempo - 0.0001
                        ElseIf (CDbl(Format(MatCal(i).Capital + MatCal(i).Interes + MatCal(i).Igv, "#0.00")) - nMontoCuota) < -1.0 Then
                            nTasaCompTempo = nTasaCompTempo + 0.0001
                        Else
                            bCorrecto = True
                        End If
                    ElseIf pnProd = 11 Or pnProd = 12 Or nCampSinInt = 1 Then
                        i = pnCuotas - 1
                        If (CDbl(Format(MatCal(i).Capital + MatCal(i).Interes + MatCal(i).Igv, "#0.00")) - nMontoCuota) > 0 Then
                            bCorrecto = False
                            nMontoCuota = nMontoCuota + 0.01
                        ElseIf (CDbl(Format(MatCal(i).Capital + MatCal(i).Interes + MatCal(i).Igv, "#0.00")) - nMontoCuota) < -1.0 Then
                            nMontoCuota = nMontoCuota - 0.01
                        Else
                            bCorrecto = True
                        End If


                    Else   ' Para Producto Mensual
                        'i = pnCuotas - 1
                        'If (CDbl(Format(MatCal(i).Capital + MatCal(i).Interes + MatCal(i).Igv, "#0.00")) - nMontoCuota) > 0 Then
                        '    bCorrecto = False
                        '    nMontoCuota = nMontoCuota + 0.01
                        'ElseIf (CDbl(Format(MatCal(i).Capital + MatCal(i).Interes + MatCal(i).Igv, "#0.00")) - nMontoCuota) < -1.0 Then
                        '    nMontoCuota = nMontoCuota - 0.01
                        'Else
                        '    bCorrecto = True
                        'End If

                        bCorrecto = True

                    End If

                Else
                    bCorrecto = True
                End If

            Loop



            If oConex.State = ConnectionState.Open Then
                oConex.Close()
            End If

            D(0) = Format(MatCal(0).Capital + MatCal(0).Interes + MatCal(0).Igv, "#,#0.00")
            GeneaCalendarioPorProductoSimCuota = D

        Catch ex As Exception
            D(0) = ex.Message
            GeneaCalendarioPorProductoSimCuota = D

        End Try


    End Function

    Public Function GeneaCalendarioPorProductoSimCuota_FLAT(ByVal pnTasa As Double, ByVal pnCuotas As Integer, ByVal pnPrestamo As Double, ByVal pnProd As Integer, ByVal pnCamSinInt As Integer) As String()

        Dim nMontoCuota As Double
        Dim nMontoCuotaNeto As Double
        Dim nGastoFijo As Double
        Dim nGastoFijoIgv As Double
        Dim D(1) As String

        Dim nTasaFactorDiario As Double
        Dim nInteresEsperado As Double
        Dim nInteresNeto As Double

        Dim nTipoPlazo As Integer
        Dim nTasaCompNeto As Double
        Dim nInteresTempo As Double
        Dim nInteresNetoTempo As Double
        Dim nTasaCompNetoTempo As Double
        Dim nTasaCompTempo As Double

        Dim nSaldoTempo As Double
        Dim InteresNetoIgv As Double

        Dim nTipoTarifario As Integer

        Dim dr As SqlClient.SqlDataReader

        Dim oConex As SqlConnection
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer

        Dim MatCal() As Cuota
        Dim MatFeriados() As Date

        Dim dFechaTempo As Date
        Dim bFeriado As Boolean

        Dim nNumCuotasTempo As Integer
        Dim dFechaPrimPago As Date
        Dim nDiasTrans As Integer

        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim nMontoIntGraciaCuota As Double

        sCadConexion = oConect.gsCadConex


        Try


            oConex = New SqlConnection(sCadConexion)

            oConex.Open()
            C = oConex.CreateCommand()
            C.CommandText = "SET DATEFORMAT DMY"
            C.ExecuteNonQuery()
            C = Nothing



            nTipoPlazo = 1
            nTasaCompNeto = 2.98

            If nCodAge = 1 Then
                nTipoTarifario = 1
            Else
                nTipoTarifario = 2
            End If
            dFechaPrimPago = CDate("01/01/2019")
            nCampSinInt = 0



            oConex.Close()

            '*********************************************
            'CARGA FERIADOS
            '*********************************************
            oConex = New SqlConnection(sCadConexion)

            oConex.Open()
            C = oConex.CreateCommand()
            C.CommandText = "SET DATEFORMAT DMY"
            C.ExecuteNonQuery()
            C = Nothing

            oComm = New SqlClient.SqlCommand("PA_RecuperaFeriados", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                i = i + 1
                ReDim Preserve MatFeriados(i)
                MatFeriados(i - 1) = dr.Item("dFecha")

            End While

            dr.Close()
            dr = Nothing

            '****************************************************************************

            If pnProd = 10 Or pnProd = 4 Or pnProd = 5 Then

                nDiasTrans = 30
                nMontoCuota = (pnPrestamo / pnCuotas) + ((pnTasa * 1.0 / 100.0) * pnCuotas * pnPrestamo) / pnCuotas + ((pnTasa * 1.0 / 100.0) * pnCuotas * pnPrestamo) * 0.18 / pnCuotas
                nMontoCuotaNeto = (pnPrestamo / pnCuotas) + ((nTasaCompNeto * 1.0 / 100.0) * pnCuotas * pnPrestamo) / pnCuotas + ((nTasaCompNeto * 1.0 / 100.0) * pnCuotas * pnPrestamo) * 0.18 / pnCuotas

                nGastoFijo = nMontoCuota - nMontoCuotaNeto
                nGastoFijo = Math.Round((nGastoFijo / 1.18), 1, MidpointRounding.AwayFromZero)
                nGastoFijoIgv = Math.Round(nGastoFijo * 0.18, 1, MidpointRounding.AwayFromZero)


            ElseIf pnProd = 6 Or pnProd = 7 Or pnProd = 8 Or pnProd = 15 Then

                nDiasTrans = 7
                nMontoCuota = (pnPrestamo / pnCuotas) + ((pnTasa * 1.0 / 100.0) * (pnCuotas / 4) * pnPrestamo) / pnCuotas + ((pnTasa * 1.0 / 100.0) * (pnCuotas / 4) * pnPrestamo) * 0.18 / pnCuotas
                nMontoCuotaNeto = (pnPrestamo / pnCuotas) + ((nTasaCompNeto * 1.0 / 100.0) * (pnCuotas / 4) * pnPrestamo) / pnCuotas + ((nTasaCompNeto * 1.0 / 100.0) * (pnCuotas / 4) * pnPrestamo) * 0.18 / pnCuotas


                nGastoFijo = nMontoCuota - nMontoCuotaNeto
                nGastoFijo = Math.Round((nGastoFijo / 1.18), 1, MidpointRounding.AwayFromZero)
                nGastoFijoIgv = Math.Round(nGastoFijo * 0.18, 1, MidpointRounding.AwayFromZero)

            ElseIf pnProd = 9 Then

                nDiasTrans = 28
                nMontoCuota = (pnPrestamo / pnCuotas) + ((pnTasa * 1.0 / 100.0) * (pnCuotas / 2) * pnPrestamo) / pnCuotas + ((pnTasa * 1.0 / 100.0) * (pnCuotas / 2) * pnPrestamo) * 0.18 / pnCuotas
                nMontoCuotaNeto = (pnPrestamo / pnCuotas) + ((nTasaCompNeto * 1.0 / 100.0) * (pnCuotas / 2) * pnPrestamo) / pnCuotas + ((nTasaCompNeto * 1.0 / 100.0) * (pnCuotas / 2) * pnPrestamo) * 0.18 / pnCuotas


                nGastoFijo = nMontoCuota - nMontoCuotaNeto
                nGastoFijo = Math.Round((nGastoFijo / 1.18), 1, MidpointRounding.AwayFromZero)
                nGastoFijoIgv = Math.Round(nGastoFijo * 0.18, 1, MidpointRounding.AwayFromZero)

            ElseIf pnProd = 11 Or pnProd = 12 Then

                nDiasTrans = 14
                nMontoCuota = (pnPrestamo / pnCuotas) + ((pnTasa * 1.0 / 100.0) * (pnCuotas / 2) * pnPrestamo) / pnCuotas + ((pnTasa * 1.0 / 100.0) * (pnCuotas / 2) * pnPrestamo) * 0.18 / pnCuotas
                nMontoCuotaNeto = (pnPrestamo / pnCuotas) + ((nTasaCompNeto * 1.0 / 100.0) * (pnCuotas / 2) * pnPrestamo) / pnCuotas + ((nTasaCompNeto * 1.0 / 100.0) * (pnCuotas / 2) * pnPrestamo) * 0.18 / pnCuotas


                nGastoFijo = nMontoCuota - nMontoCuotaNeto
                nGastoFijo = Math.Round((nGastoFijo / 1.18), 1, MidpointRounding.AwayFromZero)
                nGastoFijoIgv = Math.Round(nGastoFijo * 0.18, 1, MidpointRounding.AwayFromZero)

            Else


                nDiasTrans = 1
                nMontoCuota = (pnPrestamo / pnCuotas) + ((pnTasa * 1.0 / 100.0) * (pnCuotas / 20) * pnPrestamo) / pnCuotas + ((pnTasa * 1.0 / 100.0) * (pnCuotas / 20) * pnPrestamo) * 0.18 / pnCuotas
                nMontoCuotaNeto = (pnPrestamo / pnCuotas) + ((nTasaCompNeto * 1.0 / 100.0) * (pnCuotas / 20) * pnPrestamo) / pnCuotas + ((nTasaCompNeto * 1.0 / 100.0) * (pnCuotas / 20) * pnPrestamo) * 0.18 / pnCuotas


                nGastoFijo = nMontoCuota - nMontoCuotaNeto
                nGastoFijo = Math.Round((nGastoFijo / 1.18), 1, MidpointRounding.AwayFromZero)
                nGastoFijoIgv = Math.Round(nGastoFijo * 0.18, 1, MidpointRounding.AwayFromZero)

                'D = MontoCuotaPorProdSinIGV(Str(pnPrestamo).Trim, Str(pnCuotas).Trim, Str(Math.Round(nTasaCompNeto * 1.18, 2)).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(pnProd).Trim, Str(nDiasTrans).Trim)
                'nMontoCuotaNeto = CDbl(D(0))


            End If

            '****************************************************************************


            'DEFINE TASAS
            nTasaCompNetoTempo = nTasaCompNeto
            nTasaCompTempo = pnTasa


            Dim bCorrecto As Boolean
            Dim nDiaPago As Integer
            Dim nMesPago As Integer
            Dim bFechaEnc As Boolean
            Dim nDiasFeriado As Integer
            nDiasFeriado = 0

            bCorrecto = False
            nCampSinInt = pnCamSinInt

            nMontoIntGraciaCuota = 0
            If nCampSinInt = 1 Then
                nMontoIntGraciaCuota = Math.Round(((Math.Round((((1 + pnTasa / 100.0) ^ (7 / 30.0)) - 1), 4) * nPrestamo)) / pnCuotas, 2)
                nMontoCuota = nMontoCuota + nMontoIntGraciaCuota
            End If

            Do While Not bCorrecto


                ReDim MatCal(pnCuotas)
                dFechaTempo = CDate(Format(Now(), "dd/MM/yyyy"))

                If pnProd = 3 Then
                    dFechaTempo = dFechaTempo.AddDays(1)
                Else
                    dFechaTempo = CDate(Format(Now(), "dd/MM/yyyy"))

                End If

                '*******************************************************************
                'Si es campaña sin interes correr 1 dia si es diario
                ' o correr 7 dias si es semanal o mensual
                '*******************************************************************
                If nCampSinInt = 1 Then
                    dFechaTempo = CDate(Format(Now(), "dd/MM/yyyy"))
                    If pnProd = 6 Or pnProd = 7 Or pnProd = 8 Or pnProd = 10 Then
                        dFechaTempo = dFechaTempo.AddDays(7)
                    Else
                        dFechaTempo = dFechaTempo.AddDays(6)
                    End If

                End If

                nSaldoTempo = pnPrestamo
                nNumCuotasTempo = pnCuotas
                nDiaPago = dFechaPrimPago.Day

                For i = 0 To pnCuotas - 1

                    If pnProd = 4 Then
                        nDiasFeriado = 0
                        dFechaTempo = dFechaPrimPago
                    ElseIf pnProd = 6 Or pnProd = 7 Or pnProd = 8 Or nProd = 15 Then

                        'Restar los feriados movidos para no cambiar de dia de fecha
                        If nDiasFeriado > 0 And nDiasFeriado < 7 Then
                            dFechaTempo = dFechaTempo.AddDays(7 - nDiasFeriado)
                        Else
                            dFechaTempo = dFechaTempo.AddDays(7)
                        End If
                        'Volver a cero variable que acumula feriados
                        nDiasFeriado = 0

                    ElseIf pnProd = 5 Then
                        nDiasFeriado = 0
                        If i = 0 Then
                            dFechaTempo = dFechaPrimPago
                        Else
                            dFechaTempo = DateAdd(DateInterval.Day, 1, MatCal(i - 1).Fecha)
                            'dFechaTempo = CDate("01/" + Microsoft.VisualBasic.Right("00" + Str(dFechaTempo.Month), 2) + "/" + Str(dFechaTempo.Year).Trim)
                            nMesPago = dFechaTempo.Month
                            bFechaEnc = False
                            Do While Not bFechaEnc
                                If dFechaTempo.Day = nDiaPago Or DateDiff(DateInterval.Month, MatCal(i - 1).Fecha, dFechaTempo) >= 2 Then
                                    bFechaEnc = True
                                    Exit Do
                                End If
                                dFechaTempo = dFechaTempo.AddDays(1)
                            Loop
                        End If
                    ElseIf pnProd = 10 Then
                        nDiasFeriado = 0
                        dFechaTempo = dFechaTempo.AddDays(30)

                    ElseIf nProd = 9 Then
                        nDiasFeriado = 0
                        dFechaTempo = dFechaTempo.AddDays(28)

                    ElseIf nProd = 11 Or nProd = 12 Then
                        nDiasFeriado = 0
                        dFechaTempo = dFechaTempo.AddDays(14)

                    Else
                        nDiasFeriado = 0
                        dFechaTempo = dFechaTempo.AddDays(1)
                    End If


                    If dFechaTempo.DayOfWeek = DayOfWeek.Saturday Then
                        dFechaTempo = dFechaTempo.AddDays(1)
                        nDiasFeriado = nDiasFeriado + 1
                    End If


                    If dFechaTempo.DayOfWeek = DayOfWeek.Sunday Then
                        dFechaTempo = dFechaTempo.AddDays(1)
                        nDiasFeriado = nDiasFeriado + 1
                    End If


                    'SI es feriado incrementa en 1
                    bFeriado = False
                    For j = 0 To MatFeriados.Count - 1
                        If MatFeriados(j).Day = dFechaTempo.Day And MatFeriados(j).Month = dFechaTempo.Month And MatFeriados(j).Year = dFechaTempo.Year Then
                            bFeriado = True
                            Exit For
                        End If
                    Next
                    If bFeriado Then
                        dFechaTempo = dFechaTempo.AddDays(1)
                        nDiasFeriado = nDiasFeriado + 1
                        bFeriado = False
                    End If

                    'SI es feriado incrementa en 1
                    bFeriado = False
                    For j = 0 To MatFeriados.Count - 1
                        If MatFeriados(j).Day = dFechaTempo.Day And MatFeriados(j).Month = dFechaTempo.Month And MatFeriados(j).Year = dFechaTempo.Year Then
                            bFeriado = True
                            Exit For
                        End If
                    Next
                    If bFeriado Then
                        dFechaTempo = dFechaTempo.AddDays(1)
                        nDiasFeriado = nDiasFeriado + 1
                        bFeriado = False
                    End If

                    'SI es feriado incrementa en 1
                    bFeriado = False
                    For j = 0 To MatFeriados.Count - 1
                        If MatFeriados(j).Day = dFechaTempo.Day And MatFeriados(j).Month = dFechaTempo.Month And MatFeriados(j).Year = dFechaTempo.Year Then
                            bFeriado = True
                            Exit For
                        End If
                    Next
                    If bFeriado Then
                        dFechaTempo = dFechaTempo.AddDays(1)
                        nDiasFeriado = nDiasFeriado + 1
                        bFeriado = False
                    End If

                    If dFechaTempo.DayOfWeek = DayOfWeek.Saturday Then
                        dFechaTempo = dFechaTempo.AddDays(1)
                        nDiasFeriado = nDiasFeriado + 1
                    End If

                    If dFechaTempo.DayOfWeek = DayOfWeek.Sunday Then
                        dFechaTempo = dFechaTempo.AddDays(1)
                        nDiasFeriado = nDiasFeriado + 1
                    End If

                    MatCal(i).Fecha = CDate(Format(dFechaTempo, "dd/MM/yyyy"))
                    MatCal(i).NumCuota = i + 1

                    If pnProd = 4 Or pnProd = 9 Then

                        nInteresTempo = Format(((nTasaCompTempo * 1.0 / 100.0) * pnCuotas * pnPrestamo) / pnCuotas, "#0.00")
                        nInteresNetoTempo = Format(((nTasaCompNetoTempo * 1.0 / 100.0) * pnCuotas * pnPrestamo) / pnCuotas, "#0.00")



                    ElseIf pnProd = 6 Or pnProd = 7 Or pnProd = 8 Or pnProd = 15 Then
                        nDiasTrans = 7

                        nInteresTempo = Format(((nTasaCompTempo * 1.0 / 100.0) * (pnCuotas / 4) * pnPrestamo) / pnCuotas, "#0.00")
                        nInteresNetoTempo = Format(((nTasaCompNetoTempo * 1.0 / 100.0) * (pnCuotas / 4) * pnPrestamo) / pnCuotas, "#0.00")

                    ElseIf pnProd = 5 Or pnProd = 10 Then

                        nInteresTempo = Format(((nTasaCompTempo * 1.0 / 100.0) * (pnCuotas / 1) * pnPrestamo) / pnCuotas, "#0.00")
                        nInteresNetoTempo = Format(((nTasaCompNetoTempo * 1.0 / 100.0) * (pnCuotas / 1) * pnPrestamo) / pnCuotas, "#0.00")

                    ElseIf pnProd = 11 Or pnProd = 12 Then
                        nDiasTrans = 14

                        nInteresTempo = Format(((nTasaCompTempo * 1.0 / 100.0) * (pnCuotas / 2) * pnPrestamo) / pnCuotas, "#0.00")
                        nInteresNetoTempo = Format(((nTasaCompNetoTempo * 1.0 / 100.0) * (pnCuotas / 2) * pnPrestamo) / pnCuotas, "#0.00")


                    Else
                        nInteresTempo = Format(((nTasaCompTempo * 1.0 / 100.0) * (pnCuotas / 20) * pnPrestamo) / pnCuotas, "#0.00")
                        nInteresNetoTempo = Format(((nTasaCompNetoTempo * 1.0 / 100.0) * (pnCuotas / 20) * pnPrestamo) / pnCuotas, "#0.00")
                    End If



                    MatCal(i).Interes = nInteresTempo + nMontoIntGraciaCuota
                    MatCal(i).InteresNeto = nInteresNetoTempo
                    nGastoFijo = MatCal(i).Interes - MatCal(i).InteresNeto
                    MatCal(i).Gasto = nGastoFijo
                    MatCal(i).Igv = CDbl(Format(MatCal(i).Interes * 0.18, "#0.00"))

                    If i = (pnCuotas - 1) Then
                        MatCal(i).Capital = Format(nSaldoTempo, "#0.00")
                    Else
                        MatCal(i).Capital = Format(nMontoCuota - MatCal(i).Interes - MatCal(i).Igv, "#0.00")
                    End If
                    If MatCal(i).Capital < 0 Then
                        MatCal(i).Capital = 0
                    End If

                    nSaldoTempo = nSaldoTempo - MatCal(i).Capital

                Next i


                If pnCuotas > 1 Then
                    If pnProd = 3 Then

                        i = pnCuotas - 1
                        If (CDbl(Format(MatCal(i).Capital + MatCal(i).Interes + MatCal(i).Igv, "#0.00")) - nMontoCuota) > 0 Then
                            bCorrecto = False
                            nTasaCompTempo = nTasaCompTempo - 0.0001
                        ElseIf (CDbl(Format(MatCal(i).Capital + MatCal(i).Interes + MatCal(i).Igv, "#0.00")) - nMontoCuota) < -1.0 Then
                            nTasaCompTempo = nTasaCompTempo + 0.0001
                        Else
                            bCorrecto = True
                        End If

                    Else   ' Para Producto Mensual
                        'i = pnCuotas - 1
                        'If (CDbl(Format(MatCal(i).Capital + MatCal(i).Interes + MatCal(i).Igv, "#0.00")) - nMontoCuota) > 0 Then
                        '    bCorrecto = False
                        '    nMontoCuota = nMontoCuota + 0.01
                        'ElseIf (CDbl(Format(MatCal(i).Capital + MatCal(i).Interes + MatCal(i).Igv, "#0.00")) - nMontoCuota) < -1.0 Then
                        '    nMontoCuota = nMontoCuota - 0.01
                        'Else
                        '    bCorrecto = True
                        'End If

                        bCorrecto = True

                    End If

                Else
                    bCorrecto = True
                End If

            Loop



            If oConex.State = ConnectionState.Open Then
                oConex.Close()
            End If

            D(0) = Format(MatCal(0).Capital + MatCal(0).Interes + MatCal(0).Igv, "#0.00")
            GeneaCalendarioPorProductoSimCuota_FLAT = D

        Catch ex As Exception
            D(0) = ex.Message
            GeneaCalendarioPorProductoSimCuota_FLAT = D

        End Try



    End Function


    Public Function RecuperaPagoGrupalCabeceraIDGrupoDetalleTotal() As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_PagoGrupalCabeceraIDGrupoDetalleTotal", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu

            'oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

            Dim i As Integer = 0

            While dr.Read
                ReDim Preserve R.oPagoGrupalCabeceraDetalle(i)
                R.oPagoGrupalCabeceraDetalle(i) = New clsPagogrupalcabeceraDetalle()
                R.oPagoGrupalCabeceraDetalle(i).nID = dr.Item("nId")
                R.oPagoGrupalCabeceraDetalle(i).nCodCred = dr.Item("nCodCred")
                R.oPagoGrupalCabeceraDetalle(i).nCodAge = dr.Item("nCodAge")
                R.oPagoGrupalCabeceraDetalle(i).nMontoCuota = dr.Item("nMontoCuota")
                R.oPagoGrupalCabeceraDetalle(i).cCliente = dr.Item("cCliente")
                R.oPagoGrupalCabeceraDetalle(i).cDoc = dr.Item("cDoc")
                R.oPagoGrupalCabeceraDetalle(i).nCuotas = dr.Item("nCuotas")
                R.oPagoGrupalCabeceraDetalle(i).nCuotapendiente = dr.Item("nCuotapendiente")
                R.oPagoGrupalCabeceraDetalle(i).dFecVenc = dr.Item("dFecVenc")

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


    Public Function RecuperaPagoGrupalCabecera() As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_PagoGrupalCabecera", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@cCodUsuCob", SqlDbType.VarChar, 50).Value = sCodUsu

            'oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

            Dim i As Integer = 0

            While dr.Read
                ReDim Preserve R.oPagoGrupalCabecera(i)
                R.oPagoGrupalCabecera(i) = New clsPagogrupalcabecera()
                R.oPagoGrupalCabecera(i).nID = dr.Item("nId")
                R.oPagoGrupalCabecera(i).cNombre = dr.Item("cNombre")
                R.oPagoGrupalCabecera(i).dFechaVenc = dr.Item("dFechaVenc")
                R.oPagoGrupalCabecera(i).nNumSocios = dr.Item("nNumSocios")
                R.oPagoGrupalCabecera(i).nPrestamo = dr.Item("nPrestamo")
                R.oPagoGrupalCabecera(i).nSaldok = dr.Item("nSaldok")
                R.oPagoGrupalCabecera(i).nNumcuotas = dr.Item("nNumcuotas")
                R.oPagoGrupalCabecera(i).dFecVig = dr.Item("dFecVig")
                R.oPagoGrupalCabecera(i).nCuotasPagadas = dr.Item("nCuotasPagadas")
                R.oPagoGrupalCabecera(i).nMontoCuota = dr.Item("nMontoCuota")
                R.oPagoGrupalCabecera(i).nDiasAtraso = dr.Item("nDiasAtraso")

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

    Public Function RecuperaPersonaBase(ByVal nCodAge As Integer) As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_RecuperaPersonaBase", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge

            'oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

            Dim i As Integer = 0

            While dr.Read


                ReDim Preserve R.oPersonas(i)
                R.oPersonas(i) = New dPersona()
                R.oPersonas(i).cApellidos = dr.Item("cApellidos")
                R.oPersonas(i).cCodGenero = dr.Item("cCodGenero")
                R.oPersonas(i).cDirecDom = dr.Item("cDirecDom")
                R.oPersonas(i).cDistritoDom = dr.Item("cDistritoDom")
                R.oPersonas(i).cDoc = dr.Item("cDoc")
                R.oPersonas(i).cDocConyuge = dr.Item("cDocConyuge")
                R.oPersonas(i).cMovil1 = dr.Item("cMovil1")
                R.oPersonas(i).cMovil2 = dr.Item("cMovil2")
                R.oPersonas(i).cNomBanco = dr.Item("cNomBanco")
                R.oPersonas(i).cNombres = dr.Item("cNombres")
                R.oPersonas(i).cNroCtaAho = dr.Item("cNroCtaAho")
                R.oPersonas(i).cRefDirDom = dr.Item("cRefDirDom")
                R.oPersonas(i).cTelefFijo = dr.Item("cTelefFijo")
                R.oPersonas(i).cZonaDom = dr.Item("cZonaDom")
                R.oPersonas(i).dFechaNac = dr.Item("dFechaNac")
                R.oPersonas(i).nCodCred = dr.Item("nCodCred")
                R.oPersonas(i).nCodEstadoCiv = dr.Item("nCodEstadoCiv")
                R.oPersonas(i).nNumDep = dr.Item("nNumDep")
                R.oPersonas(i).nSaldoCap = dr.Item("nSaldoCap")
                R.oPersonas(i).nMontoMaximo = dr.Item("nMontoMaximo")
                R.oPersonas(i).nPlazoMaximo = dr.Item("nPlazoMaximo")
                R.oPersonas(i).cCategoria = dr.Item("cCategoria")
                R.oPersonas(i).nTipoDoc = dr.Item("nTipoDoc")
                R.oPersonas(i).nTipoDocConyuge = dr.Item("nTipoDocConyuge")
                R.oPersonas(i).ncodBanco = dr.Item("nCodBanco")
                R.oPersonas(i).cNroCtaAhoCCI = dr.Item("cCCI")



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


    Public Function RecuperaProyeccionAmortizacion() As outNegocio
        Dim R As New outNegocio

        'Carga Proyeccion de Amortizaciones
        Dim RProyAmor As New outNegocio
        RProyAmor = RecuperaProyeccionAmortizaciones()
        R.oProyecamortizacion = RProyAmor.oProyecamortizacion


        Return R

    End Function


    Public Function RecuperaCuotaONLINE() As outNegocio
        Dim R As New outNegocio
        Dim nTasa As Double
        Dim D As New dDesembolso
        Dim sMontoCuota As String()

        'Dim oTasa = From T In db.TarifarioAgencia
        '            Where T.nProd = nProd And T.nCodAge = nCodAge
        '            Select T

        'nTasa = oTasa.First.nTasaComp

        'If nCodAge = 1 Then
        '    sMontoCuota = D.MontoCuotaPorProd(Str(nPrestamo), Str(nCuotas), Str(nTasa), "1", "1", Str(nProd), "1")

        'Else
        '    sMontoCuota = D.MontoCuotaPorProd_flat(Str(nPrestamo), Str(nCuotas), Str(nTasa), "1", "1", Str(nProd), "1")

        'End If


        If nProd = 6 Then


            Dim oTasa = From T In db.TarifarioMicroSemanal
                        Where T.cCodCat = cCodCat And T.nCodAge = nCodAge And T.nPretamoIni <= nPrestamo And T.nPrestamoFin >= nPrestamo
                        Select T

            If oTasa.Count > 0 Then
                nTasa = oTasa.First.nTasaComp
            Else
                nTasa = 0
            End If


        ElseIf nProd = 7 Then

            Dim oTasa = From T In db.TarifarioMicroJuniorSemanal
                        Where T.cCodCat = cCodCat And T.nCodAge = nCodAge And T.nPretamoIni <= nPrestamo And T.nPrestamoFin >= nPrestamo
                        Select T

            If oTasa.Count > 0 Then
                nTasa = oTasa.First.nTasaComp
            Else
                nTasa = 0
            End If

        ElseIf nProd = 8 Then
            Dim oTasa = From T In db.TarifarioMicro
                        Where T.nCodAge = nCodAge And T.nPretamoIni <= nPrestamo And T.nPrestamoFin >= nPrestamo And T.nProd = 8
                        Select T

            If oTasa.Count > 0 Then
                nTasa = oTasa.First.nTasaComp
            Else
                nTasa = 0
            End If

        ElseIf nProd = 10 Or nProd = 9 Or nProd = 11 Or nProd = 12 Or nProd = 15 Then

            Dim oTasa = From T In db.TarifarioMicro
                        Where T.nCodAge = nCodAge And T.nPretamoIni <= nPrestamo And T.nPrestamoFin >= nPrestamo And T.nProd = nProd
                        Select T

            If oTasa.Count > 0 Then
                nTasa = oTasa.First.nTasaComp
            Else
                nTasa = 0
            End If




        ElseIf nProd = 5 Then 'mdipas 2024
            Dim oTasa = From T In db.TarifarioMicro
                        Where T.nCodAge = nCodAge And T.nPretamoIni <= nPrestamo And T.nPrestamoFin >= nPrestamo And T.nProd = nProd
                        Select T

            If oTasa.Count > 0 Then
                nTasa = oTasa.First.nTasaComp
            Else
                nTasa = 0
            End If
        Else


            Dim oTasa = From T In db.TarifarioMicro
                        Where T.cCodCat = cCodCat And T.nCodAge = nCodAge And T.nPretamoIni <= nPrestamo And T.nPrestamoFin >= nPrestamo And T.nProd = 3
                        Select T

            If oTasa.Count > 0 Then
                nTasa = oTasa.First.nTasaComp
            Else
                nTasa = 0
            End If

        End If

        If nTasa = 0 Then
            ReDim sMontoCuota(1)
            sMontoCuota(0) = "No se encontró la tasa"
        Else
            If nCodAge = 1 Then
                sMontoCuota = GeneaCalendarioPorProductoSimCuota(nTasa, nCuotas, nPrestamo, nProd, nCampSinInt)
            Else
                sMontoCuota = GeneaCalendarioPorProductoSimCuota_FLAT(nTasa, nCuotas, nPrestamo, nProd, nCampSinInt)
            End If
        End If



        R.nResultado = 1
        R.sRespuesta = sMontoCuota(0)

        Return R
    End Function

    Public Function RecuperaDatosDeCredito() As outNegocio
        Dim R As New outNegocio

        Dim oData() As dPagoCredito

        Dim oListaPagos = From P In db.vListadoPagos
                          Where P.nCodCred = nCodcred
                          Select New dPagoCredito With {.bPagadoHoy = P.bPagadoHoy, .cDireccionDom = P.cDireccionDom,
                              .cDirecNeg = P.cDirecNeg, .cDoc = P.cDoc, .cMovil = P.cMovil1,
                              .cMovil2 = P.cMovil2, .cRazonSocial = P.cRazonSocial, .cNomcliente = P.cCliente,
                              .cReferencia = "", .cZonaNegocio = P.cZonaNegocio, .cZonaPersona = P.cZonaPersona,
                              .dFechaVenc = P.sFechaVencimiento, .nCodAge = P.nCodAge, .nCodCred = P.nCodCred,
                              .nDeudaAtrasada = P.nDeudaAtrasada, .nDiasAtraso = P.nDiasAtraso,
                              .nLatitud = P.nLatitud, .nLongitud = P.nLongitud,
                              .nMonto = P.nMontoCuota,
                              .nMontoCancelaCredito = P.MontoCancelaCredito, .nMontoCancelaCuota = P.MontoCancelaCuota,
                              .nNroProxCuota = P.nNroProxCuota, .nPagoMinio = P.nPagoMinimo, .nPrestamo = P.nPrestamo,
                              .sAmpliacion = P.Ampliacion, .UltFecCob = P.UltFecCob, .cCodUsuCob = P.cCodUsuCob, .nNumPagosDelDia = P.nNumPagosDelDia,
                              .nCuotasTotales = P.nCuotasTotales, .nCuotasPagadas = P.nCuotasPagadas, .nCuotasPendientes = P.nCuotasPendientes,
                              .cCategoria = P.cCategoria, .nProd = P.nProd, .nCiclos = P.nCiclos, .nOrdenRuta = P.nOrdenRuta, .nPagopromedio = P.PagoRecurrente, .nIDGrupo = P.nIDGrupo}

        oData = oListaPagos.ToArray

        Dim i As Integer

        For i = 0 To oData.Count - 1
            oData(i).CargaCalendario()
        Next

        R.ListaPagosCredito = oData.AsQueryable

        Return R


    End Function
    Public Function RecuperaDatosBasicos() As outNegocio

        Dim R As New outNegocio
        Dim Rtemp As New outNegocio

        Try
            Rtemp = RecuperaListaPagos()
            R.ListaPagosCredito = Rtemp.ListaPagosCredito

            Dim oCatalogocodigos = From C In db.CatalogoCodigos
                                   Where C.bActivo = True
                                   Select New dCatalogocodigos With {.cDescripcion = C.cDescripcion, .cValor = C.cValor, .nCodigo = C.nCodigo}

            R.oCatalogocodigos = oCatalogocodigos

            Dim oZonas = From Z In db.vZonas
                         Select New dZonas With {.bVisible = If(String.IsNullOrEmpty(Z.bVisible), 0, 1), .cCodZona = Z.cCodZona, .cNomZona = Z.cNomZona}

            R.oZonas = oZonas

            Dim nCodAge As Integer

            Dim dUsu = From U In db.Usuarios
                       Where U.cCodUsu = sCodUsu
                       Select U

            nCodAge = dUsu.First.nCodAge


            'Dim oPersonas = From P In db.vPersonasBase
            '                Where P.nCodAge = nCodAge And P.cDoc <> "98989898"
            '                Select New dPersona With {.cApellidos = P.cApellidos, .cCodGenero = P.cGenero, .cDirecDom = P.cDireccionDom, .cDistritoDom = "", .cDoc = P.cDoc, .cDocConyuge = P.cDocConyuge, .cMovil1 = P.cMovil1, .cMovil2 = P.cMovil2, .cNomBanco = P.cNomBanco, .cNombres = P.cNombres, .cNroCtaAho = P.cctabanco, .cRefDirDom = P.cReferenciaDom, .cTelefFijo = P.cTelFijo, .cZonaDom = P.cCodZonaDom, .dFechaNac = P.dFechaNac, .nCodCred = P.nCodCred, .nCodEstadoCiv = P.nEstadoCiv, .nNumDep = P.nDependientes, .nSaldoCap = P.nSaldoK, .nMontoMaximo = P.nLineaCredito, .nPlazoMaximo = P.nCuotasMax, .cCategoria = P.cCategoria2, .nTipoDoc = P.nTipoDoc, .nTipoDocConyuge = P.nTipoDocConyuge}


            'R.oPersonas = oPersonas

            Dim RpersonaBase As New outNegocio
            RpersonaBase = RecuperaPersonaBase(nCodAge)
            R.oPersonas = RpersonaBase.oPersonas



            Dim oNegocios = From E In db.vListaNegocios
                            Where E.nCodAge = nCodAge
                            Select New dNegocio With {.cCodActEcon = E.cCodActEcon, .cCodZona = E.cCodZonaNeg, .cDirecNeg = E.cDirecNeg, .cDoc = E.cDoc,
                                .cRazonSocial = E.cRazonSocial, .cRefDirecNeg = E.cReferenciaNeg, .dFechaIniNeg = E.dFechaIniNeg, .dFechareg = E.dFechaEval,
                                .nGastosMensFam = E.nGastosFam, .nMontoVentasDiarias = E.nMontoInventario, .nNumTrab = E.nNumTrab, .nVentadialv = E.nVentadialv,
                                .nVentadiasd = E.nVentadiasd, .nComprasSemanal = E.nComprasSemanal, .nCajaEfectivo = E.nCajaEfectivo, .nGastoAlquiler = E.nGastoAlquiler,
                                .nGastoFinanciero = E.nGastoFinanciero, .nGastoPrestamistas = E.nGastoPrestamistas, .nMontoInventario = E.nMontoInventario,
                                .nMaquinYEquipos = E.nMaquinYEquipos,
                                .nCodEval = E.nCodEval,
                                .cComentariosEval = E.cComentarios, .cComentariosOtrosIngresos = E.cComentOtrosIngresos,
                                .nMontoDBueno = E.nMontoDiaBueno,
                                .nNroDBuenoSeman = E.nDiasBuenos,
                                .nMontoDiaNormal = E.nMontoDiaNormal,
                                .nOtrosIngresos = E.nOtrosIngresos,
                                .nMontoServicios = E.nCosOpeServicios,
                                .nCosOpeAlquiler = E.nCosOpeAlquiler,
                                .nTrabajaConInformales = If(E.bTrabInform, 1, 0),
                                .nNroInformales = E.nNroTotalInform,
                                .nMontoInformales = E.nMontoTotalInform,
                                .nNumHijos = E.nNroHijos,
                                .nGasFamAlimentac = E.nGasFamAlimen,
                                .nGasFamServicios = E.nGasFamServicios,
                                .nGasFamOtros = E.nGasFamOtros,
                                .nCosOpeTransp = E.nCosOpeTransp,
                                .nUtilidadBrutaTotal = E.nUtilidadBrutaTotal,
                                .nGastosNegocio = E.nGastosNegocio,
                                .nGastosTotalFam = E.nGastosTotalFam,
                                .nDeudasFinancieras = E.nDeudasFinancieras,
                                .nUtilidadNetaTotal = E.nUtilidadNetaTotal,
                                .nEndeudamientoMax = E.nEndeudamientoMax,
                                .nAsumeCuotaDiaria = E.nAsumeCuotaDiaria,
                                .nAsumeSemanal = E.nAsumeSemanal,
                                .nAsumeCatorcenal = E.nAsumeCatorcenal,
                                .nAsumeMensual = E.nAsumeMensual,
                                .nGasFamTransporte = E.nGasFamTransp,
                                .nGasFamAlquiler = E.nGasFamAlquiler
                                }




            R.oNegocios = oNegocios




            Dim oTarAge = From T In db.TarifarioAgencia
                          Select T

            R.oTarifarioAgencia = oTarAge

            Dim oTarMicro = From T In db.TarifarioMicro
                            Select T

            R.oTarifarioMicro = oTarMicro

            Dim oTarMicroSemanal = From T In db.TarifarioMicroSemanal
                                   Select T

            R.oTarifarioMicroSemanal = oTarMicroSemanal


            Dim oTarMicroJuniorSemanal = From T In db.TarifarioMicroJuniorSemanal
                                         Select T

            R.oTarifarioMicroJuniorSemanal = oTarMicroJuniorSemanal


            Dim oTarGrupal = From T In db.TarifarioGrupal
                             Select T

            R.oTarifarioGrupal = oTarGrupal


            Dim oListaVerif = From T In db.vListaClientesPorVerificar
                              Select T

            R.oListaClientesPorVerificar = oListaVerif

            Dim oListaVerifFotos = From T In db.vListaClientesPorVerificarFotos
                                   Select T

            R.oListaClientesPorVerificarFotos = oListaVerifFotos

            Dim oListaXDesemb = From T In db.vListaCreditosPorDesembolsar
                                Where T.nCodAge = nCodAge And T.cCodUsuDesem = sCodUsu
                                Select T

            R.oListaClientesPorDesembolsar = oListaXDesemb

            Dim oListaCliInactivos = From T In db.ClientesInactivos
                                     Where T.nCodAge = nCodAge
                                     Select T

            R.oListaClientesInactivos = oListaCliInactivos

            Dim oFamilia = From F In db.ArticuloFamilia
                           Where F.bEliminado = 0
                           Select F

            R.oListaArticuloFamilia = oFamilia

            Dim oMarcas = From F In db.ArticuloMarcas
                          Where F.bEliminado = 0
                          Select F

            R.oListaArticuloMarcas = oMarcas

            Dim oModelos = From F In db.ArticuloModelo
                           Where F.bEliminado = 0
                           Select F

            R.oListaArticuloModelo = oModelos

            'Carga Ampliaciones
            Dim RAmpl As New outNegocio
            RAmpl = RecuperaAmpliaciones()
            R.oAmpliaciones = RAmpl.oAmpliaciones


            Dim RPagosGrupalCab As New outNegocio
            RPagosGrupalCab = RecuperaPagoGrupalCabecera()
            R.oPagoGrupalCabecera = RPagosGrupalCab.oPagoGrupalCabecera

            Dim RPagosGrupalCabDetalle As New outNegocio
            RPagosGrupalCabDetalle = RecuperaPagoGrupalCabeceraIDGrupoDetalleTotal()
            R.oPagoGrupalCabeceraDetalle = RPagosGrupalCabDetalle.oPagoGrupalCabeceraDetalle

            'Dim oRubros = From F In db.EvaluacionRubros
            '              Select F
            'R.oNegociosRubros = RecuperaListaNegocioRubros.oNegociosRubros 

            Dim oRubros = From E In db.EvaluacionRubros
                          Select New Rubro With {.nCodEval = E.nCodEval,
                                 .nombre = E.cRubro,
                                 .ventaDiariaDN = E.nVentasDiariasDN,
                              .ventaDiariaDB = E.nVentasDiariasDB,
                              .ventaSemanalDN = E.nVentasSemanalesDN,
                              .ventaSemanalDB = E.nVentasSemanalesDB,
                              .ventaMensualDN = E.nVentasMensualesDN,
                              .ventaMensualDB = E.nVentasMensualesDB,
                              .CostoPorcDB = E.nPorcCostoDB,
                              .CostoPorcDN = E.nPorcCostoDN,
                              .CostoDB = E.nCostoDB,
                              .CostoDN = E.nCostoDN,
                              .UtilidadPorcDB = E.nPorcUtilidadDB,
                              .UtilidadPorcDN = E.nPorcUtilidadDN,
                              .UtilidadDB = E.nUtilidadDB,
                              .UtilidadDN = E.nUtilidadDN,
                              .totalUtilidadMensual = E.nTotalUtilidadMensual,
                              .totalVentaMensual = E.nTotalVentasMensual
                                }
            R.oNegociosRubros = oRubros

            Dim oEntidades = From E In db.EvaluacionDeudaFinancieras
                             Select New dEntidad With {.nCodEval = E.nCodEval,
                                 .nombre = E.cNomEntidad,
                                 .deuda = E.nSaldoDeuda
                                }
            R.oDeudasEntidades = oEntidades

            Return R
        Catch ex As Exception
            'D(0) = "0"
            'D(1) = "ERROR: " + ex.Message
            Return R
        End Try



    End Function

    Public Function RegistrarPagoGrupalLote() As outNegocio
        Dim R As New outNegocio

        Try
            TranPago.oListaPagoLoteCreditos = oListaPagoLoteCreditos
            R = TranPago.RegistrarPagoGrupalLote()

            'db.GruposComunalesDetalle.Add(oSocio)
            'Try
            '    db.SaveChanges()
            '    nResp = 1
            '    sResp = "Registro Correcto"

            'Catch ex As Exception
            '    nResp = 0
            '    sResp = ex.Message

            'End Try


        Catch ex As Exception
            R.nResultado = 0
            R.sRespuesta = ex.Message
        End Try

        R.TranPago = TranPago
        Return R

    End Function


    Public Function RegistrarPago() As outNegocio
        Dim R As New outNegocio

        Try

            R = TranPago.RegistrarPago()

            'db.GruposComunalesDetalle.Add(oSocio)
            'Try
            '    db.SaveChanges()
            '    nResp = 1
            '    sResp = "Registro Correcto"

            'Catch ex As Exception
            '    nResp = 0
            '    sResp = ex.Message

            'End Try


        Catch ex As Exception
            R.nResultado = 0
            R.sRespuesta = ex.Message
        End Try

        R.TranPago = TranPago
        Return R

    End Function

    Public Function RecuperaFotos() As outNegocio
        Dim R As New outNegocio
        Dim nCodEval As Long

        Dim oData = From F In db.EvaluacionFotos
                    Join E In db.Evaluacion On E.nCodEval Equals F.nCodEval
                    Join CP In db.CredPersonas On CP.nCodPers Equals E.nCodPers
                    Where CP.nRelacion = 10 And CP.nCodCred = nCodcred
                    Order By F.nCodEval Descending
                    Select F

        nCodEval = oData.First.nCodEval

        Dim oFotos = From F In db.EvaluacionFotos
                     Where F.nCodEval = nCodEval And F.bNoVisible = 0
                     Select F

        R.oFotosCredito = oFotos
        Return R
    End Function
    Public Function ValidaExisteUsuario() As outNegocio
        Dim R As New outNegocio

        Dim oUsu = From U In db.Usuarios
                   Where U.cCodUsu = sCodUsu
                   Select U

        If oUsu.Count > 0 Then
            R.nExisteUsuario = 1
            R.nResultado = 1
            R.sRespuesta = "Registro Correcto"

        Else
            R.nExisteUsuario = 0
            R.nResultado = 1
            R.sRespuesta = "Registro Correcto"
        End If

        Return R

    End Function
    Public Function RegistraGPS() As outNegocio
        Dim R As New outNegocio

        Try
            db.PA_Movil_RegistraGPS(nLatitud, nLongitud, sCodUsu)
            R.nResultado = 1
            R.sRespuesta = "Registro Correcto"

        Catch ex As Exception
            R.nResultado = 0
            R.sRespuesta = ex.Message
        End Try

        Return R


    End Function

    Public Function RecuperaTablasIniciales() As outNegocio
        Dim R As New outNegocio

        Try


            Dim oUsuario = From U In db.Usuarios
                           Where U.cCodUsu = sCodUsu
                           Select New dUsuario With {.cCodUsu = U.cCodUsu, .cIMEI = U.cIMEI, .cNombres = U.cNombres, .nCodAge = U.nCodAge, .sContrasenia = U.sContrasenia, .nMontoLimiteCob = U.nMontoLimiteCob, .nCierre = U.nCierre, .nHoraFin = U.nHoraFin, .nHoraIni = U.nHoraIni, .nVersion = U.nVersion}

            R.oUsuario = oUsuario.First

            If R.oUsuario.cIMEI = Nothing Then
                R.oUsuario.cIMEI = "-1"
            End If

            Dim oUsuarioRoles = From UR In db.UsuarioRoles
                                Where UR.cCodUsu = sCodUsu
                                Select New dUsuarioroles With {.cCodUsu = UR.cCodUsu, .cCodRol = UR.cCodRol}

            Dim oListaVarnegocio = From UR In db.Varnegocio
                                   Select New eVarnegocio With {.nCodVal = UR.nCodVal, .cDesVal = UR.cDesVal, .cCodVal = UR.cCodVal}

            R.oUsuarioroles = oUsuarioRoles
            R.oVarnegocio = oListaVarnegocio

            'Dim oCatalogocodigos = From C In db.CatalogoCodigos
            '                       Select New dCatalogocodigos With {.cDescripcion = C.cDescripcion, .cValor = C.cValor, .nCodigo = C.nCodigo}

            'R.oCatalogocodigos = oCatalogocodigos

            Return R

        Catch ex As Exception
            R.nResultado = 0
            R.sRespuesta = ex.Message
            Return R
        End Try

    End Function

    Public Function RecuperaListaPagosGrupales() As outNegocio
        Dim R As New outNegocio
        Dim oData() As dPagoCredito
        db.Database.CommandTimeout = 360
        Dim oListaPagos = From P In db.vListadoPagos
                          Where P.cCodUsuCob = sCodUsu Or P.cCodUsuCob2 = sCodUsu Or P.cCodUsuCobPromotor = sCodUsu
                          Select New dPagoCredito With {.bPagadoHoy = P.bPagadoHoy, .cDireccionDom = P.cDireccionDom,
                              .cDirecNeg = P.cDirecNeg, .cDoc = P.cDoc, .cMovil = P.cMovil1,
                              .cMovil2 = P.cMovil2, .cRazonSocial = P.cRazonSocial, .cNomcliente = P.cCliente,
                              .cReferencia = "", .cZonaNegocio = P.cZonaNegocio, .cZonaPersona = P.cZonaPersona,
                              .dFechaVenc = P.sFechaVencimiento, .nCodAge = P.nCodAge, .nCodCred = P.nCodCred,
                              .nDeudaAtrasada = P.nDeudaAtrasada, .nDiasAtraso = P.nDiasAtraso,
                              .nLatitud = P.nLatitud, .nLongitud = P.nLongitud,
                              .nMonto = P.nMontoCuota,
                              .nProd = P.nProd,
                              .nMontoCancelaCredito = P.MontoCancelaCredito, .nMontoCancelaCuota = P.MontoCancelaCuota,
                              .nNroProxCuota = P.nNroProxCuota, .nPagoMinio = P.nPagoMinimo, .nPrestamo = P.nPrestamo,
                              .sAmpliacion = P.Ampliacion, .UltFecCob = P.UltFecCob, .cCodUsuCob = P.cCodUsuCob, .nNumPagosDelDia = P.nNumPagosDelDia, .cCodUsuAnalista = P.cCodUsuAnalista, .nOrdenRuta = P.nOrdenRuta, .nPagopromedio = P.PagoRecurrente}

        oData = oListaPagos.ToArray

        Dim i As Integer

        For i = 0 To oData.Count - 1
            oData(i).CargaCalendario()
        Next

        R.ListaPagosCredito = oData.AsQueryable

        Return R

    End Function


    Public Function RecuperaListaPagos() As outNegocio
        Dim R As New outNegocio
        Dim oData() As dPagoCredito
        db.Database.CommandTimeout = 360
        Dim oListaPagos = From P In db.vListadoPagos
                          Where P.cCodUsuCob = sCodUsu Or P.cCodUsuCob2 = sCodUsu Or P.cCodUsuCobPromotor = sCodUsu
                          Select New dPagoCredito With {.bPagadoHoy = P.bPagadoHoy, .cDireccionDom = P.cDireccionDom,
                              .cDirecNeg = P.cDirecNeg, .cDoc = P.cDoc, .cMovil = P.cMovil1,
                              .cMovil2 = P.cMovil2, .cRazonSocial = P.cRazonSocial, .cNomcliente = P.cCliente,
                              .cReferencia = "", .cZonaNegocio = P.cZonaNegocio, .cZonaPersona = P.cZonaPersona,
                              .dFechaVenc = P.sFechaVencimiento, .nCodAge = P.nCodAge, .nCodCred = P.nCodCred,
                              .nDeudaAtrasada = P.nDeudaAtrasada, .nDiasAtraso = P.nDiasAtraso,
                              .nLatitud = P.nLatitud, .nLongitud = P.nLongitud,
                              .nMonto = P.nMontoCuota,
                              .nProd = P.nProd,
                              .nMontoCancelaCredito = P.MontoCancelaCredito, .nMontoCancelaCuota = P.MontoCancelaCuota,
                              .nNroProxCuota = P.nNroProxCuota, .nPagoMinio = P.nPagoMinimo, .nPrestamo = P.nPrestamo,
                              .sAmpliacion = P.Ampliacion, .UltFecCob = P.UltFecCob, .cCodUsuCob = P.cCodUsuCob, .nNumPagosDelDia = P.nNumPagosDelDia, .cCodUsuAnalista = P.cCodUsuAnalista, .nOrdenRuta = P.nOrdenRuta, .nPagopromedio = P.PagoRecurrente, .nIDGrupo = P.nIDGrupo}

        oData = oListaPagos.ToArray

        Dim i As Integer

        For i = 0 To oData.Count - 1
            oData(i).CargaCalendario()
        Next

        R.ListaPagosCredito = oData.AsQueryable

        Return R

    End Function


    Public Function RecuperaPoliticas() As outNegocio

        Dim R As outNegocio
        Dim S As iSesion

        S = New iSesion()

        If Not S.ValidaToken(sToken) Then
            R.nResultado = 0
            R.sRespuesta = "Token Invalido"
        End If

        R = New outNegocio()

        Dim dReglasNegocio = From U In db.sistNegocioReglas
                             Join F In db.sistFunciones On F.cNomFuncion Equals U.cFuncionAValidar
                             Where F.bEjecucionServidor = False
                             Select U

        R.ReglasdeNegocio = dReglasNegocio

        Dim dReglasNegocioParam = From U In db.sistNegocioReglasParam
                                  Join NR In db.sistNegocioReglas On U.nIDPadre Equals NR.nID
                                  Join F In db.sistFunciones On F.cNomFuncion Equals NR.cFuncionAValidar
                                  Where F.bEjecucionServidor = False
                                  Select U

        R.ReglasdeNegocioParam = dReglasNegocioParam

        Dim dReglasNegocioDependencias = From U In db.sistNegocioReglasDependencias
                                         Select U

        R.ReglasdeNegocioDependencias = dReglasNegocioDependencias

        Dim dFormulariosRoles = From U In db.sistFormulariosRoles
                                Select U

        R.ListaFormulariosRoles = dFormulariosRoles

        Dim dPantalla = From U In db.SistFormularios
                        Select U

        R.ValoresPantalla = dPantalla

        R.nResultado = 1
        R.sRespuesta = "Proceso Realizado de manera Correcta"

        RecuperaPoliticas = R

    End Function

    Public Function RecuperaSolicitud() As outNegocio
        Dim R As outNegocio
        Dim S As iSesion
        Dim i As Integer
        Dim sDato As String

        S = New iSesion()

        If Not S.ValidaToken(sToken) Then
            R.nResultado = 0
            R.sRespuesta = "Token Invalido"
        End If

        R = New outNegocio()

        If nSolicitud = 1 Then
            R.ListaPrestamoPorAprobar = db.vPrestamoPorAprobar
        End If

        'Recupera Data de Pantalla de Personas
        If nSolicitud = 2 Then

            Dim dPersona = From U In db.Persona
                           Where U.cDoc = cDocPersona
                           Select U

            'Dim dReglasNegocio = From U In db.sistNegocioReglas
            '                     Join F In db.sistFunciones On F.cNomFuncion Equals U.cFuncionAValidar
            '                     Where U.cNomForm = "persona" And F.bEjecucionServidor = False
            '                     Select U

            'Dim dReglasNegocioParam = From U In db.sistNegocioReglasParam
            '                          Join NR In db.sistNegocioReglas On U.nIDPadre Equals NR.nID
            '                          Join F In db.sistFunciones On F.cNomFuncion Equals NR.cFuncionAValidar
            '                          Where NR.cNomForm = "persona" And F.bEjecucionServidor = False
            '                          Select U


            'Dim dReglasNegocioDependencias = From U In db.sistNegocioReglasDependencias
            '                                 Select U

            'Dim dFormulariosRoles = From U In db.sistFormulariosRoles
            '                        Where U.cNomForm = "persona"
            '                        Select U

            Dim dPantalla = From U In db.SistFormularios
                            Where U.cNomForm = "persona"
                            Select U

            'R.ReglasdeNegocio = dReglasNegocio
            'R.ReglasdeNegocioParam = dReglasNegocioParam
            'R.ReglasdeNegocioDependencias = dReglasNegocioDependencias
            R.ValoresPantalla = dPantalla
            'R.ListaFormulariosRoles = dFormulariosRoles

            If dPersona.Count() > 0 Then


                For Each V In R.ValoresPantalla
                    Select Case V.cNomControl
                        Case "cboBanco"
                            V.cValor = dPersona.First.cNomBanco
                        Case "cboDistrito"
                            V.cValor = dPersona.First.cCodZonaDom
                        Case "cbodpto"
                            V.cValor = dPersona.First.cCodZonaDom
                        Case "cboEstCiv"
                            V.cValor = dPersona.First.nEstadoCiv
                        Case "cbogenero"
                            V.cValor = dPersona.First.cGenero
                        Case "cboprov"
                            V.cValor = dPersona.First.cCodZonaDom
                        Case "cboTrabEmp"
                            V.cValor = dPersona.First.cTrabEmp
                        Case "cboZona"
                            V.cValor = dPersona.First.cCodZonaDom
                        Case "DropDownList1"
                            V.cValor = dPersona.First.cAtrib1
                        Case "DropDownList2"
                            V.cValor = dPersona.First.cAtrib2
                        Case "DropDownList3"
                            V.cValor = dPersona.First.cAtrib3
                        Case "DropDownList4"
                            V.cValor = dPersona.First.cAtrib4
                        Case "DropDownList5"
                            V.cValor = dPersona.First.cAtrib5
                        Case "DropDownList6"
                            V.cValor = dPersona.First.cAtrib6
                        Case "DropDownList7"
                            V.cValor = dPersona.First.cAtrib7
                        Case "DropDownList8"
                            V.cValor = dPersona.First.cAtrib8
                        Case "DropDownList9"
                            V.cValor = dPersona.First.cAtrib9
                        Case "DropDownList10"
                            V.cValor = dPersona.First.cAtrib10
                        Case "DropDownList11"
                            V.cValor = dPersona.First.cAtrib11
                        Case "DropDownList12"
                            V.cValor = dPersona.First.cAtrib12
                        Case "TextBox1"
                            V.cValor = dPersona.First.cAtrib1
                        Case "TextBox2"
                            V.cValor = dPersona.First.cAtrib2
                        Case "TextBox3"
                            V.cValor = dPersona.First.cAtrib3
                        Case "TextBox4"
                            V.cValor = dPersona.First.cAtrib4
                        Case "TextBox5"
                            V.cValor = dPersona.First.cAtrib5
                        Case "TextBox6"
                            V.cValor = dPersona.First.cAtrib6
                        Case "TextBox7"
                            V.cValor = dPersona.First.cAtrib7
                        Case "TextBox8"
                            V.cValor = dPersona.First.cAtrib8
                        Case "TextBox9"
                            V.cValor = dPersona.First.cAtrib9
                        Case "TextBox10"
                            V.cValor = dPersona.First.cAtrib10
                        Case "TextBox11"
                            V.cValor = dPersona.First.cAtrib11
                        Case "TextBox12"
                            V.cValor = dPersona.First.cAtrib12
                        Case "txtApellidos"
                            V.cValor = dPersona.First.cApellidos
                        Case "TxtApeNomCony"
                            V.cValor = dPersona.First.cNomConyuge
                        Case "txtcargoact"
                            V.cValor = dPersona.First.cCargoAct
                        Case "Txtcel1"
                            V.cValor = dPersona.First.cMovil1
                        Case "TxtCel2"
                            V.cValor = dPersona.First.cMovil2
                        Case "txtCtaBco"
                            V.cValor = dPersona.First.cCtaBanco
                        Case "txtDep"
                            V.cValor = dPersona.First.nDependientes
                        Case "txtdiremp"
                            V.cValor = dPersona.First.cDirecTrab
                        Case "txtDoc"
                            V.cValor = dPersona.First.cDoc
                        Case "txtDocCony"
                            V.cValor = dPersona.First.cDocConyuge
                        Case "txtDomic"
                            V.cValor = dPersona.First.cDireccionDom
                        Case "TxtFecNac"
                            V.cValor = dPersona.First.dFechaNac
                        Case "txtNombres"
                            V.cValor = dPersona.First.cNombres
                        Case "txtnomempr"
                            V.cValor = dPersona.First.cNomEmp
                        Case "txtRefer"
                            V.cValor = dPersona.First.cReferenciaDom
                        Case "txtrucemp"
                            V.cValor = dPersona.First.cRUC
                        Case "txtsueldobruto"
                            V.cValor = IIf(dPersona.First.nSueldo Is Nothing, 0, dPersona.First.nSueldo)
                        Case "txttelemp"
                            V.cValor = dPersona.First.cTelefTrab
                        Case "txttelfijo"
                            V.cValor = dPersona.First.cTelFijo

                    End Select
                Next


            End If

        End If

        R.nResultado = 1
        R.sRespuesta = "Proceso Realizado de manera Correcta"


        Return R

    End Function

    Public Function RegistraPersona() As outNegocio
        Dim P As New PoliticasInternas
        Dim bRespValidaPoliticas As Boolean
        Dim sResultado As String = ""
        Dim O As New outNegocio
        Dim sAtrib1 As String
        Dim sAtrib2 As String
        Dim sAtrib3 As String
        Dim sAtrib4 As String
        Dim sAtrib5 As String
        Dim sAtrib6 As String
        Dim sAtrib7 As String
        Dim sAtrib8 As String
        Dim sAtrib9 As String
        Dim sAtrib10 As String
        Dim sAtrib11 As String
        Dim sAtrib12 As String


        P.ValoresPantalla = ValoresPantalla
        bRespValidaPoliticas = P.ValidaPoliticas(sResultado)
        O.sRespuesta = sResultado

        If Not bRespValidaPoliticas Then
            O.nResultado = 0
        Else

            Dim P1 As ObjectParameter
            P1 = New ObjectParameter("Resultado", 0)
            Dim P2 As ObjectParameter
            P2 = New ObjectParameter("nResp", 0)

            sAtrib1 = P.ValorControlPantalla("TextBox1").ToString.Trim
            If sAtrib1.Length = 0 Then
                sAtrib1 = P.ValorControlPantalla("DropDownList1").ToString.Trim
            End If

            sAtrib2 = P.ValorControlPantalla("TextBox2").ToString.Trim
            If sAtrib2.Length = 0 Then
                sAtrib2 = P.ValorControlPantalla("DropDownList2").ToString.Trim
            End If

            sAtrib3 = P.ValorControlPantalla("TextBox3").ToString.Trim
            If sAtrib3.Length = 0 Then
                sAtrib3 = P.ValorControlPantalla("DropDownList3").ToString.Trim
            End If

            sAtrib4 = P.ValorControlPantalla("TextBox4").ToString.Trim
            If sAtrib4.Length = 0 Then
                sAtrib4 = P.ValorControlPantalla("DropDownList4").ToString.Trim
            End If

            sAtrib5 = P.ValorControlPantalla("TextBox5").ToString.Trim
            If sAtrib5.Length = 0 Then
                sAtrib5 = P.ValorControlPantalla("DropDownList5").ToString.Trim
            End If

            sAtrib6 = P.ValorControlPantalla("TextBox6").ToString.Trim
            If sAtrib6.Length = 0 Then
                sAtrib6 = P.ValorControlPantalla("DropDownList6").ToString.Trim
            End If

            sAtrib7 = P.ValorControlPantalla("TextBox7").ToString.Trim
            If sAtrib7.Length = 0 Then
                sAtrib7 = P.ValorControlPantalla("DropDownList7").ToString.Trim
            End If

            sAtrib8 = P.ValorControlPantalla("TextBox8").ToString.Trim
            If sAtrib8.Length = 0 Then
                sAtrib8 = P.ValorControlPantalla("DropDownList8").ToString.Trim
            End If

            sAtrib9 = P.ValorControlPantalla("TextBox9").ToString.Trim
            If sAtrib9.Length = 0 Then
                sAtrib9 = P.ValorControlPantalla("DropDownList9").ToString.Trim
            End If

            sAtrib10 = P.ValorControlPantalla("TextBox10").ToString.Trim
            If sAtrib10.Length = 0 Then
                sAtrib10 = P.ValorControlPantalla("DropDownList10").ToString.Trim
            End If

            sAtrib11 = P.ValorControlPantalla("TextBox11").ToString.Trim
            If sAtrib11.Length = 0 Then
                sAtrib11 = P.ValorControlPantalla("DropDownList11").ToString.Trim
            End If

            sAtrib12 = P.ValorControlPantalla("TextBox12").ToString.Trim
            If sAtrib12.Length = 0 Then
                sAtrib12 = P.ValorControlPantalla("DropDownList12").ToString.Trim
            End If


            db.PA_RegistraPersona_3(P.ValorControlPantalla("txtDoc"), P.ValorControlPantalla("txtnomempr"),
                                    P.ValorControlPantalla("txtNombres"), P.ValorControlPantalla("txtApellidos"),
                                    CDate(P.ValorControlPantalla("TxtFecNac")), P.ValorControlPantalla("cbogenero"),
                                    P.ValorControlPantalla("txtDomic"), P.ValorControlPantalla("txtRefer"), P.ValorControlPantalla("cboZona"),
                                    P.ValorControlPantalla("Txtcel1"), "", P.ValorControlPantalla("TxtCel2"), "", P.ValorControlPantalla("txtDocCony"),
                                    "", P.ValorControlPantalla("TxtApeNomCony"), P.ValorControlPantalla("txttelfijo"), "", P.ValorControlPantalla("cboEstCiv"), 0, 0,
                                    P.ValorControlPantalla("txtDep"), "", P1, P2, "", P.ValorControlPantalla("txtCtaBco"), "", P.ValorControlPantalla("txtrucemp"), P.ValorControlPantalla("txtnomempr"),
                                    P.ValorControlPantalla("txttelemp"), P.ValorControlPantalla("txtdiremp"), P.ValorControlPantalla("txtcargoact"), P.ValorControlPantalla("txtsueldobruto"),
                                sAtrib1, sAtrib2, sAtrib3, sAtrib4, sAtrib5, sAtrib6, sAtrib7, sAtrib9, sAtrib9, sAtrib10, sAtrib11, sAtrib12)

            O.nResultado = 1
        End If

        RegistraPersona = O

    End Function

    Public Function RegistraDatosControl() As outNegocio
        Dim R As New outNegocio

        db.Entry(DatosPantalla).State = EntityState.Modified
        Try
            db.SaveChanges()
        Catch ex As DbUpdateConcurrencyException
            R.nResultado = 0
            R.sRespuesta = ex.Message
            Return R
        End Try

        R.nResultado = 1
        R.sRespuesta = "Proceso Realizado de manera Correcta"

        Return R
    End Function

    Public Function RecuperaListaGruposComunalesParaAprobar() As outNegocio

        Dim R As New outNegocio
        Dim ArrLista() As clsGrupoComunal
        Try


            Dim vListaGrupoComunal = From GC In db.GruposComunales
                                     Join CC In db.CatalogoCodigos On CC.cValor Equals GC.nCodEstado
                                     Join C In db.Creditos On C.nIdGrupo Equals GC.nID
                                     Join A In db.Agencias On A.nCodAge Equals C.nCodAge
                                     Where GC.bActivo = True And CC.nCodigo = 1025 And C.nEstado = 1 And C.nProd = 5
                                     Group By cNombreGrupo = GC.cNombre, cAsesor = C.cCodUsuCob, sNomAge = A.cNomAge Into MyGroup = Group
                                     Select New clsGrupoComunal With {.cNombre = cNombreGrupo, .sCodUsuCob = cAsesor, .sNomAge = sNomAge}


            ArrLista = vListaGrupoComunal.ToArray

            Dim dFecVenc As Date
            Dim nInicio As Integer
            Dim i As Integer

            dFecVenc = Now.Date
            nInicio = 0


            For i = 0 To ArrLista.Count - 1
                ArrLista(i).CargaDatos(ArrLista(i).nIdGrupo)
                nInicio = 0
                For Each s In ArrLista(i).ListaSocios
                    Dim vCalend = From Cal In db.CredCalendario
                                  Where Cal.nCodCred = s.nCodCred And Cal.nEstado = 0
                                  Group By Cal.nCodCred Into Mygroup = Group
                                  Select FecVenc = Mygroup.Min(Function(Cal) Cal.dFecVenc)
                    If nInicio = 0 Then
                        nInicio = 1
                        For Each f In vCalend
                            dFecVenc = f
                        Next
                    Else
                        For Each f In vCalend
                            If DateDiff(DateInterval.Day, f, dFecVenc) >= 0 Then
                                dFecVenc = f
                            End If
                        Next
                    End If
                Next
                ArrLista(i).nDiasAtraso = DateDiff(DateInterval.Day, dFecVenc, Now.Date)
                ArrLista(i).dFecVenc = dFecVenc

            Next

            R.ListaGrupos = ArrLista.AsQueryable
            For Each d In R.ListaGrupos
                dFecVenc = d.dFecVenc
            Next
            R.nResultado = 1
            R.sRespuesta = "Proceso Correcto"
            Return R
        Catch ex As Exception
            R.nResultado = 0
            R.sRespuesta = ex.Message
            Return R
        End Try

    End Function



    Public Function RecuperaIngresosEgresos() As outNegocio

        Dim R As outNegocio
        Dim S As iSesion
        Dim sCodRol As String
        Dim i As Integer
        Dim nCodAgeUsu As Integer

        'S = New iSesion()

        'If Not S.ValidaToken(sToken) Then
        '    R.nResultado = 0
        '    R.sRespuesta = "Token Invalido"
        'End If

        R = New outNegocio()

        'Valida si agencia esta cerrada
        Dim oData As New clsBaseDatos
        If oData.AgenciaCerrada(CInt(nCodAge), Now.Date) = 1 Then
            R.nResultado = 0
            R.sRespuesta = "La Agencia Se Encuentra Cerrada"
            Return R
        End If

        Dim oCaja As New CuadreCajaUsuario
        Dim ArrCaja() As CuadreCajaUsuario

        'Cargar todos los usuarios que pertenecen al supervisor o a la agencia
        Dim vUsu = From U In db.UsuarioRoles
                   Where U.cCodUsu.Trim = sCodUsu.Trim
                   Order By U.cCodRol Descending
                   Select U

        For Each m In vUsu
            sCodRol = m.cCodRol
            Exit For
        Next

        If sCodRol = "5" Then ' Si es Gerente

            Dim vListaAgencias = From U In db.Agencias
                                 Select U

            R.ListaAgencias = vListaAgencias

            Dim vUsuarios = From U In db.Usuarios
                            Select U
            i = 0
            For Each m In vUsuarios
                ReDim Preserve ArrCaja(i)
                oCaja = New CuadreCajaUsuario
                oCaja.CargaData(m.cCodUsu, Now.Date)
                ArrCaja(i) = oCaja
                i = i + 1
            Next


        End If


        If sCodRol = "4" Then ' Si es Adm Agencia


            Dim vAgencia = From U In db.Usuarios
                           Where U.cCodUsu.Trim = sCodUsu.Trim
                           Select U


            For Each m In vAgencia
                nCodAgeUsu = m.nCodAge
                Exit For
            Next

            Dim vListaAgencias = From U In db.Agencias
                                 Where U.nCodAge = nCodAgeUsu
                                 Select U

            R.ListaAgencias = vListaAgencias



            Dim vUsuarios = From U In db.Usuarios
                            Where U.nCodAge = nCodAgeUsu
                            Select U
            i = 0
            For Each m In vUsuarios
                ReDim Preserve ArrCaja(i)
                oCaja = New CuadreCajaUsuario
                oCaja.CargaData(m.cCodUsu, Now.Date)
                ArrCaja(i) = oCaja
                i = i + 1
            Next

        End If

        If sCodRol = "3" Then 'Si es Supervisor

            Dim vAgencia = From U In db.Usuarios
                           Where U.cCodUsu.Trim = sCodUsu.Trim
                           Select U


            For Each m In vAgencia
                nCodAgeUsu = m.nCodAge
                Exit For
            Next

            Dim vListaAgencias = From U In db.Agencias
                                 Where U.nCodAge = nCodAgeUsu
                                 Select U

            R.ListaAgencias = vListaAgencias



            Dim vUsuarios = From U In db.Usuarios
                            Where U.cCodUsuSuperv = sCodUsu
                            Select U
            i = 0
            For Each m In vUsuarios
                ReDim Preserve ArrCaja(i)
                oCaja = New CuadreCajaUsuario
                oCaja.CargaData(m.cCodUsu, Now.Date)
                ArrCaja(i) = oCaja
                i = i + 1
            Next

        End If


        R.CuadreCajaUsuarios = ArrCaja.AsQueryable()

        R.nResultado = 1
        R.sRespuesta = "Proceso Realizado de manera Correcta"

        RecuperaIngresosEgresos = R

    End Function

    Public Function RecuperaRecordCliente() As outNegocio

        Dim R As New outNegocio
        Dim oCliente As New clsPosicionCliente
        Try

            oCliente.CargaDatos(cDoc)

            R.oRecord = oCliente

            R.nResultado = 1
            R.sRespuesta = "Proceso Correcto"
            Return R
        Catch ex As Exception
            R.nResultado = 0
            R.sRespuesta = ex.Message
            Return R
        End Try

    End Function

    Public Function RecuperaListaGruposComunalesParaPagar() As outNegocio

        Dim R As New outNegocio
        Dim ArrLista() As clsGrupoComunal
        Try


            Dim vListaGrupoComunal = From GC In db.GruposComunales
                                     Join CC In db.CatalogoCodigos On CC.cValor Equals GC.nCodEstado
                                     Join C In db.Creditos On C.nIdGrupo Equals GC.nID
                                     Join A In db.Agencias On A.nCodAge Equals C.nCodAge
                                     Where GC.bActivo = True And CC.nCodigo = 1025 And C.nEstado = 4 And C.nProd = 5
                                     Group By cNombreGrupo = GC.cNombre, cAsesor = C.cCodUsuCob, sNomAge = A.cNomAge Into MyGroup = Group
                                     Select New clsGrupoComunal With {.cNombre = cNombreGrupo, .sCodUsuCob = cAsesor, .sNomAge = sNomAge}


            ArrLista = vListaGrupoComunal.ToArray

            Dim dFecVenc As Date
            Dim nInicio As Integer
            Dim i As Integer

            dFecVenc = Now.Date
            nInicio = 0


            For i = 0 To ArrLista.Count - 1
                ArrLista(i).CargaDatos(ArrLista(i).nIdGrupo)
                nInicio = 0
                For Each s In ArrLista(i).ListaSocios
                    Dim vCalend = From Cal In db.CredCalendario
                                  Where Cal.nCodCred = s.nCodCred And Cal.nEstado = 0
                                  Group By Cal.nCodCred Into Mygroup = Group
                                  Select FecVenc = Mygroup.Min(Function(Cal) Cal.dFecVenc)
                    If nInicio = 0 Then
                        nInicio = 1
                        For Each f In vCalend
                            dFecVenc = f
                        Next
                    Else
                        For Each f In vCalend
                            If DateDiff(DateInterval.Day, f, dFecVenc) >= 0 Then
                                dFecVenc = f
                            End If
                        Next
                    End If
                Next
                ArrLista(i).nDiasAtraso = DateDiff(DateInterval.Day, dFecVenc, Now.Date)
                ArrLista(i).dFecVenc = dFecVenc

            Next

            R.ListaGrupos = ArrLista.AsQueryable
            For Each d In R.ListaGrupos
                dFecVenc = d.dFecVenc
            Next
            R.nResultado = 1
            R.sRespuesta = "Proceso Correcto"
            Return R
        Catch ex As Exception
            R.nResultado = 0
            R.sRespuesta = ex.Message
            Return R
        End Try

    End Function

    Public Function RecuperaListaGruposComunales() As outNegocio

        Dim R As New outNegocio
        Dim ArrLista() As clsGrupoComunal
        Try

            Dim vListaGrupoComunal = From GC In db.GruposComunales
                                     Join CC In db.CatalogoCodigos On CC.cValor Equals GC.nCodEstado
                                     Where GC.bActivo = True And CC.nCodigo = 1025 And GC.cCodUsuCob = sCodUsu
                                     Select New clsGrupoComunal With {.cNombre = GC.cNombre, .cLugarProxReunion = GC.cLugarProximaReunion, .cNomEstado = CC.cDescripcion,
                                         .dFechaReunion = GC.dFechaProxReunion, .nCodEstado = GC.nCodEstado, .nIdGrupo = GC.nID, .nNumCiclo = GC.nNumCiclo, .nNumSocios = GC.nNumSocios, .dFecVenc = Nothing, .nDiasAtraso = 0, .nFormaPago = GC.nFormaPago
                                         }




            ArrLista = vListaGrupoComunal.ToArray

            Dim dFecVenc As Date
            Dim nInicio As Integer
            Dim i As Integer

            dFecVenc = Now.Date
            nInicio = 0


            For i = 0 To ArrLista.Count - 1
                ArrLista(i).CargaDatos(ArrLista(i).nIdGrupo)
                nInicio = 0
                For Each s In ArrLista(i).ListaSocios
                    Dim vCalend = From Cal In db.CredCalendario
                                  Where Cal.nCodCred = s.nCodCred And Cal.nEstado = 0
                                  Group By Cal.nCodCred Into Mygroup = Group
                                  Select FecVenc = Mygroup.Min(Function(Cal) Cal.dFecVenc)
                    If nInicio = 0 Then
                        nInicio = 1
                        For Each f In vCalend
                            dFecVenc = f
                        Next
                    Else
                        For Each f In vCalend
                            If DateDiff(DateInterval.Day, f, dFecVenc) >= 0 Then
                                dFecVenc = f
                            End If
                        Next
                    End If
                Next
                ArrLista(i).nDiasAtraso = DateDiff(DateInterval.Day, dFecVenc, Now.Date)
                ArrLista(i).dFecVenc = dFecVenc

            Next

            R.ListaGrupos = ArrLista.AsQueryable
            For Each d In R.ListaGrupos
                dFecVenc = d.dFecVenc
            Next
            R.nResultado = 1
            R.sRespuesta = "Proceso Correcto"
            Return R
        Catch ex As Exception
            R.nResultado = 0
            R.sRespuesta = ex.Message
            Return R
        End Try

    End Function

    Public Function RecuperaDetalleGrupoComunal() As outNegocio

        Dim R As New outNegocio
        Dim GC As New clsGrupoComunal
        Try
            GC.CargaDatos(nIdGrupo)
            R.oGrupoComunal = GC
            R.nResultado = 1
            R.sRespuesta = "Proceso Correcto"
            Return R

        Catch ex As Exception
            R.nResultado = 0
            R.sRespuesta = ex.Message
            Return R

        End Try


    End Function
    Public Function RecuperaDetalleGrupoComunalParaAprobar() As outNegocio

        Dim R As New outNegocio
        Dim GC As New clsGrupoComunal
        Try
            GC.CargaDatos_A_Aprobar(nIdGrupo)
            R.oGrupoComunal = GC
            R.nResultado = 1
            R.sRespuesta = "Proceso Correcto"
            Return R

        Catch ex As Exception
            R.nResultado = 0
            R.sRespuesta = ex.Message
            Return R

        End Try


    End Function
    Public Function RecuperaDetalleGrupoComunalParaPagos() As outNegocio

        Dim R As New outNegocio
        Dim GC As New clsGrupoComunal
        Try
            GC.CargaDatos_A_Pagar(nIdGrupo)
            R.oGrupoComunal = GC
            R.nResultado = 1
            R.sRespuesta = "Proceso Correcto"
            Return R

        Catch ex As Exception
            R.nResultado = 0
            R.sRespuesta = ex.Message
            Return R

        End Try


    End Function

    Public Function RecuperaDetalleGrupoComunalParaDesembolsar() As outNegocio

        Dim R As New outNegocio
        Dim GC As New clsGrupoComunal
        Try
            GC.CargaDatos_A_Desembolsar(nIdGrupo)
            R.oGrupoComunal = GC
            R.nResultado = 1
            R.sRespuesta = "Proceso Correcto"
            Return R

        Catch ex As Exception
            R.nResultado = 0
            R.sRespuesta = ex.Message
            Return R

        End Try


    End Function


    Public Function RegistraGrupoComunal() As outNegocio

        Dim R As New outNegocio

        Try

            oGrupoComunal.RegistraGrupoComunal(R.nResultado, R.sRespuesta)

            Return R

        Catch ex As Exception
            R.nResultado = 0
            R.sRespuesta = ex.Message
            Return R

        End Try


    End Function


    Public Function RegistraGrupoComunalSocio() As outNegocio

        Dim R As New outNegocio

        Try

            oSocio.RegistraSocio(R.nResultado, R.sRespuesta)

            Return R

        Catch ex As Exception
            R.nResultado = 0
            R.sRespuesta = ex.Message
            Return R

        End Try


    End Function
    Public Function RecuperaDatosGrupoComunal() As outNegocio
        Dim R As New outNegocio
        Try
            oGrupoComunal = New clsGrupoComunal

            oGrupoComunal.CargaDatos(nIdGrupo)
            R.oGrupoComunal = oGrupoComunal

            R.nResultado = 1
            R.sRespuesta = "Proceso Correcto"
            Return R

        Catch ex As Exception
            R.nResultado = 0
            R.sRespuesta = ex.Message
            Return R

        End Try


    End Function
    Public Function RecuperaDatosSocio() As outNegocio

        Dim R As New outNegocio


        Try

            oGrupoComunal.CargaSocio()

            If oGrupoComunal.ListaSocios.Count > 0 Then
                For Each p In oGrupoComunal.ListaSocios
                    R.oSocio = p
                    Exit For
                Next
                R.nResultado = 1
                R.sRespuesta = "Proceso Correcto"
                Return R
            Else

                R.nResultado = 0
                R.sRespuesta = "No se Encontró Socio"
                Return R

            End If

            Return R

        Catch ex As Exception
            R.nResultado = 0
            R.sRespuesta = ex.Message
            Return R

        End Try


    End Function

    Public Function BuscarGrupoComunalSocio() As outNegocio

        Dim R As New outNegocio
        Dim oSocio As New clsGrupoComunalSocio

        Try

            Dim vPersona = From P In db.Persona
                           Where P.cDoc = cDoc
                           Select P

            If vPersona.Count > 0 Then
                For Each p In vPersona
                    oSocio.cNombres = p.cNombres
                    oSocio.cApellidos = p.cApellidos
                Next
                R.oSocio = oSocio
                R.nResultado = 1
                R.sRespuesta = "Proceso Correcto"
                Return R
            Else

                R.nResultado = 0
                R.sRespuesta = "No se Encontró la Persona"
                Return R

            End If

            Return R

        Catch ex As Exception
            R.nResultado = 0
            R.sRespuesta = ex.Message
            Return R

        End Try


    End Function

    Public Function EliminaSocioGrupoComunal() As outNegocio

        Dim R As New outNegocio

        Try

            oSocio.EliminaSocio(R.nResultado, R.sRespuesta)

            Return R

        Catch ex As Exception
            R.nResultado = 0
            R.sRespuesta = ex.Message
            Return R

        End Try

    End Function

    Public Function RecuperaConsultaSolicitudes() As outNegocio

        Dim R As New outNegocio
        Dim AbonoDesemb() As dAbonoDesemb
        Dim ListaDesembolsador() As String
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

            oComm = New SqlClient.SqlCommand("PA_RecuperaBandejaConsultaSolicitudes", oConex)
            oComm.Parameters.Add("@sCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oConsultaSolicitudes(i)
                R.oConsultaSolicitudes(i) = New clsConsultaSolicitudes
                R.oConsultaSolicitudes(i).cCliente = dr.Item("cCliente")
                R.oConsultaSolicitudes(i).nCodCred = dr.Item("nCodCred")
                R.oConsultaSolicitudes(i).cDoc = dr.Item("cdoc")
                R.oConsultaSolicitudes(i).cUsuOrig = dr.Item("cUsuOrig")
                R.oConsultaSolicitudes(i).cProd = dr.Item("cProd")
                R.oConsultaSolicitudes(i).cEstado = dr.Item("cEstado")
                R.oConsultaSolicitudes(i).cComentario = dr.Item("cComentario")
                R.oConsultaSolicitudes(i).nPrestamo = dr.Item("nMontoAprobado")
                R.oConsultaSolicitudes(i).nPrestamoOrig = dr.Item("nMontoOriginal")
                R.oConsultaSolicitudes(i).nCuotas = dr.Item("nCuotaAprobado")
                R.oConsultaSolicitudes(i).nCuotasOrig = dr.Item("nCuotaOriginal")
                R.oConsultaSolicitudes(i).cComentario = dr.Item("cComentario")
                i = i + 1
            End While
            dr.Close()
            dr = Nothing


            If i > 0 Then
                i = 0
                For j = 0 To R.oConsultaSolicitudes.Count - 1
                    oComm = New SqlClient.SqlCommand("PA_RecuperaVerificacionxCredito", oConex)
                    oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = R.oConsultaSolicitudes(j).nCodCred
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()

                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

                    While dr.Read
                        ReDim Preserve R.oBandejaVerificacion(i)
                        R.oBandejaVerificacion(i) = New clsVerificacionCliente
                        R.oBandejaVerificacion(i).nCodCred = R.oConsultaSolicitudes(j).nCodCred
                        R.oBandejaVerificacion(i).nVerifCuali = dr.Item("nVerifCuali")
                        R.oBandejaVerificacion(i).nVerifRefe = dr.Item("nVerifRefe")
                        R.oBandejaVerificacion(i).nCumpleConAntNeg = dr.Item("nCumpleConAntNeg")
                        R.oBandejaVerificacion(i).nEsPropietario = dr.Item("nEsPropietario")
                        R.oBandejaVerificacion(i).nPoseeEmpleados = dr.Item("nPoseeEmpleados")
                        R.oBandejaVerificacion(i).nNroEmpleados = dr.Item("nNroEmpleados")
                        R.oBandejaVerificacion(i).nTieneRotuloVisible = dr.Item("nTieneRotuloVisible")
                        R.oBandejaVerificacion(i).nNegoCuentaInfra = dr.Item("nNegoCuentaInfra")
                        R.oBandejaVerificacion(i).nCuentaPermiOpe = dr.Item("nCuentaPermiOpe")
                        R.oBandejaVerificacion(i).nObservaAfluCli = dr.Item("nObservaAfluCli")
                        R.oBandejaVerificacion(i).nUbicaFacilAcceso = dr.Item("nUbicaFacilAcceso")
                        R.oBandejaVerificacion(i).nTieneBuenaRefeZona = dr.Item("nTieneBuenaRefeZona")
                        R.oBandejaVerificacion(i).nTieneProspBuenaActi = dr.Item("nTieneProspBuenaActi")
                        R.oBandejaVerificacion(i).nTieneHoraAtencion = dr.Item("nTieneHoraAtencion")
                        R.oBandejaVerificacion(i).nAquilaCasa = dr.Item("nAquilaCasa")
                        R.oBandejaVerificacion(i).nTiempoAquila = dr.Item("nTiempoAquila")
                        R.oBandejaVerificacion(i).nTieneDependientes = dr.Item("nTieneDependientes")
                        R.oBandejaVerificacion(i).nNroDependientes = dr.Item("nNroDependientes")
                        R.oBandejaVerificacion(i).cHoraInicio = dr.Item("cHoraInicio")
                        R.oBandejaVerificacion(i).cHoraFinal = dr.Item("cHoraFinal")
                        R.oBandejaVerificacion(i).nNivelDelincuencia = dr.Item("nNivelDelincuencia")
                        R.oBandejaVerificacion(i).cComentarioCPD = dr.Item("cComentarioCPD")

                        i = i + 1
                    End While
                    dr.Close()
                    dr = Nothing

                Next
            End If

            If i > 0 Then
                For j = 0 To R.oBandejaVerificacion.Count - 1
                    oComm = New SqlClient.SqlCommand("PA_ReferenciasVericacion", oConex)
                    oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = R.oBandejaVerificacion(j).nCodCred
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()

                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0
                    While dr.Read
                        ReDim Preserve R.oBandejaVerificacion(j).oReferencias(i)
                        R.oBandejaVerificacion(j).oReferencias(i) = New clsVerificacionReferencia
                        R.oBandejaVerificacion(j).oReferencias(i).nCodCred = R.oBandejaVerificacion(j).nCodCred
                        R.oBandejaVerificacion(j).oReferencias(i).cNombres = dr.Item("cNombres")
                        R.oBandejaVerificacion(j).oReferencias(i).cTelefono = dr.Item("cTelefono")
                        R.oBandejaVerificacion(j).oReferencias(i).nRecomendado = dr.Item("nRecomienda")
                        R.oBandejaVerificacion(j).oReferencias(i).cRecomendado = dr.Item("cRecomiendaSustento")
                        i = i + 1
                    End While
                    dr.Close()
                    dr = Nothing

                Next

                For j = 0 To R.oBandejaVerificacion.Count - 1
                    oComm = New SqlClient.SqlCommand("PA_ImagenesVerificacion", oConex)
                    oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = R.oBandejaVerificacion(j).nCodCred
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()
                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0
                    While dr.Read
                        ReDim Preserve R.oBandejaVerificacion(j).oImagenes(i)
                        R.oBandejaVerificacion(j).oImagenes(i) = New clsVerificacionClienteImagenes
                        R.oBandejaVerificacion(j).oImagenes(i).nCodCred = R.oBandejaVerificacion(j).nCodCred
                        R.oBandejaVerificacion(j).oImagenes(i).iImagen = UnicodeBytesToString(dr.Item("nFoto"))
                        R.oBandejaVerificacion(j).oImagenes(i).cTipo = dr.Item("cTipo")
                        R.oBandejaVerificacion(j).oImagenes(i).nID = i
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

    Public Function RegistrarTasacion() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim i, k As Integer
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nResp, nCodTasacion As Integer
        Dim sRespuesta As String
        Dim dr As SqlClient.SqlDataReader

        nResp = 0
        nCodTasacion = 0
        sRespuesta = ""

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Try

            oComm = New SqlClient.SqlCommand("[dbo].[PA_RegistraTasacionCliente]", oConex)
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@cDoc", SqlDbType.VarChar).Value = oTasacionRegistro.cDoc
            oComm.Parameters.Add("@cNombres", SqlDbType.VarChar).Value = oTasacionRegistro.cNombresApellidos
            oComm.Parameters.Add("@cMovil", SqlDbType.VarChar).Value = oTasacionRegistro.cMovil
            oComm.Parameters.Add("@cCorreo", SqlDbType.VarChar).Value = oTasacionRegistro.cCorreo
            oComm.Parameters.Add("@cDireccion", SqlDbType.VarChar).Value = oTasacionRegistro.cDireccion
            oComm.Parameters.Add("@cMotivoEmpenio", SqlDbType.VarChar).Value = oTasacionRegistro.cMotivoEmpeno
            oComm.Parameters.Add("@nPrestamoFinal", SqlDbType.Money).Value = oTasacionRegistro.nMontoFinal
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
            oComm.Parameters.Add("@nCodTasacion", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            nCodTasacion = oComm.Parameters("@nCodTasacion").Value
            nResp = oComm.Parameters("@nResp").Value
            sRespuesta = oComm.Parameters("@Resultado").Value

            'If (Mid(sRespuesta, 1, 2) <> "01") Then
            '    nResp = 0
            'End If
            If nResp = 0 Then
                D(0) = Str(nResp)
                D(1) = sRespuesta
                R.nResultado = CInt(D(0))
                R.sRespuesta = D(1)
                Return R
            End If

            i = 0
            For j = 0 To oTasacionRegistro.oArticulos.Count - 1

                oComm = New SqlClient.SqlCommand("[dbo].[PA_RegistrarArticulosTasacion]", oConex)
                oComm.CommandType = CommandType.StoredProcedure

                oComm.Parameters.Add("@nCodTasacion", SqlDbType.Int).Value = nCodTasacion
                oComm.Parameters.Add("@nItem", SqlDbType.Int).Value = j + 1
                oComm.Parameters.Add("@nCodFamilia", SqlDbType.Int).Value = oTasacionRegistro.oArticulos(j).nCodFamilia
                oComm.Parameters.Add("@nCodMarca", SqlDbType.Int).Value = oTasacionRegistro.oArticulos(j).nCodMarca
                oComm.Parameters.Add("@nCodModelo", SqlDbType.Int).Value = oTasacionRegistro.oArticulos(j).nCodModelo
                oComm.Parameters.Add("@nMontoSolicitado", SqlDbType.Money).Value = oTasacionRegistro.oArticulos(j).nMontoSolicitado
                oComm.Parameters.Add("@bAccesorios", SqlDbType.Int).Value = oTasacionRegistro.oArticulos(j).nCuentaAccesorio
                oComm.Parameters.Add("@nEstadoArticulo", SqlDbType.Money).Value = oTasacionRegistro.oArticulos(j).nEstadoArt
                oComm.Parameters.Add("@nAnioCompra", SqlDbType.Money).Value = oTasacionRegistro.oArticulos(j).nAnioCompra
                oComm.Parameters.Add("@nUso", SqlDbType.Money).Value = oTasacionRegistro.oArticulos(j).nUsa
                oComm.Parameters.Add("@cMarca", SqlDbType.VarChar).Value = oTasacionRegistro.oArticulos(j).cMarca
                oComm.Parameters.Add("@cModelo", SqlDbType.VarChar).Value = oTasacionRegistro.oArticulos(j).cModelo

                oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                oComm.ExecuteNonQuery()
                nResp = oComm.Parameters("@nResp").Value
                sRespuesta = oComm.Parameters("@Resultado").Value

                If nResp = 0 Then
                    D(0) = Str(nResp)
                    D(1) = sRespuesta
                    R.nResultado = CInt(D(0))
                    R.sRespuesta = D(1)
                    Return R
                End If

                For i = 0 To oTasacionRegistro.oArticulos(j).oListFotos.Count - 1
                    oComm = New SqlClient.SqlCommand("[dbo].[PA_RegistrarArticulosImagen]", oConex)
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.Parameters.Add("@nCodTasacion", SqlDbType.Int).Value = nCodTasacion
                    oComm.Parameters.Add("@nItem", SqlDbType.Int).Value = j + 1
                    oComm.Parameters.Add("@iDoc", SqlDbType.Image).Value = oTasacionRegistro.oArticulos(j).oListFotos(i).nImg

                    oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                    oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                    oComm.ExecuteNonQuery()
                    nResp = oComm.Parameters("@nResp").Value
                    sRespuesta = oComm.Parameters("@Resultado").Value

                    If nResp = 0 Then
                        D(0) = Str(nResp)
                        D(1) = sRespuesta
                        R.nResultado = CInt(D(0))
                        R.sRespuesta = D(1)
                        Return R
                    End If

                Next

            Next


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

    Public Function RecuperaRespuestaTasacion() As outNegocio
        Dim R As New outNegocio
        Dim AbonoDesemb() As dAbonoDesemb
        Dim ListaDesembolsador() As String
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
            oComm = New SqlClient.SqlCommand("PA_RecuperaBandejaRespuestaTasacion", oConex)
            oComm.Parameters.Add("@sCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oRespuestasTasacion(i)
                R.oRespuestasTasacion(i) = New clsTasacionCliente
                R.oRespuestasTasacion(i).nCodTasacion = dr.Item("nCodTasacion")
                R.oRespuestasTasacion(i).nMontoFinal = dr.Item("nPrestamoFinal")
                R.oRespuestasTasacion(i).cDoc = dr.Item("cDoc")
                R.oRespuestasTasacion(i).cNombresApellidos = dr.Item("cNombresYApe")
                R.oRespuestasTasacion(i).cMotivoEmpeno = dr.Item("cMotivoEmpenio")
                R.oRespuestasTasacion(i).cMovil = dr.Item("cMovil")
                R.oRespuestasTasacion(i).cCorreo = dr.Item("cCorreo")
                R.oRespuestasTasacion(i).cDireccion = dr.Item("cDireccion")
                R.oRespuestasTasacion(i).cComentarioTasacion = dr.Item("cComentarioTasacion")
                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            If i > 0 Then
                For j = 0 To R.oRespuestasTasacion.Count - 1
                    oComm = New SqlClient.SqlCommand("PA_RecuperaArticuloTasacion", oConex)
                    oComm.Parameters.Add("@nCodTasacion", SqlDbType.Int).Value = R.oRespuestasTasacion(j).nCodTasacion
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()

                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0
                    While dr.Read
                        ReDim Preserve R.oRespuestasTasacion(j).oArticulos(i)
                        R.oRespuestasTasacion(j).oArticulos(i) = New clsArticuloTasacion
                        R.oRespuestasTasacion(j).oArticulos(i).nCodFamilia = dr.Item("nCodFamilia")
                        R.oRespuestasTasacion(j).oArticulos(i).nCodMarca = dr.Item("nCodMarca")
                        R.oRespuestasTasacion(j).oArticulos(i).nCodModelo = dr.Item("nCodModelo")
                        R.oRespuestasTasacion(j).oArticulos(i).nCuentaAccesorio = dr.Item("bAccesorios")
                        R.oRespuestasTasacion(j).oArticulos(i).nUsa = dr.Item("nUso")
                        R.oRespuestasTasacion(j).oArticulos(i).nEstadoArt = dr.Item("nEstadoArtefacto")
                        R.oRespuestasTasacion(j).oArticulos(i).nMontoSolicitado = dr.Item("nPrestamoFinal")
                        R.oRespuestasTasacion(j).oArticulos(i).nItem = dr.Item("nItem")
                        R.oRespuestasTasacion(j).oArticulos(i).cArticuloNombre = dr.Item("cArticuloNombre")
                        i = i + 1
                    End While
                    dr.Close()
                    dr = Nothing
                Next


                For j = 0 To R.oRespuestasTasacion.Count - 1 'Fotos Articulo ************************************** 
                    For k = 0 To R.oRespuestasTasacion(j).oArticulos.Count - 1

                        oComm = New SqlClient.SqlCommand("PA_ImagenesArticuloTasacion", oConex)
                        oComm.Parameters.Add("@nCodTasacion", SqlDbType.Int).Value = R.oRespuestasTasacion(j).nCodTasacion
                        oComm.Parameters.Add("@nItem", SqlDbType.Int).Value = R.oRespuestasTasacion(j).oArticulos(k).nItem
                        oComm.CommandTimeout = 0
                        oComm.CommandType = CommandType.StoredProcedure
                        oComm.ExecuteNonQuery()
                        dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

                        i = 0
                        While dr.Read
                            ReDim Preserve R.oRespuestasTasacion(j).oArticulos(k).oListFotos(i)
                            R.oRespuestasTasacion(j).oArticulos(k).oListFotos(i) = New dFotosNegocio
                            R.oRespuestasTasacion(j).oArticulos(k).oListFotos(i).nImg = (dr.Item("nFoto"))
                            R.oRespuestasTasacion(j).oArticulos(k).oListFotos(i).nCodFoto = i
                            i = i + 1
                        End While
                        dr.Close()
                        dr = Nothing

                    Next

                Next

            End If







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

    Public Function RegistrarRespuestaTasacion() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim i, k As Integer
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nResp, nCodTasacion As Integer
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
            oComm = New SqlClient.SqlCommand("[dbo].[PA_RegistraRespuestaTasacion]", oConex)
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@nCodTasacion", SqlDbType.Int).Value = oTasacionRegistro.nCodTasacion
            oComm.Parameters.Add("@cComentarioCliente", SqlDbType.VarChar).Value = oTasacionRegistro.cComentarioTasacion
            oComm.Parameters.Add("@nEstadoRespuesta", SqlDbType.Int).Value = nCodMotivoNoDesembolso
            oComm.Parameters.Add("@cUser", SqlDbType.VarChar).Value = sCodUsu

            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            nResp = oComm.Parameters("@nResp").Value
            sRespuesta = oComm.Parameters("@Resultado").Value

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

    Public Function RegistrarExperian() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim i, k, j, l, p, q As Integer
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nResp, nCodTasacion As Integer
        Dim sRespuesta As String
        Dim dr As SqlClient.SqlDataReader
        Dim TDoc, NDoc As String
        Dim nEliminar As Integer = 0
        nResp = 0
        nCodTasacion = 0
        sRespuesta = ""
        TDoc = ""
        NDoc = ""
        Dim id As Integer

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Try

            If IsNothing(oDataExperian.respuesta.InfBas) Then
                oConex.Close()
                D(0) = "0"
                D(1) = "ERROR: NO SE ENCONTRÓ DATA"

                R.nResultado = CInt(D(0))
                R.sRespuesta = D(1)
                Return R

        End If

            If oDataExperian.respuesta.InfBas.NDoc.Length > 0 Then
                NDoc = oDataExperian.respuesta.InfBas.NDoc
            End If
            If oDataExperian.respuesta.InfBas.TDoc.Length > 0 Then
                TDoc = oDataExperian.respuesta.InfBas.TDoc
            End If
            ' INFO BASIC
            oComm = New SqlClient.SqlCommand("[dbo].[PA_RegistrarExperianInfoBas]", oConex)
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@TDoc", SqlDbType.VarChar).Value = oDataExperian.respuesta.InfBas.TDoc
            oComm.Parameters.Add("@NDoc", SqlDbType.VarChar).Value = oDataExperian.respuesta.InfBas.NDoc
            oComm.Parameters.Add("@ActEco", SqlDbType.VarChar).Value = oDataExperian.respuesta.InfBas.ActEco
            oComm.Parameters.Add("@AgeRet", SqlDbType.VarChar).Value = oDataExperian.respuesta.InfBas.AgeRet
            oComm.Parameters.Add("@ApeMat", SqlDbType.VarChar).Value = oDataExperian.respuesta.InfBas.ApeMat
            oComm.Parameters.Add("@ApePat", SqlDbType.VarChar).Value = oDataExperian.respuesta.InfBas.ApePat
            oComm.Parameters.Add("@Asi", SqlDbType.VarChar).Value = oDataExperian.respuesta.InfBas.Asi
            oComm.Parameters.Add("@CIIU", SqlDbType.VarChar).Value = oDataExperian.respuesta.InfBas.CIIU
            oComm.Parameters.Add("@EstCon", SqlDbType.VarChar).Value = oDataExperian.respuesta.InfBas.EstCon
            oComm.Parameters.Add("@EstConC", SqlDbType.VarChar).Value = oDataExperian.respuesta.InfBas.EstConC
            oComm.Parameters.Add("@EstDom", SqlDbType.VarChar).Value = oDataExperian.respuesta.InfBas.EstDom
            oComm.Parameters.Add("@EstDomC", SqlDbType.VarChar).Value = oDataExperian.respuesta.InfBas.EstDomC
            oComm.Parameters.Add("@FchInsRRPP", SqlDbType.VarChar).Value = oDataExperian.respuesta.InfBas.FchInsRRPP
            oComm.Parameters.Add("@Fol", SqlDbType.VarChar).Value = oDataExperian.respuesta.InfBas.Fol
            oComm.Parameters.Add("@IniAct", SqlDbType.VarChar).Value = oDataExperian.respuesta.InfBas.IniAct
            oComm.Parameters.Add("@Nom", SqlDbType.VarChar).Value = oDataExperian.respuesta.InfBas.Nom
            oComm.Parameters.Add("@NomCom", SqlDbType.VarChar).Value = oDataExperian.respuesta.InfBas.NomCom
            oComm.Parameters.Add("@NumParReg", SqlDbType.VarChar).Value = oDataExperian.respuesta.InfBas.NumParReg
            oComm.Parameters.Add("@RazSoc", SqlDbType.VarChar).Value = oDataExperian.respuesta.InfBas.RazSoc
            oComm.Parameters.Add("@RelNDoc", SqlDbType.VarChar).Value = oDataExperian.respuesta.InfBas.RelNDoc
            oComm.Parameters.Add("@RelTDoc", SqlDbType.VarChar).Value = oDataExperian.respuesta.InfBas.RelTDoc
            oComm.Parameters.Add("@TipCon", SqlDbType.VarChar).Value = oDataExperian.respuesta.InfBas.TipCon
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
            oComm.ExecuteNonQuery()
            nResp = oComm.Parameters("@nResp").Value
            sRespuesta = oComm.Parameters("@Resultado").Value

            If nResp = 0 Then
                D(0) = Str(nResp)
                D(1) = sRespuesta
                R.nResultado = CInt(D(0))
                R.sRespuesta = D(1)
                Return R
            End If
            ''**** END INF BASIC
            '// Sabio
            If oDataExperian.respuesta.Sabio IsNot Nothing Then
                oComm = New SqlClient.SqlCommand("[dbo].[PA_RegistrarExperianSabio]", oConex)
                oComm.CommandType = CommandType.StoredProcedure
                oComm.Parameters.Add("@TDoc", SqlDbType.VarChar).Value = TDoc
                oComm.Parameters.Add("@NDoc", SqlDbType.VarChar).Value = NDoc

                oComm.Parameters.Add("@Resul", SqlDbType.VarChar).Value = oDataExperian.respuesta.Sabio.Resul
                oComm.Parameters.Add("@ScoreSabio", SqlDbType.VarChar).Value = oDataExperian.respuesta.Sabio.ScoreSabio
                oComm.Parameters.Add("@NivSco", SqlDbType.VarChar).Value = oDataExperian.respuesta.Sabio.NivSco
                oComm.Parameters.Add("@DesSco", SqlDbType.VarChar).Value = oDataExperian.respuesta.Sabio.DesSco
                oComm.Parameters.Add("@CapacidadPago", SqlDbType.VarChar).Value = oDataExperian.respuesta.Sabio.CapacidadPago
                oComm.Parameters.Add("@CuoEst", SqlDbType.VarChar).Value = oDataExperian.respuesta.Sabio.CuoEst

                oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                oComm.ExecuteNonQuery()
                nResp = oComm.Parameters("@nResp").Value
                sRespuesta = oComm.Parameters("@Resultado").Value

                If nResp = 0 Then
                    D(0) = Str(nResp)
                    D(1) = sRespuesta
                    R.nResultado = CInt(D(0))
                    R.sRespuesta = D(1)
                    Return R
                End If
            End If
            '**** END SABIO
            '**** STAR ConRap
            If oDataExperian.respuesta.ConRap IsNot Nothing Then
                Dim nIdConRap As Integer
                oComm = New SqlClient.SqlCommand("[dbo].[PA_RegistrarExperianConRap]", oConex)
                oComm.CommandType = CommandType.StoredProcedure
                oComm.Parameters.Add("@TDoc", SqlDbType.VarChar).Value = TDoc
                oComm.Parameters.Add("@NDoc", SqlDbType.VarChar).Value = NDoc
                oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu

                oComm.Parameters.Add("@nIdConRap", SqlDbType.Int).Direction = ParameterDirection.Output

                oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                oComm.ExecuteNonQuery()
                nResp = oComm.Parameters("@nResp").Value
                sRespuesta = oComm.Parameters("@Resultado").Value
                nIdConRap = oComm.Parameters("@nIdConRap").Value

                If nResp = 0 Then
                    D(0) = Str(nResp)
                    D(1) = sRespuesta
                    R.nResultado = CInt(D(0))
                    R.sRespuesta = D(1)
                    Return R
                End If

                If oDataExperian.respuesta.ConRap.Resumen_ConRap IsNot Nothing Then
                    oComm = New SqlClient.SqlCommand("[dbo].[PA_RegistrarExperianConRap_ResumenConRap]", oConex)
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.Parameters.Add("@id_ConRap", SqlDbType.Int).Value = nIdConRap

                    oComm.Parameters.Add("@Calificativo", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.Resumen_ConRap.Calificativo
                    oComm.Parameters.Add("@DeudaImpaga", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.Resumen_ConRap.DeudaImpaga
                    oComm.Parameters.Add("@DeudaLaboral", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.Resumen_ConRap.DeudaLaboral
                    oComm.Parameters.Add("@DeudaProtestos", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.Resumen_ConRap.DeudaProtestos
                    oComm.Parameters.Add("@DeudaSBSMicrof", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.Resumen_ConRap.DeudaSBSMicrof
                    oComm.Parameters.Add("@DeudaTotal", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.Resumen_ConRap.DeudaTotal
                    oComm.Parameters.Add("@DeudaTributaria", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.Resumen_ConRap.DeudaTributaria
                    oComm.Parameters.Add("@FechaProceso", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.Resumen_ConRap.FechaProceso
                    oComm.Parameters.Add("@NEntFinNR", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.Resumen_ConRap.NEntFinNR
                    oComm.Parameters.Add("@Nota", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.Resumen_ConRap.Nota
                    oComm.Parameters.Add("@NroEntFin", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.Resumen_ConRap.NroEntFin
                    oComm.Parameters.Add("@SemaActual", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.Resumen_ConRap.SemaActual
                    oComm.Parameters.Add("@SemaPeorMejor", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.Resumen_ConRap.SemaPeorMejor
                    oComm.Parameters.Add("@SemaPrevio", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.Resumen_ConRap.SemaPrevio
                    oComm.Parameters.Add("@Semaforos", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.Resumen_ConRap.Semaforos

                    oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                    oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                    oComm.ExecuteNonQuery()
                    nResp = oComm.Parameters("@nResp").Value
                    sRespuesta = oComm.Parameters("@Resultado").Value

                    If nResp = 0 Then
                        D(0) = Str(nResp)
                        D(1) = sRespuesta
                        R.nResultado = CInt(D(0))
                        R.sRespuesta = D(1)
                        Return R
                    End If

                End If

                i = 0
                id = 0
                If oDataExperian.respuesta.ConRap.DeuSBSMicro IsNot Nothing Then
                    If oDataExperian.respuesta.ConRap.DeuSBSMicro.Count > 0 Then
                        For i = 0 To oDataExperian.respuesta.ConRap.DeuSBSMicro.Count - 1

                            oComm = New SqlClient.SqlCommand("[dbo].[PA_RegistrarExperianConRap_DeuSBSMicro]", oConex)
                            oComm.CommandType = CommandType.StoredProcedure
                            oComm.Parameters.Add("@id_ConRap", SqlDbType.Int).Value = nIdConRap

                            oComm.Parameters.Add("@Anio", SqlDbType.Int).Value = oDataExperian.respuesta.ConRap.DeuSBSMicro(i).Ano
                            oComm.Parameters.Add("@Mes", SqlDbType.Int).Value = oDataExperian.respuesta.ConRap.DeuSBSMicro(i).Mes
                            oComm.Parameters.Add("@Califi", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.DeuSBSMicro(i).Califi
                            oComm.Parameters.Add("@NroEnt", SqlDbType.Int).Value = oDataExperian.respuesta.ConRap.DeuSBSMicro(i).NroEnt

                            oComm.Parameters.Add("@nIdConRap_DeuSBSMicro", SqlDbType.Int).Direction = ParameterDirection.Output
                            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                            oComm.ExecuteNonQuery()
                            id = oComm.Parameters("@nIdConRap_DeuSBSMicro").Value
                            nResp = oComm.Parameters("@nResp").Value
                            sRespuesta = oComm.Parameters("@Resultado").Value

                            If nResp = 0 Then
                                D(0) = Str(nResp)
                                D(1) = sRespuesta
                                R.nResultado = CInt(D(0))
                                R.sRespuesta = D(1)
                                Return R
                            End If

                            ' entidad de deusbsmicro
                            If oDataExperian.respuesta.ConRap.DeuSBSMicro(i).Entidad IsNot Nothing Then

                                If oDataExperian.respuesta.ConRap.DeuSBSMicro(i).Entidad.Count > 0 Then

                                    For k = 0 To oDataExperian.respuesta.ConRap.DeuSBSMicro(i).Entidad.Count - 1

                                        oComm = New SqlClient.SqlCommand("[dbo].[PA_RegistrarExperianConRap_DeuSBSMicro_Entidad]", oConex)
                                        oComm.CommandType = CommandType.StoredProcedure
                                        oComm.Parameters.Add("@nIdConRap_DeuSBSMicro", SqlDbType.Int).Value = id
                                        oComm.Parameters.Add("@Cal", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.DeuSBSMicro(i).Entidad(k).Cal
                                        oComm.Parameters.Add("@DiaVen", SqlDbType.Int).Value = oDataExperian.respuesta.ConRap.DeuSBSMicro(i).Entidad(k).DiaVen
                                        oComm.Parameters.Add("@FchPro", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.DeuSBSMicro(i).Entidad(k).FchPro
                                        oComm.Parameters.Add("@FchRep", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.DeuSBSMicro(i).Entidad(k).FchRep
                                        oComm.Parameters.Add("@NomEnt", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.DeuSBSMicro(i).Entidad(k).NomEnt
                                        oComm.Parameters.Add("@SalDeu", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.DeuSBSMicro(i).Entidad(k).SalDeu
                                        oComm.Parameters.Add("@id_ConRap", SqlDbType.Int).Value = nIdConRap

                                        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                        oComm.ExecuteNonQuery()
                                        nResp = oComm.Parameters("@nResp").Value
                                        sRespuesta = oComm.Parameters("@Resultado").Value

                                        If nResp = 0 Then
                                            D(0) = Str(nResp)
                                            D(1) = sRespuesta
                                            R.nResultado = CInt(D(0))
                                            R.sRespuesta = D(1)
                                            Return R
                                        End If

                                    Next

                                End If
                            End If
                            ' fin

                        Next
                    End If
                End If
                ' end DeuSBSMicro


                i = 0
                id = 0
                If oDataExperian.respuesta.ConRap.IndCre IsNot Nothing Then
                    If oDataExperian.respuesta.ConRap.IndCre.Count > 0 Then
                        For i = 0 To oDataExperian.respuesta.ConRap.IndCre.Count - 1
                            oComm = New SqlClient.SqlCommand("[dbo].[PA_RegistrarExperianConRap_IndCre]", oConex)
                            oComm.CommandType = CommandType.StoredProcedure
                            oComm.Parameters.Add("@id_ConRap", SqlDbType.Int).Value = nIdConRap
                            oComm.Parameters.Add("@Indicador", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.IndCre(i).Indicador

                            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                            oComm.ExecuteNonQuery()
                            nResp = oComm.Parameters("@nResp").Value
                            sRespuesta = oComm.Parameters("@Resultado").Value

                            If nResp = 0 Then
                                D(0) = Str(nResp)
                                D(1) = sRespuesta
                                R.nResultado = CInt(D(0))
                                R.sRespuesta = D(1)
                                Return R
                            End If
                        Next
                    End If
                End If
                'fin IndCre

                i = 0
                id = 0
                If oDataExperian.respuesta.ConRap.UtiLinCre IsNot Nothing Then
                    If oDataExperian.respuesta.ConRap.UtiLinCre.Count > 0 Then
                        For i = 0 To oDataExperian.respuesta.ConRap.UtiLinCre.Count - 1
                            oComm = New SqlClient.SqlCommand("[dbo].[PA_RegistrarExperianConRap_UtiLinCre]", oConex)
                            oComm.CommandType = CommandType.StoredProcedure
                            oComm.Parameters.Add("@id_ConRap", SqlDbType.Int).Value = nIdConRap
                            oComm.Parameters.Add("@Inst", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.UtiLinCre(i).Inst
                            oComm.Parameters.Add("@LinApr", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.UtiLinCre(i).LinApr
                            oComm.Parameters.Add("@LinNoUti", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.UtiLinCre(i).LinNoUti
                            oComm.Parameters.Add("@LinUti", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.UtiLinCre(i).LinUti
                            oComm.Parameters.Add("@TipCred", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.UtiLinCre(i).TipCred

                            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                            oComm.ExecuteNonQuery()
                            nResp = oComm.Parameters("@nResp").Value
                            sRespuesta = oComm.Parameters("@Resultado").Value

                            If nResp = 0 Then
                                D(0) = Str(nResp)
                                D(1) = sRespuesta
                                R.nResultado = CInt(D(0))
                                R.sRespuesta = D(1)
                                Return R
                            End If
                        Next
                    End If
                End If
                ' fin UtiLinCre

                i = 0
                id = 0
                j = 0
                Dim nIdConRap_DetVen_Fuente As Integer
                If oDataExperian.respuesta.ConRap.DetVen IsNot Nothing Then
                    If oDataExperian.respuesta.ConRap.DetVen.Count > 0 Then
                        For i = 0 To oDataExperian.respuesta.ConRap.DetVen.Count - 1
                            oComm = New SqlCommand("[dbo].[PA_RegistrarExperianConRap_DetVen]", oConex)
                            oComm.CommandType = CommandType.StoredProcedure
                            oComm.Parameters.Add("@id_ConRap", SqlDbType.Int).Value = nIdConRap
                            oComm.Parameters.Add("@Anio", SqlDbType.Int).Value = oDataExperian.respuesta.ConRap.DetVen(i).Ano
                            oComm.Parameters.Add("@Mes", SqlDbType.Int).Value = oDataExperian.respuesta.ConRap.DetVen(i).Mes
                            oComm.Parameters.Add("@nIdConRap_DetVen", SqlDbType.Int).Direction = ParameterDirection.Output
                            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                            oComm.ExecuteNonQuery()
                            id = oComm.Parameters("@nIdConRap_DetVen").Value
                            nResp = oComm.Parameters("@nResp").Value
                            sRespuesta = oComm.Parameters("@Resultado").Value

                            If nResp = 0 Then
                                D(0) = Str(nResp)
                                D(1) = sRespuesta
                                R.nResultado = CInt(D(0))
                                R.sRespuesta = D(1)
                                Return R
                            End If


                            ' fuente de DetVen
                            If oDataExperian.respuesta.ConRap.DetVen(i).Fuentes IsNot Nothing Then

                                If oDataExperian.respuesta.ConRap.DetVen(i).Fuentes.Count > 0 Then

                                    For k = 0 To oDataExperian.respuesta.ConRap.DetVen(i).Fuentes.Count - 1
                                        nIdConRap_DetVen_Fuente = 0
                                        oComm = New SqlCommand("[dbo].[PA_RegistrarExperianConRap_DetVen_Fuente]", oConex)
                                        oComm.CommandType = CommandType.StoredProcedure
                                        oComm.Parameters.Add("@id_ConRap_DetVen", SqlDbType.Int).Value = id
                                        oComm.Parameters.Add("@IdFue", SqlDbType.Int).Value = oDataExperian.respuesta.ConRap.DetVen(i).Fuentes(k).IdFue
                                        oComm.Parameters.Add("@NomFue", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.DetVen(i).Fuentes(k).NomFue
                                        oComm.Parameters.Add("@VenTot", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.DetVen(i).Fuentes(k).VenTot
                                        oComm.Parameters.Add("@MaxDiaVen", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.DetVen(i).Fuentes(k).MaxDiaVen
                                        oComm.Parameters.Add("@id_ConRap", SqlDbType.Int).Value = nIdConRap

                                        oComm.Parameters.Add("@nIdConRap_DetVen_Fuente", SqlDbType.Int).Direction = ParameterDirection.Output
                                        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                        oComm.ExecuteNonQuery()
                                        nIdConRap_DetVen_Fuente = oComm.Parameters("@nIdConRap_DetVen_Fuente").Value
                                        nResp = oComm.Parameters("@nResp").Value
                                        sRespuesta = oComm.Parameters("@Resultado").Value

                                        If nResp = 0 Then
                                            D(0) = Str(nResp)
                                            D(1) = sRespuesta
                                            R.nResultado = CInt(D(0))
                                            R.sRespuesta = D(1)
                                            Return R
                                        End If



                                        If oDataExperian.respuesta.ConRap.DetVen(i).Fuentes(k).Entidad IsNot Nothing Then
                                            If oDataExperian.respuesta.ConRap.DetVen(i).Fuentes(k).Entidad.Count > 0 Then
                                                For j = 0 To oDataExperian.respuesta.ConRap.DetVen(i).Fuentes(k).Entidad.Count - 1
                                                    oComm = New SqlCommand("[dbo].[PA_RegistrarExperianConRap_DetVen_Fuente_Entidad]", oConex)
                                                    oComm.CommandType = CommandType.StoredProcedure
                                                    oComm.Parameters.Add("@nIdConRap_DetVen_Fuente", SqlDbType.Int).Value = nIdConRap_DetVen_Fuente
                                                    oComm.Parameters.Add("@NomEnt", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.DetVen(i).Fuentes(k).Entidad(j).NomEnt
                                                    oComm.Parameters.Add("@MontDeu", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.DetVen(i).Fuentes(k).Entidad(j).MontDeu
                                                    oComm.Parameters.Add("@DiaVen", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.DetVen(i).Fuentes(k).Entidad(j).DiaVen
                                                    oComm.Parameters.Add("@NumDoc", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.DetVen(i).Fuentes(k).Entidad(j).NumDoc
                                                    oComm.Parameters.Add("@TipEnt", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.DetVen(i).Fuentes(k).Entidad(j).TipEnt
                                                    oComm.Parameters.Add("@FlgCont", SqlDbType.VarChar).Value = oDataExperian.respuesta.ConRap.DetVen(i).Fuentes(k).Entidad(j).FlgCont
                                                    oComm.Parameters.Add("@id_ConRap", SqlDbType.Int).Value = nIdConRap

                                                    oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                                    oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                                    oComm.ExecuteNonQuery()
                                                    nResp = oComm.Parameters("@nResp").Value
                                                    sRespuesta = oComm.Parameters("@Resultado").Value

                                                    If nResp = 0 Then
                                                        D(0) = Str(nResp)
                                                        D(1) = sRespuesta
                                                        R.nResultado = CInt(D(0))
                                                        R.sRespuesta = D(1)
                                                        Return R
                                                    End If
                                                Next
                                            End If
                                        End If

                                    Next

                                End If
                            End If
                            ' fin
                        Next
                    End If
                End If
            End If
            'FIN DetVen


            '**** END ConRap

            i = 0
            id = 0
            k = 0
            j = 0
            l = 0
            p = 0
            q = 0
            'DeuOpe 
            Dim nIdDeuOpe As Integer, nIdDeuOpe_Fecha As Integer, nIdDeuOpe_Fecha_CCN1 As Integer, nIdDeuOpe_Fecha_CCN2 As Integer, nIdDeuOpe_Fecha_CCN3 As Integer
            If oDataExperian.respuesta.DeuOpe IsNot Nothing Then
                If oDataExperian.respuesta.DeuOpe.Count > 0 Then
                    For i = 0 To oDataExperian.respuesta.DeuOpe.Count - 1
                        nIdDeuOpe = 0
                        If i = 0 Then
                            nEliminar = 1
                        Else
                            nEliminar = 0
                        End If
                        oComm = New SqlClient.SqlCommand("[dbo].[PA_RegistrarExperianDeuOpe]", oConex)
                        oComm.CommandType = CommandType.StoredProcedure
                        oComm.Parameters.Add("@TDoc", SqlDbType.VarChar).Value = TDoc
                        oComm.Parameters.Add("@NDoc", SqlDbType.VarChar).Value = NDoc
                        oComm.Parameters.Add("@AcrNom", SqlDbType.VarChar).Value = oDataExperian.respuesta.DeuOpe(i).AcrNom
                        oComm.Parameters.Add("@nEliminar", SqlDbType.Int).Value = nEliminar

                        oComm.Parameters.Add("@nIdDeuOpe", SqlDbType.Int).Direction = ParameterDirection.Output

                        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                        oComm.ExecuteNonQuery()
                        nResp = oComm.Parameters("@nResp").Value
                        sRespuesta = oComm.Parameters("@Resultado").Value
                        nIdDeuOpe = oComm.Parameters("@nIdDeuOpe").Value

                        If nResp = 0 Then
                            D(0) = Str(nResp)
                            D(1) = sRespuesta
                            R.nResultado = CInt(D(0))
                            R.sRespuesta = D(1)
                            Return R
                        End If


                        If oDataExperian.respuesta.DeuOpe(i).Fecha IsNot Nothing Then
                            If oDataExperian.respuesta.DeuOpe(i).Fecha.Count > 0 Then
                                For k = 0 To oDataExperian.respuesta.DeuOpe(i).Fecha.Count - 1
                                    nIdDeuOpe_Fecha = 0
                                    oComm = New SqlClient.SqlCommand("[dbo].[PA_RegistrarExperianDeuOpe_Fecha]", oConex)
                                    oComm.CommandType = CommandType.StoredProcedure
                                    oComm.Parameters.Add("@id_DeuOpe", SqlDbType.Int).Value = nIdDeuOpe
                                    oComm.Parameters.Add("@Califi", SqlDbType.VarChar).Value = oDataExperian.respuesta.DeuOpe(i).Fecha(k).Califi
                                    oComm.Parameters.Add("@FchPro", SqlDbType.VarChar).Value = oDataExperian.respuesta.DeuOpe(i).Fecha(k).FchPro
                                    oComm.Parameters.Add("@FchPub", SqlDbType.VarChar).Value = oDataExperian.respuesta.DeuOpe(i).Fecha(k).FchPub

                                    oComm.Parameters.Add("@nIdDeuOpe_Fecha", SqlDbType.Int).Direction = ParameterDirection.Output

                                    oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                    oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                    oComm.ExecuteNonQuery()
                                    nResp = oComm.Parameters("@nResp").Value
                                    sRespuesta = oComm.Parameters("@Resultado").Value
                                    nIdDeuOpe_Fecha = oComm.Parameters("@nIdDeuOpe_Fecha").Value

                                    If nResp = 0 Then
                                        D(0) = Str(nResp)
                                        D(1) = sRespuesta
                                        R.nResultado = CInt(D(0))
                                        R.sRespuesta = D(1)
                                        Return R
                                    End If



                                    If oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1 IsNot Nothing Then
                                        If oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1.Count > 0 Then
                                            For j = 0 To oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1.Count - 1
                                                nIdDeuOpe_Fecha_CCN1 = 0
                                                oComm = New SqlClient.SqlCommand("[dbo].[PA_RegistrarExperianDeuOpe_Fecha_CCN1]", oConex)
                                                oComm.CommandType = CommandType.StoredProcedure
                                                oComm.Parameters.Add("@id_DeuOpe_Fecha", SqlDbType.Int).Value = nIdDeuOpe_Fecha
                                                oComm.Parameters.Add("@CueCon", SqlDbType.VarChar).Value = oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).CueCon
                                                oComm.Parameters.Add("@DesCueCon", SqlDbType.VarChar).Value = oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).DesCueCon
                                                oComm.Parameters.Add("@DiaVen", SqlDbType.VarChar).Value = oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).DiaVen
                                                oComm.Parameters.Add("@MonDeu", SqlDbType.VarChar).Value = oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).MonDeu
                                                oComm.Parameters.Add("@id_DeuOpe", SqlDbType.Int).Value = nIdDeuOpe
                                                oComm.Parameters.Add("@nIdDeuOpe_Fecha_CCN1", SqlDbType.Int).Direction = ParameterDirection.Output

                                                oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                                oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                                oComm.ExecuteNonQuery()
                                                nResp = oComm.Parameters("@nResp").Value
                                                sRespuesta = oComm.Parameters("@Resultado").Value
                                                nIdDeuOpe_Fecha_CCN1 = oComm.Parameters("@nIdDeuOpe_Fecha_CCN1").Value

                                                If nResp = 0 Then
                                                    D(0) = Str(nResp)
                                                    D(1) = sRespuesta
                                                    R.nResultado = CInt(D(0))
                                                    R.sRespuesta = D(1)
                                                    Return R
                                                End If


                                                If oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).CueConNiv2 IsNot Nothing Then
                                                    If oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).CueConNiv2.Count > 0 Then
                                                        For l = 0 To oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).CueConNiv2.Count - 1
                                                            nIdDeuOpe_Fecha_CCN2 = 0
                                                            oComm = New SqlClient.SqlCommand("[dbo].[PA_RegistrarExperianDeuOpe_Fecha_CCN1_CCN2]", oConex)
                                                            oComm.CommandType = CommandType.StoredProcedure
                                                            oComm.Parameters.Add("@id_DeuOpe_Fecha_CCN1", SqlDbType.Int).Value = nIdDeuOpe_Fecha_CCN1
                                                            oComm.Parameters.Add("@CueCon", SqlDbType.VarChar).Value = oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).CueConNiv2(l).CueCon
                                                            oComm.Parameters.Add("@DesCueCon", SqlDbType.VarChar).Value = oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).CueConNiv2(l).DesCueCon
                                                            oComm.Parameters.Add("@DiaVen", SqlDbType.VarChar).Value = oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).CueConNiv2(l).Diaven
                                                            oComm.Parameters.Add("@MonDeu", SqlDbType.VarChar).Value = oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).CueConNiv2(l).MonDeu
                                                            oComm.Parameters.Add("@id_DeuOpe", SqlDbType.Int).Value = nIdDeuOpe
                                                            oComm.Parameters.Add("@nIdDeuOpe_Fecha_CCN1_CCN2", SqlDbType.Int).Direction = ParameterDirection.Output

                                                            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                                            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                                            oComm.ExecuteNonQuery()
                                                            nResp = oComm.Parameters("@nResp").Value
                                                            sRespuesta = oComm.Parameters("@Resultado").Value
                                                            nIdDeuOpe_Fecha_CCN2 = oComm.Parameters("@nIdDeuOpe_Fecha_CCN1_CCN2").Value

                                                            If nResp = 0 Then
                                                                D(0) = Str(nResp)
                                                                D(1) = sRespuesta
                                                                R.nResultado = CInt(D(0))
                                                                R.sRespuesta = D(1)
                                                                Return R
                                                            End If


                                                            If oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).CueConNiv2(l).CueConNiv3 IsNot Nothing Then
                                                                If oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).CueConNiv2(l).CueConNiv3.Count > 0 Then
                                                                    For p = 0 To oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).CueConNiv2(l).CueConNiv3.Count - 1
                                                                        nIdDeuOpe_Fecha_CCN3 = 0
                                                                        oComm = New SqlClient.SqlCommand("[dbo].[PA_RegistrarExperianDeuOpe_Fecha_CCN1_CCN2_CCN3]", oConex)
                                                                        oComm.CommandType = CommandType.StoredProcedure
                                                                        oComm.Parameters.Add("@id_DeuOpe_Fecha_CCN1_CCN2", SqlDbType.Int).Value = nIdDeuOpe_Fecha_CCN2
                                                                        oComm.Parameters.Add("@CueCon", SqlDbType.VarChar).Value = oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).CueConNiv2(l).CueConNiv3(p).CueCon
                                                                        oComm.Parameters.Add("@DesCueCon", SqlDbType.VarChar).Value = oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).CueConNiv2(l).CueConNiv3(p).DesCueCon
                                                                        oComm.Parameters.Add("@DiaVen", SqlDbType.VarChar).Value = oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).CueConNiv2(l).CueConNiv3(p).Diaven
                                                                        oComm.Parameters.Add("@MonDeu", SqlDbType.VarChar).Value = oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).CueConNiv2(l).CueConNiv3(p).MonDeu
                                                                        oComm.Parameters.Add("@id_DeuOpe", SqlDbType.Int).Value = nIdDeuOpe
                                                                        oComm.Parameters.Add("@nIdDeuOpe_Fecha_CCN1_CCN2_CCN3", SqlDbType.Int).Direction = ParameterDirection.Output

                                                                        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                                                        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                                                        oComm.ExecuteNonQuery()
                                                                        nResp = oComm.Parameters("@nResp").Value
                                                                        sRespuesta = oComm.Parameters("@Resultado").Value
                                                                        nIdDeuOpe_Fecha_CCN3 = oComm.Parameters("@nIdDeuOpe_Fecha_CCN1_CCN2_CCN3").Value

                                                                        If nResp = 0 Then
                                                                            D(0) = Str(nResp)
                                                                            D(1) = sRespuesta
                                                                            R.nResultado = CInt(D(0))
                                                                            R.sRespuesta = D(1)
                                                                            Return R
                                                                        End If



                                                                        If oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).CueConNiv2(l).CueConNiv3(p).CueConNiv4 IsNot Nothing Then
                                                                            If oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).CueConNiv2(l).CueConNiv3(p).CueConNiv4.Count > 0 Then
                                                                                For q = 0 To oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).CueConNiv2(l).CueConNiv3(p).CueConNiv4.Count - 1
                                                                                    'nIdDeuOpe_Fecha_CCN3 = 0
                                                                                    oComm = New SqlCommand("[dbo].[PA_RegistrarExperianDeuOpe_Fecha_CCN1_CCN2_CCN3_CCN4]", oConex)
                                                                                    oComm.CommandType = CommandType.StoredProcedure
                                                                                    oComm.Parameters.Add("@id_DeuOpe_Fecha_CCN1_CCN2_CCN3", SqlDbType.Int).Value = nIdDeuOpe_Fecha_CCN3
                                                                                    oComm.Parameters.Add("@CueCon", SqlDbType.VarChar).Value = oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).CueConNiv2(l).CueConNiv3(p).CueConNiv4(q).CueCon
                                                                                    oComm.Parameters.Add("@DesCueCon", SqlDbType.VarChar).Value = oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).CueConNiv2(l).CueConNiv3(p).CueConNiv4(q).DesCueCon
                                                                                    oComm.Parameters.Add("@DiaVen", SqlDbType.VarChar).Value = oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).CueConNiv2(l).CueConNiv3(p).CueConNiv4(q).Diaven
                                                                                    oComm.Parameters.Add("@MonDeu", SqlDbType.VarChar).Value = oDataExperian.respuesta.DeuOpe(i).Fecha(k).CueConNiv1(j).CueConNiv2(l).CueConNiv3(p).CueConNiv4(q).MonDeu
                                                                                    oComm.Parameters.Add("@id_DeuOpe", SqlDbType.Int).Value = nIdDeuOpe
                                                                                    'oComm.Parameters.Add("@nIdDeuOpe_Fecha_CCN1_CCN2_CCN3", SqlDbType.Int).Direction = ParameterDirection.Output

                                                                                    oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                                                                    oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                                                                    oComm.ExecuteNonQuery()
                                                                                    nResp = oComm.Parameters("@nResp").Value
                                                                                    sRespuesta = oComm.Parameters("@Resultado").Value
                                                                                    'nIdDeuOpe_Fecha_CCN3 = oComm.Parameters("@nIdDeuOpe_Fecha_CCN1_CCN2_CCN3").Value

                                                                                    If nResp = 0 Then
                                                                                        D(0) = Str(nResp)
                                                                                        D(1) = sRespuesta
                                                                                        R.nResultado = CInt(D(0))
                                                                                        R.sRespuesta = D(1)
                                                                                        Return R
                                                                                    End If

                                                                                Next
                                                                            End If
                                                                        End If


                                                                    Next
                                                                End If
                                                            End If

                                                        Next
                                                    End If
                                                End If


                                            Next
                                        End If
                                    End If


                                Next
                            End If
                        End If


                    Next
                End If
            End If
            '**** END DeuOpe
            'OtroProd
            If oDataExperian.respuesta.OtroProd IsNot Nothing Then
                If oDataExperian.respuesta.OtroProd.Count > 0 Then
                    For i = 0 To oDataExperian.respuesta.OtroProd.Count - 1
                        If i = 0 Then
                            nEliminar = 1
                        Else
                            nEliminar = 0
                        End If
                        oComm = New SqlClient.SqlCommand("[dbo].[PA_RegistrarExperianOtroProd]", oConex)
                        oComm.CommandType = CommandType.StoredProcedure
                        oComm.Parameters.Add("@TDoc", SqlDbType.VarChar).Value = TDoc
                        oComm.Parameters.Add("@NDoc", SqlDbType.VarChar).Value = NDoc
                        oComm.Parameters.Add("@CodPrd", SqlDbType.VarChar).Value = oDataExperian.respuesta.OtroProd(i).CodPrd
                        oComm.Parameters.Add("@NomPrd", SqlDbType.VarChar).Value = oDataExperian.respuesta.OtroProd(i).NomPrd
                        oComm.Parameters.Add("@Valor", SqlDbType.VarChar).Value = oDataExperian.respuesta.OtroProd(i).Valor
                        oComm.Parameters.Add("@Desc1", SqlDbType.VarChar).Value = oDataExperian.respuesta.OtroProd(i).Desc1
                        oComm.Parameters.Add("@Desc2", SqlDbType.VarChar).Value = oDataExperian.respuesta.OtroProd(i).Desc2
                        oComm.Parameters.Add("@nEliminar", SqlDbType.Int).Value = nEliminar
                        'oComm.Parameters.Add("@nIdDeuOpe", SqlDbType.Int).Direction = ParameterDirection.Output

                        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                        oComm.ExecuteNonQuery()
                        nResp = oComm.Parameters("@nResp").Value
                        sRespuesta = oComm.Parameters("@Resultado").Value
                        'nIdDeuOpe = oComm.Parameters("@nIdDeuOpe").Value

                        If nResp = 0 Then
                            D(0) = Str(nResp)
                            D(1) = sRespuesta
                            R.nResultado = CInt(D(0))
                            R.sRespuesta = D(1)
                            Return R
                        End If
                    Next
                End If
            End If
            '**** END OtroProd

            'HueCons
            If oDataExperian.respuesta.HueCons IsNot Nothing Then
                If oDataExperian.respuesta.HueCons.Count > 0 Then
                    For i = 0 To oDataExperian.respuesta.HueCons.Count - 1
                        If i = 0 Then
                            nEliminar = 1
                        Else
                            nEliminar = 0
                        End If
                        oComm = New SqlCommand("[dbo].[PA_RegistrarExperianHueCons]", oConex)
                        oComm.CommandType = CommandType.StoredProcedure
                        oComm.Parameters.Add("@TDoc", SqlDbType.VarChar).Value = TDoc
                        oComm.Parameters.Add("@NDoc", SqlDbType.VarChar).Value = NDoc

                        oComm.Parameters.Add("@ActEco", SqlDbType.VarChar).Value = oDataExperian.respuesta.HueCons(i).ActEco
                        oComm.Parameters.Add("@CanCon", SqlDbType.VarChar).Value = oDataExperian.respuesta.HueCons(i).CanCon
                        oComm.Parameters.Add("@MesAnio", SqlDbType.VarChar).Value = oDataExperian.respuesta.HueCons(i).MesAnio
                        oComm.Parameters.Add("@nEliminar", SqlDbType.Int).Value = nEliminar
                        'oComm.Parameters.Add("@nIdDeuOpe", SqlDbType.Int).Direction = ParameterDirection.Output

                        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                        oComm.ExecuteNonQuery()
                        nResp = oComm.Parameters("@nResp").Value
                        sRespuesta = oComm.Parameters("@Resultado").Value
                        'nIdDeuOpe = oComm.Parameters("@nIdDeuOpe").Value

                        If nResp = 0 Then
                            D(0) = Str(nResp)
                            D(1) = sRespuesta
                            R.nResultado = CInt(D(0))
                            R.sRespuesta = D(1)
                            Return R
                        End If
                    Next
                End If
            End If
            '**** END HueCons

            'DetHueCons
            If oDataExperian.respuesta.DetHueCons IsNot Nothing Then
                If oDataExperian.respuesta.DetHueCons.Count > 0 Then
                    For i = 0 To oDataExperian.respuesta.DetHueCons.Count - 1
                        If i = 0 Then
                            nEliminar = 1
                        Else
                            nEliminar = 0
                        End If
                        oComm = New SqlCommand("[dbo].[PA_RegistrarExperianDetHueCons]", oConex)
                        oComm.CommandType = CommandType.StoredProcedure
                        oComm.Parameters.Add("@TDoc", SqlDbType.VarChar).Value = TDoc
                        oComm.Parameters.Add("@NDoc", SqlDbType.VarChar).Value = NDoc
                        oComm.Parameters.Add("@ActEco", SqlDbType.VarChar).Value = oDataExperian.respuesta.DetHueCons(i).ActEco
                        oComm.Parameters.Add("@CanCon", SqlDbType.VarChar).Value = oDataExperian.respuesta.DetHueCons(i).CanCon
                        oComm.Parameters.Add("@EntNom", SqlDbType.VarChar).Value = oDataExperian.respuesta.DetHueCons(i).EntNom
                        oComm.Parameters.Add("@FchUCo", SqlDbType.VarChar).Value = oDataExperian.respuesta.DetHueCons(i).FchUCo
                        oComm.Parameters.Add("@MesAnio", SqlDbType.VarChar).Value = oDataExperian.respuesta.DetHueCons(i).MesAnio
                        oComm.Parameters.Add("@nEliminar", SqlDbType.Int).Value = nEliminar
                        'oComm.Parameters.Add("@nIdDeuOpe", SqlDbType.Int).Direction = ParameterDirection.Output

                        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                        oComm.ExecuteNonQuery()
                        nResp = oComm.Parameters("@nResp").Value
                        sRespuesta = oComm.Parameters("@Resultado").Value
                        'nIdDeuOpe = oComm.Parameters("@nIdDeuOpe").Value

                        If nResp = 0 Then
                            D(0) = Str(nResp)
                            D(1) = sRespuesta
                            R.nResultado = CInt(D(0))
                            R.sRespuesta = D(1)
                            Return R
                        End If
                    Next
                End If
            End If
            '**** END DetHueCons

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

    Public Function RecuperarExperian() As outNegocio
        Dim R As New outNegocio
        Dim AbonoDesemb() As dAbonoDesemb
        Dim ListaDesembolsador() As String
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

            oComm = New SqlClient.SqlCommand("PA_ObtenerComportamientoPago", oConex)
            oComm.Parameters.Add("@TipoDoc", SqlDbType.VarChar, 5).Value = "D"
            oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 10).Value = cDoc
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oCatalogocodigos_2(i)
                R.oCatalogocodigos_2(i) = New clsCatalogocodigos
                R.oCatalogocodigos_2(i).cDescripcion = dr.Item("cDesc")
                R.oCatalogocodigos_2(i).cValor = dr.Item("cPuntos")
                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            '***********************
            oComm = New SqlCommand("PA_RecuperaEntidadExperian", oConex)
            oComm.Parameters.Add("@TipoDoc", SqlDbType.VarChar, 5).Value = "D"
            oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 10).Value = cDoc
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read
                ReDim Preserve R.oCreditosAAprobar(i)
                R.oCreditosAAprobar(i) = New clsCreditosAAprobar
                R.oCreditosAAprobar(i).cDoc = dr.Item("cEntidad")
                R.oCreditosAAprobar(i).Cliente = dr.Item("cDiaVenc")
                R.oCreditosAAprobar(i).credito = dr.Item("cSalDeu")
                R.oCreditosAAprobar(i).ESTADO = dr.Item("cCalificacion")
                i = i + 1
            End While
            dr.Close()
            dr = Nothing
            '**************


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
    Public Function RecuperaConsultaBonos() As outNegocio
        Dim R As New outNegocio
        Dim AbonoDesemb() As dAbonoDesemb
        Dim ListaDesembolsador() As String
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


            oComm = New SqlCommand("PA_RecuperaConsultaBono", oConex)
            oComm.Parameters.Add("@sCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oConsultaBonos(i)
                R.oConsultaBonos(i) = New clsBono
                R.oConsultaBonos(i).cComentario = dr.Item("cComentario")
                R.oConsultaBonos(i).cConcepto = dr.Item("cConcepto")
                R.oConsultaBonos(i).nMeta = dr.Item("nMeta")
                R.oConsultaBonos(i).nMetaReal = dr.Item("nMetaReal")
                R.oConsultaBonos(i).nMontoBono = dr.Item("nMontoBono")
                R.oConsultaBonos(i).nComision = dr.Item("nComision")
                i = i + 1
            End While
            dr.Close()
            dr = Nothing




            oConex.Close()
            D(0) = "1"
            D(1) = "Consulta Correcta"
        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message
        End Try
        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))
        Return R

    End Function

    Public Function RecuperaConsultaClienteTasaPreferencial() As outNegocio
        Dim R As New outNegocio
        R.oDatosAprobacion = New clsDatosAprobacion
        Dim AbonoDesemb() As dAbonoDesemb
        Dim ListaDesembolsador() As String
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


            oComm = New SqlCommand("PA_RecuperaClienteTasaPreferencial", oConex)
            oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 20).Value = cDoc
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                R.oDatosAprobacion.nPrestamo = CDbl(dr.Item("nPrestamo")).ToString
                R.oDatosAprobacion.cApeNom = dr.Item("cCliente")
                R.oDatosAprobacion.Tasa = CDbl(dr.Item("nTasaComp")).ToString
                R.oDatosAprobacion.nCuotas = CDbl(dr.Item("nCuotas")).ToString
                R.oDatosAprobacion.nMontoPagoUltimo = CDbl(dr.Item("nMontoCuota")).ToString
                R.oDatosAprobacion.cNomProd = dr.Item("cProducto")
                R.oDatosAprobacion.nCodCred = dr.Item("nCodCred").ToString
                R.oDatosAprobacion.nCodAge = dr.Item("nCodAge").ToString
            End While
            dr.Close()
            dr = Nothing

            'R.oDatosAprobacion.nPrestamo = oComm.Parameters("nPrestamo").Value
            'R.oDatosAprobacion.cApeNom = oComm.Parameters("cCliente").Value
            'R.oDatosAprobacion.Tasa = Str(oComm.Parameters("nTasaComp").Value)
            'R.oDatosAprobacion.nCuotas = Str(oComm.Parameters("nCuotas").Value)
            'R.oDatosAprobacion.nMontoPagoUltimo = Str(oComm.Parameters("nMontoCuota").Value)
            'R.oDatosAprobacion.cNomProd = oComm.Parameters("cProducto").Value
            'R.oDatosAprobacion.nCodCred = Str(oComm.Parameters("nCodCred").Value)
            'R.oDatosAprobacion.nCodAge = Str(oComm.Parameters("nCodAge").Value)

            oConex.Close()
            D(0) = "1"
            D(1) = "Consulta Correcta"
        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message
        End Try
        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))
        Return R

    End Function

    Public Function RegistrarSolicitudTasaPreferencial() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim i, k As Integer
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nResp, nCodTasacion As Integer
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
            oComm = New SqlClient.SqlCommand("[dbo].[PA_RegistraSolicitudTasaPreferencial]", oConex)
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = oTasacionRegistro.nCodTasacion
            oComm.Parameters.Add("@nTasaSol", SqlDbType.Money).Value = oTasacionRegistro.nMontoFinal
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu

            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            nResp = oComm.Parameters("@nResp").Value
            sRespuesta = oComm.Parameters("@Resultado").Value

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

    Public Function RecuperaConsultaClienteTasaPreferencialaAprobar() As outNegocio
        Dim R As New outNegocio
        R.oDatosAprobacion = New clsDatosAprobacion
        Dim AbonoDesemb() As dAbonoDesemb
        Dim ListaDesembolsador() As String
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


            oComm = New SqlCommand("PA_RecuperaClienteTasaPreferencialaAprobar", oConex)
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 100).Value = sCodUsu
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oTasaPreferencial(i)
                R.oTasaPreferencial(i) = New clsTasaPreferencial
                R.oTasaPreferencial(i).cCliente = dr.Item("cCliente")
                R.oTasaPreferencial(i).cDni = dr.Item("cDni")
                R.oTasaPreferencial(i).cProducto = dr.Item("cProducto")
                R.oTasaPreferencial(i).cTasa = CDbl(dr.Item("nTasa")).ToString
                R.oTasaPreferencial(i).cNuevaTasa = CDbl(dr.Item("nNuevaTasa")).ToString
                R.oTasaPreferencial(i).cMonto = CDbl(dr.Item("nMonto")).ToString
                R.oTasaPreferencial(i).nNumSolicitud = dr.Item("nNumSolicitud")
                R.oTasaPreferencial(i).cMontoCuota = CDbl(dr.Item("nMontoCuota")).ToString
                R.oTasaPreferencial(i).cFechaSol = dr.Item("dFechaSolicitud")
                R.oTasaPreferencial(i).cCuotas = CInt(dr.Item("nCuotas")).ToString
                i = i + 1
            End While
            dr.Close()
            dr = Nothing


            oConex.Close()
            D(0) = "1"
            D(1) = "Consulta Correcta"
        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message
        End Try
        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))
        Return R

    End Function

    Public Function RegistrarAprobacionTasaPreferencial() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim i, k As Integer
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nResp, nCodTasacion As Integer
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
            oComm = New SqlClient.SqlCommand("[dbo].[PA_RegistraAprobacionTasaPreferencial]", oConex)
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@nNumSol", SqlDbType.Int).Value = nCodcred
            oComm.Parameters.Add("@nEstadoAprob", SqlDbType.Int).Value = nCodRol
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu

            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            nResp = oComm.Parameters("@nResp").Value
            sRespuesta = oComm.Parameters("@Resultado").Value

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

    Public Function RecuperaBandejaAprobacionGrupo() As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_LimpiaTempoAprobacionCredito", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()


            'For k = 0 To oBandejaAprobacion.Count - 1

            '    oComm = New SqlClient.SqlCommand("PA_RegistraTempoAprobacionCredito", oConex)
            '    oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            '    oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = oBandejaAprobacion(k).ncodcred
            '    oComm.CommandTimeout = 0
            '    oComm.CommandType = CommandType.StoredProcedure
            '    oComm.ExecuteNonQuery()

            'Next

            oComm = New SqlClient.SqlCommand("PA_RecuperaBandejaAprobacionGrupo", oConex)
            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
            oComm.Parameters.Add("@pnIDGrupo", SqlDbType.Int).Value = nCodRol

            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oAprobaciones(i)
                R.oAprobaciones(i) = New clsAprobacionNueva

                R.oAprobaciones(i).cDoc = dr.Item("cDoc")
                R.oAprobaciones(i).cCliente = dr.Item("cCliente")
                R.oAprobaciones(i).cMovil1 = dr.Item("cMovil1")
                R.oAprobaciones(i).cGenero = dr.Item("cGenero")
                R.oAprobaciones(i).cEstadoCiv = dr.Item("cEstadoCiv")
                R.oAprobaciones(i).cCategoria2 = dr.Item("cCategoria2")
                R.oAprobaciones(i).cCodUsuCob = dr.Item("cCodUsuCob")
                R.oAprobaciones(i).cDistritoDom = dr.Item("cDistritoDom")
                R.oAprobaciones(i).cZonaDom = dr.Item("cZonaDom")
                R.oAprobaciones(i).cDirecccionDom = dr.Item("cDirecccionDom")
                R.oAprobaciones(i).nLatitud = dr.Item("nLatitud")
                R.oAprobaciones(i).nLongitud = dr.Item("nLongitud")
                R.oAprobaciones(i).ccodusu = dr.Item("ccodusu")
                R.oAprobaciones(i).dFecha = dr.Item("dFecha")
                R.oAprobaciones(i).ncodcred = dr.Item("ncodcred")
                R.oAprobaciones(i).cRazonSocial = dr.Item("cRazonSocial")
                R.oAprobaciones(i).cActividadEcon = dr.Item("cActividadEcon")
                R.oAprobaciones(i).nLatNeg = dr.Item("nLatNeg")
                R.oAprobaciones(i).nLongNeg = dr.Item("nLongNeg")
                R.oAprobaciones(i).nVentasDiarias = dr.Item("nVentasDiarias")
                R.oAprobaciones(i).cDistritoNeg = dr.Item("cDistritoNeg")
                R.oAprobaciones(i).cZonaNeg = dr.Item("cZonaNeg")
                R.oAprobaciones(i).cDireccionNeg = dr.Item("cDireccionNeg")
                R.oAprobaciones(i).cResultadoCall = dr.Item("cResultadoCall")
                R.oAprobaciones(i).dFechaCall = dr.Item("dFechaCall")
                R.oAprobaciones(i).cResultadoVerif = dr.Item("cResultadoVerif")
                R.oAprobaciones(i).cUsuVerificador = dr.Item("cUsuVerificador")
                R.oAprobaciones(i).cComentarioSustento = dr.Item("cComentarioSustento")
                R.oAprobaciones(i).nPrestamoMaximo = dr.Item("nPrestamoMaximo")
                R.oAprobaciones(i).nTasaAsignada = dr.Item("nTasaAsignada")
                R.oAprobaciones(i).nPrestamo = dr.Item("nPrestamo")
                R.oAprobaciones(i).nCuotas = dr.Item("nCuotas")
                R.oAprobaciones(i).nMontoCuota = dr.Item("nMontoCuota")
                R.oAprobaciones(i).cTipoAprobacion = dr.Item("cTipoAprobacion")
                R.oAprobaciones(i).nPagoPromedio = dr.Item("nPagoPromedio")
                R.oAprobaciones(i).nMontoCuotaAnt = dr.Item("nMontoCuotaAnt")
                R.oAprobaciones(i).nPrestamoAnt = dr.Item("nPrestamoAnt")
                R.oAprobaciones(i).nProd = dr.Item("nProd")
                R.oAprobaciones(i).nCodAge = dr.Item("nCodAge")
                R.oAprobaciones(i).cProducto = dr.Item("cProducto")
                R.oAprobaciones(i).nVerificacion = dr.Item("nVerificacion")
                R.oAprobaciones(i).nExcepcion = dr.Item("nExcepcion")
                R.oAprobaciones(i).cComentarioExcepcion = dr.Item("cComentarioExcepcion")

                i = i + 1

            End While
            dr.Close()
            dr = Nothing

            If i > 0 Then

                For j = 0 To R.oAprobaciones.Count - 1

                    oComm = New SqlClient.SqlCommand("PA_RecuperaRecordCliente", oConex)
                    oComm.Parameters.Add("@psDoc", SqlDbType.Int).Value = R.oAprobaciones(j).cDoc
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()

                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0

                    While dr.Read
                        ReDim Preserve R.oAprobaciones(j).oRecord(i)
                        R.oAprobaciones(j).oRecord(i) = New clsRecordNuevo
                        R.oAprobaciones(j).oRecord(i).ncodcred = dr.Item("ncodcred")
                        R.oAprobaciones(j).oRecord(i).cEstado = dr.Item("cEstado")
                        R.oAprobaciones(j).oRecord(i).nDiasAtraso = dr.Item("nDiasAtraso")
                        R.oAprobaciones(j).oRecord(i).dFecVig = dr.Item("dFecVig")
                        R.oAprobaciones(j).oRecord(i).dFecUltPago = dr.Item("dFecUltPago")
                        R.oAprobaciones(j).oRecord(i).nDiasAtrasoProm = dr.Item("nDiasAtrasoProm")
                        R.oAprobaciones(j).oRecord(i).nPrestamo = dr.Item("nPrestamo")
                        R.oAprobaciones(j).oRecord(i).nMontoCuota = dr.Item("nMontoCuota")
                        R.oAprobaciones(j).oRecord(i).nMontoPagoPromedio = dr.Item("nMontoPagoPromedio")
                        R.oAprobaciones(j).oRecord(i).nTasaComp = dr.Item("nTasaComp")
                        R.oAprobaciones(j).oRecord(i).nCuotasPag = dr.Item("nCuotasPagadas")
                        R.oAprobaciones(j).oRecord(i).nCuotasPend = dr.Item("nCuotasPendientes")

                        i = i + 1

                    End While
                    dr.Close()
                    dr = Nothing

                Next


                For j = 0 To R.oAprobaciones.Count - 1

                    oComm = New SqlClient.SqlCommand("PA_RecuperaHistoriaSBS", oConex)
                    oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = R.oAprobaciones(j).cDoc
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()

                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0

                    While dr.Read
                        ReDim Preserve R.oAprobaciones(j).oSBS(i)
                        R.oAprobaciones(j).oSBS(i) = New clsSBSHist
                        R.oAprobaciones(j).oSBS(i).Fec_rep = dr.Item("Fec_rep")
                        R.oAprobaciones(j).oSBS(i).Calif_0 = dr.Item("Calif_0")
                        R.oAprobaciones(j).oSBS(i).Calif_1 = dr.Item("Calif_1")
                        R.oAprobaciones(j).oSBS(i).Calif_2 = dr.Item("Calif_2")
                        R.oAprobaciones(j).oSBS(i).Calif_3 = dr.Item("Calif_3")
                        R.oAprobaciones(j).oSBS(i).Calif_4 = dr.Item("Calif_4")
                        R.oAprobaciones(j).oSBS(i).Can_Ents = dr.Item("Can_Ents")

                        i = i + 1

                    End While
                    dr.Close()
                    dr = Nothing

                Next



                For j = 0 To R.oAprobaciones.Count - 1

                    oComm = New SqlClient.SqlCommand("PA_DetalleSBS", oConex)
                    oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = R.oAprobaciones(j).cDoc
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()

                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0

                    While dr.Read
                        ReDim Preserve R.oAprobaciones(j).oSBSDetalle(i)
                        R.oAprobaciones(j).oSBSDetalle(i) = New clsSBSHistDet
                        R.oAprobaciones(j).oSBSDetalle(i).nombre = dr.Item("nombre")
                        R.oAprobaciones(j).oSBSDetalle(i).nDeuda = dr.Item("nDeuda")
                        R.oAprobaciones(j).oSBSDetalle(i).clasificacion = dr.Item("clasificacion")

                        i = i + 1

                    End While
                    dr.Close()
                    dr = Nothing

                Next


                For j = 0 To R.oAprobaciones.Count - 1

                    oComm = New SqlClient.SqlCommand("PA_ImagenesNegocioAprobacion", oConex)
                    oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = R.oAprobaciones(j).ncodcred
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()

                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0

                    While dr.Read
                        ReDim Preserve R.oAprobaciones(j).oImagenes(i)
                        R.oAprobaciones(j).oImagenes(i) = New clsVerificacionClienteImagenes
                        R.oAprobaciones(j).oImagenes(i).nCodCred = R.oAprobaciones(j).ncodcred
                        R.oAprobaciones(j).oImagenes(i).iImagen = UnicodeBytesToString(dr.Item("nFoto"))
                        R.oAprobaciones(j).oImagenes(i).cTipo = dr.Item("cTipo")
                        R.oAprobaciones(j).oImagenes(i).nID = i

                        i = i + 1

                    End While
                    dr.Close()
                    dr = Nothing


                Next


            End If


            'Carga Gestores de Cobranza de Administrador
            oComm = New SqlClient.SqlCommand("PA_RecuperaGestoresDeCobranzaDeAdm", oConex)
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu

            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            R.oUsuariosEquipo = Nothing

            While dr.Read

                ReDim Preserve R.oUsuariosEquipo(i)
                R.oUsuariosEquipo(i) = New clsUsuariosEquipo

                R.oUsuariosEquipo(i).cCodUsu = dr.Item("cCodUsu")
                R.oUsuariosEquipo(i).cRol = dr.Item("cRol")

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

    Public Function RecuperaGruposAprobacion() As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlCommand("PA_RecuperaGrupos_Aprobacion", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read
                ReDim Preserve R.oCatalogocodigos_2(i)
                R.oCatalogocodigos_2(i) = New clsCatalogocodigos()
                R.oCatalogocodigos_2(i).cValor = dr.Item("cValor")
                R.oCatalogocodigos_2(i).cDescripcion = dr.Item("cDescripcion")
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

    Public Function ValidaRegistroVoucher() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Try


            oComm = New SqlCommand("[dbo].[PA_ValidateRegistrarDeposito]", oConex)
            oComm.Parameters.Add("@psCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@cNumOperacion", SqlDbType.VarChar, 50).Value = cCodCat
            oComm.Parameters.Add("@psResult", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@pnResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            D(0) = Str(oComm.Parameters("@pnResp").Value).Trim
            D(1) = oComm.Parameters("@psResult").Value

            oConex.Close()

        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message

        End Try


        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))

        Return R

    End Function

    Public Function RecuperaGruposDesembolso() As outNegocio
        Dim R As New outNegocio
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

            oComm = New SqlCommand("PA_RecuperaGruposDesembolso", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read
                ReDim Preserve R.oCatalogocodigos_2(i)
                R.oCatalogocodigos_2(i) = New clsCatalogocodigos()
                R.oCatalogocodigos_2(i).cValor = dr.Item("cValor")
                R.oCatalogocodigos_2(i).cDescripcion = dr.Item("cDescripcion")
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

    Public Function RecuperaCreditosXAsignarAsesorGrupal() As outNegocio
        Dim R As New outNegocio
        Dim ListaAsesor() As String
        Dim ListaDesembolsador() As String
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

            oComm = New SqlCommand("PA_RecuperaDesembolsadores", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read
                ReDim Preserve R.ListaDesembolsadoresAbonar(i)
                R.ListaDesembolsadoresAbonar(i) = UCase(dr.Item("cValor"))
                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            oComm = New SqlCommand("PA_RecuperaAsesoresAsigna", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read
                ReDim Preserve R.oUsuarios(i)
                R.oUsuarios(i) = UCase(dr.Item("cValor"))
                i = i + 1
            End While
            dr.Close()
            dr = Nothing


            oComm = New SqlCommand("PA_RecuperaCreditosXAsignarAsesorGrupal", oConex)
            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
            oComm.Parameters.Add("@nIDGrupo", SqlDbType.Int).Value = nIdGrupo
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oAsigAbonAsesorGrupal(i)
                R.oAsigAbonAsesorGrupal(i) = New clsAsigAbonAsesorGrupal
                R.oAsigAbonAsesorGrupal(i).cDoc = dr.Item("cdoc")
                R.oAsigAbonAsesorGrupal(i).cCliente = dr.Item("cCliente")
                R.oAsigAbonAsesorGrupal(i).nCodAge = dr.Item("nCodAge")
                R.oAsigAbonAsesorGrupal(i).nCodCred = dr.Item("nCodCred")
                R.oAsigAbonAsesorGrupal(i).nPrestamo = dr.Item("nPrestamo")
                R.oAsigAbonAsesorGrupal(i).nDeudaVigente = dr.Item("nDeudaVigente")
                R.oAsigAbonAsesorGrupal(i).nPrestamoNeto = dr.Item("nPrestamoNeto")
                R.oAsigAbonAsesorGrupal(i).cDistritoNeg = dr.Item("cDistritoNeg")
                R.oAsigAbonAsesorGrupal(i).cZonaNeg = dr.Item("cZonaNeg")
                R.oAsigAbonAsesorGrupal(i).cDistritoDom = dr.Item("cDistritoDom")
                R.oAsigAbonAsesorGrupal(i).cZonaDom = dr.Item("cZonaDom")
                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            oConex.Close()
            D(0) = "1"
            D(1) = "Consulta Correcta"
        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message
        End Try
        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))
        Return R

    End Function

    Public Function RecuperaCreditosXAbonarADesembGrupal() As outNegocio
        Dim R As New outNegocio
        Dim ListaAsesor() As String
        Dim ListaDesembolsador() As String
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

            oComm = New SqlCommand("PA_RecuperaCreditosXAbonarADesembGrupal", oConex)
            oComm.Parameters.Add("@psCodDesemb", SqlDbType.VarChar).Value = cDocPersona
            oComm.Parameters.Add("@nIdGrupo", SqlDbType.Int).Value = nIdGrupo
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oAsigAbonAsesorGrupal(i)
                R.oAsigAbonAsesorGrupal(i) = New clsAsigAbonAsesorGrupal
                R.oAsigAbonAsesorGrupal(i).cDoc = dr.Item("cdoc")
                R.oAsigAbonAsesorGrupal(i).cCliente = dr.Item("cCliente")
                R.oAsigAbonAsesorGrupal(i).nCodAge = dr.Item("nCodAge")
                R.oAsigAbonAsesorGrupal(i).nCodCred = dr.Item("nCodCred")
                R.oAsigAbonAsesorGrupal(i).nPrestamo = dr.Item("nPrestamo")
                R.oAsigAbonAsesorGrupal(i).nDeudaVigente = dr.Item("nDeudaVigente")
                R.oAsigAbonAsesorGrupal(i).nPrestamoNeto = dr.Item("nPrestamoNeto")
                R.oAsigAbonAsesorGrupal(i).cDistritoNeg = ""
                R.oAsigAbonAsesorGrupal(i).cZonaNeg = ""
                R.oAsigAbonAsesorGrupal(i).cDistritoDom = ""
                R.oAsigAbonAsesorGrupal(i).cZonaDom = ""
                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            oConex.Close()
            D(0) = "1"
            D(1) = "Consulta Correcta"
        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message
        End Try
        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))
        Return R

    End Function

    Public Function RecuperaConsultaSolicitudesxCred() As outNegocio

        Dim R As New outNegocio
        Dim i, k As Integer
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlCommand
        Dim dr As SqlDataReader

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Try
            i = 0
            For k = 0 To oBandejaDesembolsoNuevo.Count - 1
                oComm = New SqlCommand("PA_RecuperaBandejaConsultaSolicitudesxCred", oConex)
                oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = oBandejaDesembolsoNuevo(k).nCodCred
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.ExecuteNonQuery()
                dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

                While dr.Read
                    ReDim Preserve R.oConsultaSolicitudes(i)
                    R.oConsultaSolicitudes(i) = New clsConsultaSolicitudes
                    R.oConsultaSolicitudes(i).cCliente = dr.Item("cCliente")
                    R.oConsultaSolicitudes(i).nCodCred = dr.Item("nCodCred")
                    R.oConsultaSolicitudes(i).cDoc = dr.Item("cdoc")
                    R.oConsultaSolicitudes(i).cUsuOrig = dr.Item("cUsuOrig")
                    R.oConsultaSolicitudes(i).cProd = dr.Item("cProd")
                    R.oConsultaSolicitudes(i).cEstado = dr.Item("cEstado")
                    R.oConsultaSolicitudes(i).cComentario = dr.Item("cComentario")
                    R.oConsultaSolicitudes(i).nPrestamo = dr.Item("nMontoAprobado")
                    R.oConsultaSolicitudes(i).nPrestamoOrig = dr.Item("nMontoOriginal")
                    R.oConsultaSolicitudes(i).nCuotas = dr.Item("nCuotaAprobado")
                    R.oConsultaSolicitudes(i).nCuotasOrig = dr.Item("nCuotaOriginal")
                    R.oConsultaSolicitudes(i).cComentario = dr.Item("cComentario")
                    i = i + 1
                End While
                dr.Close()
                dr = Nothing
            Next


            If i > 0 Then
                i = 0
                For j = 0 To R.oConsultaSolicitudes.Count - 1
                    oComm = New SqlCommand("PA_RecuperaVerificacionxCredito", oConex)
                    oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = R.oConsultaSolicitudes(j).nCodCred
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()

                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

                    While dr.Read
                        ReDim Preserve R.oBandejaVerificacion(i)
                        R.oBandejaVerificacion(i) = New clsVerificacionCliente
                        R.oBandejaVerificacion(i).cDoc = R.oConsultaSolicitudes(j).cDoc
                        R.oBandejaVerificacion(i).nCodCred = R.oConsultaSolicitudes(j).nCodCred
                        R.oBandejaVerificacion(i).nVerifCuali = dr.Item("nVerifCuali")
                        R.oBandejaVerificacion(i).nVerifRefe = dr.Item("nVerifRefe")
                        R.oBandejaVerificacion(i).nCumpleConAntNeg = dr.Item("nCumpleConAntNeg")
                        R.oBandejaVerificacion(i).nEsPropietario = dr.Item("nEsPropietario")
                        R.oBandejaVerificacion(i).nPoseeEmpleados = dr.Item("nPoseeEmpleados")
                        R.oBandejaVerificacion(i).nNroEmpleados = dr.Item("nNroEmpleados")
                        R.oBandejaVerificacion(i).nTieneRotuloVisible = dr.Item("nTieneRotuloVisible")
                        R.oBandejaVerificacion(i).nNegoCuentaInfra = dr.Item("nNegoCuentaInfra")
                        R.oBandejaVerificacion(i).nCuentaPermiOpe = dr.Item("nCuentaPermiOpe")
                        R.oBandejaVerificacion(i).nObservaAfluCli = dr.Item("nObservaAfluCli")
                        R.oBandejaVerificacion(i).nUbicaFacilAcceso = dr.Item("nUbicaFacilAcceso")
                        R.oBandejaVerificacion(i).nTieneBuenaRefeZona = dr.Item("nTieneBuenaRefeZona")
                        R.oBandejaVerificacion(i).nTieneProspBuenaActi = dr.Item("nTieneProspBuenaActi")
                        R.oBandejaVerificacion(i).nTieneHoraAtencion = dr.Item("nTieneHoraAtencion")
                        R.oBandejaVerificacion(i).nAquilaCasa = dr.Item("nAquilaCasa")
                        R.oBandejaVerificacion(i).nTiempoAquila = dr.Item("nTiempoAquila")
                        R.oBandejaVerificacion(i).nTieneDependientes = dr.Item("nTieneDependientes")
                        R.oBandejaVerificacion(i).nNroDependientes = dr.Item("nNroDependientes")
                        R.oBandejaVerificacion(i).cHoraInicio = dr.Item("cHoraInicio")
                        R.oBandejaVerificacion(i).cHoraFinal = dr.Item("cHoraFinal")
                        R.oBandejaVerificacion(i).nNivelDelincuencia = dr.Item("nNivelDelincuencia")
                        R.oBandejaVerificacion(i).cComentarioCPD = dr.Item("cComentarioCPD")
                        i = i + 1
                    End While
                    dr.Close()
                    dr = Nothing
                Next
            End If

            If i > 0 Then
                For j = 0 To R.oBandejaVerificacion.Count - 1
                    oComm = New SqlCommand("PA_ReferenciasVericacion", oConex)
                    oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = R.oBandejaVerificacion(j).nCodCred
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()

                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0
                    While dr.Read
                        ReDim Preserve R.oBandejaVerificacion(j).oReferencias(i)
                        R.oBandejaVerificacion(j).oReferencias(i) = New clsVerificacionReferencia
                        R.oBandejaVerificacion(j).oReferencias(i).nCodCred = R.oBandejaVerificacion(j).nCodCred
                        R.oBandejaVerificacion(j).oReferencias(i).cNombres = dr.Item("cNombres")
                        R.oBandejaVerificacion(j).oReferencias(i).cTelefono = dr.Item("cTelefono")
                        R.oBandejaVerificacion(j).oReferencias(i).nRecomendado = dr.Item("nRecomienda")
                        R.oBandejaVerificacion(j).oReferencias(i).cRecomendado = dr.Item("cRecomiendaSustento")
                        i = i + 1
                    End While
                    dr.Close()
                    dr = Nothing

                Next
                For j = 0 To R.oBandejaVerificacion.Count - 1
                    oComm = New SqlCommand("PA_ImagenesVerificacion", oConex)
                    oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = R.oBandejaVerificacion(j).nCodCred
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()
                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0
                    While dr.Read
                        ReDim Preserve R.oBandejaVerificacion(j).oImagenes(i)
                        R.oBandejaVerificacion(j).oImagenes(i) = New clsVerificacionClienteImagenes
                        R.oBandejaVerificacion(j).oImagenes(i).nCodCred = R.oBandejaVerificacion(j).nCodCred
                        R.oBandejaVerificacion(j).oImagenes(i).iImagen = UnicodeBytesToString(dr.Item("nFoto"))
                        R.oBandejaVerificacion(j).oImagenes(i).cTipo = dr.Item("cTipo")
                        R.oBandejaVerificacion(j).oImagenes(i).nID = i
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

    Public Function ActualizarVersionAppUsuario() As outNegocio

        Dim R As New outNegocio
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlCommand

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Try

            oComm = New SqlCommand("PA_ActualizaVersionAppUsuario", oConex)
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@nVersionApp", SqlDbType.Money).Value = nTasa
            oComm.Parameters.Add("@nResultado", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@sRespuesta", SqlDbType.VarChar, 150).Direction = ParameterDirection.Output

            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            R.nResultado = oComm.Parameters("@nResultado").Value
            R.sRespuesta = oComm.Parameters("@sRespuesta").Value

            oConex.Close()

            'D(0) = "1"
            'D(1) = "Gestión Correcta"
            D(0) = Str(R.nResultado).Trim
            D(1) = R.sRespuesta

        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message
        End Try

        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))

        Return R

    End Function

    ' CONSULTA EXPERIAN EXTERNO
    Public Function getVarnegocio(ByVal nCodVal As Integer) As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlCommand
        Dim nResp As Integer
        Dim sRespuesta As String
        Dim dr As SqlDataReader
        Dim i, j As Integer

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Try

            oComm = New SqlCommand("PA_DevuelveVarNegocio", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@pnCodVar", SqlDbType.Int).Value = nCodVal
            oComm.Parameters.Add("@psValor", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            sRespuesta = oComm.Parameters("@psValor").Value

            nResp = 1

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

    Public Function ValidaDocumentoExperian() As outNegocio
        Dim R As New outNegocio
        Dim AbonoDesemb() As dAbonoDesemb
        Dim ListaDesembolsador() As String
        Dim i As Integer
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlCommand
        Dim dr As SqlDataReader

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Try
            oComm = New SqlClient.SqlCommand("PA_ValidaDocumentoExperianLocal", oConex)
            oComm.Parameters.Add("@TipoDoc", SqlDbType.VarChar, 50).Value = "D"
            oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = cDocPersona

            oComm.Parameters.Add("@psResult", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@pnResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()
            D(0) = Str(oComm.Parameters("@pnResp").Value).Trim
            D(1) = oComm.Parameters("@psResult").Value
            oConex.Close()
        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message
        End Try
        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))
        Return R

    End Function

    Public Function ConsultaDocumentoExperian() As outNegocio
        Dim R As New outNegocio
        Dim R2 As New ResponseExperianBasic
        Dim D(2) As String
        Dim Client_id = "V63yUUOrxLtkGVFwNfq1ylXCstbGbnP3"
        Dim Client_secret = "RC2LmnVMelpuiHlG"
        Dim ruta = "https://api.experian.com.pe"
        Dim Gx_Key = "0D910906E0E5AC631B82977E1C54F087"
        Dim Gx_usuario = "20602421806"

        Try

            Dim validaEx = ValidaDocumentoExperian()
            If validaEx.nResultado = 1 Then

                'OBTENGO LA INFO
                cDoc = cDocPersona
                Dim infoEx = RecuperarExperian()
                If infoEx.nResultado = 1 Then
                    R2.sDocumento = cDocPersona
                    R2.sEntidades = infoEx.oCatalogocodigos_2(0).cDescripcion
                    R2.sCalificacion = infoEx.oCatalogocodigos_2(1).cDescripcion
                    R2.nScore = CInt(infoEx.oCatalogocodigos_2(8).cValor)

                    R2.nResultado = 1
                    R2.sRespuesta = "Resultado Correcto."
                    R.nResultado = 1
                    R.sRespuesta = "Resultado Correcto."

                End If

            Else

                'OBTENGO TOKEN EXPERIAN
                Dim request = New TokenRequest()
                request.username = getVarnegocio(18).sRespuesta
                request.password = getVarnegocio(19).sRespuesta

                Dim json_envio = JsonConvert.SerializeObject(request, Formatting.Indented)
                Console.WriteLine(json_envio)
                Dim json_sincodificar = Encoding.UTF8.GetBytes(json_envio)
                Dim json_utf8 = Encoding.UTF8.GetString(json_sincodificar)
                Dim json_de_respuesta = sendJson(ruta, json_utf8, Client_id, Client_secret, "iam/oauth2/v1/basic/token")
                Dim leer_respuesta = JsonConvert.DeserializeObject(Of ResponseToken)(json_de_respuesta)

                If Not IsNothing(leer_respuesta.access_token) Then

                    'OBTENGO EXPERIAN
                    Dim request2 = New DocumentoClienteRequest()
                    request2.TipDoc = "D"
                    request2.NroDoc = cDocPersona

                    Dim json_envio2 = JsonConvert.SerializeObject(request2, Formatting.Indented)
                    Console.WriteLine(json_envio2)
                    Dim json_sincodificar2 = Encoding.UTF8.GetBytes(json_envio2)
                    Dim json_utf82 = Encoding.UTF8.GetString(json_sincodificar2)

                    Dim json_de_respuesta2 = sendJson2(ruta, json_utf82, leer_respuesta.access_token,
                              Client_id, Client_secret, "modular/credit-history/v1/hdc/gethistory",
                              Gx_Key, Gx_usuario, request.username)
                    Dim leer_respuesta2 = JsonConvert.DeserializeObject(Of ResponseHistorial)(json_de_respuesta2)


                    If Not IsNothing(leer_respuesta2) Then
                        'REGISTRO EXPERIAN
                        oDataExperian = leer_respuesta2
                        Dim registroEx = RegistrarExperian()

                        If registroEx.nResultado = 1 Then
                            'OBTENGO LA INFO
                            cDoc = cDocPersona
                            Dim infoEx = RecuperarExperian()
                            If infoEx.nResultado = 1 Then
                                R2.sDocumento = cDocPersona
                                R2.sEntidades = infoEx.oCatalogocodigos_2(0).cDescripcion
                                R2.sCalificacion = infoEx.oCatalogocodigos_2(1).cDescripcion
                                R2.nScore = CInt(infoEx.oCatalogocodigos_2(8).cValor)
                                R2.nResultado = 1
                                R2.sRespuesta = "Resultado Correcto Y Registrado."
                                R.nResultado = 1
                                R.sRespuesta = "Resultado Correcto Y Registrado."
                            End If

                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message
        End Try
        R.experianBasic = R2
        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)
        Return R
    End Function

    Function sendJson2(ruta As String, json As String, token As String, Client_id As String, Client_secret As String, metodo As String, Gx_Key As String, Gx_usuario As String, Gx_email As String) As String
        Using wc As New WebClient
            Try
                wc.Headers.Add("Authorization", "Bearer " & token)
                wc.Headers.Add("Content-Type", "application/json")
                wc.Headers.Add("Client_id", Client_id)
                wc.Headers.Add("Client_secret", Client_secret)
                wc.Headers.Add("Gx_email", Gx_email)
                wc.Headers.Add("Gx_usuario", Gx_usuario)
                wc.Headers.Add("Gx_Key", Gx_Key)
                Dim respuesta = wc.UploadString(ruta + "/" + metodo, "POST", json)
                Return respuesta
            Catch x As WebException
                Return New StreamReader(x.Response.GetResponseStream).ReadToEnd
            End Try
        End Using
    End Function

    Function sendJson(ruta As String, json As String, Client_id As String, Client_secret As String, metodo As String) As String
        Using wc As New WebClient
            Try
                wc.Headers.Add("Content-Type", "application/json")
                wc.Headers.Add("Client_id", Client_id)
                wc.Headers.Add("Client_secret", Client_secret)
                Dim respuesta = wc.UploadString(ruta + "/" + metodo, "POST", json)
                Return respuesta
            Catch x As WebException
                Return New StreamReader(x.Response.GetResponseStream).ReadToEnd
            End Try
        End Using
    End Function
    ' FIN CONSULTA EXPERIAN EXTERNO

    Public Function RegistrarHuellaDigital() As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlCommand
        Dim nResp As Integer
        Dim sRespuesta As String
        Dim dr As SqlDataReader

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
            oComm = New SqlCommand("[dbo].[PA_RegistrarHuellaDigital]", oConex)
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@cDoc", SqlDbType.VarChar).Value = cDoc
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
            oComm.Parameters.Add("@iDoc", SqlDbType.Image).Value = nImg
            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            nResp = oComm.Parameters("@nResp").Value
            sRespuesta = oComm.Parameters("@Resultado").Value

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

    'CONSULTA EQUIFAX 
    Public Function ConsultaDocumentoGOLDENCASH() As outNegocio
        Dim R As New outNegocio
        Dim R2 As New ResponseEquifaxBasic
        Dim D(2) As String
        Dim Client_id = "F6xo6EVSrteWAsdNa59e7Mk9cLXdyzpq"
        Dim Client_secret = "jrg3TsXqYaKh5czh"
        Dim ruta = "https://api.latam.equifax.com/"
        Dim Gx_Key = "0D910906E0E5AC631B82977E1C54F087"
        Dim Gx_usuario = "20602421806"
        Dim scope = "https://api.latam.equifax.com/datos-comerciales/transaction"
        Dim authorization = "RjZ4bzZFVlNydGVXQXNkTmE1OWU3TWs5Y0xYZHl6cHE6anJnM1RzWHFZYUtoNWN6aA=="

        Try

            Dim validaEx = ValidaDocumentoEquifax()
            If validaEx.nResultado = 1 Then
                'OBTENGO LA INFO
                cDoc = cDocPersona
                Dim infoEx = RecuperarEquifax()
                If infoEx.nResultado = 1 Then
                    R2.sDocumento = cDocPersona
                    '*
                    R2.ResumenFinanciero = New ResumenFinanciero
                    R2.ResumenFinanciero = infoEx.equifaxInformacion.ResumenFinanciero
                    R2.SistemaFinanciero = New SistemaFinancieroSYNRegulado
                    R2.SistemaFinanciero = infoEx.equifaxInformacion.SistemaFinanciero
                    R2.ResumenFlags = New ResumenFlags
                    R2.ResumenFlags = infoEx.equifaxInformacion.ResumenFlags
                    R2.ResumenScoreRP3 = New ResumenScoreRP3
                    R2.ResumenScoreRP3 = infoEx.equifaxInformacion.ResumenScoreRP3
                    '*
                    R2.nResultado = 1
                    R2.sRespuesta = "Resultado Correcto."
                    R.nResultado = 1
                    R.sRespuesta = "Resultado Correcto."
                End If

            Else

                '
                '

                'OBTENGO TOKEN EQUIFAX
                Dim request = New TokenEquifaxRequest()
                request.grant_type = "client_credentials" 'getVarnegocio(18).sRespuesta
                request.scope = scope 'getVarnegocio(19).sRespuesta

                Dim json_envio = JsonConvert.SerializeObject(request, Formatting.Indented)
                Console.WriteLine(json_envio)
                Dim json_sincodificar = Encoding.UTF8.GetBytes(json_envio)
                Dim json_utf8 = Encoding.UTF8.GetString(json_sincodificar)
                Dim json_de_respuesta = sendJsonGolden(ruta, json_utf8, authorization, "v2/oauth/token")
                Dim leer_respuesta = JsonConvert.DeserializeObject(Of ResponseToken)(json_de_respuesta)

                If Not IsNothing(leer_respuesta.access_token) Then

                    'OBTENGO EQUIFAX
                    Dim request2 = New RequestClienteEquifax() 'DocumentoClienteRequest()
                    request2.applicants = New Applicants()
                    request2.productData = New ProductData()
                    request2.applicants.primaryConsumer = New PrimaryConsumer()
                    request2.applicants.primaryConsumer.personalInformation = New PersonalInformation()

                    request2.applicants.primaryConsumer.personalInformation.id = cDocPersona
                    request2.applicants.primaryConsumer.personalInformation.tipoPersona = 1
                    request2.applicants.primaryConsumer.personalInformation.tipoDocumento = 1

                    request2.productData.billTo = "146300B001"
                    request2.productData.shipTo = "146300B001S0001"
                    request2.productData.productName = "PEREPORT"
                    request2.productData.productOrch = "CREDITREPORTV2"
                    request2.productData.configuration = "Config"
                    request2.productData.customer = "PEREPGOLDENCASH"
                    request2.productData.model = "CreditReportGOLDENCASH"

                    Dim json_envio2 = JsonConvert.SerializeObject(request2, Formatting.Indented)
                    Console.WriteLine(json_envio2)
                    Dim json_sincodificar2 = Encoding.UTF8.GetBytes(json_envio2)
                    Dim json_utf82 = Encoding.UTF8.GetString(json_sincodificar2)

                    Dim json_de_respuesta2 = sendJsonGolden2(ruta, json_utf82, leer_respuesta.access_token, "datos-comerciales/transaction/execute")
                    Dim leer_respuesta2 = JsonConvert.DeserializeObject(Of ResponseHistorialEquifax)(json_de_respuesta2)


                    If Not IsNothing(leer_respuesta2) Then
                        'REGISTRO EQUIFAX
                        oDataEquifax = leer_respuesta2
                        Dim registroEx = RegistrarEquifax(json_de_respuesta2)                        '

                        If registroEx.nResultado = 1 Then
                            'OBTENGO LA INFO
                            cDoc = cDocPersona
                            Dim infoEx = RecuperarEquifax()
                            If infoEx.nResultado = 1 Then
                                R2.sDocumento = cDocPersona
                                '*
                                R2.ResumenFinanciero = New ResumenFinanciero
                                R2.ResumenFinanciero = infoEx.equifaxInformacion.ResumenFinanciero
                                R2.SistemaFinanciero = New SistemaFinancieroSYNRegulado
                                R2.SistemaFinanciero = infoEx.equifaxInformacion.SistemaFinanciero
                                R2.ResumenFlags = New ResumenFlags
                                R2.ResumenFlags = infoEx.equifaxInformacion.ResumenFlags
                                R2.ResumenScoreRP3 = New ResumenScoreRP3
                                R2.ResumenScoreRP3 = infoEx.equifaxInformacion.ResumenScoreRP3
                                '*
                                R2.nResultado = 1
                                R2.sRespuesta = "Resultado Correcto Y Registrado."
                                R.nResultado = 1
                                R.sRespuesta = "Resultado Correcto Y Registrado."
                            End If

                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message
        End Try
        R.equifaxInformacion = R2
        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)
        Return R
    End Function

    Function sendJsonGolden(ruta As String, json As String, authorization As String, metodo As String) As String
        'Using wc As New WebClient
        Using client As New WebClient
            Try

                'wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded")
                'wc.Headers.Add("Authorization", "Basic " & authorization)
                'Dim respuesta = wc.UploadString(ruta + "/" + metodo, "POST", "")





                ' 1. Set Authorization Header for Basic Auth
                'Dim authHeader As String = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{"application/x-www-form-urlencoded"}:{authorization}"))
                client.Headers.Add("Content-Type", "application/x-www-form-urlencoded")
                client.Headers.Add("Authorization", "Basic " & authorization)

                ' 2. The UploadValues method sets the Content-Type header automatically to x-www-form-urlencoded.

                ' 3. Prepare the request body parameters
                Dim postData As New NameValueCollection()
                postData.Add("grant_type", "client_credentials")
                postData.Add("scope", "https://api.latam.equifax.com/datos-comerciales/transaction")

                ' 4. Send the POST request with the form data
                Dim responseBytes As Byte() = client.UploadValues(ruta + "/" + metodo, "POST", postData)

                ' Convert response bytes to string for processing
                Dim responseString As String = Encoding.UTF8.GetString(responseBytes)


                Return responseString


            Catch x As WebException
                Return New StreamReader(x.Response.GetResponseStream).ReadToEnd
            End Try
        End Using
    End Function

    Function sendJsonGolden2(ruta As String, json As String, token As String, metodo As String) As String
        Using wc As New WebClient
            wc.Encoding = Encoding.UTF8
            Try
                wc.Headers.Add("Authorization", "Bearer " & token)
                wc.Headers.Add("Content-Type", "application/json")
                Dim respuesta = wc.UploadString(ruta + "/" + metodo, "POST", json)
                Return respuesta
            Catch x As WebException
                Return New StreamReader(x.Response.GetResponseStream).ReadToEnd
            End Try
        End Using
    End Function

    Public Function RegistrarEquifax(jsonConsulta As String) As outNegocio
        Dim R As New outNegocio
        Dim D(2) As String
        Dim i, k, j, l, p, q As Integer
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nResp, nCodTasacion As Integer
        Dim sRespuesta As String
        Dim dr As SqlClient.SqlDataReader
        Dim TDoc, NDoc As String
        Dim nEliminar As Integer = 0
        nResp = 0
        nCodTasacion = 0
        sRespuesta = ""
        TDoc = ""
        NDoc = ""
        Dim id As Integer
        Dim id_Avalados As Integer
        Dim id_inicial As Integer
        Dim id_ResumenFlags As Integer

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Try

            If oDataEquifax.applicants.primaryConsumer.personalInformation.id.Length > 0 Then
                NDoc = oDataEquifax.applicants.primaryConsumer.personalInformation.id
            End If
            If oDataEquifax.applicants.primaryConsumer.personalInformation.tipoDocumento.Length > 0 Then
                TDoc = oDataEquifax.applicants.primaryConsumer.personalInformation.tipoDocumento
            End If
            ' Informacion Personal
            oComm = New SqlCommand("[dbo].[PA_RegistrarInformacionPersonal]", oConex)
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@tipoDocumento", SqlDbType.VarChar).Value = TDoc
            oComm.Parameters.Add("@id", SqlDbType.VarChar).Value = NDoc
            oComm.Parameters.Add("@tipoPersona", SqlDbType.VarChar).Value = oDataEquifax.applicants.primaryConsumer.personalInformation.tipoPersona.NullToEmpty()
            oComm.Parameters.Add("@jsonConsulta", SqlDbType.VarChar).Value = jsonConsulta
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu

            oComm.Parameters.Add("@id_inicial", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
            oComm.ExecuteNonQuery()
            nResp = oComm.Parameters("@nResp").Value
            sRespuesta = oComm.Parameters("@Resultado").Value
            id_inicial = oComm.Parameters("@id_inicial").Value

            If nResp = 0 Then
                D(0) = Str(nResp)
                D(1) = sRespuesta
                R.nResultado = CInt(D(0))
                R.sRespuesta = D(1)
                Return R
            End If
            ''**** Informacion Personal

            If oDataEquifax.applicants.primaryConsumer.interconnectResponse IsNot Nothing Then

                If oDataEquifax.applicants.primaryConsumer.interconnectResponse.Count > 0 Then

                    '/RESUMEN CONSULTA
                    If oDataEquifax.applicants.primaryConsumer.interconnectResponse(0).Codigo = 100 Then
                        If oDataEquifax.applicants.primaryConsumer.interconnectResponse(0).TieneError Then
                            'ERROR DE CONSULTA
                            R.nResultado = oDataEquifax.applicants.primaryConsumer.interconnectResponse(0).DetallesError.CodigoError
                            R.sRespuesta = oDataEquifax.applicants.primaryConsumer.interconnectResponse(0).DetallesError.MensajeError.NullToEmpty()
                            Return R
                        End If
                    End If
                    'END RESUMEN CONSULTA

                    For i = 0 To oDataEquifax.applicants.primaryConsumer.interconnectResponse.Count - 1

                        '// AVALADOS 870
                        If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Codigo = 870 Then

                            oComm = New SqlCommand("[dbo].[PA_RegistrarAvalados]", oConex)
                            oComm.CommandType = CommandType.StoredProcedure
                            oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial

                            oComm.Parameters.Add("@CodigoModulo", SqlDbType.VarChar).Value = "870" 'oDataExperian.respuesta.Sabio.Resul
                            oComm.Parameters.Add("@NombreModulo", SqlDbType.VarChar).Value = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Nombre.NullToEmpty()
                            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu

                            oComm.Parameters.Add("@id_Avalados", SqlDbType.Int).Direction = ParameterDirection.Output

                            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                            oComm.ExecuteNonQuery()
                            nResp = oComm.Parameters("@nResp").Value
                            sRespuesta = oComm.Parameters("@Resultado").Value
                            id_Avalados = oComm.Parameters("@id_Avalados").Value

                            If nResp = 0 Then
                                D(0) = Str(nResp)
                                D(1) = sRespuesta
                                R.nResultado = CInt(D(0))
                                R.sRespuesta = D(1)
                                Return R
                            End If

                            '--
                            j = 0
                            If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.Avalados.Avalado IsNot Nothing Then
                                For j = 0 To oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.Avalados.Avalado.Count - 1

                                    Dim avalado As New Avalado
                                    avalado = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.Avalados.Avalado(j)

                                    oComm = New SqlCommand("[dbo].[PA_RegistrarAvalados_Avalo]", oConex)
                                    oComm.CommandType = CommandType.StoredProcedure
                                    oComm.Parameters.Add("@id_Avalados", SqlDbType.Int).Value = id_Avalados
                                    oComm.Parameters.Add("@TipoDocumento", SqlDbType.VarChar).Value = avalado.TipoDocumento.NullToEmpty()
                                    oComm.Parameters.Add("@NumeroDocumento", SqlDbType.VarChar).Value = avalado.NumeroDocumento.NullToEmpty()
                                    oComm.Parameters.Add("@Nombre", SqlDbType.VarChar).Value = avalado.Nombre.NullToEmpty()
                                    oComm.Parameters.Add("@UltimaClasificacion", SqlDbType.VarChar).Value = avalado.UltimaClasificacion.NullToEmpty()
                                    oComm.Parameters.Add("@UltimoPeriodo", SqlDbType.VarChar).Value = avalado.UltimoPeriodo.NullToEmpty()

                                    oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                    oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                    oComm.ExecuteNonQuery()
                                    nResp = oComm.Parameters("@nResp").Value
                                    sRespuesta = oComm.Parameters("@Resultado").Value


                                    If nResp = 0 Then
                                        D(0) = Str(nResp)
                                        D(1) = sRespuesta
                                        R.nResultado = CInt(D(0))
                                        R.sRespuesta = D(1)
                                        Return R
                                    End If

                                Next
                            End If

                        End If
                        '**** END AVALADOS

                        'RESUMEN FLAG - 865
                        If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Codigo = 865 Then
                            oComm = New SqlCommand("[dbo].[PA_RegistrarResumenFlags]", oConex)
                            oComm.CommandType = CommandType.StoredProcedure
                            oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial

                            oComm.Parameters.Add("@CodigoModulo", SqlDbType.VarChar).Value = "865" 'oDataExperian.respuesta.Sabio.Resul
                            oComm.Parameters.Add("@NombreModulo", SqlDbType.VarChar).Value = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Nombre.NullToEmpty()
                            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu

                            oComm.Parameters.Add("@id_ResumenFlags", SqlDbType.Int).Direction = ParameterDirection.Output

                            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                            oComm.ExecuteNonQuery()
                            nResp = oComm.Parameters("@nResp").Value
                            sRespuesta = oComm.Parameters("@Resultado").Value
                            id_ResumenFlags = oComm.Parameters("@id_ResumenFlags").Value

                            If nResp = 0 Then
                                D(0) = Str(nResp)
                                D(1) = sRespuesta
                                R.nResultado = CInt(D(0))
                                R.sRespuesta = D(1)
                                Return R
                            End If

                            '** [ResumenComportamiento]

                            Dim id_ResumenComportamiento As Integer = 0
                            If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenFlags.ResumenComportamiento IsNot Nothing Then

                                Dim resumenComportamiento As New ResumenComportamiento
                                resumenComportamiento = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenFlags.ResumenComportamiento

                                oComm = New SqlCommand("[dbo].[PA_RegistrarResumenComportamiento]", oConex)
                                oComm.CommandType = CommandType.StoredProcedure
                                oComm.Parameters.Add("@id_ResumenFlags", SqlDbType.Int).Value = id_ResumenFlags
                                oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial

                                oComm.Parameters.Add("@TipoDocumento", SqlDbType.VarChar).Value = resumenComportamiento.TipoDocumento.NullToEmpty()
                                oComm.Parameters.Add("@NumeroDocumento", SqlDbType.VarChar).Value = resumenComportamiento.NumeroDocumento.NullToEmpty()
                                oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                oComm.Parameters.Add("@id_ResumenComportamiento", SqlDbType.Int).Direction = ParameterDirection.Output

                                oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                oComm.ExecuteNonQuery()
                                nResp = oComm.Parameters("@nResp").Value
                                sRespuesta = oComm.Parameters("@Resultado").Value
                                id_ResumenComportamiento = oComm.Parameters("@id_ResumenComportamiento").Value

                                If nResp = 0 Then
                                    D(0) = Str(nResp)
                                    D(1) = sRespuesta
                                    R.nResultado = CInt(D(0))
                                    R.sRespuesta = D(1)
                                    Return R
                                End If

                                '***** RESUMEN DEUDA
                                If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenFlags.ResumenComportamiento.ResumenDeuda IsNot Nothing Then

                                    Dim resumenDeuda As New ResumenDeuda
                                    resumenDeuda = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenFlags.ResumenComportamiento.ResumenDeuda

                                    oComm = New SqlCommand("[dbo].[PA_RegistrarResumenDeuda]", oConex)
                                    oComm.CommandType = CommandType.StoredProcedure
                                    oComm.Parameters.Add("@id_ResumenComportamiento", SqlDbType.Int).Value = id_ResumenComportamiento
                                    oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial

                                    oComm.Parameters.Add("@Periodo", SqlDbType.VarChar).Value = resumenDeuda.Periodo.NullToEmpty()
                                    oComm.Parameters.Add("@DeudaTotal", SqlDbType.Decimal).Value = resumenDeuda.DeudaTotal
                                    oComm.Parameters.Add("@DeudaDirecta", SqlDbType.Decimal).Value = resumenDeuda.DeudaDirecta
                                    oComm.Parameters.Add("@DeudaIndirecta", SqlDbType.Decimal).Value = resumenDeuda.DeudaIndirecta
                                    oComm.Parameters.Add("@PorcentajeDeudaNormal", SqlDbType.Decimal).Value = resumenDeuda.PorcentajeDeudaNormal
                                    oComm.Parameters.Add("@PorcentajeDeudaPotencial", SqlDbType.Decimal).Value = resumenDeuda.PorcentajeDeudaPotencial
                                    oComm.Parameters.Add("@PorcentajeDeudaDeficiente", SqlDbType.Decimal).Value = resumenDeuda.PorcentajeDeudaDeficiente
                                    oComm.Parameters.Add("@PorcentajeDeudaEnRiesgo", SqlDbType.Decimal).Value = resumenDeuda.PorcentajeDeudaEnRiesgo
                                    oComm.Parameters.Add("@PorcentajeDeudaPerdida", SqlDbType.Decimal).Value = resumenDeuda.PorcentajeDeudaPerdida
                                    oComm.Parameters.Add("@Variacion", SqlDbType.Decimal).Value = resumenDeuda.Variacion

                                    oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                    'oComm.Parameters.Add("@id_ResumenComportamiento", SqlDbType.Int).Direction = ParameterDirection.Output

                                    oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                    oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                    oComm.ExecuteNonQuery()
                                    nResp = oComm.Parameters("@nResp").Value
                                    sRespuesta = oComm.Parameters("@Resultado").Value
                                    'id_ResumenComportamiento = oComm.Parameters("@id_ResumenComportamiento").Value

                                    If nResp = 0 Then
                                        D(0) = Str(nResp)
                                        D(1) = sRespuesta
                                        R.nResultado = CInt(D(0))
                                        R.sRespuesta = D(1)
                                        Return R
                                    End If
                                End If
                                '*********
                                '***** RESUMEN SCORE HISTORICO
                                Dim id_ResumenScoreHistorico As Integer = 0
                                If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenFlags.ResumenComportamiento.ResumenScoreHistorico IsNot Nothing Then
                                    Dim resumenScoreHistorico As New ResumenScoreHistorico
                                    resumenScoreHistorico = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenFlags.ResumenComportamiento.ResumenScoreHistorico

                                    oComm = New SqlCommand("[dbo].[PA_RegistrarResumenScoreHistorico]", oConex)
                                    oComm.CommandType = CommandType.StoredProcedure
                                    oComm.Parameters.Add("@id_ResumenComportamiento", SqlDbType.Int).Value = id_ResumenComportamiento
                                    oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial

                                    oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                    oComm.Parameters.Add("@id_ResumenScoreHistorico", SqlDbType.Int).Direction = ParameterDirection.Output

                                    oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                    oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                    oComm.ExecuteNonQuery()
                                    nResp = oComm.Parameters("@nResp").Value
                                    sRespuesta = oComm.Parameters("@Resultado").Value
                                    id_ResumenScoreHistorico = oComm.Parameters("@id_ResumenScoreHistorico").Value

                                    If nResp = 0 Then
                                        D(0) = Str(nResp)
                                        D(1) = sRespuesta
                                        R.nResultado = CInt(D(0))
                                        R.sRespuesta = D(1)
                                        Return R
                                    End If

                                    'SCORE ACTUAL
                                    If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.
                                        ResumenFlags.ResumenComportamiento.ResumenScoreHistorico.ScoreActual IsNot Nothing Then
                                        Dim scoreActual As New ScoreData
                                        scoreActual = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenFlags.
                                            ResumenComportamiento.ResumenScoreHistorico.ScoreActual
                                        oComm = New SqlCommand("[dbo].[PA_RegistrarScore_ResumenScoreHistorico]", oConex)
                                        oComm.CommandType = CommandType.StoredProcedure
                                        oComm.Parameters.Add("@id_ResumenScoreHistorico", SqlDbType.Int).Value = id_ResumenScoreHistorico
                                        oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial
                                        oComm.Parameters.Add("@Periodo", SqlDbType.VarChar).Value = scoreActual.Periodo.NullToEmpty()
                                        oComm.Parameters.Add("@Riesgo", SqlDbType.VarChar).Value = scoreActual.Riesgo.NullToEmpty()
                                        oComm.Parameters.Add("@MotivoSinScore", SqlDbType.VarChar).Value = scoreActual.MotivoSinScore.NullToEmpty()
                                        oComm.Parameters.Add("@nTipo", SqlDbType.Int).Value = 1
                                        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                        oComm.ExecuteNonQuery()
                                        nResp = oComm.Parameters("@nResp").Value
                                        sRespuesta = oComm.Parameters("@Resultado").Value
                                        If nResp = 0 Then
                                            D(0) = Str(nResp)
                                            D(1) = sRespuesta
                                            R.nResultado = CInt(D(0))
                                            R.sRespuesta = D(1)
                                            Return R
                                        End If
                                    End If
                                    '***************
                                    'SCORE ANTERIOR
                                    If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.
                                        ResumenFlags.ResumenComportamiento.ResumenScoreHistorico.ScoreAnterior IsNot Nothing Then
                                        Dim scoreActual As New ScoreData
                                        scoreActual = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenFlags.
                                            ResumenComportamiento.ResumenScoreHistorico.ScoreAnterior
                                        oComm = New SqlCommand("[dbo].[PA_RegistrarScore_ResumenScoreHistorico]", oConex)
                                        oComm.CommandType = CommandType.StoredProcedure
                                        oComm.Parameters.Add("@id_ResumenScoreHistorico", SqlDbType.Int).Value = id_ResumenScoreHistorico
                                        oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial
                                        oComm.Parameters.Add("@Periodo", SqlDbType.VarChar).Value = scoreActual.Periodo.NullToEmpty()
                                        oComm.Parameters.Add("@Riesgo", SqlDbType.VarChar).Value = scoreActual.Riesgo.NullToEmpty()
                                        oComm.Parameters.Add("@MotivoSinScore", SqlDbType.VarChar).Value = scoreActual.MotivoSinScore.NullToEmpty()
                                        oComm.Parameters.Add("@nTipo", SqlDbType.Int).Value = 2
                                        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                        oComm.ExecuteNonQuery()
                                        nResp = oComm.Parameters("@nResp").Value
                                        sRespuesta = oComm.Parameters("@Resultado").Value
                                        If nResp = 0 Then
                                            D(0) = Str(nResp)
                                            D(1) = sRespuesta
                                            R.nResultado = CInt(D(0))
                                            R.sRespuesta = D(1)
                                            Return R
                                        End If
                                    End If
                                    '***************
                                    'SCORE HACE 12 MESES
                                    If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.
                                        ResumenFlags.ResumenComportamiento.ResumenScoreHistorico.ScoreHace12Meses IsNot Nothing Then
                                        Dim scoreActual As New ScoreData
                                        scoreActual = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenFlags.
                                            ResumenComportamiento.ResumenScoreHistorico.ScoreHace12Meses
                                        oComm = New SqlCommand("[dbo].[PA_RegistrarScore_ResumenScoreHistorico]", oConex)
                                        oComm.CommandType = CommandType.StoredProcedure
                                        oComm.Parameters.Add("@id_ResumenScoreHistorico", SqlDbType.Int).Value = id_ResumenScoreHistorico
                                        oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial
                                        oComm.Parameters.Add("@Periodo", SqlDbType.VarChar).Value = scoreActual.Periodo.NullToEmpty()
                                        oComm.Parameters.Add("@Riesgo", SqlDbType.VarChar).Value = scoreActual.Riesgo.NullToEmpty()
                                        oComm.Parameters.Add("@MotivoSinScore", SqlDbType.VarChar).Value = scoreActual.MotivoSinScore.NullToEmpty()
                                        oComm.Parameters.Add("@nTipo", SqlDbType.Int).Value = 3
                                        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                        oComm.ExecuteNonQuery()
                                        nResp = oComm.Parameters("@nResp").Value
                                        sRespuesta = oComm.Parameters("@Resultado").Value
                                        If nResp = 0 Then
                                            D(0) = Str(nResp)
                                            D(1) = sRespuesta
                                            R.nResultado = CInt(D(0))
                                            R.sRespuesta = D(1)
                                            Return R
                                        End If
                                    End If
                                    '***************

                                End If
                                '*********

                            End If

                            '***** RESUMEN BLOQUE FLAGS
                            Dim id_ResumenBloqueFlags As Integer = 0
                            If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenFlags.ResumenBloqueFlags IsNot Nothing Then
                                Dim resumenBloqueFlags As New ResumenBloqueFlags
                                resumenBloqueFlags = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenFlags.ResumenBloqueFlags

                                oComm = New SqlCommand("[dbo].[PA_RegistrarResumenBloqueFlags]", oConex)
                                oComm.CommandType = CommandType.StoredProcedure
                                oComm.Parameters.Add("@id_ResumenFlags", SqlDbType.Int).Value = id_ResumenFlags
                                oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial

                                oComm.Parameters.Add("@TarjetaCredito", SqlDbType.VarChar).Value = resumenBloqueFlags.TarjetaCredito.NullToEmpty()
                                oComm.Parameters.Add("@LineaDeCredito", SqlDbType.VarChar).Value = resumenBloqueFlags.LineaDeCredito.NullToEmpty()
                                oComm.Parameters.Add("@CreditoHipotecario", SqlDbType.VarChar).Value = resumenBloqueFlags.CreditoHipotecario.NullToEmpty()
                                oComm.Parameters.Add("@BuenPagadorDeServicios", SqlDbType.VarChar).Value = resumenBloqueFlags.BuenPagadorDeServicios.NullToEmpty()
                                oComm.Parameters.Add("@EstaEnInfocorp", SqlDbType.VarChar).Value = resumenBloqueFlags.EstaEnInfocorp.NullToEmpty()
                                oComm.Parameters.Add("@AvalAvalado", SqlDbType.VarChar).Value = resumenBloqueFlags.AvalAvalado.NullToEmpty()
                                oComm.Parameters.Add("@RepresentanteLegal", SqlDbType.VarChar).Value = resumenBloqueFlags.RepresentanteLegal.NullToEmpty()
                                oComm.Parameters.Add("@GastoMensualEstimado", SqlDbType.Decimal).Value = resumenBloqueFlags.GastoMensualEstimado
                                oComm.Parameters.Add("@PosibleRestringido", SqlDbType.VarChar).Value = resumenBloqueFlags.PosibleRestringido.NullToEmpty()
                                oComm.Parameters.Add("@TieneAuto", SqlDbType.VarChar).Value = resumenBloqueFlags.TieneAuto.NullToEmpty()
                                oComm.Parameters.Add("@EntidadesQueConsultaron", SqlDbType.Int).Value = resumenBloqueFlags.EntidadesQueConsultaron
                                oComm.Parameters.Add("@Homonimos", SqlDbType.VarChar).Value = resumenBloqueFlags.Homonimos.NullToEmpty()
                                oComm.Parameters.Add("@ComercioExterior", SqlDbType.VarChar).Value = resumenBloqueFlags.ComercioExterior.NullToEmpty()
                                oComm.Parameters.Add("@DeudaPrevisional", SqlDbType.VarChar).Value = resumenBloqueFlags.DeudaPrevisional.NullToEmpty()
                                oComm.Parameters.Add("@AlertaPep", SqlDbType.VarChar).Value = resumenBloqueFlags.AlertaPep.NullToEmpty()
                                oComm.Parameters.Add("@AlertaRedam", SqlDbType.VarChar).Value = resumenBloqueFlags.AlertaRedam.NullToEmpty()
                                oComm.Parameters.Add("@ReactivaPeru", SqlDbType.VarChar).Value = resumenBloqueFlags.ReactivaPeru.NullToEmpty()
                                oComm.Parameters.Add("@EsInquilinoMoroso", SqlDbType.VarChar).Value = resumenBloqueFlags.EsInquilinoMoroso.NullToEmpty()


                                oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                oComm.Parameters.Add("@id_ResumenBloqueFlags", SqlDbType.Int).Direction = ParameterDirection.Output

                                oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                oComm.ExecuteNonQuery()
                                nResp = oComm.Parameters("@nResp").Value
                                sRespuesta = oComm.Parameters("@Resultado").Value
                                id_ResumenBloqueFlags = oComm.Parameters("@id_ResumenBloqueFlags").Value

                                If nResp = 0 Then
                                    D(0) = Str(nResp)
                                    D(1) = sRespuesta
                                    R.nResultado = CInt(D(0))
                                    R.sRespuesta = D(1)
                                    Return R
                                End If



                                'Reactiva Peru Info
                                If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenFlags.ResumenBloqueFlags.ReactivaPeruInfo IsNot Nothing Then
                                    Dim reactivaPeruInfo As New ReactivaPeruInfo
                                    reactivaPeruInfo = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenFlags.ResumenBloqueFlags.ReactivaPeruInfo

                                    oComm = New SqlCommand("[dbo].[PA_RegistrarReactivaPeruInfo]", oConex)
                                    oComm.CommandType = CommandType.StoredProcedure
                                    oComm.Parameters.Add("@id_ResumenBloqueFlags", SqlDbType.Int).Value = id_ResumenBloqueFlags
                                    oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial

                                    oComm.Parameters.Add("@Fecha", SqlDbType.VarChar).Value = reactivaPeruInfo.Fecha.NullToEmpty()
                                    oComm.Parameters.Add("@Monto", SqlDbType.VarChar).Value = reactivaPeruInfo.Monto.NullToEmpty()

                                    oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu

                                    oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                    oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                    oComm.ExecuteNonQuery()
                                    nResp = oComm.Parameters("@nResp").Value
                                    sRespuesta = oComm.Parameters("@Resultado").Value

                                    If nResp = 0 Then
                                        D(0) = Str(nResp)
                                        D(1) = sRespuesta
                                        R.nResultado = CInt(D(0))
                                        R.sRespuesta = D(1)
                                        Return R
                                    End If
                                End If
                                '****************

                            End If
                            '*****************************



                        End If
                        'END RESUMEN FLAG

                        'SCORE PREDICTIVO CON VARIABLES 822
                        Dim id_ResumenScoreRP3_Cabecera As Integer = 0
                        If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Codigo = 822 Then
                            oComm = New SqlCommand("[dbo].[PA_RegistrarResumenScoreRP3_Cabecera]", oConex)
                            oComm.CommandType = CommandType.StoredProcedure
                            oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial

                            oComm.Parameters.Add("@CodigoModulo", SqlDbType.VarChar).Value = "822"
                            oComm.Parameters.Add("@NombreModulo", SqlDbType.VarChar).Value = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Nombre.NullToEmpty()

                            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu

                            oComm.Parameters.Add("@id_ResumenScoreRP3_Cabecera", SqlDbType.Int).Direction = ParameterDirection.Output

                            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                            oComm.ExecuteNonQuery()
                            nResp = oComm.Parameters("@nResp").Value
                            sRespuesta = oComm.Parameters("@Resultado").Value
                            id_ResumenScoreRP3_Cabecera = oComm.Parameters("@id_ResumenScoreRP3_Cabecera").Value

                            If nResp = 0 Then
                                D(0) = Str(nResp)
                                D(1) = sRespuesta
                                R.nResultado = CInt(D(0))
                                R.sRespuesta = D(1)
                                Return R
                            End If


                            j = 0
                            Dim id_ResumenScoreRP3 As Integer = 0
                            If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenScoreRP3 IsNot Nothing Then
                                Dim resumenScoreRP3 As New ResumenScoreRP3
                                resumenScoreRP3 = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenScoreRP3

                                oComm = New SqlCommand("[dbo].[PA_RegistrarResumenScoreRP3]", oConex)
                                oComm.CommandType = CommandType.StoredProcedure
                                oComm.Parameters.Add("@id_ResumenScoreRP3_Cabecera", SqlDbType.Int).Value = id_ResumenScoreRP3_Cabecera
                                oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial

                                oComm.Parameters.Add("@Puntaje", SqlDbType.Int).Value = resumenScoreRP3.Puntaje
                                oComm.Parameters.Add("@NivelRiesgo", SqlDbType.VarChar).Value = resumenScoreRP3.NivelRiesgo.NullToEmpty()
                                oComm.Parameters.Add("@Conclusion", SqlDbType.VarChar).Value = resumenScoreRP3.Conclusion.NullToEmpty()

                                oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                oComm.Parameters.Add("@id_ResumenScoreRP3", SqlDbType.Int).Direction = ParameterDirection.Output

                                oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                oComm.ExecuteNonQuery()
                                nResp = oComm.Parameters("@nResp").Value
                                sRespuesta = oComm.Parameters("@Resultado").Value
                                id_ResumenScoreRP3 = oComm.Parameters("@id_ResumenScoreRP3").Value

                                If nResp = 0 Then
                                    D(0) = Str(nResp)
                                    D(1) = sRespuesta
                                    R.nResultado = CInt(D(0))
                                    R.sRespuesta = D(1)
                                    Return R
                                End If

                                Dim id_PrincipalesVariables As Integer = 0
                                If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenScoreRP3.PrincipalesVariables IsNot Nothing Then

                                    oComm = New SqlCommand("[dbo].[PA_RegistrarPrincipalesVariables]", oConex)
                                    oComm.CommandType = CommandType.StoredProcedure
                                    oComm.Parameters.Add("@id_ResumenScoreRP3", SqlDbType.Int).Value = id_ResumenScoreRP3
                                    oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial

                                    oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                    oComm.Parameters.Add("@id_PrincipalesVariables", SqlDbType.Int).Direction = ParameterDirection.Output

                                    oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                    oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                    oComm.ExecuteNonQuery()
                                    nResp = oComm.Parameters("@nResp").Value
                                    sRespuesta = oComm.Parameters("@Resultado").Value
                                    id_PrincipalesVariables = oComm.Parameters("@id_PrincipalesVariables").Value

                                    If nResp = 0 Then
                                        D(0) = Str(nResp)
                                        D(1) = sRespuesta
                                        R.nResultado = CInt(D(0))
                                        R.sRespuesta = D(1)
                                        Return R
                                    End If


                                    If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenScoreRP3.PrincipalesVariables.PresentaInformacionSFR IsNot Nothing Then
                                        Dim variable As VariableValor
                                        variable = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenScoreRP3.PrincipalesVariables.PresentaInformacionSFR
                                        oComm = New SqlCommand("[dbo].[PA_RegistrarVariableValor]", oConex)
                                        oComm.CommandType = CommandType.StoredProcedure
                                        oComm.Parameters.Add("@id_PrincipalesVariables", SqlDbType.Int).Value = id_PrincipalesVariables
                                        oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial
                                        oComm.Parameters.Add("@Variable", SqlDbType.VarChar).Value = variable.Variable.NullToEmpty
                                        oComm.Parameters.Add("@Valor", SqlDbType.VarChar).Value = variable.Valor.NullToEmpty
                                        oComm.Parameters.Add("@nTipo", SqlDbType.Int).Value = 1
                                        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                        oComm.ExecuteNonQuery()
                                        nResp = oComm.Parameters("@nResp").Value
                                        sRespuesta = oComm.Parameters("@Resultado").Value

                                        If nResp = 0 Then
                                            D(0) = Str(nResp)
                                            D(1) = sRespuesta
                                            R.nResultado = CInt(D(0))
                                            R.sRespuesta = D(1)
                                            Return R
                                        End If
                                    End If
                                    If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenScoreRP3.PrincipalesVariables.PresentaDeudas IsNot Nothing Then
                                        Dim variable As VariableValor
                                        variable = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenScoreRP3.PrincipalesVariables.PresentaDeudas
                                        oComm = New SqlCommand("[dbo].[PA_RegistrarVariableValor]", oConex)
                                        oComm.CommandType = CommandType.StoredProcedure
                                        oComm.Parameters.Add("@id_PrincipalesVariables", SqlDbType.Int).Value = id_PrincipalesVariables
                                        oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial
                                        oComm.Parameters.Add("@Variable", SqlDbType.VarChar).Value = variable.Variable.NullToEmpty
                                        oComm.Parameters.Add("@Valor", SqlDbType.VarChar).Value = variable.Valor.NullToEmpty
                                        oComm.Parameters.Add("@nTipo", SqlDbType.Int).Value = 2
                                        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                        oComm.ExecuteNonQuery()
                                        nResp = oComm.Parameters("@nResp").Value
                                        sRespuesta = oComm.Parameters("@Resultado").Value
                                        If nResp = 0 Then
                                            D(0) = Str(nResp)
                                            D(1) = sRespuesta
                                            R.nResultado = CInt(D(0))
                                            R.sRespuesta = D(1)
                                            Return R
                                        End If
                                    End If
                                    If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenScoreRP3.PrincipalesVariables.InformacionReportadaInfocorp IsNot Nothing Then
                                        Dim variable As VariableValor
                                        variable = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenScoreRP3.PrincipalesVariables.InformacionReportadaInfocorp
                                        oComm = New SqlCommand("[dbo].[PA_RegistrarVariableValor]", oConex)
                                        oComm.CommandType = CommandType.StoredProcedure
                                        oComm.Parameters.Add("@id_PrincipalesVariables", SqlDbType.Int).Value = id_PrincipalesVariables
                                        oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial
                                        oComm.Parameters.Add("@Variable", SqlDbType.VarChar).Value = variable.Variable.NullToEmpty
                                        oComm.Parameters.Add("@Valor", SqlDbType.VarChar).Value = variable.Valor.NullToEmpty
                                        oComm.Parameters.Add("@nTipo", SqlDbType.Int).Value = 3
                                        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                        oComm.ExecuteNonQuery()
                                        nResp = oComm.Parameters("@nResp").Value
                                        sRespuesta = oComm.Parameters("@Resultado").Value
                                        If nResp = 0 Then
                                            D(0) = Str(nResp)
                                            D(1) = sRespuesta
                                            R.nResultado = CInt(D(0))
                                            R.sRespuesta = D(1)
                                            Return R
                                        End If
                                    End If
                                    If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenScoreRP3.PrincipalesVariables.MalComportamientoCrediticio IsNot Nothing Then
                                        Dim variable As VariableValor
                                        variable = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenScoreRP3.PrincipalesVariables.MalComportamientoCrediticio
                                        oComm = New SqlCommand("[dbo].[PA_RegistrarVariableValor]", oConex)
                                        oComm.CommandType = CommandType.StoredProcedure
                                        oComm.Parameters.Add("@id_PrincipalesVariables", SqlDbType.Int).Value = id_PrincipalesVariables
                                        oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial
                                        oComm.Parameters.Add("@Variable", SqlDbType.VarChar).Value = variable.Variable.NullToEmpty
                                        oComm.Parameters.Add("@Valor", SqlDbType.VarChar).Value = variable.Valor.NullToEmpty
                                        oComm.Parameters.Add("@nTipo", SqlDbType.Int).Value = 4
                                        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                        oComm.ExecuteNonQuery()
                                        nResp = oComm.Parameters("@nResp").Value
                                        sRespuesta = oComm.Parameters("@Resultado").Value
                                        If nResp = 0 Then
                                            D(0) = Str(nResp)
                                            D(1) = sRespuesta
                                            R.nResultado = CInt(D(0))
                                            R.sRespuesta = D(1)
                                            Return R
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        'END SCORE PREDICTIVO CON VARIABLES

                        'COMPORTAMIENTO GENERAL 861
                        Dim id_ComportamientoGeneral As Integer = 0
                        If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Codigo = 861 Then
                            oComm = New SqlCommand("[dbo].[PA_RegistrarComportamientoGeneral]", oConex)
                            oComm.CommandType = CommandType.StoredProcedure
                            oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial
                            oComm.Parameters.Add("@CodigoModulo", SqlDbType.VarChar).Value = "861"
                            oComm.Parameters.Add("@NombreModulo", SqlDbType.VarChar).Value = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Nombre.NullToEmpty()
                            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                            oComm.Parameters.Add("@id_ComportamientoGeneral", SqlDbType.Int).Direction = ParameterDirection.Output
                            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                            oComm.ExecuteNonQuery()
                            nResp = oComm.Parameters("@nResp").Value
                            sRespuesta = oComm.Parameters("@Resultado").Value
                            id_ComportamientoGeneral = oComm.Parameters("@id_ComportamientoGeneral").Value
                            If nResp = 0 Then
                                D(0) = Str(nResp)
                                D(1) = sRespuesta
                                R.nResultado = CInt(D(0))
                                R.sRespuesta = D(1)
                                Return R
                            End If
                            '2
                            Dim id_SistemaFinanciero As Integer = 0
                            If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.SistemaFinanciero IsNot Nothing Then
                                oComm = New SqlCommand("[dbo].[PA_RegistrarSistemaFinanciero]", oConex)
                                oComm.CommandType = CommandType.StoredProcedure
                                oComm.Parameters.Add("@id_ComportamientoGeneral", SqlDbType.Int).Value = id_ComportamientoGeneral
                                oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial
                                oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                oComm.Parameters.Add("@id_SistemaFinanciero", SqlDbType.Int).Direction = ParameterDirection.Output
                                oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                oComm.ExecuteNonQuery()
                                nResp = oComm.Parameters("@nResp").Value
                                sRespuesta = oComm.Parameters("@Resultado").Value
                                id_SistemaFinanciero = oComm.Parameters("@id_SistemaFinanciero").Value
                                If nResp = 0 Then
                                    D(0) = Str(nResp)
                                    D(1) = sRespuesta
                                    R.nResultado = CInt(D(0))
                                    R.sRespuesta = D(1)
                                    Return R
                                End If
                                '3
                                j = 0
                                If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.SistemaFinanciero.SFDetalle IsNot Nothing Then
                                    If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.SistemaFinanciero.SFDetalle.Count > 0 Then
                                        Dim listSFDetalle As New SFDetalle()
                                        For j = 0 To oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.SistemaFinanciero.SFDetalle.Count - 1
                                            oComm = New SqlCommand("[dbo].[PA_RegistrarSistemaFinanciero_SFDetalle]", oConex)
                                            oComm.CommandType = CommandType.StoredProcedure
                                            oComm.Parameters.Add("@id_SistemaFinanciero", SqlDbType.Int).Value = id_SistemaFinanciero
                                            oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial

                                            oComm.Parameters.Add("@Periodo", SqlDbType.VarChar).Value = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.SistemaFinanciero.SFDetalle(j).Periodo.NullToEmpty

                                            '**************** validar el NOTHING para INT

                                            oComm.Parameters.Add("@DiasAtrasoMinimo", SqlDbType.Int).Value = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.SistemaFinanciero.SFDetalle(j).DiasAtrasoMinimo.NullToZero
                                            oComm.Parameters.Add("@DiasAtrasoMaximo", SqlDbType.Int).Value = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.SistemaFinanciero.SFDetalle(j).DiasAtrasoMaximo.NullToZero
                                            oComm.Parameters.Add("@Riesgo", SqlDbType.VarChar).Value = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.SistemaFinanciero.SFDetalle(j).Riesgo.NullToEmpty

                                            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                            oComm.ExecuteNonQuery()
                                            nResp = oComm.Parameters("@nResp").Value
                                            sRespuesta = oComm.Parameters("@Resultado").Value

                                            If nResp = 0 Then
                                                D(0) = Str(nResp)
                                                D(1) = sRespuesta
                                                R.nResultado = CInt(D(0))
                                                R.sRespuesta = D(1)
                                                Return R
                                            End If
                                        Next
                                    End If
                                End If
                            End If
                            '2
                            Dim id_OtrasDeudasImpagas As Integer = 0
                            j = 0
                            If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.OtrasDeudasImpagas IsNot Nothing Then
                                oComm = New SqlCommand("[dbo].[PA_RegistrarOtrasDeudasImpagas]", oConex)
                                oComm.CommandType = CommandType.StoredProcedure
                                oComm.Parameters.Add("@id_ComportamientoGeneral", SqlDbType.Int).Value = id_ComportamientoGeneral
                                oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial
                                oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                oComm.Parameters.Add("@id_OtrasDeudasImpagas", SqlDbType.Int).Direction = ParameterDirection.Output
                                oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                oComm.ExecuteNonQuery()
                                nResp = oComm.Parameters("@nResp").Value
                                sRespuesta = oComm.Parameters("@Resultado").Value
                                id_OtrasDeudasImpagas = oComm.Parameters("@id_OtrasDeudasImpagas").Value
                                If nResp = 0 Then
                                    D(0) = Str(nResp)
                                    D(1) = sRespuesta
                                    R.nResultado = CInt(D(0))
                                    R.sRespuesta = D(1)
                                    Return R
                                End If

                                Dim id_ODIDetalle As Integer = 0
                                If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.OtrasDeudasImpagas.ODIDetalle IsNot Nothing Then
                                    If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.OtrasDeudasImpagas.ODIDetalle.Count > 0 Then
                                        For j = 0 To oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.OtrasDeudasImpagas.ODIDetalle.Count - 1
                                            oComm = New SqlCommand("[dbo].[PA_RegistrarOtrasDeudasImpagas_ODIDetalle]", oConex)
                                            oComm.CommandType = CommandType.StoredProcedure
                                            oComm.Parameters.Add("@id_OtrasDeudasImpagas", SqlDbType.Int).Value = id_OtrasDeudasImpagas
                                            oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial

                                            oComm.Parameters.Add("@Periodo", SqlDbType.VarChar).Value = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.OtrasDeudasImpagas.ODIDetalle(j).Periodo.NullToEmpty
                                            oComm.Parameters.Add("@Riesgo", SqlDbType.VarChar).Value = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.OtrasDeudasImpagas.ODIDetalle(j).Riesgo.NullToEmpty

                                            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                            oComm.Parameters.Add("@id_ODIDetalle", SqlDbType.Int).Direction = ParameterDirection.Output
                                            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                            oComm.ExecuteNonQuery()
                                            nResp = oComm.Parameters("@nResp").Value
                                            sRespuesta = oComm.Parameters("@Resultado").Value
                                            id_ODIDetalle = oComm.Parameters("@id_ODIDetalle").Value
                                            If nResp = 0 Then
                                                D(0) = Str(nResp)
                                                D(1) = sRespuesta
                                                R.nResultado = CInt(D(0))
                                                R.sRespuesta = D(1)
                                                Return R
                                            End If

                                            If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.OtrasDeudasImpagas.ODIDetalle(j).DeudasImpagasProtestadas.MontoEnDolares IsNot Nothing And oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.OtrasDeudasImpagas.ODIDetalle(j).DeudasImpagasProtestadas.MontoEnSoles IsNot Nothing Then
                                                oComm = New SqlCommand("[dbo].[PA_RegistrarODIDetalle_Deudas]", oConex)
                                                oComm.CommandType = CommandType.StoredProcedure
                                                oComm.Parameters.Add("@id_ODIDetalle", SqlDbType.Int).Value = id_ODIDetalle
                                                oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial

                                                oComm.Parameters.Add("@MontoEnSoles", SqlDbType.Decimal).Value = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.OtrasDeudasImpagas.ODIDetalle(j).DeudasImpagasProtestadas.MontoEnSoles
                                                oComm.Parameters.Add("@MontoEnDolares", SqlDbType.Decimal).Value = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.OtrasDeudasImpagas.ODIDetalle(j).DeudasImpagasProtestadas.MontoEnDolares
                                                oComm.Parameters.Add("@nTipo", SqlDbType.Int).Value = 1
                                                oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                                oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                                oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                                oComm.ExecuteNonQuery()
                                                nResp = oComm.Parameters("@nResp").Value
                                                sRespuesta = oComm.Parameters("@Resultado").Value

                                                If nResp = 0 Then
                                                    D(0) = Str(nResp)
                                                    D(1) = sRespuesta
                                                    R.nResultado = CInt(D(0))
                                                    R.sRespuesta = D(1)
                                                    Return R
                                                End If
                                            End If
                                            If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.OtrasDeudasImpagas.ODIDetalle(j).DeudasImpagasInfocorp.MontoEnDolares IsNot Nothing And oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.OtrasDeudasImpagas.ODIDetalle(j).DeudasImpagasInfocorp.MontoEnSoles IsNot Nothing Then
                                                oComm = New SqlCommand("[dbo].[PA_RegistrarODIDetalle_Deudas]", oConex)
                                                oComm.CommandType = CommandType.StoredProcedure
                                                oComm.Parameters.Add("@id_ODIDetalle", SqlDbType.Int).Value = id_ODIDetalle
                                                oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial

                                                oComm.Parameters.Add("@MontoEnSoles", SqlDbType.Decimal).Value = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.OtrasDeudasImpagas.ODIDetalle(j).DeudasImpagasInfocorp.MontoEnSoles
                                                oComm.Parameters.Add("@MontoEnDolares", SqlDbType.Decimal).Value = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.OtrasDeudasImpagas.ODIDetalle(j).DeudasImpagasInfocorp.MontoEnDolares
                                                oComm.Parameters.Add("@nTipo", SqlDbType.Int).Value = 2
                                                oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                                oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                                oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                                oComm.ExecuteNonQuery()
                                                nResp = oComm.Parameters("@nResp").Value
                                                sRespuesta = oComm.Parameters("@Resultado").Value

                                                If nResp = 0 Then
                                                    D(0) = Str(nResp)
                                                    D(1) = sRespuesta
                                                    R.nResultado = CInt(D(0))
                                                    R.sRespuesta = D(1)
                                                    Return R
                                                End If
                                            End If
                                            If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.OtrasDeudasImpagas.ODIDetalle(j).DeudasImpagasSunat.MontoEnDolares IsNot Nothing And oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.OtrasDeudasImpagas.ODIDetalle(j).DeudasImpagasSunat.MontoEnSoles IsNot Nothing Then
                                                oComm = New SqlCommand("[dbo].[PA_RegistrarODIDetalle_Deudas]", oConex)
                                                oComm.CommandType = CommandType.StoredProcedure
                                                oComm.Parameters.Add("@id_ODIDetalle", SqlDbType.Int).Value = id_ODIDetalle
                                                oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial

                                                oComm.Parameters.Add("@MontoEnSoles", SqlDbType.Decimal).Value = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.OtrasDeudasImpagas.ODIDetalle(j).DeudasImpagasSunat.MontoEnSoles
                                                oComm.Parameters.Add("@MontoEnDolares", SqlDbType.Decimal).Value = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.OtrasDeudasImpagas.ODIDetalle(j).DeudasImpagasSunat.MontoEnDolares
                                                oComm.Parameters.Add("@nTipo", SqlDbType.Int).Value = 3
                                                oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                                oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                                oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                                oComm.ExecuteNonQuery()
                                                nResp = oComm.Parameters("@nResp").Value
                                                sRespuesta = oComm.Parameters("@Resultado").Value

                                                If nResp = 0 Then
                                                    D(0) = Str(nResp)
                                                    D(1) = sRespuesta
                                                    R.nResultado = CInt(D(0))
                                                    R.sRespuesta = D(1)
                                                    Return R
                                                End If
                                            End If
                                            If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.OtrasDeudasImpagas.ODIDetalle(j).DeudasImpagasPrevisional.MontoEnDolares IsNot Nothing And oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.OtrasDeudasImpagas.ODIDetalle(j).DeudasImpagasPrevisional.MontoEnSoles IsNot Nothing Then
                                                oComm = New SqlCommand("[dbo].[PA_RegistrarODIDetalle_Deudas]", oConex)
                                                oComm.CommandType = CommandType.StoredProcedure
                                                oComm.Parameters.Add("@id_ODIDetalle", SqlDbType.Int).Value = id_ODIDetalle
                                                oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial

                                                oComm.Parameters.Add("@MontoEnSoles", SqlDbType.Decimal).Value = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.OtrasDeudasImpagas.ODIDetalle(j).DeudasImpagasPrevisional.MontoEnSoles
                                                oComm.Parameters.Add("@MontoEnDolares", SqlDbType.Decimal).Value = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ComportamientoGeneral.OtrasDeudasImpagas.ODIDetalle(j).DeudasImpagasPrevisional.MontoEnDolares
                                                oComm.Parameters.Add("@nTipo", SqlDbType.Int).Value = 4
                                                oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                                oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                                oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                                oComm.ExecuteNonQuery()
                                                nResp = oComm.Parameters("@nResp").Value
                                                sRespuesta = oComm.Parameters("@Resultado").Value

                                                If nResp = 0 Then
                                                    D(0) = Str(nResp)
                                                    D(1) = sRespuesta
                                                    R.nResultado = CInt(D(0))
                                                    R.sRespuesta = D(1)
                                                    Return R
                                                End If
                                            End If


                                        Next
                                    End If
                                End If

                            End If

                        End If
                        'END COMPORTAMIENTO GENERAL

                        '****** RESUMEN FINANCIERO 859
                        Dim id_ResumenFinanciero As Integer = 0
                        If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Codigo = 859 Then
                            oComm = New SqlCommand("[dbo].[PA_RegistrarResumenFinanciero]", oConex)
                            oComm.CommandType = CommandType.StoredProcedure
                            oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial
                            oComm.Parameters.Add("@CodigoModulo", SqlDbType.VarChar).Value = "859"
                            oComm.Parameters.Add("@NombreModulo", SqlDbType.VarChar).Value = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Nombre.NullToEmpty()

                            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                            oComm.Parameters.Add("@id_ResumenFinanciero", SqlDbType.Int).Direction = ParameterDirection.Output
                            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                            oComm.ExecuteNonQuery()
                            nResp = oComm.Parameters("@nResp").Value
                            sRespuesta = oComm.Parameters("@Resultado").Value
                            id_ResumenFinanciero = oComm.Parameters("@id_ResumenFinanciero").Value
                            If nResp = 0 Then
                                D(0) = Str(nResp)
                                D(1) = sRespuesta
                                R.nResultado = CInt(D(0))
                                R.sRespuesta = D(1)
                                Return R
                            End If
                            '** 2
                            Dim id_DeudasUltimoPeriodo As Integer = 0
                            j = 0
                            If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenFinanciero.DeudasUltimoPeriodo IsNot Nothing Then
                                Dim deudasUltimoPeriodo As New DeudasUltimoPeriodo
                                deudasUltimoPeriodo = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenFinanciero.DeudasUltimoPeriodo

                                oComm = New SqlCommand("[dbo].[PA_RegistrarDeudasUltimoPeriodo]", oConex)
                                oComm.CommandType = CommandType.StoredProcedure
                                oComm.Parameters.Add("@id_ResumenFinanciero", SqlDbType.Int).Value = id_ResumenFinanciero
                                oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial

                                oComm.Parameters.Add("@Periodo", SqlDbType.VarChar).Value = deudasUltimoPeriodo.periodo.NullToEmpty()
                                oComm.Parameters.Add("@MesActual", SqlDbType.VarChar).Value = deudasUltimoPeriodo.MesActual.NullToEmpty()
                                oComm.Parameters.Add("@MesAnterior", SqlDbType.VarChar).Value = deudasUltimoPeriodo.MesAnterior.NullToEmpty()
                                oComm.Parameters.Add("@AnioAnterior", SqlDbType.VarChar).Value = deudasUltimoPeriodo.AnioAnterior.NullToEmpty()

                                oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                oComm.Parameters.Add("@id_DeudasUltimoPeriodo", SqlDbType.Int).Direction = ParameterDirection.Output
                                oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                oComm.ExecuteNonQuery()
                                nResp = oComm.Parameters("@nResp").Value
                                sRespuesta = oComm.Parameters("@Resultado").Value
                                id_DeudasUltimoPeriodo = oComm.Parameters("@id_DeudasUltimoPeriodo").Value
                                If nResp = 0 Then
                                    D(0) = Str(nResp)
                                    D(1) = sRespuesta
                                    R.nResultado = CInt(D(0))
                                    R.sRespuesta = D(1)
                                    Return R
                                End If

                                '*** 3 DEUDAS LISTA
                                Dim id_DeudasUltimoPeriodo_Deudas As Integer = 0
                                For j = 0 To oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenFinanciero.DeudasUltimoPeriodo.Deudas.Count - 1

                                    Dim deudadUP As New DeudaUP
                                    deudadUP = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenFinanciero.DeudasUltimoPeriodo.Deudas(j)

                                    oComm = New SqlCommand("[dbo].[PA_RegistrarDeudasUltimoPeriodo_Deudas]", oConex)
                                    oComm.CommandType = CommandType.StoredProcedure
                                    oComm.Parameters.Add("@id_DeudasUltimoPeriodo", SqlDbType.Int).Value = id_DeudasUltimoPeriodo
                                    oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial

                                    oComm.Parameters.Add("@Entidad", SqlDbType.VarChar).Value = deudadUP.Entidad.NullToEmpty()
                                    oComm.Parameters.Add("@SistemaFinanciero", SqlDbType.VarChar).Value = deudadUP.SistemaFinanciero.NullToEmpty()
                                    oComm.Parameters.Add("@CalificacionMesActual", SqlDbType.VarChar).Value = deudadUP.CalificacionMesActual.NullToEmpty()
                                    oComm.Parameters.Add("@MontoSubTotalMesActual", SqlDbType.Decimal).Value = deudadUP.MontoSubTotalMesActual 'NullToZero()
                                    oComm.Parameters.Add("@CalificacionMesAnterior", SqlDbType.VarChar).Value = deudadUP.CalificacionMesAnterior.NullToEmpty()
                                    oComm.Parameters.Add("@MontoSubTotalMesAnterior", SqlDbType.Decimal).Value = deudadUP.MontoSubTotalMesAnterior '.NullToZero()
                                    oComm.Parameters.Add("@CalificacionAnioAnterior", SqlDbType.VarChar).Value = deudadUP.CalificacionAnioAnterior.NullToEmpty()
                                    oComm.Parameters.Add("@MontoSubTotalAnioAnterior", SqlDbType.Decimal).Value = deudadUP.MontoSubTotalAnioAnterior '.NullToZero()

                                    oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                    oComm.Parameters.Add("@id_DeudasUltimoPeriodo_Deudas", SqlDbType.Int).Direction = ParameterDirection.Output
                                    oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                    oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                    oComm.ExecuteNonQuery()
                                    nResp = oComm.Parameters("@nResp").Value
                                    sRespuesta = oComm.Parameters("@Resultado").Value
                                    id_DeudasUltimoPeriodo_Deudas = oComm.Parameters("@id_DeudasUltimoPeriodo_Deudas").Value

                                    If nResp = 0 Then
                                        D(0) = Str(nResp)
                                        D(1) = sRespuesta
                                        R.nResultado = CInt(D(0))
                                        R.sRespuesta = D(1)
                                        Return R
                                    End If

                                    Dim z As Integer = 0
                                    For z = 0 To oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenFinanciero.DeudasUltimoPeriodo.Deudas(j).Productos.Count - 1
                                        Dim prod As New Producto
                                        prod = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenFinanciero.DeudasUltimoPeriodo.Deudas(j).Productos(z)

                                        oComm = New SqlCommand("[dbo].[PA_RegistrarDeudasUltimoPeriodoDeudas_Productos]", oConex)
                                        oComm.CommandType = CommandType.StoredProcedure
                                        oComm.Parameters.Add("@id_DeudasUltimoPeriodo_Deudas", SqlDbType.Int).Value = id_DeudasUltimoPeriodo_Deudas
                                        oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial

                                        oComm.Parameters.Add("@Tipo", SqlDbType.VarChar).Value = prod.Tipo.NullToEmpty()
                                        oComm.Parameters.Add("@Descripcion", SqlDbType.VarChar).Value = prod.Descripcion.NullToEmpty()
                                        oComm.Parameters.Add("@MontoMesActual", SqlDbType.Decimal).Value = prod.MontoMesActual
                                        oComm.Parameters.Add("@MontoMesAnterior", SqlDbType.Decimal).Value = prod.MontoMesAnterior
                                        oComm.Parameters.Add("@MontoAnioAnterior", SqlDbType.Decimal).Value = prod.MontoAnioAnterior
                                        oComm.Parameters.Add("@DiasAtraso", SqlDbType.VarChar).Value = prod.DiasAtraso.NullToEmpty()

                                        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                        oComm.ExecuteNonQuery()
                                        nResp = oComm.Parameters("@nResp").Value
                                        sRespuesta = oComm.Parameters("@Resultado").Value


                                        If nResp = 0 Then
                                            D(0) = Str(nResp)
                                            D(1) = sRespuesta
                                            R.nResultado = CInt(D(0))
                                            R.sRespuesta = D(1)
                                            Return R
                                        End If
                                    Next


                                Next
                                '*****

                                '*** 3 TOTALES
                                If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenFinanciero.DeudasUltimoPeriodo.Totales IsNot Nothing Then
                                    Dim totales As New Totales
                                    totales = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.ResumenFinanciero.DeudasUltimoPeriodo.Totales

                                    oComm = New SqlCommand("[dbo].[PA_RegistrarDeudasUltimoPeriodo_Totales]", oConex)
                                    oComm.CommandType = CommandType.StoredProcedure
                                    oComm.Parameters.Add("@id_DeudasUltimoPeriodo", SqlDbType.Int).Value = id_DeudasUltimoPeriodo
                                    oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial

                                    oComm.Parameters.Add("@MontoTotalMesActual", SqlDbType.Decimal).Value = totales.MontoTotalMesActual '.NullToZero()
                                    oComm.Parameters.Add("@MontoTotalMesAnterior", SqlDbType.Decimal).Value = totales.MontoTotalMesAnterior '.NullToZero()
                                    oComm.Parameters.Add("@MontoTotalAnioAnterior", SqlDbType.Decimal).Value = totales.MontoTotalAnioAnterior '.NullToZero()

                                    oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                    oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                    oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                    oComm.ExecuteNonQuery()
                                    nResp = oComm.Parameters("@nResp").Value
                                    sRespuesta = oComm.Parameters("@Resultado").Value


                                    If nResp = 0 Then
                                        D(0) = Str(nResp)
                                        D(1) = sRespuesta
                                        R.nResultado = CInt(D(0))
                                        R.sRespuesta = D(1)
                                        Return R
                                    End If
                                End If

                                '*****

                            End If
                        End If
                        '*****************************

                        '****** SISTEMA FINANCIERO REGULADO (SBS) Y NO REGULADO (MICROFINANZAS) 873
                        Dim id_SistemaFinancieroReguladoYNoRegulado As Integer
                        If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Codigo = 873 Then
                            oComm = New SqlCommand("[dbo].[PA_RegistrarSistemaFinancieroReguladoYNoRegulado]", oConex)
                            oComm.CommandType = CommandType.StoredProcedure
                            oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial
                            oComm.Parameters.Add("@CodigoModulo", SqlDbType.VarChar).Value = "873"
                            oComm.Parameters.Add("@NombreModulo", SqlDbType.VarChar).Value = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Nombre.NullToEmpty()

                            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                            oComm.Parameters.Add("@id_SistemaFinancieroReguladoYNoRegulado", SqlDbType.Int).Direction = ParameterDirection.Output
                            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                            oComm.ExecuteNonQuery()
                            nResp = oComm.Parameters("@nResp").Value
                            sRespuesta = oComm.Parameters("@Resultado").Value
                            id_SistemaFinancieroReguladoYNoRegulado = oComm.Parameters("@id_SistemaFinancieroReguladoYNoRegulado").Value
                            If nResp = 0 Then
                                D(0) = Str(nResp)
                                D(1) = sRespuesta
                                R.nResultado = CInt(D(0))
                                R.sRespuesta = D(1)
                                Return R
                            End If


                            '2
                            Dim id_ResumenComportamientoPago As Integer = 0
                            If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.SistemaFinanciero.ResumenComportamientoPago IsNot Nothing Then
                                oComm = New SqlCommand("[dbo].[PA_RegistrarResumenComportamientoPago]", oConex)
                                oComm.CommandType = CommandType.StoredProcedure
                                oComm.Parameters.Add("@id_SistemaFinancieroReguladoYNoRegulado", SqlDbType.Int).Value = id_SistemaFinancieroReguladoYNoRegulado
                                oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial

                                oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                oComm.Parameters.Add("@id_ResumenComportamientoPago", SqlDbType.Int).Direction = ParameterDirection.Output
                                oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                oComm.ExecuteNonQuery()
                                nResp = oComm.Parameters("@nResp").Value
                                sRespuesta = oComm.Parameters("@Resultado").Value
                                id_ResumenComportamientoPago = oComm.Parameters("@id_ResumenComportamientoPago").Value
                                If nResp = 0 Then
                                    D(0) = Str(nResp)
                                    D(1) = sRespuesta
                                    R.nResultado = CInt(D(0))
                                    R.sRespuesta = D(1)
                                    Return R
                                End If

                                '*SEMAFORO
                                j = 0
                                If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.SistemaFinanciero.ResumenComportamientoPago.Semaforo IsNot Nothing Then
                                    For j = 0 To oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.SistemaFinanciero.ResumenComportamientoPago.Semaforo.Count - 1
                                        Dim semaforo As New Semaforo
                                        semaforo = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.SistemaFinanciero.ResumenComportamientoPago.Semaforo(j)

                                        oComm = New SqlCommand("[dbo].[PA_RegistrarSemaforo]", oConex)
                                        oComm.CommandType = CommandType.StoredProcedure
                                        oComm.Parameters.Add("@id_ResumenComportamientoPago", SqlDbType.Int).Value = id_ResumenComportamientoPago
                                        oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial

                                        oComm.Parameters.Add("@periodo", SqlDbType.VarChar).Value = semaforo.periodo.NullToEmpty()
                                        oComm.Parameters.Add("@NoTieneImpagos", SqlDbType.Bit).Value = semaforo.NoTieneImpagos
                                        oComm.Parameters.Add("@TieneDeudasAtrasadas", SqlDbType.Bit).Value = semaforo.TieneDeudasAtrasadas
                                        oComm.Parameters.Add("@TieneDeudasImpagasInfocorp", SqlDbType.Bit).Value = semaforo.TieneDeudasImpagasInfocorp
                                        oComm.Parameters.Add("@InformacionNoDisponible", SqlDbType.Bit).Value = semaforo.InformacionNoDisponible
                                        oComm.Parameters.Add("@DiasAtraso", SqlDbType.Int).Value = semaforo.DiasAtraso.NullToZero()

                                        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                        oComm.ExecuteNonQuery()
                                        nResp = oComm.Parameters("@nResp").Value
                                        sRespuesta = oComm.Parameters("@Resultado").Value

                                        If nResp = 0 Then
                                            D(0) = Str(nResp)
                                            D(1) = sRespuesta
                                            R.nResultado = CInt(D(0))
                                            R.sRespuesta = D(1)
                                            Return R
                                        End If
                                    Next
                                End If

                            End If

                            '2
                            Dim id_DeudasHistoricas As Integer
                            If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.SistemaFinanciero.DeudasHistoricas IsNot Nothing Then
                                oComm = New SqlCommand("[dbo].[PA_RegistrarDeudasHistoricas]", oConex)
                                oComm.CommandType = CommandType.StoredProcedure
                                oComm.Parameters.Add("@id_SistemaFinancieroReguladoYNoRegulado", SqlDbType.Int).Value = id_SistemaFinancieroReguladoYNoRegulado
                                oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial

                                oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                oComm.Parameters.Add("@id_DeudasHistoricas", SqlDbType.Int).Direction = ParameterDirection.Output
                                oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                oComm.ExecuteNonQuery()
                                nResp = oComm.Parameters("@nResp").Value
                                sRespuesta = oComm.Parameters("@Resultado").Value
                                id_DeudasHistoricas = oComm.Parameters("@id_DeudasHistoricas").Value
                                If nResp = 0 Then
                                    D(0) = Str(nResp)
                                    D(1) = sRespuesta
                                    R.nResultado = CInt(D(0))
                                    R.sRespuesta = D(1)
                                    Return R
                                End If


                                '3
                                j = 0
                                If oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.SistemaFinanciero.DeudasHistoricas.Deuda IsNot Nothing Then
                                    For j = 0 To oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.SistemaFinanciero.DeudasHistoricas.Deuda.Count - 1
                                        Dim deudasHistoricasDeuda As New DeudasHistoricasDeuda
                                        deudasHistoricasDeuda = oDataEquifax.applicants.primaryConsumer.interconnectResponse(i).Data.SistemaFinanciero.DeudasHistoricas.Deuda(j)

                                        oComm = New SqlCommand("[dbo].[PA_RegistrarDeudasHistoricas_Deuda]", oConex)
                                        oComm.CommandType = CommandType.StoredProcedure
                                        oComm.Parameters.Add("@id_DeudasHistoricas", SqlDbType.Int).Value = id_DeudasHistoricas
                                        oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_inicial

                                        oComm.Parameters.Add("@periodo", SqlDbType.VarChar).Value = deudasHistoricasDeuda.periodo.NullToEmpty()
                                        oComm.Parameters.Add("@Calificacion", SqlDbType.VarChar).Value = deudasHistoricasDeuda.Calificacion.NullToEmpty()
                                        oComm.Parameters.Add("@Porcentaje", SqlDbType.VarChar).Value = deudasHistoricasDeuda.Porcentaje.NullToEmpty()
                                        oComm.Parameters.Add("@NroEntidades", SqlDbType.Int).Value = deudasHistoricasDeuda.NroEntidades.NullToZero()
                                        oComm.Parameters.Add("@DeudaVigente", SqlDbType.Decimal).Value = deudasHistoricasDeuda.DeudaVigente
                                        oComm.Parameters.Add("@DeudaAtrasada", SqlDbType.Decimal).Value = deudasHistoricasDeuda.DeudaAtrasada
                                        oComm.Parameters.Add("@DeudaVencida", SqlDbType.Decimal).Value = deudasHistoricasDeuda.DeudaVencida
                                        oComm.Parameters.Add("@DeudaRefinanciada", SqlDbType.Decimal).Value = deudasHistoricasDeuda.DeudaRefinanciada
                                        oComm.Parameters.Add("@DeudaReestructurada", SqlDbType.Decimal).Value = deudasHistoricasDeuda.DeudaReestructurada
                                        oComm.Parameters.Add("@DeudaJudicial", SqlDbType.Decimal).Value = deudasHistoricasDeuda.DeudaJudicial
                                        oComm.Parameters.Add("@DeudaCastigada", SqlDbType.Decimal).Value = deudasHistoricasDeuda.DeudaCastigada
                                        oComm.Parameters.Add("@DeudaTotal", SqlDbType.Decimal).Value = deudasHistoricasDeuda.DeudaTotal
                                        oComm.Parameters.Add("@DiasAtraso", SqlDbType.Int).Value = deudasHistoricasDeuda.DiasAtraso

                                        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar).Value = sCodUsu
                                        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
                                        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
                                        oComm.ExecuteNonQuery()
                                        nResp = oComm.Parameters("@nResp").Value
                                        sRespuesta = oComm.Parameters("@Resultado").Value

                                        If nResp = 0 Then
                                            D(0) = Str(nResp)
                                            D(1) = sRespuesta
                                            R.nResultado = CInt(D(0))
                                            R.sRespuesta = D(1)
                                            Return R
                                        End If
                                    Next
                                End If

                            End If

                        End If
                        '*******************************


                    Next


                End If

            End If

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

    Public Function ValidaDocumentoEquifax() As outNegocio
        Dim R As New outNegocio
        Dim AbonoDesemb() As dAbonoDesemb
        Dim ListaDesembolsador() As String
        Dim i As Integer
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlCommand
        Dim dr As SqlDataReader

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Try
            oComm = New SqlClient.SqlCommand("PA_ValidaDocumentoEquifaxLocal", oConex)
            oComm.Parameters.Add("@TipoDoc", SqlDbType.VarChar, 50).Value = "1"
            oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = cDocPersona

            oComm.Parameters.Add("@psResult", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@pnResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()
            D(0) = Str(oComm.Parameters("@pnResp").Value).Trim
            D(1) = oComm.Parameters("@psResult").Value
            oConex.Close()
        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message
        End Try
        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))
        Return R

    End Function

    Public Function RecuperarEquifax() As outNegocio
        Dim R As New outNegocio
        Dim AbonoDesemb() As dAbonoDesemb
        Dim ListaDesembolsador() As String
        Dim i, k, j As Integer
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlCommand
        Dim dr As SqlDataReader
        Dim id_ip As Integer

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Try

            oComm = New SqlCommand("[dbo].[PA_ObtenerInformacionEquifax]", oConex)
            oComm.Parameters.Add("@TipoDoc", SqlDbType.VarChar, 5).Value = "1"
            oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 10).Value = cDocPersona
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oCatalogocodigos_2(0)
                R.oCatalogocodigos_2(0) = New clsCatalogocodigos
                'R.oCatalogocodigos_2(0).cDescripcion = dr.Item("id")
                'R.oCatalogocodigos_2(0).cValor = dr.Item("tipoDocumento")
                id_ip = dr.Item("id_ip")
                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            R.equifaxInformacion = New ResponseEquifaxBasic
            '*********************** 
            'RESUMEN FINANCIERO: SISTEMA FINANCIERO (SBS) Y NO REGULADO (MICROFINANZAS) 859
            Dim id_ResumenFinanciero As Integer = 0
            oComm = New SqlCommand("[dbo].[PA_ObtenerResumenFinanciero]", oConex)
            oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read
                R.equifaxInformacion.ResumenFinanciero = New ResumenFinanciero
                id_ResumenFinanciero = dr.Item("id_ResumenFinanciero")
                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            If id_ResumenFinanciero > 0 Then
                oComm = New SqlCommand("[dbo].[PA_ObtenerDeudasUltimoPeriodo]", oConex)
                oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.ExecuteNonQuery()
                dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                i = 0
                While dr.Read
                    R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo = New DeudasUltimoPeriodo
                    R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.periodo = dr.Item("Periodo")
                    R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.MesActual = dr.Item("MesActual")
                    R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.MesAnterior = dr.Item("MesAnterior")
                    R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.AnioAnterior = dr.Item("AnioAnterior")
                    i = i + 1
                End While
                dr.Close()
                dr = Nothing

                oComm = New SqlCommand("[dbo].[PA_ObtenerDeudasUltimoPeriodo_Deudas]", oConex)
                oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.ExecuteNonQuery()
                dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                i = 0
                While dr.Read
                    ReDim Preserve R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Deudas(i)
                    R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Deudas(i) = New DeudaUP
                    R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Deudas(i).id_DeudasUltimoPeriodo_Deudas = dr.Item("id_DeudasUltimoPeriodo_Deudas")
                    R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Deudas(i).Entidad = dr.Item("Entidad")
                    R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Deudas(i).SistemaFinanciero = dr.Item("SistemaFinanciero")
                    R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Deudas(i).CalificacionMesActual = dr.Item("CalificacionMesActual")
                    R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Deudas(i).MontoSubTotalMesActual = CDbl(dr.Item("MontoSubTotalMesActual"))
                    R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Deudas(i).CalificacionMesAnterior = dr.Item("CalificacionMesAnterior")
                    R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Deudas(i).MontoSubTotalMesAnterior = CDbl(dr.Item("MontoSubTotalMesAnterior"))
                    R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Deudas(i).CalificacionAnioAnterior = dr.Item("CalificacionAnioAnterior")
                    R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Deudas(i).MontoSubTotalAnioAnterior = CDbl(dr.Item("MontoSubTotalAnioAnterior"))
                    i = i + 1
                End While
                dr.Close()
                dr = Nothing


                k = 0
                i = 0
                j = 0
                For i = 0 To R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Deudas.Count - 1

                    oComm = New SqlCommand("[dbo].[PA_ObtenerDeudasUltimoPeriodo_Deudas_Producto]", oConex)
                    oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
                    oComm.Parameters.Add("@id_DeudasUltimoPeriodo_Deudas", SqlDbType.Int).Value = R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Deudas(i).id_DeudasUltimoPeriodo_Deudas
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()
                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    k = 0
                    While dr.Read
                        ReDim Preserve R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Deudas(i).Productos(k)
                        R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Deudas(i).Productos(k) = New Producto
                        R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Deudas(i).Productos(k).Tipo = dr.Item("Tipo")
                        R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Deudas(i).Productos(k).Descripcion = dr.Item("Descripcion")
                        R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Deudas(i).Productos(k).MontoMesActual = CDbl(dr.Item("MontoMesActual"))
                        R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Deudas(i).Productos(k).MontoMesAnterior = CDbl(dr.Item("MontoMesAnterior"))
                        R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Deudas(i).Productos(k).MontoAnioAnterior = CDbl(dr.Item("MontoAnioAnterior"))
                        R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Deudas(i).Productos(k).DiasAtraso = dr.Item("DiasAtraso")
                        k = k + 1
                    End While
                    dr.Close()
                    dr = Nothing

                Next

                oComm = New SqlCommand("[dbo].[PA_ObtenerDeudasUltimoPeriodo_Totales]", oConex)
                oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.ExecuteNonQuery()
                dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                i = 0
                While dr.Read
                    R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Totales = New Totales
                    R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Totales.MontoTotalMesActual = CDbl(dr.Item("MontoTotalMesActual"))
                    R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Totales.MontoTotalMesAnterior = CDbl(dr.Item("MontoTotalMesAnterior"))
                    R.equifaxInformacion.ResumenFinanciero.DeudasUltimoPeriodo.Totales.MontoTotalAnioAnterior = CDbl(dr.Item("MontoTotalAnioAnterior"))
                    i = i + 1
                End While
                dr.Close()
                dr = Nothing

            End If
            '**************

            '*********************************************************
            'SISTEMA FINANCIERO REGULADO (SBS) Y NO REGULADO (MICROFINANZAS) 873
            Dim id_SistemaFinancieroReguladoYNoRegulado As Integer = 0
            oComm = New SqlCommand("[dbo].[PA_Obtener_SistemaFinancieroReguladoYNoRegulado]", oConex)
            oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read
                R.equifaxInformacion.SistemaFinanciero = New SistemaFinancieroSYNRegulado
                id_SistemaFinancieroReguladoYNoRegulado = dr.Item("id_SistemaFinancieroReguladoYNoRegulado")
                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            Dim id_ResumenComportamientoPago As Integer = 0
            If id_SistemaFinancieroReguladoYNoRegulado > 0 Then
                oComm = New SqlCommand("[dbo].[PA_Obtener_ResumenComportamientoPago]", oConex)
                oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.ExecuteNonQuery()
                dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                i = 0
                While dr.Read
                    R.equifaxInformacion.SistemaFinanciero.ResumenComportamientoPago = New ResumenComportamientoPago
                    id_ResumenComportamientoPago = dr.Item("id_ResumenComportamientoPago")
                    i = i + 1
                End While
                dr.Close()
                dr = Nothing

                If id_ResumenComportamientoPago > 0 Then

                    oComm = New SqlCommand("[dbo].[PA_Obtener_ResumenComportamientoPago_Semaforo]", oConex)
                    oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()
                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0
                    While dr.Read
                        ReDim Preserve R.equifaxInformacion.SistemaFinanciero.ResumenComportamientoPago.Semaforo(i)
                        R.equifaxInformacion.SistemaFinanciero.ResumenComportamientoPago.Semaforo(i) = New Semaforo
                        R.equifaxInformacion.SistemaFinanciero.ResumenComportamientoPago.Semaforo(i).periodo = dr.Item("periodo")
                        R.equifaxInformacion.SistemaFinanciero.ResumenComportamientoPago.Semaforo(i).NoTieneImpagos = dr.Item("NoTieneImpagos")
                        R.equifaxInformacion.SistemaFinanciero.ResumenComportamientoPago.Semaforo(i).TieneDeudasAtrasadas = dr.Item("TieneDeudasAtrasadas")
                        R.equifaxInformacion.SistemaFinanciero.ResumenComportamientoPago.Semaforo(i).TieneDeudasImpagasInfocorp = dr.Item("TieneDeudasImpagasInfocorp")
                        R.equifaxInformacion.SistemaFinanciero.ResumenComportamientoPago.Semaforo(i).InformacionNoDisponible = dr.Item("InformacionNoDisponible")
                        R.equifaxInformacion.SistemaFinanciero.ResumenComportamientoPago.Semaforo(i).DiasAtraso = dr.Item("DiasAtraso")
                        i = i + 1
                    End While
                    dr.Close()
                    dr = Nothing

                End If


                '2
                Dim id_DeudasHistoricas As Integer = 0
                oComm = New SqlCommand("[dbo].[PA_Obtener_DeudasHistoricas]", oConex)
                oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.ExecuteNonQuery()
                dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                i = 0
                While dr.Read
                    R.equifaxInformacion.SistemaFinanciero.DeudasHistoricas = New DeudasHistoricas
                    id_DeudasHistoricas = dr.Item("id_DeudasHistoricas")
                    i = i + 1
                End While
                dr.Close()
                dr = Nothing

                If id_DeudasHistoricas > 0 Then

                    oComm = New SqlCommand("[dbo].[PA_Obtener_DeudasHistoricas_Deuda]", oConex)
                    oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()
                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0
                    While dr.Read
                        ReDim Preserve R.equifaxInformacion.SistemaFinanciero.DeudasHistoricas.Deuda(i)
                        R.equifaxInformacion.SistemaFinanciero.DeudasHistoricas.Deuda(i) = New DeudasHistoricasDeuda
                        R.equifaxInformacion.SistemaFinanciero.DeudasHistoricas.Deuda(i).periodo = dr.Item("periodo")
                        R.equifaxInformacion.SistemaFinanciero.DeudasHistoricas.Deuda(i).Calificacion = dr.Item("Calificacion")
                        R.equifaxInformacion.SistemaFinanciero.DeudasHistoricas.Deuda(i).Porcentaje = dr.Item("Porcentaje")
                        R.equifaxInformacion.SistemaFinanciero.DeudasHistoricas.Deuda(i).NroEntidades = dr.Item("NroEntidades")
                        R.equifaxInformacion.SistemaFinanciero.DeudasHistoricas.Deuda(i).DeudaVigente = CDbl(dr.Item("DeudaVigente"))
                        R.equifaxInformacion.SistemaFinanciero.DeudasHistoricas.Deuda(i).DeudaAtrasada = CDbl(dr.Item("DeudaAtrasada"))
                        R.equifaxInformacion.SistemaFinanciero.DeudasHistoricas.Deuda(i).DeudaVencida = CDbl(dr.Item("DeudaVencida"))
                        R.equifaxInformacion.SistemaFinanciero.DeudasHistoricas.Deuda(i).DeudaRefinanciada = CDbl(dr.Item("DeudaRefinanciada"))
                        R.equifaxInformacion.SistemaFinanciero.DeudasHistoricas.Deuda(i).DeudaReestructurada = CDbl(dr.Item("DeudaReestructurada"))
                        R.equifaxInformacion.SistemaFinanciero.DeudasHistoricas.Deuda(i).DeudaJudicial = CDbl(dr.Item("DeudaJudicial"))
                        R.equifaxInformacion.SistemaFinanciero.DeudasHistoricas.Deuda(i).DeudaCastigada = CDbl(dr.Item("DeudaCastigada"))
                        R.equifaxInformacion.SistemaFinanciero.DeudasHistoricas.Deuda(i).DeudaTotal = CDbl(dr.Item("DeudaTotal"))
                        R.equifaxInformacion.SistemaFinanciero.DeudasHistoricas.Deuda(i).DiasAtraso = dr.Item("DiasAtraso")

                        i = i + 1
                    End While
                    dr.Close()
                    dr = Nothing

                End If


            End If
            '**********************************************
            '*********************************************************
            'Resumen Flag 865
            Dim id_ResumenFlags As Integer = 0
            oComm = New SqlCommand("[dbo].[PA_Obtener_ResumenFlags]", oConex)
            oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read
                R.equifaxInformacion.ResumenFlags = New ResumenFlags
                id_ResumenFlags = dr.Item("id_ResumenFlags")
                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            Dim id_ResumenComportamiento As Integer = 0
            If id_ResumenFlags > 0 Then
                oComm = New SqlCommand("[dbo].[PA_Obtener_ResumenComportamiento]", oConex)
                oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.ExecuteNonQuery()
                dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                i = 0
                While dr.Read
                    R.equifaxInformacion.ResumenFlags.ResumenComportamiento = New ResumenComportamiento
                    R.equifaxInformacion.ResumenFlags.ResumenComportamiento.TipoDocumento = dr.Item("TipoDocumento")
                    R.equifaxInformacion.ResumenFlags.ResumenComportamiento.NumeroDocumento = dr.Item("NumeroDocumento")
                    id_ResumenComportamiento = dr.Item("id_ResumenComportamiento")
                    i = i + 1
                End While
                dr.Close()
                dr = Nothing


                oComm = New SqlCommand("[dbo].[PA_Obtener_ResumenComportamiento_ResumenDeuda]", oConex)
                oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.ExecuteNonQuery()
                dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                i = 0
                While dr.Read
                    R.equifaxInformacion.ResumenFlags.ResumenComportamiento.ResumenDeuda = New ResumenDeuda
                    R.equifaxInformacion.ResumenFlags.ResumenComportamiento.ResumenDeuda.Periodo = dr.Item("Periodo")
                    R.equifaxInformacion.ResumenFlags.ResumenComportamiento.ResumenDeuda.DeudaTotal = (dr.Item("DeudaTotal"))
                    R.equifaxInformacion.ResumenFlags.ResumenComportamiento.ResumenDeuda.DeudaDirecta = (dr.Item("DeudaDirecta"))
                    R.equifaxInformacion.ResumenFlags.ResumenComportamiento.ResumenDeuda.DeudaIndirecta = (dr.Item("DeudaIndirecta"))
                    R.equifaxInformacion.ResumenFlags.ResumenComportamiento.ResumenDeuda.PorcentajeDeudaNormal = dr.Item("PorcentajeDeudaNormal")
                    R.equifaxInformacion.ResumenFlags.ResumenComportamiento.ResumenDeuda.PorcentajeDeudaPotencial = dr.Item("PorcentajeDeudaPotencial")
                    R.equifaxInformacion.ResumenFlags.ResumenComportamiento.ResumenDeuda.PorcentajeDeudaDeficiente = dr.Item("PorcentajeDeudaDeficiente")
                    R.equifaxInformacion.ResumenFlags.ResumenComportamiento.ResumenDeuda.PorcentajeDeudaEnRiesgo = dr.Item("PorcentajeDeudaEnRiesgo")
                    R.equifaxInformacion.ResumenFlags.ResumenComportamiento.ResumenDeuda.PorcentajeDeudaPerdida = dr.Item("PorcentajeDeudaPerdida")
                    R.equifaxInformacion.ResumenFlags.ResumenComportamiento.ResumenDeuda.Variacion = dr.Item("Variacion")
                    i = i + 1
                End While
                dr.Close()
                dr = Nothing

                Dim id_ResumenScoreHistorico As Integer = 0
                oComm = New SqlCommand("[dbo].[PA_Obtener_ResumenComportamiento_ResumenScoreHistorico]", oConex)
                oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.ExecuteNonQuery()
                dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                i = 0
                While dr.Read
                    R.equifaxInformacion.ResumenFlags.ResumenComportamiento.ResumenScoreHistorico = New ResumenScoreHistorico
                    id_ResumenScoreHistorico = dr.Item("id_ResumenScoreHistorico")
                    i = i + 1
                End While
                dr.Close()
                dr = Nothing

                If id_ResumenScoreHistorico > 0 Then
                    oComm = New SqlCommand("[dbo].[PA_Obtener_ResumenComportamiento_ResumenScoreHistorico_ScoreActual]", oConex)
                    oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()
                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0
                    While dr.Read
                        R.equifaxInformacion.ResumenFlags.ResumenComportamiento.ResumenScoreHistorico.ScoreActual = New ScoreData
                        R.equifaxInformacion.ResumenFlags.ResumenComportamiento.ResumenScoreHistorico.ScoreActual.Periodo = dr.Item("Periodo")
                        R.equifaxInformacion.ResumenFlags.ResumenComportamiento.ResumenScoreHistorico.ScoreActual.Riesgo = dr.Item("Riesgo")
                        R.equifaxInformacion.ResumenFlags.ResumenComportamiento.ResumenScoreHistorico.ScoreActual.MotivoSinScore = dr.Item("MotivoSinScore")
                        i = i + 1
                    End While
                    dr.Close()
                    dr = Nothing

                    oComm = New SqlCommand("[dbo].[PA_Obtener_ResumenComportamiento_ResumenScoreHistorico_ScoreAnterior]", oConex)
                    oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()
                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0
                    While dr.Read
                        R.equifaxInformacion.ResumenFlags.ResumenComportamiento.ResumenScoreHistorico.ScoreAnterior = New ScoreData
                        R.equifaxInformacion.ResumenFlags.ResumenComportamiento.ResumenScoreHistorico.ScoreAnterior.Periodo = dr.Item("Periodo")
                        R.equifaxInformacion.ResumenFlags.ResumenComportamiento.ResumenScoreHistorico.ScoreAnterior.Riesgo = dr.Item("Riesgo")
                        R.equifaxInformacion.ResumenFlags.ResumenComportamiento.ResumenScoreHistorico.ScoreAnterior.MotivoSinScore = dr.Item("MotivoSinScore")
                        i = i + 1
                    End While
                    dr.Close()
                    dr = Nothing

                    oComm = New SqlCommand("[dbo].[PA_Obtener_ResumenComportamiento_ResumenScoreHistorico_ScoreHace12Meses]", oConex)
                    oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()
                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0
                    While dr.Read
                        R.equifaxInformacion.ResumenFlags.ResumenComportamiento.ResumenScoreHistorico.ScoreHace12Meses = New ScoreData
                        R.equifaxInformacion.ResumenFlags.ResumenComportamiento.ResumenScoreHistorico.ScoreHace12Meses.Periodo = dr.Item("Periodo")
                        R.equifaxInformacion.ResumenFlags.ResumenComportamiento.ResumenScoreHistorico.ScoreHace12Meses.Riesgo = dr.Item("Riesgo")
                        R.equifaxInformacion.ResumenFlags.ResumenComportamiento.ResumenScoreHistorico.ScoreHace12Meses.MotivoSinScore = dr.Item("MotivoSinScore")
                        i = i + 1
                    End While
                    dr.Close()
                    dr = Nothing

                End If


                oComm = New SqlCommand("[dbo].[PA_Obtener_ResumenBloqueFlags]", oConex)
                oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.ExecuteNonQuery()
                dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                i = 0
                While dr.Read
                    R.equifaxInformacion.ResumenFlags.ResumenBloqueFlags = New ResumenBloqueFlags
                    R.equifaxInformacion.ResumenFlags.ResumenBloqueFlags.TarjetaCredito = dr.Item("TarjetaCredito")
                    R.equifaxInformacion.ResumenFlags.ResumenBloqueFlags.LineaDeCredito = dr.Item("LineaDeCredito")
                    R.equifaxInformacion.ResumenFlags.ResumenBloqueFlags.CreditoHipotecario = dr.Item("CreditoHipotecario")
                    R.equifaxInformacion.ResumenFlags.ResumenBloqueFlags.BuenPagadorDeServicios = dr.Item("BuenPagadorDeServicios")
                    R.equifaxInformacion.ResumenFlags.ResumenBloqueFlags.EstaEnInfocorp = dr.Item("EstaEnInfocorp")
                    R.equifaxInformacion.ResumenFlags.ResumenBloqueFlags.AvalAvalado = dr.Item("AvalAvalado")
                    R.equifaxInformacion.ResumenFlags.ResumenBloqueFlags.RepresentanteLegal = dr.Item("RepresentanteLegal")
                    R.equifaxInformacion.ResumenFlags.ResumenBloqueFlags.GastoMensualEstimado = dr.Item("GastoMensualEstimado")
                    R.equifaxInformacion.ResumenFlags.ResumenBloqueFlags.PosibleRestringido = dr.Item("PosibleRestringido")
                    R.equifaxInformacion.ResumenFlags.ResumenBloqueFlags.TieneAuto = dr.Item("TieneAuto")
                    R.equifaxInformacion.ResumenFlags.ResumenBloqueFlags.EntidadesQueConsultaron = dr.Item("EntidadesQueConsultaron")
                    R.equifaxInformacion.ResumenFlags.ResumenBloqueFlags.Homonimos = dr.Item("Homonimos")
                    R.equifaxInformacion.ResumenFlags.ResumenBloqueFlags.ComercioExterior = dr.Item("ComercioExterior")
                    R.equifaxInformacion.ResumenFlags.ResumenBloqueFlags.DeudaPrevisional = dr.Item("DeudaPrevisional")
                    R.equifaxInformacion.ResumenFlags.ResumenBloqueFlags.AlertaPep = dr.Item("AlertaPep")
                    R.equifaxInformacion.ResumenFlags.ResumenBloqueFlags.AlertaRedam = dr.Item("AlertaRedam")
                    R.equifaxInformacion.ResumenFlags.ResumenBloqueFlags.ReactivaPeru = dr.Item("ReactivaPeru")
                    R.equifaxInformacion.ResumenFlags.ResumenBloqueFlags.EsInquilinoMoroso = dr.Item("EsInquilinoMoroso")

                    'id_ResumenScoreHistorico = dr.Item("id_ResumenScoreHistorico")
                    i = i + 1
                End While
                dr.Close()
                dr = Nothing

                oComm = New SqlCommand("[dbo].[PA_Obtener_ResumenBloqueFlags_ReactivaPeruInfo]", oConex)
                oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.ExecuteNonQuery()
                dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                i = 0
                While dr.Read
                    R.equifaxInformacion.ResumenFlags.ResumenBloqueFlags.ReactivaPeruInfo = New ReactivaPeruInfo
                    R.equifaxInformacion.ResumenFlags.ResumenBloqueFlags.ReactivaPeruInfo.Fecha = dr.Item("Fecha")
                    R.equifaxInformacion.ResumenFlags.ResumenBloqueFlags.ReactivaPeruInfo.Monto = dr.Item("Monto")
                    'id_ResumenScoreHistorico = dr.Item("id_ResumenScoreHistorico")
                    i = i + 1
                End While
                dr.Close()
                dr = Nothing

            End If
            '*************************************************************
            '*********************************************************
            'SCORE PREDICTIVO CON VARIABLES 822
            Dim id_ResumenScoreRP3_Cabecera As Integer = 0
            oComm = New SqlCommand("[dbo].[PA_Obtener_ResumenScoreRP3_Cabecera]", oConex)
            oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read
                'R.equifaxInformacion.ResumenScoreRP3 = New ResumenScoreRP3
                id_ResumenScoreRP3_Cabecera = dr.Item("id_ResumenScoreRP3_Cabecera")
                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            If id_ResumenScoreRP3_Cabecera > 0 Then
                oComm = New SqlCommand("[dbo].[PA_Obtener_ResumenScoreRP3]", oConex)
                oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.ExecuteNonQuery()
                dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                i = 0
                While dr.Read
                    R.equifaxInformacion.ResumenScoreRP3 = New ResumenScoreRP3
                    R.equifaxInformacion.ResumenScoreRP3.Puntaje = dr.Item("Puntaje")
                    R.equifaxInformacion.ResumenScoreRP3.NivelRiesgo = dr.Item("NivelRiesgo")
                    R.equifaxInformacion.ResumenScoreRP3.Conclusion = dr.Item("Conclusion")
                    i = i + 1
                End While
                dr.Close()
                dr = Nothing

                Dim id_PrincipalesVariables As Integer = 0
                oComm = New SqlCommand("[dbo].[PA_Obtener_ResumenScoreRP3_PrincipalesVariables]", oConex)
                oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.ExecuteNonQuery()
                dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                i = 0
                While dr.Read
                    R.equifaxInformacion.ResumenScoreRP3.PrincipalesVariables = New PrincipalesVariables
                    id_PrincipalesVariables = dr.Item("id_PrincipalesVariables")
                    i = i + 1
                End While
                dr.Close()
                dr = Nothing

                If id_PrincipalesVariables > 0 Then
                    oComm = New SqlCommand("[dbo].[PA_Obtener_ResumenScoreRP3_PrincipalesVariables_Variables]", oConex)
                    oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
                    oComm.Parameters.Add("@nTipo", SqlDbType.Int).Value = 1
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()
                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0
                    While dr.Read
                        R.equifaxInformacion.ResumenScoreRP3.PrincipalesVariables.PresentaInformacionSFR = New VariableValor
                        R.equifaxInformacion.ResumenScoreRP3.PrincipalesVariables.PresentaInformacionSFR.Variable = dr.Item("Variable")
                        R.equifaxInformacion.ResumenScoreRP3.PrincipalesVariables.PresentaInformacionSFR.Valor = dr.Item("Valor")
                        i = i + 1
                    End While
                    dr.Close()
                    dr = Nothing

                    oComm = New SqlCommand("[dbo].[PA_Obtener_ResumenScoreRP3_PrincipalesVariables_Variables]", oConex)
                    oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
                    oComm.Parameters.Add("@nTipo", SqlDbType.Int).Value = 2
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()
                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0
                    While dr.Read
                        R.equifaxInformacion.ResumenScoreRP3.PrincipalesVariables.PresentaDeudas = New VariableValor
                        R.equifaxInformacion.ResumenScoreRP3.PrincipalesVariables.PresentaDeudas.Variable = dr.Item("Variable")
                        R.equifaxInformacion.ResumenScoreRP3.PrincipalesVariables.PresentaDeudas.Valor = dr.Item("Valor")
                        i = i + 1
                    End While
                    dr.Close()
                    dr = Nothing

                    oComm = New SqlCommand("[dbo].[PA_Obtener_ResumenScoreRP3_PrincipalesVariables_Variables]", oConex)
                    oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
                    oComm.Parameters.Add("@nTipo", SqlDbType.Int).Value = 3
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()
                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0
                    While dr.Read
                        R.equifaxInformacion.ResumenScoreRP3.PrincipalesVariables.InformacionReportadaInfocorp = New VariableValor
                        R.equifaxInformacion.ResumenScoreRP3.PrincipalesVariables.InformacionReportadaInfocorp.Variable = dr.Item("Variable")
                        R.equifaxInformacion.ResumenScoreRP3.PrincipalesVariables.InformacionReportadaInfocorp.Valor = dr.Item("Valor")
                        i = i + 1
                    End While
                    dr.Close()
                    dr = Nothing

                    oComm = New SqlCommand("[dbo].[PA_Obtener_ResumenScoreRP3_PrincipalesVariables_Variables]", oConex)
                    oComm.Parameters.Add("@id_ip", SqlDbType.Int).Value = id_ip
                    oComm.Parameters.Add("@nTipo", SqlDbType.Int).Value = 4
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()
                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0
                    While dr.Read
                        R.equifaxInformacion.ResumenScoreRP3.PrincipalesVariables.MalComportamientoCrediticio = New VariableValor
                        R.equifaxInformacion.ResumenScoreRP3.PrincipalesVariables.MalComportamientoCrediticio.Variable = dr.Item("Variable")
                        R.equifaxInformacion.ResumenScoreRP3.PrincipalesVariables.MalComportamientoCrediticio.Valor = dr.Item("Valor")
                        i = i + 1
                    End While
                    dr.Close()
                    dr = Nothing

                End If

            End If


            '*********************************************************




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
    'END CONSULTA EQUIFAX 



    Public Function RecuperaCanceladosHoy() As outNegocio
        Dim R As New outNegocio
        Dim AbonoDesemb() As dAbonoDesemb
        Dim ListaDesembolsador() As String
        Dim i As Integer
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


            oComm = New SqlClient.SqlCommand("PA_RecuperaCreditosCanceladosEnElDia", oConex)
            oComm.Parameters.Add("@cCodUsuCob", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                ReDim Preserve R.oBandejaCancelados(i)
                R.oBandejaCancelados(i) = New clsCanceladosHoy

                R.oBandejaCancelados(i).dFecha = Format(dr.Item("dFecha"), "dd/MM/yyyy")
                R.oBandejaCancelados(i).ncodcred = Str(dr.Item("ncodcred"))
                R.oBandejaCancelados(i).cDoc = dr.Item("cDoc")
                R.oBandejaCancelados(i).cCliente = dr.Item("cCliente")
                R.oBandejaCancelados(i).cMovil1 = dr.Item("cMovil1")
                R.oBandejaCancelados(i).nPrestamo = Format(dr.Item("nPrestamo"), "#,#0.00")
                R.oBandejaCancelados(i).cDireccion = dr.Item("cDireccion")
                R.oBandejaCancelados(i).nDiasAtraso = Str(dr.Item("nDiasAtraso"))

                i = i + 1

            End While
            dr.Close()
            dr = Nothing


            oConex.Close()

            D(0) = "1"
            D(1) = Str(i) + " Registros Encontrados"

        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message

        End Try


        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))

        Return R

    End Function

    Public Function RecuperaExperianData() As outNegocio
        Dim R As New outNegocio
        Dim nID As Long
        Dim i As Integer
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
            cDocPersona = cDoc
            ConsultaDocumentoExperian()


            oComm = New SqlClient.SqlCommand("PA_RecuperaExperianCabecera", oConex)
            oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = cDoc
            oComm.Parameters.Add("@cApePat", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@cApeMat", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@cNombres", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nID", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nScore", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@cNivScore", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@cCapPago", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@cResultado", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@cCalif", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@cSemaforo", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nNumEnt", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nDeudaTotal", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@dFechaProc", SqlDbType.Date).Direction = ParameterDirection.Output


            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()
            i = 0

            ReDim Preserve R.oBandejaExperian(i)
            R.oBandejaExperian(i) = New clsExperian
            R.oBandejaExperian(i).cDoc = cDoc
            R.oBandejaExperian(i).cNomApe = oComm.Parameters("@cApePat").Value + " " + oComm.Parameters("@cApeMat").Value + " " + oComm.Parameters("@cNombres").Value
            R.oBandejaExperian(i).nScore = oComm.Parameters("@nScore").Value
            R.oBandejaExperian(i).cNivScore = oComm.Parameters("@cNivScore").Value
            R.oBandejaExperian(i).cCapPago = oComm.Parameters("@cCapPago").Value
            R.oBandejaExperian(i).cResultado = oComm.Parameters("@cResultado").Value
            R.oBandejaExperian(i).cCalif = oComm.Parameters("@cCalif").Value
            R.oBandejaExperian(i).cSemaforo = oComm.Parameters("@cSemaforo").Value
            R.oBandejaExperian(i).nNumEnt = oComm.Parameters("@nNumEnt").Value
            R.oBandejaExperian(i).nDeudaTotal = oComm.Parameters("@nDeudaTotal").Value
            R.oBandejaExperian(i).dFechaProc = oComm.Parameters("@dFechaProc").Value
            nID = oComm.Parameters("@nID").Value


            oComm = New SqlClient.SqlCommand("PA_RecuperaExperianDetalle", oConex)
            oComm.Parameters.Add("@nID", SqlDbType.VarChar, 50).Value = nID


            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

            i = 0

            While dr.Read
                ReDim Preserve R.oBandejaExperian(0).oEntidadesDetalle(i)
                R.oBandejaExperian(0).oEntidadesDetalle(i) = New clsExperianEntDet

                R.oBandejaExperian(0).oEntidadesDetalle(i).cNomEnt = dr.Item("NomEnt")
                R.oBandejaExperian(0).oEntidadesDetalle(i).nSaldoDeuda = CDbl(dr.Item("SalDeu"))
                R.oBandejaExperian(0).oEntidadesDetalle(i).cCalif = dr.Item("Cal")
                R.oBandejaExperian(0).oEntidadesDetalle(i).nDiasVenc = CInt(dr.Item("DiaVen"))
                R.oBandejaExperian(0).oEntidadesDetalle(i).dFecRepo = CDate(dr.Item("FchRep"))
                R.oBandejaExperian(0).oEntidadesDetalle(i).dFecProc = CDate(dr.Item("FchPro"))


                i = i + 1

            End While
            dr.Close()
            dr = Nothing


            oConex.Close()

            D(0) = "1"
            D(1) = Str(i) + " Registros Encontrados"

        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message

        End Try


        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))

        Return R

    End Function

    Public Function RecuperaConsultaSolicitudesxCred_Consulta() As outNegocio

        Dim R As New outNegocio
        Dim i, k As Integer
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlCommand
        Dim dr As SqlDataReader

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Try
            i = 0
            For k = 0 To oBandejaDesembolsoNuevo.Count - 1
                oComm = New SqlCommand("PA_RecuperaBandejaConsultaSolicitudesxCred", oConex)
                oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = oBandejaDesembolsoNuevo(k).nCodCred
                oComm.Parameters.Add("@nTipoConsulta", SqlDbType.Int).Value = 1
                oComm.Parameters.Add("@cDoc", SqlDbType.VarChar).Value = cDoc
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.ExecuteNonQuery()
                dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

                While dr.Read
                    ReDim Preserve R.oConsultaSolicitudes(i)
                    R.oConsultaSolicitudes(i) = New clsConsultaSolicitudes
                    R.oConsultaSolicitudes(i).cCliente = dr.Item("cCliente")
                    R.oConsultaSolicitudes(i).nCodCred = dr.Item("nCodCred")
                    R.oConsultaSolicitudes(i).cDoc = dr.Item("cdoc")
                    R.oConsultaSolicitudes(i).cUsuOrig = dr.Item("cUsuOrig")
                    R.oConsultaSolicitudes(i).cProd = dr.Item("cProd")
                    R.oConsultaSolicitudes(i).cEstado = dr.Item("cEstado")
                    R.oConsultaSolicitudes(i).cComentario = dr.Item("cComentario")
                    R.oConsultaSolicitudes(i).nPrestamo = dr.Item("nMontoAprobado")
                    R.oConsultaSolicitudes(i).nPrestamoOrig = dr.Item("nMontoOriginal")
                    R.oConsultaSolicitudes(i).nCuotas = dr.Item("nCuotaAprobado")
                    R.oConsultaSolicitudes(i).nCuotasOrig = dr.Item("nCuotaOriginal")
                    R.oConsultaSolicitudes(i).cComentario = dr.Item("cComentario")
                    i = i + 1
                End While
                dr.Close()
                dr = Nothing
            Next


            If i > 0 Then
                i = 0
                For j = 0 To R.oConsultaSolicitudes.Count - 1
                    oComm = New SqlCommand("PA_RecuperaVerificacionxCredito", oConex)
                    oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = R.oConsultaSolicitudes(j).nCodCred
                    oComm.Parameters.Add("@nTipoConsulta", SqlDbType.Int).Value = 1
                    oComm.Parameters.Add("@cDoc", SqlDbType.VarChar).Value = cDoc
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()

                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

                    While dr.Read
                        ReDim Preserve R.oBandejaVerificacion(i)
                        R.oBandejaVerificacion(i) = New clsVerificacionCliente
                        R.oBandejaVerificacion(i).cDoc = R.oConsultaSolicitudes(j).cDoc
                        R.oBandejaVerificacion(i).nCodCred = R.oConsultaSolicitudes(j).nCodCred
                        R.oBandejaVerificacion(i).nVerifCuali = dr.Item("nVerifCuali")
                        R.oBandejaVerificacion(i).nVerifRefe = dr.Item("nVerifRefe")
                        R.oBandejaVerificacion(i).nCumpleConAntNeg = dr.Item("nCumpleConAntNeg")
                        R.oBandejaVerificacion(i).nEsPropietario = dr.Item("nEsPropietario")
                        R.oBandejaVerificacion(i).nPoseeEmpleados = dr.Item("nPoseeEmpleados")
                        R.oBandejaVerificacion(i).nNroEmpleados = dr.Item("nNroEmpleados")
                        R.oBandejaVerificacion(i).nTieneRotuloVisible = dr.Item("nTieneRotuloVisible")
                        R.oBandejaVerificacion(i).nNegoCuentaInfra = dr.Item("nNegoCuentaInfra")
                        R.oBandejaVerificacion(i).nCuentaPermiOpe = dr.Item("nCuentaPermiOpe")
                        R.oBandejaVerificacion(i).nObservaAfluCli = dr.Item("nObservaAfluCli")
                        R.oBandejaVerificacion(i).nUbicaFacilAcceso = dr.Item("nUbicaFacilAcceso")
                        R.oBandejaVerificacion(i).nTieneBuenaRefeZona = dr.Item("nTieneBuenaRefeZona")
                        R.oBandejaVerificacion(i).nTieneProspBuenaActi = dr.Item("nTieneProspBuenaActi")
                        R.oBandejaVerificacion(i).nTieneHoraAtencion = dr.Item("nTieneHoraAtencion")
                        R.oBandejaVerificacion(i).nAquilaCasa = dr.Item("nAquilaCasa")
                        R.oBandejaVerificacion(i).nTiempoAquila = dr.Item("nTiempoAquila")
                        R.oBandejaVerificacion(i).nTieneDependientes = dr.Item("nTieneDependientes")
                        R.oBandejaVerificacion(i).nNroDependientes = dr.Item("nNroDependientes")
                        R.oBandejaVerificacion(i).cHoraInicio = dr.Item("cHoraInicio")
                        R.oBandejaVerificacion(i).cHoraFinal = dr.Item("cHoraFinal")
                        R.oBandejaVerificacion(i).nNivelDelincuencia = dr.Item("nNivelDelincuencia")
                        R.oBandejaVerificacion(i).cComentarioCPD = dr.Item("cComentarioCPD")
                        i = i + 1
                    End While
                    dr.Close()
                    dr = Nothing
                Next
            End If

            If i > 0 Then
                For j = 0 To R.oBandejaVerificacion.Count - 1
                    oComm = New SqlCommand("PA_ReferenciasVericacion", oConex)
                    oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = R.oBandejaVerificacion(j).nCodCred
                    oComm.Parameters.Add("@nTipoConsulta", SqlDbType.Int).Value = 1
                    oComm.Parameters.Add("@cDoc", SqlDbType.VarChar).Value = cDoc
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()

                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0
                    While dr.Read
                        ReDim Preserve R.oBandejaVerificacion(j).oReferencias(i)
                        R.oBandejaVerificacion(j).oReferencias(i) = New clsVerificacionReferencia
                        R.oBandejaVerificacion(j).oReferencias(i).nCodCred = R.oBandejaVerificacion(j).nCodCred
                        R.oBandejaVerificacion(j).oReferencias(i).cNombres = dr.Item("cNombres")
                        R.oBandejaVerificacion(j).oReferencias(i).cTelefono = dr.Item("cTelefono")
                        R.oBandejaVerificacion(j).oReferencias(i).nRecomendado = dr.Item("nRecomienda")
                        R.oBandejaVerificacion(j).oReferencias(i).cRecomendado = dr.Item("cRecomiendaSustento")
                        i = i + 1
                    End While
                    dr.Close()
                    dr = Nothing

                Next
                For j = 0 To R.oBandejaVerificacion.Count - 1
                    oComm = New SqlCommand("PA_ImagenesVerificacion", oConex)
                    oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = R.oBandejaVerificacion(j).nCodCred
                    oComm.Parameters.Add("@nTipoConsulta", SqlDbType.Int).Value = 1
                    oComm.Parameters.Add("@cDoc", SqlDbType.VarChar).Value = cDoc
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.ExecuteNonQuery()
                    dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
                    i = 0
                    While dr.Read
                        ReDim Preserve R.oBandejaVerificacion(j).oImagenes(i)
                        R.oBandejaVerificacion(j).oImagenes(i) = New clsVerificacionClienteImagenes
                        R.oBandejaVerificacion(j).oImagenes(i).nCodCred = R.oBandejaVerificacion(j).nCodCred
                        R.oBandejaVerificacion(j).oImagenes(i).iImagen = UnicodeBytesToString(dr.Item("nFoto"))
                        R.oBandejaVerificacion(j).oImagenes(i).cTipo = dr.Item("cTipo")
                        R.oBandejaVerificacion(j).oImagenes(i).nID = i
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


End Class
