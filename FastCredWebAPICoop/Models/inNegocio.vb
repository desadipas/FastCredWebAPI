Imports System.Data
Imports System.Data.Entity
Imports System.Data.Entity.Core.Objects
Imports System.Data.Entity.Infrastructure

Imports System
Imports System.Collections.Generic
Imports FastCredWebAPI.FastCredWebAPI
Imports System.Data.SqlClient

Public Class inNegocio

    Private db As New DBNEGOCIOEntities
    Public Property sToken As String

    Public Property nSolicitud As Integer

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

    Public Property ListaAbonos As dAbonoDesemb()

    Public Property nPrestamo As Double
    Public Property nCuotas As Integer

    Public Property nProd As Integer
    Public Property cCodCat As String
    Public Property oVoucher As clsvoucher
    Public Property oFirmaContrato As clsFirmaContrato
    Public Property nOperador As Integer
    Public Property cMovil As String
    Public Property sNroTran As String
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

    Public Property oPresolicitudApp As clsCoopPreSol

    Public oListaFeriados() As Date

    Public Property oSolicitudGrupal As clsCoopSolicitudGrupal

    Public Property sClave As String
    Public Property sClaveNueva As String
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



    Public Sub RegistraIntegranteGrupoSiNoExisteComoSocio(ByVal pnCodGrupo As Integer, ByVal nCodPers As Integer)

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nNumCiclo As Integer

        sCadConexion = oConect.gsCadConexDBSolida

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        oComm = New SqlClient.SqlCommand("PA_App_RegistraIntegranteGrupoSiNoExiste", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@ncodage", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@nCodGrupo", SqlDbType.Int).Value = pnCodGrupo
        oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = nCodPers


        oComm.ExecuteNonQuery()


    End Sub


    Public Function DevuelveCodigoDeGrupo(ByVal pnCodGrupo As Integer) As Integer

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nNumCiclo As Integer

        sCadConexion = oConect.gsCadConexDBSolida

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        oComm = New SqlClient.SqlCommand("App_Cred_ObtenerCodigoGrupo", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@pnCodGrp", SqlDbType.Int).Value = pnCodGrupo
        oComm.Parameters.Add("@pnCicloResp", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()
        nNumCiclo = oComm.Parameters("@pnCicloResp").Value

        DevuelveCodigoDeGrupo = nNumCiclo

        Return nNumCiclo

    End Function




    Public Function DevuelveCicloDeGrupo(ByVal pnCodGrupo As Integer) As Integer

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nNumCiclo As Integer

        sCadConexion = oConect.gsCadConexDBSolida

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        oComm = New SqlClient.SqlCommand("App_Cred_RecuperaCiclodeGrupo", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@pnCodGrp", SqlDbType.Int).Value = pnCodGrupo
        oComm.Parameters.Add("@pnCicloResp", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()
        nNumCiclo = oComm.Parameters("@pnCicloResp").Value

        DevuelveCicloDeGrupo = nNumCiclo

        Return nNumCiclo

    End Function


    Public Function DevuelvenCodPers(ByVal psDoc As String) As Integer

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nCodPers As Integer

        sCadConexion = oConect.gsCadConexDBSolida

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        oComm = New SqlClient.SqlCommand("PA_AppDevuelveCodPersonaTitular", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = psDoc
        oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()
        nCodPers = oComm.Parameters("@nCodPers").Value

        DevuelvenCodPers = nCodPers


    End Function

    Public Function RegistraEvaluacioGrupal() As outNegocio
        Dim D(2) As String
        Dim Cred() As Integer
        Dim R As New outNegocio
        Dim SolGrp As New ComTipos.ClsWCredSolGrp
        Dim oCredito As ComCredito.clsCredito
        Dim nReg As Integer
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nResp As Integer
        Dim nRespArr As Integer()
        Dim sRespuesta As String
        Dim dr As SqlClient.SqlDataReader
        Dim i, j As Integer
        Dim nCodPers As Integer

        nResp = 0
        sRespuesta = ""

        sCadConexion = oConect.gsCadConexDBSolida

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Try

            '******************************************************************
            '*** REGISTRA PERSONAS
            '*******************************************************************
            Dim nCicloGrupo As Integer
            Dim nCodGrupoReal As Integer
            nCicloGrupo = DevuelveCicloDeGrupo(oSolicitudGrupal.nCodGrupo)
            nCodGrupoReal = DevuelveCodigoDeGrupo(oSolicitudGrupal.nCodGrupo)


            For i = 0 To oSolicitudGrupal.oCoopSolicitudGrupalSocio.Count - 1

                R = RegistraPersonaIntegranteGrupo(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).oCoopPersona, nCodGrupoReal)

                If R.nResultado = 0 Then
                    Return R
                End If

            Next


            oComm = New SqlClient.SqlCommand("PA_AppDevuelveCodPersonaAnalista", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@psCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()
            nCodPers = oComm.Parameters("@nCodPers").Value


            '********************************************************************************
            '****REGISTRA SOLICITUD DE PRESTAMO
            '********************************************************************************

            SolGrp.Datos = New ComTipos.clsDataset

            SolGrp.Datos.AgregaColumna("nProd")
            SolGrp.Datos.AgregaColumna("nSubProd")
            SolGrp.Datos.AgregaColumna("nCuotas")
            SolGrp.Datos.AgregaColumna("nPlazo")
            SolGrp.Datos.AgregaColumna("nDiaVen")
            SolGrp.Datos.AgregaColumna("dFecDes")
            SolGrp.Datos.AgregaColumna("nCond")
            SolGrp.Datos.AgregaColumna("nDiasGra")
            SolGrp.Datos.AgregaColumna("TipoGracia")
            SolGrp.Datos.AgregaColumna("nCodGrp")
            SolGrp.Datos.AgregaColumna("nCiclo")
            SolGrp.Datos.AgregaColumna("nCamp")
            SolGrp.Datos.AgregaColumna("nDestino")
            SolGrp.Datos.AgregaColumna("nAnalista")
            SolGrp.Datos.AgregaColumna("nCodAge")
            SolGrp.Datos.AgregaColumna("nTipoGra")
            SolGrp.Datos.AgregaColumna("nCodCred")
            SolGrp.Datos.AgregaColumna("nCodModalidad")
            SolGrp.Datos.AgregaColumna("nTipoPeriodo")
            SolGrp.Datos.AgregaColumna("nMoneda")
            SolGrp.Datos.AgregaColumna("nOficina")

            SolGrp.Datos.AddRegistro()

            SolGrp.Datos.AsignarDato("10", "nProd")
            SolGrp.Datos.AsignarDato("2", "nSubProd")
            SolGrp.Datos.AsignarDato(Str(oSolicitudGrupal.nNumCuotas), "nCuotas")
            SolGrp.Datos.AsignarDato(Str(oSolicitudGrupal.nPlazo), "nPlazo")
            SolGrp.Datos.AsignarDato(Str(oSolicitudGrupal.nDiaPago), "nDiaVen")
            SolGrp.Datos.AsignarDato(Format(Now(), "dd/MM/yyyy"), "dFecDes")
            SolGrp.Datos.AsignarDato("1", "nCond")
            SolGrp.Datos.AsignarDato("0", "nDiasGra")
            SolGrp.Datos.AsignarDato("0", "TipoGracia")

            SolGrp.Datos.AsignarDato(Str(nCodGrupoReal), "nCodGrp")
            SolGrp.Datos.AsignarDato(Str(nCicloGrupo), "nCiclo")
            SolGrp.Datos.AsignarDato("0", "nCamp")
            SolGrp.Datos.AsignarDato("2", "nDestino")
            SolGrp.Datos.AsignarDato(Str(nCodPers), "nAnalista")
            SolGrp.Datos.AsignarDato(Str(nCodAge), "nCodAge")
            SolGrp.Datos.AsignarDato("1", "nTipoGra")
            SolGrp.Datos.AsignarDato("2", "nCodCred")
            SolGrp.Datos.AsignarDato("1", "nCodModalidad")
            SolGrp.Datos.AsignarDato(Str(oSolicitudGrupal.nTipoPeriodo), "nTipoPeriodo")
            SolGrp.Datos.AsignarDato("1", "nMoneda")
            SolGrp.Datos.AsignarDato("0", "nOficina")

            SolGrp.Integrantes = New ComTipos.clsDataset

            SolGrp.Integrantes.AgregaColumna("nCodPers")
            SolGrp.Integrantes.AgregaColumna("cPersNombre")
            SolGrp.Integrantes.AgregaColumna("nPrestamo")
            SolGrp.Integrantes.AgregaColumna("nCodAct")
            SolGrp.Integrantes.AgregaColumna("cActividad")
            SolGrp.Integrantes.AgregaColumna("nCoord")
            SolGrp.Integrantes.AgregaColumna("nTasaComp")
            SolGrp.Integrantes.AgregaColumna("nMontoCuota")
            SolGrp.Integrantes.AgregaColumna("nLinea")
            SolGrp.Integrantes.AgregaColumna("nPrestamoAnt")
            SolGrp.Integrantes.AgregaColumna("cDocumento") 'ARCV 08-11-2010



            For i = 0 To oSolicitudGrupal.oCoopSolicitudGrupalSocio.Count - 1
                nReg = SolGrp.Integrantes.AddRegistro()
                SolGrp.Integrantes.Seek(nReg)

                SolGrp.Integrantes.AsignarDato(Str(DevuelvenCodPers(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).cDoc)), "nCodPers")
                SolGrp.Integrantes.AsignarDato("", "cPersNombre")
                SolGrp.Integrantes.AsignarDato(Format(CDbl(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).nMonto), "#,0.00"), "nPrestamo")
                SolGrp.Integrantes.AsignarDato(Str(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).oCoopPersona.nCiiu), "nCodAct")
                SolGrp.Integrantes.AsignarDato(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).cActividad, "cActividad")
                SolGrp.Integrantes.AsignarDato("0", "nCoord")
                SolGrp.Integrantes.AsignarDato(Format(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).nTasa, "#0.00"), "nTasaComp")
                SolGrp.Integrantes.AsignarDato(Format(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).nMontoCuota, "#0.00"), "nMontoCuota")
                SolGrp.Integrantes.AsignarDato("0", "nLinea")
                SolGrp.Integrantes.AsignarDato("0", "nPrestamoAnt")

                SolGrp.Integrantes.AsignarDato(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).cDoc, "cDocumento")

            Next

            oCredito = New ComCredito.clsCredito
            nRespArr = oCredito.RegistrarSolicitudGrupoWIN(SolGrp, 1, 0)
            oCredito = Nothing


            '********************************************************************************
            '**** REGISTRA CREDITO GRUPAL
            '********************************************************************************

            Dim oEVal As New ComTipos.clsDataset

            oEVal.AgregaColumna("nCodPers")
            oEVal.AgregaColumna("cActividad")
            oEVal.AgregaColumna("nVentasIngre")
            oEVal.AgregaColumna("nCostoVenta")
            oEVal.AgregaColumna("nCostoOperativo")
            oEVal.AgregaColumna("nUtilidadNeta")
            oEVal.AgregaColumna("nOtrosIngre")
            oEVal.AgregaColumna("nGastosFinan")
            oEVal.AgregaColumna("nGastosFam")
            oEVal.AgregaColumna("nExcedente")
            oEVal.AgregaColumna("nDisponible")
            oEVal.AgregaColumna("nInventario")
            oEVal.AgregaColumna("nCtaCobrar")
            oEVal.AgregaColumna("nActivoCorrien")
            oEVal.AgregaColumna("nActivoFijo")
            oEVal.AgregaColumna("nPasivoCorrien")
            oEVal.AgregaColumna("nPasivoNoCorrien")

            Dim nUtilidadNeta As Double
            Dim nExcedente As Double

            For i = 0 To oSolicitudGrupal.oCoopSolicitudGrupalSocio.Count - 1
                oEVal.AddRegistro()
                oEVal.AsignarDato(Str(DevuelvenCodPers(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).cDoc)), "nCodPers")
                oEVal.AsignarDato(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).cActividad, "cActividad")
                oEVal.AsignarDato(Format(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).nVentas, "#0.00"), "nVentasIngre")
                oEVal.AsignarDato(Format(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).nCostoVentas, "#0.00"), "nCostoVenta")
                oEVal.AsignarDato(Format(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).nCostosOperativos, "#0.00"), "nCostoOperativo")
                nUtilidadNeta = oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).nVentas - oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).nCostoVentas - oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).nCostosOperativos
                oEVal.AsignarDato(Format(nUtilidadNeta, "#0.00"), "nUtilidadNeta")
                oEVal.AsignarDato(Format(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).nOtrosIngresos, "#0.00"), "nOtrosIngre")
                oEVal.AsignarDato(Format(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).nGastosFinancieros, "#0.00"), "nGastosFinan")
                oEVal.AsignarDato(Format(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).nGastosFamiliares, "#0.00"), "nGastosFam")
                nExcedente = nUtilidadNeta + oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).nOtrosIngresos - oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).nGastosFinancieros - oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).nGastosFamiliares
                oEVal.AsignarDato(Format(nExcedente, "#0.00"), "nExcedente")
                oEVal.AsignarDato(Format(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).nDisponibles, "#0.00"), "nDisponible")
                oEVal.AsignarDato(Format(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).nInventarios, "#0.00"), "nInventario")
                oEVal.AsignarDato(Format(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).nCtasxCobrar, "#0.00"), "nCtaCobrar")
                oEVal.AsignarDato(Format(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).nActivoCorrien, "#0.00"), "nActivoCorrien")
                oEVal.AsignarDato(Format(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).nActivoFijo, "#0.00"), "nActivoFijo")
                oEVal.AsignarDato(Format(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).nPasivoCorriente, "#0.00"), "nPasivoCorrien")
                oEVal.AsignarDato(Format(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).nPasivoNOCorriente, "#0.00"), "nPasivoNoCorrien")

            Next


            Dim E As ComPersona.clsEvaluacion
            E = New ComPersona.clsEvaluacion
            Call E.PersEvalActualizaEvaluacionGrupal(0, nCodGrupoReal, nCicloGrupo, Now(), oEVal)
            E = Nothing


            '********************************************************************************
            '**** REGISTRA GARANTIA GRUPAL
            '********************************************************************************
            Dim oDatosGaran As ComTipos.clsDataset
            oDatosGaran = New ComTipos.clsDataset

            oDatosGaran.AgregaColumna("nCodCred")
            oDatosGaran.AgregaColumna("cNombre")
            oDatosGaran.AgregaColumna("nPrestamo")
            oDatosGaran.AgregaColumna("nMonto")
            oDatosGaran.AgregaColumna("nCodAge")
            oDatosGaran.AgregaColumna("nCodPers")
            oDatosGaran.AgregaColumna("nMoneda")

            For i = 0 To oSolicitudGrupal.oCoopSolicitudGrupalSocio.Count - 1

                oDatosGaran.AddRegistro()
                oDatosGaran.AsignarDato(Str(nRespArr(i + 1)), "nCodCred")
                oDatosGaran.AsignarDato("", "cNombre")
                oDatosGaran.AsignarDato(Str(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).nMonto), "nPrestamo")
                oDatosGaran.AsignarDato(Str(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).nMonto), "nMonto")
                oDatosGaran.AsignarDato(Str(nCodAge), "nCodAge")
                oDatosGaran.AsignarDato(Str(DevuelvenCodPers(oSolicitudGrupal.oCoopSolicitudGrupalSocio(i).cDoc)), "nCodPers")
                oDatosGaran.AsignarDato("1", "nMoneda")

            Next

            Dim oObjGar As ComCredito.clsCredito
            oObjGar = New ComCredito.clsCredito

            oObjGar.Graba_GarantiaGrupal(Now(), 1, oDatosGaran, nCodGrupoReal, nCicloGrupo)

            oObjGar = Nothing


            D(0) = "1"
            D(1) = "REGISTRO CORRECTO"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try

        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function

    Public Function RegistraDataEvaluaciondApp() As outNegocio

        Dim R As New outNegocio


        R = EliminaDataPersona()
        If (R.nResultado = 0) Then
            Return R
        End If

        R = RegistraPersonaApp()
        If (R.nResultado = 0) Then
            Return R
        End If

        R = RegistraSolicitudApp()
        If (R.nResultado = 0) Then
            Return R
        End If


        If Not IsNothing(oPresolicitudApp.oCoopSolicitud.oEvaluacion) Then
            R = RegistraEvaluacionPymeADetalleBalance()
            If (R.nResultado = 0) Then
                Return R
            End If

            R = RegistraEvaluacionPymeSemovientes()
            If (R.nResultado = 0) Then
                Return R
            End If

            R = RegistraEvaluacionPymeEvaDeudas()
            If (R.nResultado = 0) Then
                Return R
            End If

            R = RegistraEvaluacionPymeOtrosNegocios()
            If (R.nResultado = 0) Then
                Return R
            End If

            R = RegistraEvaluacionPymeIngresosEgresosEstac()
            If (R.nResultado = 0) Then
                Return R
            End If

            R = RegistraEvaluacionPymeEvalGiros()
            If (R.nResultado = 0) Then
                Return R
            End If

            R = RegistraEvaluacionPymeEvalGirosCosteoProd()
            If (R.nResultado = 0) Then
                Return R
            End If

            R = RegistraEvaluacionPymeEvalGiroInventario()
            If (R.nResultado = 0) Then
                Return R
            End If

            R = RegistraEvaluacionPymeEvalGiroGastosOper()
            If (R.nResultado = 0) Then
                Return R
            End If

            R = RegistraEvaluacionPymeEvalGiroVentas()
            If (R.nResultado = 0) Then
                Return R
            End If

            R = RegistraEvaluacionAgricola()
            If (R.nResultado = 0) Then
                Return R
            End If

            R = RegistraEvaluacionAgricolaProyVentas()
            If (R.nResultado = 0) Then
                Return R
            End If

            R = RegistraEvaluacionAgricolaProyCostos()
            If (R.nResultado = 0) Then
                Return R
            End If

        End If

        If Not IsNothing(oPresolicitudApp.oCoopSolicitud.oCoopevalunidfam) Then

            R = RegistraEvaluacionUnidadFamiliar()
            If (R.nResultado = 0) Then
                Return R
            End If

            R = RegistraEvaluacionUniFamDatosPersonales()
            If (R.nResultado = 0) Then
                Return R
            End If

            R = RegistraEvaluacionUniFamDetBalance()
            If (R.nResultado = 0) Then
                Return R
            End If

            R = RegistraEvaluacionUniFamDeudas()
            If (R.nResultado = 0) Then
                Return R
            End If

            R = RegistraEvaluacionUniFamEducacion()
            If (R.nResultado = 0) Then
                Return R
            End If

            R = RegistraEvaluacionUniFamIngresos()
            If (R.nResultado = 0) Then
                Return R
            End If

            R = RegistraEvaluacionUniFamIngresosEstacionales()
            If (R.nResultado = 0) Then
                Return R
            End If

            R = RegistraEvaluacionUniFamOtrosNegocios()
            If (R.nResultado = 0) Then
                Return R
            End If

            R = RegistraEvaluacionUniFamGastosFam()
            If (R.nResultado = 0) Then
                Return R
            End If

            R = RegistraEvaluacionUniFamComentarios()
            If (R.nResultado = 0) Then
                Return R
            End If

            R = RegistraEvaluacionUniFamFotos()
            If (R.nResultado = 0) Then
                Return R
            End If

            R = RegistraEvaluacionUniFamReferencias()
            If (R.nResultado = 0) Then
                Return R
            End If


        End If

        If oPresolicitudApp.oCoopSolicitud.nOperacion = 10 And oPresolicitudApp.oCoopSolicitud.nSubProd = 5 Then
            R = RegistraEvaluacionCreditoDiario()
            If (R.nResultado = 0) Then
                Return R
            End If

            R = RegistraEvaluacionDiariaFotos()
            If (R.nResultado = 0) Then
                Return R
            End If
        End If

        If oPresolicitudApp.oCoopSolicitud.nOperacion = 10 And (oPresolicitudApp.oCoopSolicitud.nSubProd = 1 Or oPresolicitudApp.oCoopSolicitud.nSubProd = 2) Then
            R = RegistraEvaluacionCreditoGrupal()
            If (R.nResultado = 0) Then
                Return R
            End If
        End If


        R = RegistraEvaluacionPymeApp()
        If (R.nResultado = 0) Then
            Return R
        End If


        R = RegistraPreSolicApp()
        If (R.nResultado = 0) Then
            Return R
        End If



        R.cDoc = oPresolicitudApp.cNroDoc
        Return R

    End Function



    Public Function EliminaDataPersona() As outNegocio

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



            oComm = New SqlClient.SqlCommand("App_EliminaDataPersona", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = oPresolicitudApp.cNroDoc

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

    Public Function ObtieneGastoTotalCal(ByVal pnPrestamo As Double, ByVal pnMoneda As Integer,
        ByVal pnProd As Integer, ByVal pnSubProd As Integer, Optional ByVal pnSegInd As Integer = 1) As Double

        Dim R As ComTipos.clsDataset

        R = New ComTipos.clsDataset
        R.CargaTabla(, , "CredW_RecuperaGastosParaSimulador " & pnPrestamo & "," & pnMoneda & "," & pnProd & "," & pnSubProd & "," & pnSegInd)

        ObtieneGastoTotalCal = CDbl(R.ObtenerDato("nValorGasto"))

        R = Nothing

    End Function
    Public Function RecuperaFeriadosDBSolidaCore(ByVal pdFechaDesem As Date) As ComTipos.clsDataset
        Dim D As ComTipos.clsDataset
        D = New ComTipos.clsDataset
        D.CargaTabla(, , "Cred_DevuelveFeriados '" & Format(pdFechaDesem, "dd/MM/yyyy") & "'")
        RecuperaFeriadosDBSolidaCore = D
        D = Nothing
    End Function
    Public Sub ActualizaEstadoAPropuesta(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand
        Dim oCal As ComTipos.ClsCalendario
        Dim oCalData As ComTipos.clsDataset
        Dim nMontoGasto As Double
        Dim oFeriado As ComTipos.clsDataset
        Dim nPlazo As Integer

        nMontoGasto = ObtieneGastoTotalCal(oPresolicitudApp.oCoopSolicitud.nMonto, 1, oPresolicitudApp.oCoopSolicitud.nOperacion, oPresolicitudApp.oCoopSolicitud.nSubProd)
        oFeriado = RecuperaFeriadosDBSolidaCore(oPresolicitudApp.oCoopSolicitud.dFecDesem)

        If oPresolicitudApp.oCoopSolicitud.nOperacion = 10 And oPresolicitudApp.oCoopSolicitud.nSubProd = 5 Then
            nPlazo = 1
        Else
            nPlazo = 30
        End If

        oCal = New ComTipos.ClsCalendario
        oCalData = oCal.GeneraCalendarioDataSet(nMontoGasto, oFeriado, pnCodAge, pnCodCred,
            1, oPresolicitudApp.oCoopSolicitud.nMonto, oPresolicitudApp.oCoopSolicitud.nNumCuo, oPresolicitudApp.oCoopSolicitud.nTasa,
            oPresolicitudApp.oCoopSolicitud.nDiaPago, oPresolicitudApp.oCoopSolicitud.dFecDesem, 0, 0, 0,
            oPresolicitudApp.oCoopSolicitud.nOperacion, oPresolicitudApp.oCoopSolicitud.nSubProd, 0, oPresolicitudApp.oCoopSolicitud.nPeriodo, nPlazo, 0, False, 0, oPresolicitudApp.oCoopSolicitud.dFecPrimCuota, False, False)

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("App_ActualizaEstadoAPropuesta", objCon.DameConexion)
        '---------------

        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int)
        oComm.Parameters("@nCodCred").Value = pnCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int)
        oComm.Parameters("@nCodAge").Value = pnCodAge
        oComm.Parameters.Add("@Cal_XML", SqlDbType.Xml)
        oComm.Parameters("@Cal_XML").Value = oCalData.DevuelveXML
        oComm.ExecuteNonQuery()

        objCon.CierraConexion()
        objCon = Nothing
        oComm = Nothing
        oCal = Nothing
        oCalData = Nothing



    End Sub

    Public Function RegistraEvaluacionCreditoGrupal() As outNegocio

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


            With oPresolicitudApp.oCoopSolicitud

                If Not IsNothing(oPresolicitudApp.oCoopSolicitud) Then

                    oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalCreditoGrupal", oConex)
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure

                    oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = oPresolicitudApp.cNroDoc
                    oComm.Parameters.Add("@nTotalventas", SqlDbType.Money).Value = .oCoopevalgrupal.nTotalventas
                    oComm.Parameters.Add("@nTotalcompras", SqlDbType.Money).Value = .oCoopevalgrupal.nTotalcompras
                    oComm.Parameters.Add("@nGanancia", SqlDbType.Money).Value = .oCoopevalgrupal.nGanancia
                    oComm.Parameters.Add("@nInventario", SqlDbType.Money).Value = .oCoopevalgrupal.nInventario
                    oComm.Parameters.Add("@nMontosol", SqlDbType.Money).Value = .oCoopevalgrupal.nMontosol
                    oComm.Parameters.Add("@nPlazo", SqlDbType.Int).Value = .oCoopevalgrupal.nPlazo
                    oComm.Parameters.Add("@nCuota", SqlDbType.Money).Value = .oCoopevalgrupal.nCuota
                    oComm.Parameters.Add("@nTotalventaslun", SqlDbType.Money).Value = .oCoopevalgrupal.nTotalventaslun
                    oComm.Parameters.Add("@nTotalcompraslun", SqlDbType.Money).Value = .oCoopevalgrupal.nTotalcompraslun
                    oComm.Parameters.Add("@ncajalun", SqlDbType.Money).Value = .oCoopevalgrupal.ncajalun
                    oComm.Parameters.Add("@nTotalventasmar", SqlDbType.Money).Value = .oCoopevalgrupal.nTotalventasmar
                    oComm.Parameters.Add("@nTotalcomprasmar", SqlDbType.Money).Value = .oCoopevalgrupal.nTotalcomprasmar
                    oComm.Parameters.Add("@ncajamar", SqlDbType.Money).Value = .oCoopevalgrupal.ncajamar
                    oComm.Parameters.Add("@nTotalventasmie", SqlDbType.Money).Value = .oCoopevalgrupal.nTotalventasmie
                    oComm.Parameters.Add("@nTotalcomprasmie", SqlDbType.Money).Value = .oCoopevalgrupal.nTotalcomprasmie
                    oComm.Parameters.Add("@ncajamie", SqlDbType.Money).Value = .oCoopevalgrupal.ncajamie
                    oComm.Parameters.Add("@nTotalventasjue", SqlDbType.Money).Value = .oCoopevalgrupal.nTotalventasjue
                    oComm.Parameters.Add("@nTotalcomprasjue", SqlDbType.Money).Value = .oCoopevalgrupal.nTotalcomprasjue
                    oComm.Parameters.Add("@ncajajue", SqlDbType.Money).Value = .oCoopevalgrupal.ncajajue
                    oComm.Parameters.Add("@nTotalventasvie", SqlDbType.Money).Value = .oCoopevalgrupal.nTotalventasvie
                    oComm.Parameters.Add("@nTotalcomprasvie", SqlDbType.Money).Value = .oCoopevalgrupal.nTotalcomprasvie
                    oComm.Parameters.Add("@ncajavie", SqlDbType.Money).Value = .oCoopevalgrupal.ncajavie
                    oComm.Parameters.Add("@nTotalventassab", SqlDbType.Money).Value = .oCoopevalgrupal.nTotalventassab
                    oComm.Parameters.Add("@nTotalcomprassab", SqlDbType.Money).Value = .oCoopevalgrupal.nTotalcomprassab
                    oComm.Parameters.Add("@ncajasab", SqlDbType.Money).Value = .oCoopevalgrupal.ncajasab
                    oComm.Parameters.Add("@nTotalventasdom", SqlDbType.Money).Value = .oCoopevalgrupal.nTotalventasdom
                    oComm.Parameters.Add("@nTotalcomprasdom", SqlDbType.Money).Value = .oCoopevalgrupal.nTotalcomprasdom
                    oComm.Parameters.Add("@ncajadom", SqlDbType.Money).Value = .oCoopevalgrupal.ncajadom
                    oComm.Parameters.Add("@nalimentacion", SqlDbType.Money).Value = .oCoopevalgrupal.nalimentacion
                    oComm.Parameters.Add("@nalquiler", SqlDbType.Money).Value = .oCoopevalgrupal.nalquiler
                    oComm.Parameters.Add("@nvestido", SqlDbType.Money).Value = .oCoopevalgrupal.nvestido
                    oComm.Parameters.Add("@neducacion", SqlDbType.Money).Value = .oCoopevalgrupal.neducacion
                    oComm.Parameters.Add("@nsalud", SqlDbType.Money).Value = .oCoopevalgrupal.nsalud
                    oComm.Parameters.Add("@ntransporte", SqlDbType.Money).Value = .oCoopevalgrupal.ntransporte
                    oComm.Parameters.Add("@nservicios", SqlDbType.Money).Value = .oCoopevalgrupal.nservicios
                    oComm.Parameters.Add("@npension", SqlDbType.Money).Value = .oCoopevalgrupal.npension
                    oComm.Parameters.Add("@nrecreacion", SqlDbType.Money).Value = .oCoopevalgrupal.nrecreacion
                    oComm.Parameters.Add("@notros", SqlDbType.Money).Value = .oCoopevalgrupal.notros
                    oComm.Parameters.Add("@ncuotatarjeta", SqlDbType.Money).Value = .oCoopevalgrupal.ncuotatarjeta
                    oComm.Parameters.Add("@ncuotacasacom", SqlDbType.Money).Value = .oCoopevalgrupal.ncuotacasacom
                    oComm.Parameters.Add("@ncuotapresifi", SqlDbType.Money).Value = .oCoopevalgrupal.ncuotapresifi
                    oComm.Parameters.Add("@ncuotaprestamista", SqlDbType.Money).Value = .oCoopevalgrupal.ncuotaprestamista
                    oComm.Parameters.Add("@ncuotajuntas", SqlDbType.Money).Value = .oCoopevalgrupal.ncuotajuntas
                    oComm.Parameters.Add("@notrosingr", SqlDbType.Money).Value = .oCoopevalgrupal.notrosingr
                    oComm.Parameters.Add("@notrosegresos", SqlDbType.Money).Value = .oCoopevalgrupal.notrosegresos
                    If IsNothing(.oCoopevalgrupal.cComentario) Then
                        oComm.Parameters.Add("@cComentario", SqlDbType.VarChar, 200).Value = ""
                    Else
                        oComm.Parameters.Add("@cComentario", SqlDbType.VarChar, 200).Value = .oCoopevalgrupal.cComentario
                    End If


                    oComm.Parameters.Add("@nTiempolaborando", SqlDbType.Int).Value = .oCoopevalgrupal.nTiempolaborando
                    oComm.Parameters.Add("@nFrecuenciaingresos", SqlDbType.Int).Value = .oCoopevalgrupal.nFrecuenciaingresos
                    oComm.Parameters.Add("@nReferenciaspersonales", SqlDbType.Int).Value = .oCoopevalgrupal.nReferenciaspersonales

                    oComm.ExecuteNonQuery()

                End If

            End With

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

    Public Function RegistraEvaluacionCreditoDiario() As outNegocio

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


            With oPresolicitudApp.oCoopSolicitud

                If Not IsNothing(oPresolicitudApp.oCoopSolicitud) Then

                    oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalCreditoDiario", oConex)
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure

                    oComm.Parameters.Add("@nIDEval", SqlDbType.Int).Direction = ParameterDirection.Output
                    oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = oPresolicitudApp.cNroDoc
                    oComm.Parameters.Add("@nTotalventas", SqlDbType.Money).Value = .oCoopevaldiario.nTotalventas
                    oComm.Parameters.Add("@nTotalcompras", SqlDbType.Money).Value = .oCoopevaldiario.nTotalcompras
                    oComm.Parameters.Add("@nGanancia", SqlDbType.Money).Value = .oCoopevaldiario.nGanancia
                    oComm.Parameters.Add("@nInventario", SqlDbType.Money).Value = .oCoopevaldiario.nInventario
                    oComm.Parameters.Add("@nMontosol", SqlDbType.Money).Value = .oCoopevaldiario.nMontosol
                    oComm.Parameters.Add("@nPlazo", SqlDbType.Int).Value = .oCoopevaldiario.nPlazo
                    oComm.Parameters.Add("@nCuota", SqlDbType.Money).Value = .oCoopevaldiario.nCuota
                    oComm.Parameters.Add("@nTotalventaslun", SqlDbType.Money).Value = .oCoopevaldiario.nTotalventaslun
                    oComm.Parameters.Add("@nTotalcompraslun", SqlDbType.Money).Value = .oCoopevaldiario.nTotalcompraslun
                    oComm.Parameters.Add("@ncajalun", SqlDbType.Money).Value = .oCoopevaldiario.ncajalun
                    oComm.Parameters.Add("@nTotalventasmar", SqlDbType.Money).Value = .oCoopevaldiario.nTotalventasmar
                    oComm.Parameters.Add("@nTotalcomprasmar", SqlDbType.Money).Value = .oCoopevaldiario.nTotalcomprasmar
                    oComm.Parameters.Add("@ncajamar", SqlDbType.Money).Value = .oCoopevaldiario.ncajamar
                    oComm.Parameters.Add("@nTotalventasmie", SqlDbType.Money).Value = .oCoopevaldiario.nTotalventasmie
                    oComm.Parameters.Add("@nTotalcomprasmie", SqlDbType.Money).Value = .oCoopevaldiario.nTotalcomprasmie
                    oComm.Parameters.Add("@ncajamie", SqlDbType.Money).Value = .oCoopevaldiario.ncajamie
                    oComm.Parameters.Add("@nTotalventasjue", SqlDbType.Money).Value = .oCoopevaldiario.nTotalventasjue
                    oComm.Parameters.Add("@nTotalcomprasjue", SqlDbType.Money).Value = .oCoopevaldiario.nTotalcomprasjue
                    oComm.Parameters.Add("@ncajajue", SqlDbType.Money).Value = .oCoopevaldiario.ncajajue
                    oComm.Parameters.Add("@nTotalventasvie", SqlDbType.Money).Value = .oCoopevaldiario.nTotalventasvie
                    oComm.Parameters.Add("@nTotalcomprasvie", SqlDbType.Money).Value = .oCoopevaldiario.nTotalcomprasvie
                    oComm.Parameters.Add("@ncajavie", SqlDbType.Money).Value = .oCoopevaldiario.ncajavie
                    oComm.Parameters.Add("@nTotalventassab", SqlDbType.Money).Value = .oCoopevaldiario.nTotalventassab
                    oComm.Parameters.Add("@nTotalcomprassab", SqlDbType.Money).Value = .oCoopevaldiario.nTotalcomprassab
                    oComm.Parameters.Add("@ncajasab", SqlDbType.Money).Value = .oCoopevaldiario.ncajasab
                    oComm.Parameters.Add("@nTotalventasdom", SqlDbType.Money).Value = .oCoopevaldiario.nTotalventasdom
                    oComm.Parameters.Add("@nTotalcomprasdom", SqlDbType.Money).Value = .oCoopevaldiario.nTotalcomprasdom
                    oComm.Parameters.Add("@ncajadom", SqlDbType.Money).Value = .oCoopevaldiario.ncajadom
                    oComm.Parameters.Add("@nalimentacion", SqlDbType.Money).Value = .oCoopevaldiario.nalimentacion
                    oComm.Parameters.Add("@nalquiler", SqlDbType.Money).Value = .oCoopevaldiario.nalquiler
                    oComm.Parameters.Add("@nvestido", SqlDbType.Money).Value = .oCoopevaldiario.nvestido
                    oComm.Parameters.Add("@neducacion", SqlDbType.Money).Value = .oCoopevaldiario.neducacion
                    oComm.Parameters.Add("@nsalud", SqlDbType.Money).Value = .oCoopevaldiario.nsalud
                    oComm.Parameters.Add("@ntransporte", SqlDbType.Money).Value = .oCoopevaldiario.ntransporte
                    oComm.Parameters.Add("@nservicios", SqlDbType.Money).Value = .oCoopevaldiario.nservicios
                    oComm.Parameters.Add("@npension", SqlDbType.Money).Value = .oCoopevaldiario.npension
                    oComm.Parameters.Add("@nrecreacion", SqlDbType.Money).Value = .oCoopevaldiario.nrecreacion
                    oComm.Parameters.Add("@notros", SqlDbType.Money).Value = .oCoopevaldiario.notros
                    oComm.Parameters.Add("@ncuotatarjeta", SqlDbType.Money).Value = .oCoopevaldiario.ncuotatarjeta
                    oComm.Parameters.Add("@ncuotacasacom", SqlDbType.Money).Value = .oCoopevaldiario.ncuotacasacom
                    oComm.Parameters.Add("@ncuotapresifi", SqlDbType.Money).Value = .oCoopevaldiario.ncuotapresifi
                    oComm.Parameters.Add("@ncuotaprestamista", SqlDbType.Money).Value = .oCoopevaldiario.ncuotaprestamista
                    oComm.Parameters.Add("@ncuotajuntas", SqlDbType.Money).Value = .oCoopevaldiario.ncuotajuntas
                    oComm.Parameters.Add("@notrosingr", SqlDbType.Money).Value = .oCoopevaldiario.notrosingr
                    oComm.Parameters.Add("@notrosegresos", SqlDbType.Money).Value = .oCoopevaldiario.notrosegresos
                    If IsNothing(.oCoopevaldiario.cComentario) Then
                        oComm.Parameters.Add("@cComentario", SqlDbType.VarChar, 200).Value = ""
                    Else
                        oComm.Parameters.Add("@cComentario", SqlDbType.VarChar, 200).Value = .oCoopevaldiario.cComentario
                    End If
                    oComm.Parameters.Add("@nLatitud", SqlDbType.Float).Value = .oCoopevaldiario.nLatitud
                    oComm.Parameters.Add("@nLongitud", SqlDbType.Float).Value = .oCoopevaldiario.nLongitud

                    oComm.Parameters.Add("@nNumEnt", SqlDbType.Int).Value = .oCoopevaldiario.nNumEnt
                    oComm.Parameters.Add("@nNumDepen", SqlDbType.Int).Value = .oCoopevaldiario.nNumDep
                    oComm.Parameters.Add("@nAniosLaborando", SqlDbType.Int).Value = .oCoopevaldiario.nAniosLaborando
                    oComm.Parameters.Add("@nFrencuenciaIngresos", SqlDbType.Int).Value = .oCoopevaldiario.nFrecuenciaIngresos
                    oComm.Parameters.Add("@nReferencias", SqlDbType.Int).Value = .oCoopevaldiario.nReferencias

                    oComm.ExecuteNonQuery()




                End If

            End With

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


    Public Function RegistraEvaluacionUniFamReferencias() As outNegocio

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


            With oPresolicitudApp.oCoopSolicitud.oCoopevalunidfam
                If Not IsNothing(.oEvalfamreferencias) Then

                    For i = 0 To .oEvalfamreferencias.Count - 1

                        oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalUniFamReferencias", oConex)
                        oComm.CommandTimeout = 0
                        oComm.CommandType = CommandType.StoredProcedure

                        oComm.Parameters.Add("@nIDEvalfam", SqlDbType.Int).Value = .nIDEvalUniFam
                        oComm.Parameters.Add("@nTipoReferencia", SqlDbType.Int).Value = .oEvalfamreferencias(i).nTipoReferencia
                        oComm.Parameters.Add("@cReferencia", SqlDbType.VarChar, 150).Value = .oEvalfamreferencias(i).cReferencia


                        oComm.ExecuteNonQuery()

                    Next
                End If

            End With

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

    Public Function RegistraEvaluacionDiariaFotos() As outNegocio

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


            With oPresolicitudApp.oCoopSolicitud.oCoopevaldiario
                If Not IsNothing(.oFotos) Then

                    For i = 0 To .oFotos.Count - 1

                        oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalDiarioFotos", oConex)
                        oComm.CommandTimeout = 0
                        oComm.CommandType = CommandType.StoredProcedure

                        oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = .oFotos(i).cDoc
                        oComm.Parameters.Add("@cNomfoto", SqlDbType.VarChar, 100).Value = .oFotos(i).cNomfoto
                        oComm.Parameters.Add("@iFoto", SqlDbType.Image).Value = UnicodeStringToBytes(.oFotos(i).iFoto)

                        oComm.ExecuteNonQuery()

                    Next
                End If

            End With

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


    Public Function RegistraEvaluacionUniFamFotos() As outNegocio

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


            With oPresolicitudApp.oCoopSolicitud.oCoopevalunidfam
                If Not IsNothing(.oEvalfamfotos) Then

                    For i = 0 To .oEvalfamfotos.Count - 1

                        oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalUniFamFotos", oConex)
                        oComm.CommandTimeout = 0
                        oComm.CommandType = CommandType.StoredProcedure

                        oComm.Parameters.Add("@nIDEvalfam", SqlDbType.Int).Value = .nIDEvalUniFam
                        oComm.Parameters.Add("@cNomfoto", SqlDbType.VarChar, 100).Value = .oEvalfamfotos(i).cNomfoto
                        oComm.Parameters.Add("@iFoto", SqlDbType.Image).Value = UnicodeStringToBytes(.oEvalfamfotos(i).iFoto)

                        oComm.ExecuteNonQuery()

                    Next
                End If

            End With

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


    Public Function RegistraEvaluacionUniFamComentarios() As outNegocio

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


            With oPresolicitudApp.oCoopSolicitud.oCoopevalunidfam
                If Not IsNothing(.oEvalfamcomentarios) Then

                    For i = 0 To .oEvalfamcomentarios.Count - 1

                        oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalUniFamComentarios", oConex)
                        oComm.CommandTimeout = 0
                        oComm.CommandType = CommandType.StoredProcedure

                        oComm.Parameters.Add("@nIDEvalfam", SqlDbType.Int).Value = .nIDEvalUniFam
                        oComm.Parameters.Add("@nTipocomentario", SqlDbType.Int).Value = .oEvalfamcomentarios(i).nTipocomentario
                        oComm.Parameters.Add("@cComentario", SqlDbType.VarChar, 200).Value = .oEvalfamcomentarios(i).cComentario

                        oComm.ExecuteNonQuery()

                    Next
                End If

            End With

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



    Public Function RegistraEvaluacionUniFamGastosFam() As outNegocio

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


            With oPresolicitudApp.oCoopSolicitud.oCoopevalunidfam
                If Not IsNothing(.oCoopevalgastosfam) Then

                    For i = 0 To .oCoopevalgastosfam.Count - 1

                        oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalUniFamGastosFam", oConex)
                        oComm.CommandTimeout = 0
                        oComm.CommandType = CommandType.StoredProcedure

                        oComm.Parameters.Add("@nIDEvalUniFam", SqlDbType.Int).Value = .nIDEvalUniFam
                        oComm.Parameters.Add("@nCodGasto", SqlDbType.Int).Value = .oCoopevalgastosfam(i).nCodGasto
                        oComm.Parameters.Add("@cDescripcion", SqlDbType.VarChar, 150).Value = .oCoopevalgastosfam(i).cDescripcion
                        oComm.Parameters.Add("@nMontoMin", SqlDbType.Money).Value = .oCoopevalgastosfam(i).nMontoMin
                        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = .oCoopevalgastosfam(i).nMonto

                        oComm.ExecuteNonQuery()

                    Next
                End If

            End With

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


    Public Function RegistraEvaluacionUniFamOtrosNegocios() As outNegocio

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


            With oPresolicitudApp.oCoopSolicitud.oCoopevalunidfam
                If Not IsNothing(.oCoopevalfamotrosneg) Then

                    For i = 0 To .oCoopevalfamotrosneg.Count - 1

                        oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalUniFamOtrosNegocios", oConex)
                        oComm.CommandTimeout = 0
                        oComm.CommandType = CommandType.StoredProcedure

                        oComm.Parameters.Add("@nIDEvalUniFam", SqlDbType.Int).Value = .nIDEvalUniFam
                        oComm.Parameters.Add("@nTipoIntervin", SqlDbType.Int).Value = .oCoopevalfamotrosneg(i).nTipoIntervin
                        oComm.Parameters.Add("@cDescripcion", SqlDbType.VarChar, 150).Value = .oCoopevalfamotrosneg(i).cDescripcion
                        oComm.Parameters.Add("@nVentas", SqlDbType.Money).Value = .oCoopevalfamotrosneg(i).nVentas
                        oComm.Parameters.Add("@nCostoVentas", SqlDbType.Money).Value = .oCoopevalfamotrosneg(i).nCostoVentas
                        oComm.Parameters.Add("@nGastosOper", SqlDbType.Money).Value = .oCoopevalfamotrosneg(i).nGastosOper
                        oComm.Parameters.Add("@nGastosFinanc", SqlDbType.Money).Value = .oCoopevalfamotrosneg(i).nGastosFinanc
                        oComm.Parameters.Add("@nMoneda", SqlDbType.Int).Value = .oCoopevalfamotrosneg(i).nMoneda


                        oComm.ExecuteNonQuery()

                    Next
                End If

            End With

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


    Public Function RegistraEvaluacionUniFamIngresosEstacionales() As outNegocio

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


            With oPresolicitudApp.oCoopSolicitud.oCoopevalunidfam
                If Not IsNothing(.oCoopevalfamingrestac) Then

                    For i = 0 To .oCoopevalfamingrestac.Count - 1

                        oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalUniFamIngresosEstacionales", oConex)
                        oComm.CommandTimeout = 0
                        oComm.CommandType = CommandType.StoredProcedure

                        oComm.Parameters.Add("@nIDEvalUniFam", SqlDbType.Int).Value = .nIDEvalUniFam
                        oComm.Parameters.Add("@nTipoIngrEgr", SqlDbType.Int).Value = .oCoopevalfamingrestac(i).nTipoIngrEgr
                        oComm.Parameters.Add("@dMes", SqlDbType.Date).Value = .oCoopevalfamingrestac(i).dMes
                        oComm.Parameters.Add("@cDescripcion", SqlDbType.VarChar, 150).Value = .oCoopevalfamingrestac(i).cDescripcion
                        oComm.Parameters.Add("@nMoneda", SqlDbType.Int).Value = .oCoopevalfamingrestac(i).nMoneda
                        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = .oCoopevalfamingrestac(i).nMonto

                        oComm.ExecuteNonQuery()

                    Next
                End If

            End With

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


    Public Function RegistraEvaluacionUniFamIngresos() As outNegocio

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


            With oPresolicitudApp.oCoopSolicitud.oCoopevalunidfam
                If Not IsNothing(.oCoopevalfamingresos) Then

                    For i = 0 To .oCoopevalfamingresos.Count - 1

                        oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalUniFamIngresos", oConex)
                        oComm.CommandTimeout = 0
                        oComm.CommandType = CommandType.StoredProcedure

                        oComm.Parameters.Add("@nIDEvalUniFam", SqlDbType.Int).Value = .nIDEvalUniFam
                        oComm.Parameters.Add("@nTipoIntervin", SqlDbType.Int).Value = .oCoopevalfamingresos(i).nTipoIntervin
                        oComm.Parameters.Add("@cIngresos", SqlDbType.VarChar, 100).Value = .oCoopevalfamingresos(i).cIngresos
                        oComm.Parameters.Add("@nMoneda", SqlDbType.Int).Value = .oCoopevalfamingresos(i).nMoneda
                        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = .oCoopevalfamingresos(i).nMonto

                        oComm.ExecuteNonQuery()

                    Next
                End If

            End With

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



    Public Function RegistraEvaluacionUniFamEducacion() As outNegocio

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


            With oPresolicitudApp.oCoopSolicitud.oCoopevalunidfam
                If Not IsNothing(.oCoopevalfameducacion) Then

                    For i = 0 To .oCoopevalfameducacion.Count - 1


                        oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalUniFamEduacion", oConex)
                        oComm.CommandTimeout = 0
                        oComm.CommandType = CommandType.StoredProcedure

                        oComm.Parameters.Add("@nIDEvalUniFam", SqlDbType.Int).Value = .nIDEvalUniFam
                        oComm.Parameters.Add("@nTipoNivel", SqlDbType.VarChar, 100).Value = .oCoopevalfameducacion(i).nTipoNivel
                        oComm.Parameters.Add("@nCantidad", SqlDbType.Int).Value = .oCoopevalfameducacion(i).nCantidad
                        oComm.Parameters.Add("@nMontoMinimo", SqlDbType.Int).Value = .oCoopevalfameducacion(i).nMontoMinimo
                        oComm.Parameters.Add("@nMonto", SqlDbType.Int).Value = .oCoopevalfameducacion(i).nMonto

                        oComm.ExecuteNonQuery()

                    Next
                End If

            End With

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




    Public Function RegistraEvaluacionUniFamDeudas() As outNegocio

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


            With oPresolicitudApp.oCoopSolicitud.oCoopevalunidfam
                If Not IsNothing(.oCoopevalfamdeudas) Then
                    For i = 0 To .oCoopevalfamdeudas.Count - 1


                        oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalUniFamDeudas", oConex)
                        oComm.CommandTimeout = 0
                        oComm.CommandType = CommandType.StoredProcedure

                        oComm.Parameters.Add("@nIDEvalUniFam", SqlDbType.Int).Value = .nIDEvalUniFam
                        oComm.Parameters.Add("@cEntidad", SqlDbType.VarChar, 100).Value = .oCoopevalfamdeudas(i).cEntidad
                        oComm.Parameters.Add("@nTipocred", SqlDbType.Int).Value = .oCoopevalfamdeudas(i).nTipocred
                        oComm.Parameters.Add("@nMoneda", SqlDbType.Int).Value = .oCoopevalfamdeudas(i).nMoneda
                        oComm.Parameters.Add("@nCuotas", SqlDbType.Int).Value = .oCoopevalfamdeudas(i).nCuotas
                        oComm.Parameters.Add("@nCuotasPag", SqlDbType.Int).Value = .oCoopevalfamdeudas(i).nCuotasPag
                        oComm.Parameters.Add("@nCuotasPend", SqlDbType.Int).Value = .oCoopevalfamdeudas(i).nCuotasPend
                        oComm.Parameters.Add("@nMontoCuota", SqlDbType.Money).Value = .oCoopevalfamdeudas(i).nMontoCuota
                        oComm.Parameters.Add("@nSaldoCap", SqlDbType.Money).Value = .oCoopevalfamdeudas(i).nSaldoCap


                        oComm.ExecuteNonQuery()

                    Next
                End If

            End With

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




    Public Function RegistraEvaluacionUniFamDetBalance() As outNegocio

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


            With oPresolicitudApp.oCoopSolicitud.oCoopevalunidfam
                If Not IsNothing(.oCoopevalfamdetbal) Then

                    For i = 0 To .oCoopevalfamdetbal.Count - 1


                        oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalUniFamDetBalance", oConex)
                        oComm.CommandTimeout = 0
                        oComm.CommandType = CommandType.StoredProcedure

                        oComm.Parameters.Add("@nIDEvalUniFam", SqlDbType.Int).Value = .nIDEvalUniFam
                        oComm.Parameters.Add("@nCodCuenta", SqlDbType.Int).Value = .oCoopevalfamdetbal(i).nCodCuenta
                        oComm.Parameters.Add("@cCuenta", SqlDbType.VarChar, 150).Value = .oCoopevalfamdetbal(i).cCuenta
                        oComm.Parameters.Add("@nMoneda", SqlDbType.Int).Value = .oCoopevalfamdetbal(i).nMoneda
                        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = .oCoopevalfamdetbal(i).nMonto


                        oComm.ExecuteNonQuery()

                    Next
                End If

            End With

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


    Public Function RegistraEvaluacionUniFamDatosPersonales() As outNegocio

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


            With oPresolicitudApp.oCoopSolicitud.oCoopevalunidfam

                If Not IsNothing(.oCoopevalfamdatospers) Then


                    For i = 0 To .oCoopevalfamdatospers.Count() - 1

                        oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalUniFamDatosPersonales", oConex)
                        oComm.CommandTimeout = 0
                        oComm.CommandType = CommandType.StoredProcedure

                        oComm.Parameters.Add("@nIDEvalUniFam", SqlDbType.Int).Value = .nIDEvalUniFam
                        oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = .oCoopevalfamdatospers(i).nCodPers
                        oComm.Parameters.Add("@nTipoRelacionTit", SqlDbType.Int).Value = .oCoopevalfamdatospers(i).nTipoRelacionTit
                        oComm.Parameters.Add("@cEmpleador", SqlDbType.VarChar, 100).Value = .oCoopevalfamdatospers(i).cEmpleador
                        oComm.Parameters.Add("@nAniosExper", SqlDbType.Int).Value = .oCoopevalfamdatospers(i).nAniosExper
                        oComm.Parameters.Add("@nAniosServicio", SqlDbType.Int).Value = .oCoopevalfamdatospers(i).nAniosServicio
                        oComm.Parameters.Add("@nCondicionLab", SqlDbType.Int).Value = .oCoopevalfamdatospers(i).nCondicionLab
                        oComm.Parameters.Add("@nTipoContrato", SqlDbType.Int).Value = .oCoopevalfamdatospers(i).nTipoContrato
                        oComm.Parameters.Add("@nTipoSustentoIngr", SqlDbType.Int).Value = .oCoopevalfamdatospers(i).nTipoSustentoIngr

                        oComm.ExecuteNonQuery()

                    Next
                End If
            End With

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

    Public Function RegistraEvaluacionUnidadFamiliar() As outNegocio

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


            With oPresolicitudApp.oCoopSolicitud

                oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalUnidadFamiliar", oConex)
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure

                oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = .cDoc
                If Year(.oCoopevalunidfam.dFechaEval) = 1 Then
                    oComm.Parameters.Add("@dFechaEval", SqlDbType.Date).Value = Now()
                Else
                    oComm.Parameters.Add("@dFechaEval", SqlDbType.Date).Value = .oCoopevalunidfam.dFechaEval
                End If
                If Year(.oCoopevalunidfam.dFechaEval) = 1 Then
                    oComm.Parameters.Add("@dFechaVencEval", SqlDbType.Date).Value = Now.AddDays(180)
                Else
                    oComm.Parameters.Add("@dFechaVencEval", SqlDbType.Date).Value = .oCoopevalunidfam.dFechaVencEval
                End If

                oComm.Parameters.Add("@nDependientes", SqlDbType.Int).Value = .oCoopevalunidfam.nDependientes
                oComm.Parameters.Add("@nNumPadres", SqlDbType.Int).Value = .oCoopevalunidfam.nNumPadres
                oComm.Parameters.Add("@nNumHijos", SqlDbType.Int).Value = .oCoopevalunidfam.nNumHijos
                oComm.Parameters.Add("@nNumTerceros", SqlDbType.Int).Value = .oCoopevalunidfam.nNumTerceros
                oComm.Parameters.Add("@nTipoCambio", SqlDbType.Money).Value = .oCoopevalunidfam.nTipoCambio
                oComm.Parameters.Add("@nIDEvalUniFam", SqlDbType.Money).Direction = ParameterDirection.Output

                oComm.ExecuteNonQuery()
                .oCoopevalunidfam.nIDEvalUniFam = oComm.Parameters("@nIDEvalUniFam").Value


            End With
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

    Public Function RegistraEvaluacionAgricolaProyCostos() As outNegocio

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


            With oPresolicitudApp.oCoopSolicitud.oEvaluacion
                For i = 0 To .oEvalAgricola.Count - 1
                    For j = 0 To .oEvalAgricola(i).oEvalAgricolaProyVentas.Count - 1

                        oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalAgricolaProyCostos", oConex)
                        oComm.CommandTimeout = 0
                        oComm.CommandType = CommandType.StoredProcedure

                        oComm.Parameters.Add("@nIDProducto", SqlDbType.Int).Value = .oEvalAgricola(i).nIDProducto
                        oComm.Parameters.Add("@nIDEVal", SqlDbType.Int).Value = .nIDEval
                        oComm.Parameters.Add("@cRecurso", SqlDbType.VarChar, 100).Value = .oEvalAgricola(i).oevalagricolaproycostos(j).cRecurso
                        oComm.Parameters.Add("@nCodUnidad", SqlDbType.Int).Value = .oEvalAgricola(i).oevalagricolaproycostos(j).nCodUnidad
                        oComm.Parameters.Add("@nPrecioUnit", SqlDbType.Money).Value = .oEvalAgricola(i).oevalagricolaproycostos(j).nPrecioUnit
                        oComm.Parameters.Add("@nCantidad", SqlDbType.Int).Value = .oEvalAgricola(i).oevalagricolaproycostos(j).nCantidad
                        oComm.Parameters.Add("@nPeriodo", SqlDbType.Int).Value = .oEvalAgricola(i).oevalagricolaproycostos(j).nPeriodo
                        oComm.Parameters.Add("@dMesPeriodo", SqlDbType.Date).Value = .oEvalAgricola(i).oevalagricolaproycostos(j).dMesPeriodo

                        oComm.ExecuteNonQuery()

                    Next
                Next


            End With
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



    Public Function RegistraEvaluacionAgricolaProyVentas() As outNegocio

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


            With oPresolicitudApp.oCoopSolicitud.oEvaluacion
                For i = 0 To .oEvalAgricola.Count - 1
                    For j = 0 To .oEvalAgricola(i).oEvalAgricolaProyVentas.Count - 1

                        oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalAgricolaProyVentas", oConex)
                        oComm.CommandTimeout = 0
                        oComm.CommandType = CommandType.StoredProcedure

                        oComm.Parameters.Add("@nIDProducto", SqlDbType.Int).Value = .oEvalAgricola(i).nIDProducto
                        oComm.Parameters.Add("@nIDEVal", SqlDbType.Int).Value = .nIDEval
                        oComm.Parameters.Add("@nPeriodo", SqlDbType.Int).Value = .oEvalAgricola(i).oEvalAgricolaProyVentas(j).nPeriodo
                        oComm.Parameters.Add("@dMesPeriodo", SqlDbType.Date).Value = .oEvalAgricola(i).oEvalAgricolaProyVentas(j).dMesPeriodo
                        oComm.Parameters.Add("@nPorcVentas", SqlDbType.Money).Value = .oEvalAgricola(i).oEvalAgricolaProyVentas(j).nPorcVentas
                        oComm.Parameters.Add("@nMontoVenta", SqlDbType.Money).Value = .oEvalAgricola(i).oEvalAgricolaProyVentas(j).nMontoVenta

                        oComm.ExecuteNonQuery()

                    Next
                Next


            End With
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


    Public Function RegistraEvaluacionAgricola() As outNegocio

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


        Try


            With oPresolicitudApp.oCoopSolicitud.oEvaluacion
                For i = 0 To .oEvalAgricola.Count - 1

                    oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalAgricola", oConex)
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure

                    oComm.Parameters.Add("@nIDEval", SqlDbType.Int).Value = .nIDEval
                    oComm.Parameters.Add("@cProducto", SqlDbType.VarChar, 100).Value = .oEvalAgricola(i).cProducto
                    oComm.Parameters.Add("@nUnidMedVenta", SqlDbType.Int).Value = .oEvalAgricola(i).nUnidMedVenta
                    oComm.Parameters.Add("@nUnidMedSiembra", SqlDbType.Int).Value = .oEvalAgricola(i).nUnidMedSiembra
                    oComm.Parameters.Add("@nRendEsperado", SqlDbType.Money).Value = .oEvalAgricola(i).nRendEsperado
                    oComm.Parameters.Add("@nCantCultivar", SqlDbType.Money).Value = .oEvalAgricola(i).nCantCultivar
                    oComm.Parameters.Add("@nVentaEsperada", SqlDbType.Money).Value = .oEvalAgricola(i).nVentaEsperada
                    oComm.Parameters.Add("@nCicloProductivo", SqlDbType.Int).Value = .oEvalAgricola(i).nCicloProductivo
                    oComm.Parameters.Add("@dMesInicio", SqlDbType.Date).Value = .oEvalAgricola(i).dMesInicio
                    oComm.Parameters.Add("@nPreXQuintalCalidadA", SqlDbType.Money).Value = .oEvalAgricola(i).nPreXQuintalCalidadA
                    oComm.Parameters.Add("@nPorcentajeCalidadA", SqlDbType.Money).Value = .oEvalAgricola(i).nPorcentajeCalidadA
                    oComm.Parameters.Add("@nPreXQuintalCalidadB", SqlDbType.Money).Value = .oEvalAgricola(i).nPreXQuintalCalidadB
                    oComm.Parameters.Add("@nPorcentajeCalidadB", SqlDbType.Money).Value = .oEvalAgricola(i).nPorcentajeCalidadB
                    oComm.Parameters.Add("@nPreXQuintalCalidadC", SqlDbType.Money).Value = .oEvalAgricola(i).nPreXQuintalCalidadC
                    oComm.Parameters.Add("@nPorcentajeCalidadC", SqlDbType.Money).Value = .oEvalAgricola(i).nPorcentajeCalidadC
                    oComm.Parameters.Add("@nPreXQuintalCalidadD", SqlDbType.Money).Value = .oEvalAgricola(i).nPreXQuintalCalidadD
                    oComm.Parameters.Add("@nPorcentajeCalidadD", SqlDbType.Money).Value = .oEvalAgricola(i).nPorcentajeCalidadD
                    oComm.Parameters.Add("@nBroza", SqlDbType.Money).Value = .oEvalAgricola(i).nBroza
                    oComm.Parameters.Add("@nIDEvalAgricola", SqlDbType.Int).Direction = ParameterDirection.Output

                    oComm.ExecuteNonQuery()

                    .oEvalAgricola(i).nIDProducto = oComm.Parameters("@nIDEvalAgricola").Value

                Next


            End With
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


    Public Function RegistraEvaluacionPymeEvalGiroVentas() As outNegocio

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


        Try

            For z = 0 To oPresolicitudApp.oCoopSolicitud.oEvaluacion.oEvalgiros.Count - 1

                With oPresolicitudApp.oCoopSolicitud.oEvaluacion
                    For i = 0 To .oEvalgiros(z).oEvalventas.Count - 1

                        oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalGiroVentas", oConex)
                        oComm.CommandTimeout = 0
                        oComm.CommandType = CommandType.StoredProcedure

                        oComm.Parameters.Add("@nIDEval", SqlDbType.Int).Value = .nIDEval
                        oComm.Parameters.Add("@nIDGiro", SqlDbType.Int).Value = .oEvalgiros(z).nIDGiro
                        oComm.Parameters.Add("@nCodDia", SqlDbType.Int).Value = .oEvalgiros(z).oEvalventas(i).nCodDia
                        oComm.Parameters.Add("@cDia", SqlDbType.VarChar, 50).Value = .oEvalgiros(z).oEvalventas(i).cDia
                        oComm.Parameters.Add("@nVentaMin", SqlDbType.Money).Value = .oEvalgiros(z).oEvalventas(i).nVentaMin
                        oComm.Parameters.Add("@nVentaMax", SqlDbType.Money).Value = .oEvalgiros(z).oEvalventas(i).nVentaMax

                        oComm.ExecuteNonQuery()

                    Next


                End With
            Next
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


    Public Function RegistraEvaluacionPymeEvalGiroGastosOper() As outNegocio

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
        Dim i, j, z As Integer

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
            For z = 0 To oPresolicitudApp.oCoopSolicitud.oEvaluacion.oEvalgiros.Count - 1

                With oPresolicitudApp.oCoopSolicitud.oEvaluacion
                    For i = 0 To .oEvalgiros(z).oEvalgasoper.Count - 1

                        oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalGiroGastosOper", oConex)
                        oComm.CommandTimeout = 0
                        oComm.CommandType = CommandType.StoredProcedure

                        oComm.Parameters.Add("@nIDEval", SqlDbType.Int).Value = .nIDEval
                        oComm.Parameters.Add("@nIDGiro", SqlDbType.Int).Value = .oEvalgiros(z).nIDGiro
                        oComm.Parameters.Add("@nCodgasto", SqlDbType.Int).Value = .oEvalgiros(z).oEvalgasoper(i).nCodGasto
                        oComm.Parameters.Add("@cGasto", SqlDbType.VarChar, 100).Value = .oEvalgiros(z).oEvalgasoper(i).cGasto
                        oComm.Parameters.Add("@nMoneda", SqlDbType.Int).Value = .oEvalgiros(z).oEvalgasoper(i).nMoneda
                        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = .oEvalgiros(z).oEvalgasoper(i).nMonto

                        oComm.ExecuteNonQuery()

                    Next


                End With
            Next
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

    Public Function RegistraEvaluacionPymeEvalGiroInventario() As outNegocio

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
        Dim i, j, z As Integer

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

            For z = 0 To oPresolicitudApp.oCoopSolicitud.oEvaluacion.oEvalgiros.Count - 1
                With oPresolicitudApp.oCoopSolicitud.oEvaluacion
                    For i = 0 To .oEvalgiros(z).oEvalinventario.Count - 1

                        oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalGiroInventario", oConex)
                        oComm.CommandTimeout = 0
                        oComm.CommandType = CommandType.StoredProcedure

                        oComm.Parameters.Add("@nIDEval", SqlDbType.Int).Value = .nIDEval
                        oComm.Parameters.Add("@nIDGiro", SqlDbType.Int).Value = .oEvalgiros(z).nIDGiro
                        oComm.Parameters.Add("@nTipoInv", SqlDbType.Int).Value = .oEvalgiros(z).oEvalinventario(i).nTipoInv
                        oComm.Parameters.Add("@cProducto", SqlDbType.VarChar, 100).Value = .oEvalgiros(z).oEvalinventario(i).cProducto
                        oComm.Parameters.Add("@nCodUnidad", SqlDbType.Int).Value = .oEvalgiros(z).oEvalinventario(i).nCodUnidad
                        oComm.Parameters.Add("@nCantidad", SqlDbType.Int).Value = .oEvalgiros(z).oEvalinventario(i).nCantidad
                        oComm.Parameters.Add("@nPreUnit", SqlDbType.Money).Value = .oEvalgiros(z).oEvalinventario(i).nPreUnit
                        oComm.Parameters.Add("@nMoneda", SqlDbType.Int).Value = .oEvalgiros(z).oEvalinventario(i).nMoneda

                        oComm.ExecuteNonQuery()

                    Next


                End With
            Next
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

    Public Function RegistraEvaluacionPymeEvalGirosCosteoProd() As outNegocio

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
        Dim i, j, z As Integer

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

            For z = 0 To oPresolicitudApp.oCoopSolicitud.oEvaluacion.oEvalgiros.Count - 1

                With oPresolicitudApp.oCoopSolicitud.oEvaluacion.oEvalgiros(z)

                    For i = 0 To .oEvalcostoventa.Count - 1

                        oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalGiroCostoVenta", oConex)
                        oComm.CommandTimeout = 0
                        oComm.CommandType = CommandType.StoredProcedure

                        oComm.Parameters.Add("@nIDEval", SqlDbType.Int).Value = .nIDEVal
                        oComm.Parameters.Add("@nIDGiro", SqlDbType.Int).Value = .nIDGiro
                        oComm.Parameters.Add("@cProducto", SqlDbType.VarChar, 100).Value = .oEvalcostoventa(i).cProducto
                        oComm.Parameters.Add("@nPreCompra", SqlDbType.Money).Value = .oEvalcostoventa(i).nPreCompra
                        oComm.Parameters.Add("@nPreVenta", SqlDbType.Money).Value = .oEvalcostoventa(i).nPreVenta
                        oComm.Parameters.Add("@nCodUnidad", SqlDbType.Int).Value = .oEvalcostoventa(i).nCodUnidad
                        oComm.Parameters.Add("@nPorcProporVentas", SqlDbType.Money).Value = .oEvalcostoventa(i).nPorcProporVentas
                        oComm.Parameters.Add("@nCantidadCostear", SqlDbType.Int).Value = .oEvalcostoventa(i).nCantidadCostear
                        oComm.Parameters.Add("@nRelaCompVenta", SqlDbType.Money).Value = .oEvalcostoventa(i).nRelaCompVenta

                        oComm.Parameters.Add("@nIDCosVen", SqlDbType.Int).Direction = ParameterDirection.Output

                        oComm.ExecuteNonQuery()
                        .oEvalcostoventa(i).nIDCosVen = oComm.Parameters("@nIDCosVen").Value

                        For j = 0 To .oEvalcostoventa(i).oEvalCosteoProdDet.Count - 1

                            oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalCosteoProductoDetalle", oConex)
                            oComm.CommandTimeout = 0
                            oComm.CommandType = CommandType.StoredProcedure

                            oComm.Parameters.Add("@nIDEval", SqlDbType.Int).Value = .nIDEVal
                            oComm.Parameters.Add("@nIDGiro", SqlDbType.Int).Value = .nIDGiro
                            oComm.Parameters.Add("@nIDCosVen", SqlDbType.Int).Value = .oEvalcostoventa(i).nIDCosVen
                            oComm.Parameters.Add("@cInsumo", SqlDbType.VarChar, 100).Value = .oEvalcostoventa(i).oEvalCosteoProdDet(j).cInsumo
                            oComm.Parameters.Add("@nCantidad", SqlDbType.Int).Value = .oEvalcostoventa(i).oEvalCosteoProdDet(j).nCantidad
                            oComm.Parameters.Add("@nCodUnidad", SqlDbType.Money).Value = .oEvalcostoventa(i).oEvalCosteoProdDet(j).nCodUnidad
                            oComm.Parameters.Add("@nPrecioCompra", SqlDbType.Money).Value = .oEvalcostoventa(i).oEvalCosteoProdDet(j).nPrecioCompra
                            oComm.Parameters.Add("@nSubTotal", SqlDbType.Money).Value = .oEvalcostoventa(i).oEvalCosteoProdDet(j).nSubTotal

                            oComm.ExecuteNonQuery()

                        Next

                    Next



                End With
            Next
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



    Public Function RegistraEvaluacionPymeEvalGiros() As outNegocio

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


        Try


            With oPresolicitudApp.oCoopSolicitud.oEvaluacion
                For i = 0 To .oEvalgiros.Count - 1

                    oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalGiros", oConex)
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure

                    oComm.Parameters.Add("@nIDEval", SqlDbType.Int).Value = .nIDEval
                    oComm.Parameters.Add("@nCodGiro", SqlDbType.Int).Value = .oEvalgiros(i).nCodGiro
                    oComm.Parameters.Add("@cDescripcion", SqlDbType.VarChar, 150).Value = .oEvalgiros(i).cDescripcion
                    oComm.Parameters.Add("@nVentas", SqlDbType.Money).Value = .oEvalgiros(i).nVentas
                    oComm.Parameters.Add("@nCostoVentas", SqlDbType.Money).Value = .oEvalgiros(i).nCostoVentas
                    oComm.Parameters.Add("@nMargen", SqlDbType.Money).Value = .oEvalgiros(i).nMargen
                    oComm.Parameters.Add("@nPorcPart", SqlDbType.Money).Value = .oEvalgiros(i).nPorcPart
                    oComm.Parameters.Add("@nUtilidad", SqlDbType.Money).Value = .oEvalgiros(i).nUtilidad
                    oComm.Parameters.Add("@nGasOper", SqlDbType.Money).Value = .oEvalgiros(i).nGasOper
                    oComm.Parameters.Add("@nPorcCred", SqlDbType.Money).Value = .oEvalgiros(i).nPorcCred
                    oComm.Parameters.Add("@nVentaCred", SqlDbType.Money).Value = .oEvalgiros(i).nVentaCred
                    oComm.Parameters.Add("@nProdTermin", SqlDbType.Money).Value = .oEvalgiros(i).nProdTermin
                    oComm.Parameters.Add("@nProdProceso", SqlDbType.Money).Value = .oEvalgiros(i).nProdProceso
                    oComm.Parameters.Add("@nMateriales", SqlDbType.Money).Value = .oEvalgiros(i).nMateriales
                    oComm.Parameters.Add("@nIDGiro", SqlDbType.Int).Direction = ParameterDirection.Output

                    oComm.ExecuteNonQuery()
                    .oEvalgiros(i).nIDGiro = oComm.Parameters("@nIDGiro").Value

                Next


            End With
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

    Public Function RegistraEvaluacionPymeIngresosEgresosEstac() As outNegocio

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


        Try


            With oPresolicitudApp.oCoopSolicitud.oEvaluacion
                For i = 0 To .oEvalIngEgrEstac.Count - 1

                    oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalingregreestac", oConex)
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure

                    oComm.Parameters.Add("@nIDEval", SqlDbType.Int).Value = .nIDEval
                    oComm.Parameters.Add("@dMes", SqlDbType.Date).Value = .oEvalIngEgrEstac(i).dMes
                    oComm.Parameters.Add("@nTipoIngEgr", SqlDbType.Int).Value = .oEvalIngEgrEstac(i).nTipoIngEgr
                    oComm.Parameters.Add("@cDescripcion", SqlDbType.VarChar, 150).Value = .oEvalIngEgrEstac(i).cDescripcion
                    oComm.Parameters.Add("@nMoneda", SqlDbType.Int).Value = .oEvalIngEgrEstac(i).nMoneda
                    oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = .oEvalIngEgrEstac(i).nMonto
                    oComm.Parameters.Add("@nMontoMN", SqlDbType.Money).Value = .oEvalIngEgrEstac(i).nMontoMN

                    oComm.ExecuteNonQuery()

                Next

            End With



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

    Public Function RegistraEvaluacionPymeOtrosNegocios() As outNegocio

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


        Try

            With oPresolicitudApp.oCoopSolicitud.oEvaluacion
                For i = 0 To .oevalotrosneg.Count - 1

                    oComm = New SqlClient.SqlCommand("PA_CoopregistraOtrosNeg", oConex)
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure


                    oComm.Parameters.Add("@nIDEval", SqlDbType.Int).Value = .nIDEval
                    oComm.Parameters.Add("@cDescripcion", SqlDbType.VarChar, 100).Value = .oevalotrosneg(i).cDescripcion
                    oComm.Parameters.Add("@nVentas", SqlDbType.Money).Value = .oevalotrosneg(i).nVentas
                    oComm.Parameters.Add("@nCostoVentas", SqlDbType.Money).Value = .oevalotrosneg(i).nCostoVentas
                    oComm.Parameters.Add("@nGastosOper", SqlDbType.Money).Value = .oevalotrosneg(i).nGastosOper
                    oComm.Parameters.Add("@nGastosFinanc", SqlDbType.Money).Value = .oevalotrosneg(i).nGastosFinanc
                    oComm.Parameters.Add("@nUtilidad", SqlDbType.Money).Value = .oevalotrosneg(i).nUtilidad
                    oComm.Parameters.Add("@nMoneda", SqlDbType.Int).Value = .oevalotrosneg(i).nMoneda
                    oComm.Parameters.Add("@nUtilidadMN", SqlDbType.Money).Value = .oevalotrosneg(i).nUtilidadMN

                    oComm.ExecuteNonQuery()

                Next

            End With



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


    Public Function RegistraEvaluacionPymeEvaDeudas() As outNegocio

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


        Try


            With oPresolicitudApp.oCoopSolicitud.oEvaluacion
                For i = 0 To .oEvalDeudas.Count - 1

                    oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalDeudas", oConex)
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure

                    oComm.Parameters.Add("@nIDEval", SqlDbType.Int).Value = .nIDEval
                    oComm.Parameters.Add("@cEntidad", SqlDbType.VarChar, 100).Value = .oEvalDeudas(i).cEntidad
                    oComm.Parameters.Add("@nTipoCred", SqlDbType.Int).Value = .oEvalDeudas(i).nTipoCred
                    oComm.Parameters.Add("@nMoneda", SqlDbType.Int).Value = .oEvalDeudas(i).nMoneda
                    oComm.Parameters.Add("@nCuotas", SqlDbType.Int).Value = .oEvalDeudas(i).nCuotas
                    oComm.Parameters.Add("@nCuotasPag", SqlDbType.Int).Value = .oEvalDeudas(i).nCuotasPag
                    oComm.Parameters.Add("@nCuotasPend", SqlDbType.Int).Value = .oEvalDeudas(i).nCuotasPend
                    oComm.Parameters.Add("@nMontoCuota", SqlDbType.Money).Value = .oEvalDeudas(i).nMontoCuota
                    oComm.Parameters.Add("@nMontoCuotaMN", SqlDbType.Money).Value = .oEvalDeudas(i).nMontoCuotaMN
                    oComm.Parameters.Add("@nSaldoCap", SqlDbType.Money).Value = .oEvalDeudas(i).nSaldoCap
                    oComm.Parameters.Add("@nSaldoCapMN", SqlDbType.Money).Value = .oEvalDeudas(i).nSaldoCapMN
                    oComm.Parameters.Add("@nSaldoCapCP", SqlDbType.Money).Value = .oEvalDeudas(i).nSaldoCapCP
                    oComm.Parameters.Add("@nSaldoCapLP", SqlDbType.Money).Value = .oEvalDeudas(i).nSaldoCapLP

                    oComm.ExecuteNonQuery()

                Next

            End With



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


    Public Function RegistraEvaluacionPymeSemovientes() As outNegocio

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


        Try
            If IsNothing(oPresolicitudApp.oCoopSolicitud.oEvaluacion) Then
                nResp = 1
                sRespuesta = "REGISTRO CORRECTO"
            End If


            With oPresolicitudApp.oCoopSolicitud.oEvaluacion
                For i = 0 To .oEvalsemovientes.Count - 1

                    oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalSemovientes", oConex)
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure

                    oComm.Parameters.Add("@nIDEval", SqlDbType.Int).Value = .nIDEval
                    oComm.Parameters.Add("@cDescripcion", SqlDbType.VarChar, 150).Value = .oEvalsemovientes(i).cDescripcion
                    oComm.Parameters.Add("@nUnidades", SqlDbType.Int).Value = .oEvalsemovientes(i).nUnidades
                    oComm.Parameters.Add("@nCantidad", SqlDbType.Int).Value = .oEvalsemovientes(i).nCantidad
                    oComm.Parameters.Add("@nPrecioUnit", SqlDbType.Money).Value = .oEvalsemovientes(i).nPrecioUnit
                    oComm.Parameters.Add("@nMoneda", SqlDbType.Int).Value = .oEvalsemovientes(i).nMoneda
                    oComm.Parameters.Add("@nImporte", SqlDbType.Money).Value = .oEvalsemovientes(i).nImporte
                    oComm.Parameters.Add("@nImporteMN", SqlDbType.Money).Value = .oEvalsemovientes(i).nImporteMN

                    oComm.ExecuteNonQuery()

                Next

            End With



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


    Public Function RegistraEvaluacionPymeADetalleBalance() As outNegocio

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


        Try


            oComm = New SqlClient.SqlCommand("PA_CoopRegistraEvalPyme", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = oPresolicitudApp.oCoopSolicitud.oCoopPersona.nID
            oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = oPresolicitudApp.oCoopSolicitud.oCoopPersona.cDoc
            oComm.Parameters.Add("@dFecRegistro", SqlDbType.Date).Value = oPresolicitudApp.oCoopSolicitud.oEvaluacion.dFecRegistro
            oComm.Parameters.Add("@dFecVenc", SqlDbType.Date).Value = oPresolicitudApp.oCoopSolicitud.oEvaluacion.dFecRegistro.AddDays(180)
            oComm.Parameters.Add("@nTipCam", SqlDbType.Money).Value = oPresolicitudApp.oCoopSolicitud.oEvaluacion.nTipCam
            oComm.Parameters.Add("@dFecEEFF", SqlDbType.Date).Value = oPresolicitudApp.oCoopSolicitud.oEvaluacion.dFecEEFF
            oComm.Parameters.Add("@nTipoPropiedad", SqlDbType.Int).Value = oPresolicitudApp.oCoopSolicitud.oEvaluacion.nTipoPropiedad
            oComm.Parameters.Add("@nTiempoUsoLoc", SqlDbType.Int).Value = oPresolicitudApp.oCoopSolicitud.oEvaluacion.nTiempoUsoLoc
            oComm.Parameters.Add("@nTiempoExpNeg", SqlDbType.Int).Value = oPresolicitudApp.oCoopSolicitud.oEvaluacion.nTiempoExpNeg
            oComm.Parameters.Add("@nHorasDedic", SqlDbType.Int).Value = oPresolicitudApp.oCoopSolicitud.oEvaluacion.nHorasDedic
            oComm.Parameters.Add("@nHorasAtencion", SqlDbType.Int).Value = oPresolicitudApp.oCoopSolicitud.oEvaluacion.nHorasAtencion
            oComm.Parameters.Add("@nNumColab", SqlDbType.Int).Value = oPresolicitudApp.oCoopSolicitud.oEvaluacion.nNumColab
            oComm.Parameters.Add("@nIDEval", SqlDbType.Int).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            oPresolicitudApp.oCoopSolicitud.oEvaluacion.nIDEval = oComm.Parameters("@nIDEval").Value

            With oPresolicitudApp.oCoopSolicitud.oEvaluacion
                For i = 0 To .oEvalDetBalance.Count - 1

                    oComm = New SqlClient.SqlCommand("PA_CoopRegistraDetalleBalance", oConex)
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure

                    oComm.Parameters.Add("@nIDEval", SqlDbType.Int).Value = .nIDEval
                    oComm.Parameters.Add("@nCodConcepto", SqlDbType.Int).Value = .oEvalDetBalance(i).nCodConcepto
                    oComm.Parameters.Add("@cDescripcion", SqlDbType.VarChar, 150).Value = .oEvalDetBalance(i).cDescripcion
                    oComm.Parameters.Add("@nMoneda", SqlDbType.Int).Value = .oEvalDetBalance(i).nMoneda
                    oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = .oEvalDetBalance(i).nMonto
                    oComm.Parameters.Add("@nMontoMN", SqlDbType.Money).Value = .oEvalDetBalance(i).nMontoMN

                    oComm.ExecuteNonQuery()

                Next

            End With



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


    Public Function RegistraEvaluacionPymeApp() As outNegocio

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
        Dim nCodPersTitutal As Long
        Dim nCodAgeTit As Integer
        Dim nCodCredTit As Integer
        Dim nErrorEval As Integer
        Dim nCodEvalResp As Integer

        nResp = 0
        sRespuesta = ""

        sCadConexion = oConect.gsCadConexDBSolida

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try

            oComm = New SqlClient.SqlCommand("PA_ActualizaRelaCompraVenta", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = oPresolicitudApp.oCoopSolicitud.oEvaluacion.cDoc

            oComm.ExecuteNonQuery()


            With oPresolicitudApp.oCoopSolicitud

                oComm = New SqlClient.SqlCommand("App_DevuelveCopersonaTitutalYCredito", oConex)
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure

                oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = .oEvaluacion.cDoc
                oComm.Parameters.Add("@nCodAgeUsu", SqlDbType.Int).Value = nCodAge
                oComm.Parameters.Add("@nCodpers", SqlDbType.Int).Direction = ParameterDirection.Output
                oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Direction = ParameterDirection.Output
                oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Direction = ParameterDirection.Output

                oComm.ExecuteNonQuery()

                nCodPersTitutal = oComm.Parameters("@nCodpers").Value
                nCodCredTit = oComm.Parameters("@nCodCred").Value
                nCodAgeTit = oComm.Parameters("@nCodAge").Value



                oComm = New SqlClient.SqlCommand("App_Eval_PersEvaluacion_Insertar", oConex)
                    oComm.CommandTimeout = 0
                    oComm.CommandType = CommandType.StoredProcedure

                    oComm.Parameters.Add("@dFecEval", SqlDbType.Date).Value = .oEvaluacion.dFecEEFF
                    oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = nCodPersTitutal
                    oComm.Parameters.Add("@cDescripcion", SqlDbType.VarChar, 200).Value = .oEvaluacion.cDescripcion
                    oComm.Parameters.Add("@cNomNegocio", SqlDbType.VarChar, 50).Value = .oEvaluacion.cNomNegocio
                    oComm.Parameters.Add("@cDireccion", SqlDbType.VarChar, 50).Value = .oEvaluacion.cDireccion
                    oComm.Parameters.Add("@cTelefono", SqlDbType.VarChar, 50).Value = .oEvaluacion.cTelefono
                    oComm.Parameters.Add("@nTipoFteIngreso", SqlDbType.VarChar, 50).Value = .oEvaluacion.nTipoFuente
                    oComm.Parameters.Add("@cCodZona", SqlDbType.VarChar, 50).Value = .oEvaluacion.cCodZona
                    oComm.Parameters.Add("@bValidaParamFecha", SqlDbType.Bit).Value = 0
                    oComm.Parameters.Add("@nError", SqlDbType.Int).Direction = ParameterDirection.Output
                    oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Direction = ParameterDirection.Output
                    oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCredTit
                    oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAgeTit
                    oComm.Parameters.Add("@nAnioxExperiencia", SqlDbType.Int).Value = oPresolicitudApp.oCoopSolicitud.oEvaluacion.nTiempoExpNeg
                    oComm.Parameters.Add("@nCantEntidades", SqlDbType.Int).Value = oPresolicitudApp.oCoopSolicitud.oEvaluacion.oEvalDeudas.Count
                    oComm.Parameters.Add("@nCalificacionExterna", SqlDbType.Int).Value = oPresolicitudApp.oCoopSolicitud.oEvaluacion.nCalificacionExterna
                    oComm.Parameters.Add("@nLatitud", SqlDbType.Float).Value = oPresolicitudApp.oCoopSolicitud.oEvaluacion.nLatitud
                    oComm.Parameters.Add("@nLongitud", SqlDbType.Float).Value = oPresolicitudApp.oCoopSolicitud.oEvaluacion.nLongitud

                    oComm.ExecuteNonQuery()

                nErrorEval = oComm.Parameters("@nError").Value
                nCodEvalResp = oComm.Parameters("@nCodEval").Value

            End With
            If nErrorEval > 0 Then
                nResp = 0
            Else
                nResp = 1
            End If


            '********************************* Inserta Balance *********************************************            

            oComm = New SqlClient.SqlCommand("App_RecuperaDatosBalance", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@psDoc", SqlDbType.Int).Value = oPresolicitudApp.oCoopSolicitud.oEvaluacion.cDoc
            oComm.Parameters.Add("@nDisponible", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nAdelantoProveedores", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nCuentasXCobrar", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nInventario", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nMercaderiaTran", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nBienesMueb", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nBienesInmue", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nCtasPagar", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nAdelantoCli", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nOtrasCtasPag", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nPasivoNoC", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nCtasPagarLP", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nOtrasCPagarLP", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nCapital", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nReservas", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nDistribUtilidad", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nMateriales", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nMatPrima", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nMatProceso", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nProdTerminado", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nInversiones", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nDeudaActivoF", SqlDbType.Money).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            Dim nDisponible As Double
            nDisponible = oComm.Parameters("@nDisponible").Value

            Dim nAdelantoProveedores As Double
            nAdelantoProveedores = oComm.Parameters("@nAdelantoProveedores").Value

            Dim nCuentasXCobrar As Double
            nCuentasXCobrar = oComm.Parameters("@nCuentasXCobrar").Value

            Dim nInventario As Double
            nInventario = oComm.Parameters("@nInventario").Value

            Dim nMercaderiaTran As Double
            nMercaderiaTran = oComm.Parameters("@nMercaderiaTran").Value

            Dim nBienesMueb As Double
            nBienesMueb = oComm.Parameters("@nBienesMueb").Value

            Dim nBienesInmue As Double
            nBienesInmue = oComm.Parameters("@nBienesInmue").Value

            Dim nCtasPagar As Double
            nCtasPagar = oComm.Parameters("@nCtasPagar").Value

            Dim nAdelantoCli As Double
            nAdelantoCli = oComm.Parameters("@nAdelantoCli").Value


            Dim nOtrasCtasPag As Double
            nOtrasCtasPag = oComm.Parameters("@nOtrasCtasPag").Value

            Dim nPasivoNoC As Double
            nPasivoNoC = oComm.Parameters("@nPasivoNoC").Value

            Dim nCtasPagarLP As Double
            nCtasPagarLP = oComm.Parameters("@nCtasPagarLP").Value

            Dim nOtrasCPagarLP As Double
            nOtrasCPagarLP = oComm.Parameters("@nOtrasCPagarLP").Value

            Dim nCapital As Double
            nCapital = oComm.Parameters("@nCapital").Value

            Dim nReservas As Double
            nReservas = oComm.Parameters("@nReservas").Value

            Dim nDistribUtilidad As Double
            nDistribUtilidad = oComm.Parameters("@nDistribUtilidad").Value

            Dim nMateriales As Double
            nMateriales = oComm.Parameters("@nMateriales").Value

            Dim nMatPrima As Double
            nMatPrima = oComm.Parameters("@nMatPrima").Value

            Dim nMatProceso As Double
            nMatProceso = oComm.Parameters("@nMatProceso").Value

            Dim nProdTerminado As Double
            nProdTerminado = oComm.Parameters("@nProdTerminado").Value

            Dim nInversiones As Double
            nInversiones = oComm.Parameters("@nInversiones").Value

            Dim nDeudaActivoF As Double
            nDeudaActivoF = oComm.Parameters("@nDeudaActivoF").Value


            oComm = New SqlClient.SqlCommand("Eval_PersEvalBalGen_Grabar_NEW", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEvalResp
            oComm.Parameters.Add("@nDisponible", SqlDbType.Money).Value = nDisponible
            oComm.Parameters.Add("@nAdelantoProv", SqlDbType.Money).Value = nAdelantoProveedores
            oComm.Parameters.Add("@nCtasCobrar", SqlDbType.Money).Value = nCuentasXCobrar
            oComm.Parameters.Add("@nInventario", SqlDbType.Money).Value = nInventario
            oComm.Parameters.Add("@nMercaderiaTran", SqlDbType.Money).Value = nMercaderiaTran
            oComm.Parameters.Add("@nBienesMueb", SqlDbType.Money).Value = nBienesMueb
            oComm.Parameters.Add("@nBienesInmue", SqlDbType.Money).Value = nBienesInmue
            oComm.Parameters.Add("@nCtasPagar", SqlDbType.Money).Value = nCtasPagar
            oComm.Parameters.Add("@nAdelantoCli", SqlDbType.Money).Value = nAdelantoCli
            oComm.Parameters.Add("@nOtrasCtasPag", SqlDbType.Money).Value = nOtrasCtasPag
            oComm.Parameters.Add("@nPasivoNoC", SqlDbType.Money).Value = nPasivoNoC

            oComm.Parameters.Add("@nCtasPagarLP", SqlDbType.Money).Value = nCtasPagarLP
            oComm.Parameters.Add("@nOtrasCPagarLP", SqlDbType.Money).Value = nOtrasCPagarLP
            oComm.Parameters.Add("@nCapital", SqlDbType.Money).Value = nCapital
            oComm.Parameters.Add("@nReservas", SqlDbType.Money).Value = nReservas
            oComm.Parameters.Add("@nDistribUtilidad", SqlDbType.Money).Value = nDistribUtilidad

            '02/08/2012
            oComm.Parameters.Add("@nMateriales", SqlDbType.Money).Value = nMateriales
            oComm.Parameters.Add("@nMatPrima", SqlDbType.Money).Value = nMatPrima
            oComm.Parameters.Add("@nMatProceso", SqlDbType.Money).Value = nMatProceso
            oComm.Parameters.Add("@nProdTerminado", SqlDbType.Money).Value = nProdTerminado

            oComm.Parameters.Add("@nInversiones", SqlDbType.Money).Value = nInversiones
            oComm.Parameters.Add("@nDeudaActivoF", SqlDbType.Money).Value = nDeudaActivoF

            oComm.Parameters.Add("@bInversiones", SqlDbType.Bit).Value = 1
            oComm.Parameters.Add("@bMuebles", SqlDbType.Bit).Value = 1
            oComm.Parameters.Add("@bInmuebles", SqlDbType.Bit).Value = 1

            oComm.ExecuteNonQuery()



            '********************************* Inserta Estado de Ganancias y Perdidas *********************************************            
            oComm = New SqlClient.SqlCommand("App_RecuperaEstadoGanaciasPerdidas", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = oPresolicitudApp.oCoopSolicitud.oEvaluacion.cDoc
            oComm.Parameters.Add("@nVentaContado", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nVentaCredito", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nPromDiasOtorCred", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nCompraContado", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nCompraCredito", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nPromDiasPagCred", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nVentasMensuales", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nCostoVentas", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nGastoPersonal", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nServicios", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nAlquilerLocal", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nTrasporteMer", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nImpuestos", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nGastosFinacNeg", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nOtrosGastos", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nOtrosIngresosFam", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nTotalGastosFam", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nUtiNegSecundario", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nOtrosGastosFam", SqlDbType.Money).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            Dim nVentaContado As Double
            nVentaContado = oComm.Parameters("@nVentaContado").Value

            Dim nVentaCredito As Double
            nVentaCredito = oComm.Parameters("@nVentaCredito").Value

            Dim nPromDiasOtorCred As Double
            nPromDiasOtorCred = oComm.Parameters("@nPromDiasOtorCred").Value

            Dim nCompraContado As Double
            nCompraContado = oComm.Parameters("@nCompraContado").Value

            Dim nCompraCredito As Double
            nCompraCredito = oComm.Parameters("@nCompraCredito").Value

            Dim nPromDiasPagCred As Double
            nPromDiasPagCred = oComm.Parameters("@nPromDiasPagCred").Value

            Dim nVentasMensuales As Double
            nVentasMensuales = oComm.Parameters("@nVentasMensuales").Value

            Dim nCostoVentas As Double
            nCostoVentas = oComm.Parameters("@nCostoVentas").Value

            Dim nGastoPersonal As Double
            nGastoPersonal = oComm.Parameters("@nGastoPersonal").Value

            Dim nServicios As Double
            nServicios = oComm.Parameters("@nServicios").Value

            Dim nAlquilerLocal As Double
            nAlquilerLocal = oComm.Parameters("@nAlquilerLocal").Value

            Dim nTrasporteMer As Double
            nTrasporteMer = oComm.Parameters("@nTrasporteMer").Value

            Dim nImpuestos As Double
            nImpuestos = oComm.Parameters("@nImpuestos").Value

            Dim nGastosFinacNeg As Double
            nGastosFinacNeg = oComm.Parameters("@nGastosFinacNeg").Value

            Dim nOtrosGastos As Double
            nOtrosGastos = oComm.Parameters("@nOtrosGastos").Value

            Dim nOtrosIngresosFam As Double
            nOtrosIngresosFam = oComm.Parameters("@nOtrosIngresosFam").Value

            Dim nTotalGastosFam As Double
            nTotalGastosFam = oComm.Parameters("@nTotalGastosFam").Value

            Dim nUtiNegSecundario As Double
            nUtiNegSecundario = oComm.Parameters("@nUtiNegSecundario").Value

            Dim nOtrosGastosFam As Double
            nOtrosGastosFam = oComm.Parameters("@nOtrosGastosFam").Value



            oComm = New SqlClient.SqlCommand("Eval_PersEvalEstGanPerd_Grabar", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEvalResp
            oComm.Parameters.Add("@nVentaContado", SqlDbType.Money).Value = nVentaContado
            oComm.Parameters.Add("@nVentaCredito", SqlDbType.Money).Value = nVentaCredito
            oComm.Parameters.Add("@nPromDiasOtorCred", SqlDbType.Int).Value = nPromDiasOtorCred
            oComm.Parameters.Add("@nCompraContado", SqlDbType.Money).Value = nCompraContado
            oComm.Parameters.Add("@nCompraCredito", SqlDbType.Money).Value = nCompraCredito
            oComm.Parameters.Add("@nPromDiasPagCred", SqlDbType.Int).Value = nPromDiasPagCred
            oComm.Parameters.Add("@nVentasMensuales", SqlDbType.Money).Value = nVentasMensuales
            oComm.Parameters.Add("@nCostoVentas", SqlDbType.Money).Value = nCostoVentas
            oComm.Parameters.Add("@nGastoPersonal", SqlDbType.Money).Value = nGastoPersonal
            oComm.Parameters.Add("@nServicios", SqlDbType.Money).Value = nServicios
            oComm.Parameters.Add("@nAlquilerLocal", SqlDbType.Money).Value = nAlquilerLocal
            oComm.Parameters.Add("@nTrasporteMer", SqlDbType.Money).Value = nTrasporteMer
            oComm.Parameters.Add("@nImpuestos", SqlDbType.Money).Value = nImpuestos
            oComm.Parameters.Add("@nGastosFinacNeg", SqlDbType.Money).Value = nGastosFinacNeg
            oComm.Parameters.Add("@nOtrosGastos", SqlDbType.Money).Value = nOtrosGastos
            oComm.Parameters.Add("@nOtrosIngresosFam", SqlDbType.Money).Value = nOtrosIngresosFam
            oComm.Parameters.Add("@nTotalGastosFam", SqlDbType.Money).Value = nTotalGastosFam
            oComm.Parameters.Add("@nUtiNegSecundario", SqlDbType.Money).Value = nUtiNegSecundario
            oComm.Parameters.Add("@nOtrosGastosFam", SqlDbType.Money).Value = nOtrosGastosFam
            oComm.Parameters.Add("@bManual", SqlDbType.Bit).Value = 1

            oComm.ExecuteNonQuery()

            '**************************** REGISTRA EVALUACION UNIDAD FAMILIAR ***********************************************

            oComm = New SqlClient.SqlCommand("App_RecuperaDatosUnidadFamiliar", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = oPresolicitudApp.oCoopSolicitud.oEvaluacion.cDoc
            oComm.Parameters.Add("@nIngAdi", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nOtrosIng", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nAlimentac", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nAlqCasa", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nVesCal", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nEducacion", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nSalud", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nTransporte", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nServicios", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nOtros", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nCuotaMens", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nJuntas", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nNumDep", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@cEdades", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@cComenOtrIng", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nIngCony", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nTitIng", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nConyIng", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nAlqIng", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nRemesasIng", SqlDbType.Money).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            Dim nIngAdi As Double
            nIngAdi = oComm.Parameters("@nIngAdi").Value

            Dim nOtrosIng As Double
            nOtrosIng = oComm.Parameters("@nOtrosIng").Value

            Dim nAlimentac As Double
            nAlimentac = oComm.Parameters("@nAlimentac").Value

            Dim nAlqCasa As Double
            nAlqCasa = oComm.Parameters("@nAlqCasa").Value

            Dim nVesCal As Double
            nVesCal = oComm.Parameters("@nVesCal").Value

            Dim nEducacion As Double
            nEducacion = oComm.Parameters("@nEducacion").Value

            Dim nSalud As Double
            nSalud = oComm.Parameters("@nSalud").Value

            Dim nTransporte As Double
            nTransporte = oComm.Parameters("@nTransporte").Value

            Dim nServiciosUniFam As Double
            nServiciosUniFam = oComm.Parameters("@nServicios").Value

            Dim nOtros As Double
            nOtros = oComm.Parameters("@nOtros").Value

            Dim nCuotaMens As Double
            nCuotaMens = oComm.Parameters("@nCuotaMens").Value

            Dim nJuntas As Double
            nJuntas = oComm.Parameters("@nJuntas").Value

            Dim nNumDep As Integer
            nNumDep = oComm.Parameters("@nNumDep").Value

            Dim cEdades As String
            cEdades = oComm.Parameters("@cEdades").Value

            Dim cComenOtrIng As String
            cComenOtrIng = oComm.Parameters("@cComenOtrIng").Value

            Dim nIngCony As Double
            nIngCony = oComm.Parameters("@nIngCony").Value

            Dim nTitIng As Double
            nTitIng = oComm.Parameters("@nTitIng").Value

            Dim nConyIng As Double
            nConyIng = oComm.Parameters("@nConyIng").Value

            Dim nAlqIng As Double
            nAlqIng = oComm.Parameters("@nAlqIng").Value

            Dim nRemesasIng As Double
            nRemesasIng = oComm.Parameters("@nRemesasIng").Value


            oComm = New SqlClient.SqlCommand("Cred_RegistraEvalUniFam", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEvalResp
            oComm.Parameters.Add("@nDisCli", SqlDbType.Int).Value = 0
            oComm.Parameters.Add("@nExpNeg", SqlDbType.Int).Value = 0
            oComm.Parameters.Add("@nRieZonNeg", SqlDbType.Int).Value = 0
            oComm.Parameters.Add("@nRefCom", SqlDbType.Int).Value = 0
            oComm.Parameters.Add("@nComp", SqlDbType.Int).Value = 0
            oComm.Parameters.Add("@nCueCob", SqlDbType.Int).Value = 0
            oComm.Parameters.Add("@nGradInst", SqlDbType.Int).Value = 0
            oComm.Parameters.Add("@nEdadTit", SqlDbType.Int).Value = 0
            oComm.Parameters.Add("@nRieZonDom", SqlDbType.Int).Value = 0
            oComm.Parameters.Add("@nSolFam", SqlDbType.Int).Value = 0
            oComm.Parameters.Add("@nAntCre", SqlDbType.Int).Value = 0
            oComm.Parameters.Add("@nIngAdi", SqlDbType.Money).Value = nIngAdi
            oComm.Parameters.Add("@nOtrosIng", SqlDbType.Money).Value = nOtrosIng
            oComm.Parameters.Add("@nAlimentac", SqlDbType.Money).Value = nAlimentac
            oComm.Parameters.Add("@nAlqCasa", SqlDbType.Money).Value = nAlqCasa
            oComm.Parameters.Add("@nVesCal", SqlDbType.Money).Value = nVesCal
            oComm.Parameters.Add("@nEducacion", SqlDbType.Money).Value = nEducacion
            oComm.Parameters.Add("@nSalud", SqlDbType.Money).Value = nSalud
            oComm.Parameters.Add("@nTransporte", SqlDbType.Money).Value = nTransporte
            oComm.Parameters.Add("@nServicios", SqlDbType.Money).Value = nServiciosUniFam
            oComm.Parameters.Add("@nOtros", SqlDbType.Money).Value = nOtros
            oComm.Parameters.Add("@nCuotaMens", SqlDbType.Money).Value = nCuotaMens
            oComm.Parameters.Add("@nJuntas", SqlDbType.Money).Value = nJuntas
            oComm.Parameters.Add("@nNumDep", SqlDbType.Int).Value = nNumDep
            oComm.Parameters.Add("@cEdades", SqlDbType.VarChar, 1000).Value = cEdades
            oComm.Parameters.Add("@cComenOtrIng", SqlDbType.VarChar, 1000).Value = cComenOtrIng
            oComm.Parameters.Add("@nIngCony", SqlDbType.Money).Value = nIngCony
            oComm.Parameters.Add("@nTitIng", SqlDbType.Money).Value = nTitIng
            oComm.Parameters.Add("@nConyIng", SqlDbType.Money).Value = nConyIng
            oComm.Parameters.Add("@nAlqIng", SqlDbType.Money).Value = nAlqIng
            oComm.Parameters.Add("@nRemesasIng", SqlDbType.Money).Value = nRemesasIng
            oComm.Parameters.Add("@nServiciosExtras", SqlDbType.Money).Value = 0
            oComm.Parameters.Add("@nOtrosImp", SqlDbType.Money).Value = 0
            oComm.Parameters.Add("@nPagosOtros", SqlDbType.Money).Value = 0

            oComm.ExecuteNonQuery()

            '**************************** REGISTRA EVALUACION VENTAS ***********************************************

            oComm = New SqlClient.SqlCommand("PA_ActualizaEvaluacionVentas", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@pnCodEval", SqlDbType.Int).Value = nCodEvalResp
            oComm.Parameters.Add("@sDoc", SqlDbType.VarChar, 50).Value = oPresolicitudApp.oCoopSolicitud.oEvaluacion.cDoc

            oComm.ExecuteNonQuery()



            '**************************** REGISTRA EVALUACION CONSUMO ***********************************************
            oComm = New SqlClient.SqlCommand("App_ActualizaEvaluacionConsumo", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@pnIDEval", SqlDbType.Int).Value = nCodEvalResp
            oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = oPresolicitudApp.oCoopSolicitud.oEvaluacion.cDoc

            oComm.ExecuteNonQuery()

            '**************************** REGISTRA EVALUACION DIARIO ***********************************************
            oComm = New SqlClient.SqlCommand("App_ActualizaEvaluacionDiario", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEvalResp
            oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = oPresolicitudApp.oCoopSolicitud.oEvaluacion.cDoc

            oComm.ExecuteNonQuery()


            '**************************** MIGRA DEUDAS ***********************************************
            oComm = New SqlClient.SqlCommand("App_MigraDeudas", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEvalResp

            oComm.ExecuteNonQuery()


            '************************************************************************
            '***************  Si es Diario va de frente  a Aprobacion
            '************************************************************************

            If oPresolicitudApp.oCoopSolicitud.nOperacion = 10 And oPresolicitudApp.oCoopSolicitud.nSubProd = 5 Then
                ActualizaEstadoAPropuesta(nCodAgeTit, nCodCredTit)
            End If



            '**************************** MIGRA TABLAS COOPAC A DBSOLIDA ***********************************************
            oComm = New SqlClient.SqlCommand("PA_MigraTablasAppMovilDetalle", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEvalResp
            oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = oPresolicitudApp.oCoopSolicitud.cDoc

            oComm.ExecuteNonQuery()



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

    Public Function RegistraPersonaIntegranteGrupo(ByVal oPersona As clsCoopPersona, ByVal nCodGrupo As Integer) As outNegocio

        Dim oData As clsDataset
        Dim oDocumentos As clsDataset

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

        oData = New clsDataset
        oDocumentos = New clsDataset

        oDocumentos.AgregaColumna("nTipoDoc")
        oDocumentos.AgregaColumna("cValor")
        oDocumentos.AgregaColumna("cNomCod")
        oDocumentos.AddRegistro()

        oData.AgregaColumna("nTipoPersona")
        oData.AgregaColumna("cCodSBS")
        oData.AgregaColumna("nEstadoCivil")
        oData.AgregaColumna("cApePat")
        oData.AgregaColumna("cApeMat")
        oData.AgregaColumna("cApeCas")
        oData.AgregaColumna("cNombres")
        oData.AgregaColumna("cCIIU")
        oData.AgregaColumna("dFecNac")
        oData.AgregaColumna("nRelacIns")
        oData.AgregaColumna("nSexo")
        oData.AgregaColumna("nNacionalidad")
        oData.AgregaColumna("cDNI")
        oData.AgregaColumna("nEstado")
        oData.AgregaColumna("nGradoIns")
        oData.AgregaColumna("nProfes")
        oData.AgregaColumna("cMovil")
        oData.AgregaColumna("cMail")
        oData.AgregaColumna("cRPM_RPC")
        oData.AgregaColumna("cTelefono")
        oData.AgregaColumna("nVehiculo")
        oData.AgregaColumna("nEstInst")
        oData.AgregaColumna("nEstUni")

        oData.AgregaColumna("nDirTipo1")
        oData.AgregaColumna("nDirTipo2")
        oData.AgregaColumna("nDirTipo3")
        oData.AgregaColumna("nDirTipo4")
        oData.AgregaColumna("cDirValor1")
        oData.AgregaColumna("cDirValor2")
        oData.AgregaColumna("cDirValor3")
        oData.AgregaColumna("cDirValor4")
        oData.AgregaColumna("cDirReferencia")
        oData.AgregaColumna("cZona")
        oData.AgregaColumna("cLugarNac")

        'Datos de Vivienda 04/01/10...Carlos Pua
        oData.AgregaColumna("cTiempo")
        oData.AgregaColumna("nTipoPropiedad")
        oData.AgregaColumna("nTipoPropietario")
        oData.AgregaColumna("nTipoZona")
        oData.AgregaColumna("nTipoBarrio")
        oData.AgregaColumna("nTipoVivienda")
        oData.AgregaColumna("nTipoMaterial")
        oData.AgregaColumna("nEstadoVivienda")
        oData.AgregaColumna("cNombrePropietario")
        oData.AgregaColumna("cFonoPropietario")
        '-----

        'Direccion del trabajo de la persona 05/01/10...Carlos Pua
        oData.AgregaColumna("nDirTrabTip1")
        oData.AgregaColumna("nDirTrabTip2")
        oData.AgregaColumna("nDirTrabTip3")
        oData.AgregaColumna("nDirTrabTip4")
        oData.AgregaColumna("cDirTrab1")
        oData.AgregaColumna("cDirTrab2")
        oData.AgregaColumna("cDirTrab3")
        oData.AgregaColumna("cDirTrab4")
        oData.AgregaColumna("cDirTrabRef")
        oData.AgregaColumna("cZonaTrab")

        oData.AgregaColumna("cNomEmpresa")
        oData.AgregaColumna("cRucEmpresa")
        oData.AgregaColumna("cTelefEmpresa")
        '-----


        'Informacion del Negocio 06/01/10...Carlos Pua
        oData.AgregaColumna("nDirNegTip1")
        oData.AgregaColumna("nDirNegTip2")
        oData.AgregaColumna("nDirNegTip3")
        oData.AgregaColumna("nDirNegTip4")
        oData.AgregaColumna("cDirNeg1")
        oData.AgregaColumna("cDirNeg2")
        oData.AgregaColumna("cDirNeg3")
        oData.AgregaColumna("cDirNeg4")
        oData.AgregaColumna("cDirNegRef")
        oData.AgregaColumna("cZonaNeg")

        oData.AgregaColumna("cNomNegocio")
        oData.AgregaColumna("cRucNegocio")
        oData.AgregaColumna("cTelefNegocio")

        oData.AgregaColumna("cTiempoNeg")
        oData.AgregaColumna("cTiempoActualNeg")
        oData.AgregaColumna("cNomPropNeg")
        oData.AgregaColumna("cTelefNeg")

        oData.AgregaColumna("nLocal")
        oData.AgregaColumna("cCIIUNeg")
        oData.AgregaColumna("cActUNeg")
        '-----

        oData.AgregaColumna("nTipoFteIngreso")
        oData.AgregaColumna("cSuministroElectrico")
        oData.AgregaColumna("bApeCas")

        oData.AddRegistro()

        With oPersona

            oData.AsignarDato("1", "nTipoPersona")
            oData.AsignarDato("", "cCodSBS")
            oData.AsignarDato(Str(.nEstadoCiv), "nEstadoCivil")
            oData.AsignarDato(.cApePat, "cApePat")
            oData.AsignarDato(.cApeMat, "cApeMat")
            oData.AsignarDato(.cApeCasada, "cApeCas")
            oData.AsignarDato(.cPrimerNombre + .cSegundoNombre, "cNombres")
            oData.AsignarDato(Str(.nCiiu), "cCIIU")
            oData.AsignarDato(Format(.dFecNac, "dd/MM/yyyy"), "dFecNac")
            oData.AsignarDato("0", "nRelacIns")
            oData.AsignarDato(Str(.nSexo), "nSexo")
            oData.AsignarDato(Str(.nPais), "nNacionalidad")
            oData.AsignarDato(.cDoc, "cDNI")
            oData.AsignarDato("1", "nEstado")
            oData.AsignarDato(Str(.nNivelInstruc), "nGradoIns")
            oData.AsignarDato(Str(.nProfesion), "nProfes")
            oData.AsignarDato(.cMovil1, "cMovil")
            oData.AsignarDato(.cEmailprinc, "cMail")
            oData.AsignarDato(.cMovil2, "cRPM_RPC")
            oData.AsignarDato(.cTeleffijo, "cTelefono")
            If .nvehpropio = 1 Then
                oData.AsignarDato("True", "nVehiculo")
            Else
                oData.AsignarDato("False", "nVehiculo")
            End If
            oData.AsignarDato(Str(.nhijoseninst), "nEstInst")
            oData.AsignarDato(Str(.nhijosenuniv), "nEstUni")

            For i = 0 To .oCooppersonadireccion.Count - 1
                If .oCooppersonadireccion(i).ntipodirec = 2 Then
                    oData.AsignarDato("1", "nDirTipo1")
                    oData.AsignarDato("1", "nDirTipo2")
                    oData.AsignarDato("1", "nDirTipo3")
                    oData.AsignarDato("1", "nDirTipo4")
                    oData.AsignarDato(.oCooppersonadireccion(i).cDireccion, "cDirValor1")
                    oData.AsignarDato("", "cDirValor2")
                    oData.AsignarDato("", "cDirValor3")
                    oData.AsignarDato("", "cDirValor4")
                    oData.AsignarDato(.oCooppersonadireccion(i).cReferencia, "cDirReferencia")
                    oData.AsignarDato(.oCooppersonadireccion(i).cCodlocalidad, "cZona")
                End If
                If .oCooppersonadireccion(i).ntipodirec = 1 Then
                    oData.AsignarDato("1", "nDirNegTip1")
                    oData.AsignarDato("1", "nDirNegTip2")
                    oData.AsignarDato("1", "nDirNegTip3")
                    oData.AsignarDato("1", "nDirNegTip4")
                    oData.AsignarDato(.oCooppersonadireccion(i).cDireccion, "cDirNeg1")
                    oData.AsignarDato("", "cDirNeg2")
                    oData.AsignarDato("", "cDirNeg3")
                    oData.AsignarDato("", "cDirNeg4")
                    oData.AsignarDato(.oCooppersonadireccion(i).cReferencia, "cDirNegRef")
                    oData.AsignarDato(.oCooppersonadireccion(i).cCodlocalidad, "cZonaNeg")
                End If


            Next

            oData.AsignarDato(.cCodZona, "cLugarNac")


            'Datos de Vivienda 
            oData.AsignarDato(Str(.ntiemporesid), "cTiempo")
            oData.AsignarDato(Str(.ncbopropiedad), "nTipoPropiedad")
            oData.AsignarDato(Str(.ncbopropietario), "nTipoPropietario")
            oData.AsignarDato(Str(.ncbotipozona), "nTipoZona")
            oData.AsignarDato(Str(.ncbobarrio), "nTipoBarrio")
            oData.AsignarDato(Str(.ncbotipovivienda), "nTipoVivienda")
            oData.AsignarDato(Str(.ncbomaterialcontruc), "nTipoMaterial")
            oData.AsignarDato(Str(.ncboestadoviv), "nEstadoVivienda")
            oData.AsignarDato("", "cNombrePropietario")
            oData.AsignarDato("", "cFonoPropietario")
            '-----

            oData.AsignarDato("1", "nDirTrabTip1")
            oData.AsignarDato("1", "nDirTrabTip2")
            oData.AsignarDato("1", "nDirTrabTip3")
            oData.AsignarDato("1", "nDirTrabTip4")
            oData.AsignarDato("1", "cDirTrab1")
            oData.AsignarDato("1", "cDirTrab2")
            oData.AsignarDato("1", "cDirTrab3")
            oData.AsignarDato("1", "cDirTrab4")
            oData.AsignarDato("1", "cDirTrabRef")
            oData.AsignarDato("1", "cZonaTrab")

            oData.AsignarDato("", "cNomEmpresa")
            oData.AsignarDato("", "cRucEmpresa")
            oData.AsignarDato("", "cTelefEmpresa")
            '-----


            'Informacion del Negocio 
            oData.AsignarDato("", "cNomNegocio")
            oData.AsignarDato("", "cRucNegocio")
            oData.AsignarDato("", "cTelefNegocio")

            oData.AsignarDato("0", "cTiempoNeg")
            oData.AsignarDato("", "cTiempoActualNeg")
            oData.AsignarDato("", "cNomPropNeg")
            oData.AsignarDato("", "cTelefNeg")

            oData.AsignarDato("", "nLocal")
            oData.AsignarDato("", "cCIIUNeg")
            oData.AsignarDato("", "cActUNeg")
            '-----

            oData.AsignarDato(.nClasific, "nTipoFteIngreso")
            oData.AsignarDato(.numsuministro, "cSuministroElectrico")
            oData.AsignarDato("0", "bApeCas")

            oDocumentos.AsignarDato("1", "nTipoDoc")
            oDocumentos.AsignarDato(.cDoc, "cValor")
            oDocumentos.AsignarDato("DNI", "cNomCod")

        End With



        nResp = 0
        sRespuesta = ""

        sCadConexion = oConect.gsCadConexDBSolida

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try

            oComm = New SqlClient.SqlCommand("App_Pers_RegistraPersona", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = oPersona.cDoc
            oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = 1
            oComm.Parameters.Add("@pData", SqlDbType.Xml).Value = oData.DevuelveXML
            oComm.Parameters.Add("@pDocs", SqlDbType.Xml).Value = oDocumentos.DevuelveXML
            oComm.Parameters.Add("@pnCodPersNew", SqlDbType.Int).Direction = ParameterDirection.Output

            oComm.Parameters.Add("@pResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@psError", SqlDbType.VarChar, 250).Direction = ParameterDirection.Output
            oComm.ExecuteNonQuery()

            oPersona.nID = oComm.Parameters("@pnCodPersNew").Value

            RegistraIntegranteGrupoSiNoExisteComoSocio(nCodGrupo, oPersona.nID)


            nResp = oComm.Parameters("@pResp").Value
            sRespuesta = oComm.Parameters("@psError").Value



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



    Public Function RegistraPersonaApp() As outNegocio

        Dim oData As clsDataset
        Dim oDocumentos As clsDataset

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

        oData = New clsDataset
        oDocumentos = New clsDataset

        oDocumentos.AgregaColumna("nTipoDoc")
        oDocumentos.AgregaColumna("cValor")
        oDocumentos.AgregaColumna("cNomCod")
        oDocumentos.AddRegistro()

        oData.AgregaColumna("nTipoPersona")
        oData.AgregaColumna("cCodSBS")
        oData.AgregaColumna("nEstadoCivil")
        oData.AgregaColumna("cApePat")
        oData.AgregaColumna("cApeMat")
        oData.AgregaColumna("cApeCas")
        oData.AgregaColumna("cNombres")
        oData.AgregaColumna("cCIIU")
        oData.AgregaColumna("dFecNac")
        oData.AgregaColumna("nRelacIns")
        oData.AgregaColumna("nSexo")
        oData.AgregaColumna("nNacionalidad")
        oData.AgregaColumna("cDNI")
        oData.AgregaColumna("nEstado")
        oData.AgregaColumna("nGradoIns")
        oData.AgregaColumna("nProfes")
        oData.AgregaColumna("cMovil")
        oData.AgregaColumna("cMail")
        oData.AgregaColumna("cRPM_RPC")
        oData.AgregaColumna("cTelefono")
        oData.AgregaColumna("nVehiculo")
        oData.AgregaColumna("nEstInst")
        oData.AgregaColumna("nEstUni")

        oData.AgregaColumna("nDirTipo1")
        oData.AgregaColumna("nDirTipo2")
        oData.AgregaColumna("nDirTipo3")
        oData.AgregaColumna("nDirTipo4")
        oData.AgregaColumna("cDirValor1")
        oData.AgregaColumna("cDirValor2")
        oData.AgregaColumna("cDirValor3")
        oData.AgregaColumna("cDirValor4")
        oData.AgregaColumna("cDirReferencia")
        oData.AgregaColumna("cZona")
        oData.AgregaColumna("cLugarNac")

        'Datos de Vivienda 04/01/10...Carlos Pua
        oData.AgregaColumna("cTiempo")
        oData.AgregaColumna("nTipoPropiedad")
        oData.AgregaColumna("nTipoPropietario")
        oData.AgregaColumna("nTipoZona")
        oData.AgregaColumna("nTipoBarrio")
        oData.AgregaColumna("nTipoVivienda")
        oData.AgregaColumna("nTipoMaterial")
        oData.AgregaColumna("nEstadoVivienda")
        oData.AgregaColumna("cNombrePropietario")
        oData.AgregaColumna("cFonoPropietario")
        '-----

        'Direccion del trabajo de la persona 05/01/10...Carlos Pua
        oData.AgregaColumna("nDirTrabTip1")
        oData.AgregaColumna("nDirTrabTip2")
        oData.AgregaColumna("nDirTrabTip3")
        oData.AgregaColumna("nDirTrabTip4")
        oData.AgregaColumna("cDirTrab1")
        oData.AgregaColumna("cDirTrab2")
        oData.AgregaColumna("cDirTrab3")
        oData.AgregaColumna("cDirTrab4")
        oData.AgregaColumna("cDirTrabRef")
        oData.AgregaColumna("cZonaTrab")

        oData.AgregaColumna("cNomEmpresa")
        oData.AgregaColumna("cRucEmpresa")
        oData.AgregaColumna("cTelefEmpresa")
        '-----


        'Informacion del Negocio 06/01/10...Carlos Pua
        oData.AgregaColumna("nDirNegTip1")
        oData.AgregaColumna("nDirNegTip2")
        oData.AgregaColumna("nDirNegTip3")
        oData.AgregaColumna("nDirNegTip4")
        oData.AgregaColumna("cDirNeg1")
        oData.AgregaColumna("cDirNeg2")
        oData.AgregaColumna("cDirNeg3")
        oData.AgregaColumna("cDirNeg4")
        oData.AgregaColumna("cDirNegRef")
        oData.AgregaColumna("cZonaNeg")

        oData.AgregaColumna("cNomNegocio")
        oData.AgregaColumna("cRucNegocio")
        oData.AgregaColumna("cTelefNegocio")

        oData.AgregaColumna("cTiempoNeg")
        oData.AgregaColumna("cTiempoActualNeg")
        oData.AgregaColumna("cNomPropNeg")
        oData.AgregaColumna("cTelefNeg")

        oData.AgregaColumna("nLocal")
        oData.AgregaColumna("cCIIUNeg")
        oData.AgregaColumna("cActUNeg")
        '-----

        oData.AgregaColumna("nTipoFteIngreso")
        oData.AgregaColumna("cSuministroElectrico")
        oData.AgregaColumna("bApeCas")

        oData.AddRegistro()

        With oPresolicitudApp.oCoopSolicitud.oCoopPersona

            oData.AsignarDato("1", "nTipoPersona")
            oData.AsignarDato("", "cCodSBS")
            oData.AsignarDato(Str(.nEstadoCiv), "nEstadoCivil")
            oData.AsignarDato(.cApePat, "cApePat")
            oData.AsignarDato(.cApeMat, "cApeMat")
            oData.AsignarDato(.cApeCasada, "cApeCas")
            oData.AsignarDato(.cPrimerNombre + .cSegundoNombre, "cNombres")
            oData.AsignarDato(Str(.nCiiu), "cCIIU")
            oData.AsignarDato(Format(.dFecNac, "dd/MM/yyyy"), "dFecNac")
            oData.AsignarDato("0", "nRelacIns")
            oData.AsignarDato(Str(.nSexo), "nSexo")
            oData.AsignarDato(Str(.nPais), "nNacionalidad")
            oData.AsignarDato(.cDoc, "cDNI")
            oData.AsignarDato("1", "nEstado")
            oData.AsignarDato(Str(.nNivelInstruc), "nGradoIns")
            oData.AsignarDato(Str(.nProfesion), "nProfes")
            oData.AsignarDato(.cMovil1, "cMovil")
            oData.AsignarDato(.cEmailprinc, "cMail")
            oData.AsignarDato(.cMovil2, "cRPM_RPC")
            oData.AsignarDato(.cTeleffijo, "cTelefono")
            If .nvehpropio = 1 Then
                oData.AsignarDato("True", "nVehiculo")
            Else
                oData.AsignarDato("False", "nVehiculo")
            End If
            oData.AsignarDato(Str(.nhijoseninst), "nEstInst")
            oData.AsignarDato(Str(.nhijosenuniv), "nEstUni")

            For i = 0 To .oCooppersonadireccion.Count - 1
                If .oCooppersonadireccion(i).ntipodirec = 2 Then
                    oData.AsignarDato("1", "nDirTipo1")
                    oData.AsignarDato("1", "nDirTipo2")
                    oData.AsignarDato("1", "nDirTipo3")
                    oData.AsignarDato("1", "nDirTipo4")
                    oData.AsignarDato(.oCooppersonadireccion(i).cDireccion, "cDirValor1")
                    oData.AsignarDato("", "cDirValor2")
                    oData.AsignarDato("", "cDirValor3")
                    oData.AsignarDato("", "cDirValor4")
                    oData.AsignarDato(.oCooppersonadireccion(i).cReferencia, "cDirReferencia")
                    oData.AsignarDato(.oCooppersonadireccion(i).cCodlocalidad, "cZona")
                End If
                If .oCooppersonadireccion(i).ntipodirec = 1 Then
                    oData.AsignarDato("1", "nDirNegTip1")
                    oData.AsignarDato("1", "nDirNegTip2")
                    oData.AsignarDato("1", "nDirNegTip3")
                    oData.AsignarDato("1", "nDirNegTip4")
                    oData.AsignarDato(.oCooppersonadireccion(i).cDireccion, "cDirNeg1")
                    oData.AsignarDato("", "cDirNeg2")
                    oData.AsignarDato("", "cDirNeg3")
                    oData.AsignarDato("", "cDirNeg4")
                    oData.AsignarDato(.oCooppersonadireccion(i).cReferencia, "cDirNegRef")
                    oData.AsignarDato(.oCooppersonadireccion(i).cCodlocalidad, "cZonaNeg")
                End If


            Next

            oData.AsignarDato(.cCodZona, "cLugarNac")


            'Datos de Vivienda 
            oData.AsignarDato(Str(.ntiemporesid), "cTiempo")
            oData.AsignarDato(Str(.ncbopropiedad), "nTipoPropiedad")
            oData.AsignarDato(Str(.ncbopropietario), "nTipoPropietario")
            oData.AsignarDato(Str(.ncbotipozona), "nTipoZona")
            oData.AsignarDato(Str(.ncbobarrio), "nTipoBarrio")
            oData.AsignarDato(Str(.ncbotipovivienda), "nTipoVivienda")
            oData.AsignarDato(Str(.ncbomaterialcontruc), "nTipoMaterial")
            oData.AsignarDato(Str(.ncboestadoviv), "nEstadoVivienda")
            oData.AsignarDato("", "cNombrePropietario")
            oData.AsignarDato("", "cFonoPropietario")
            '-----

            oData.AsignarDato("1", "nDirTrabTip1")
            oData.AsignarDato("1", "nDirTrabTip2")
            oData.AsignarDato("1", "nDirTrabTip3")
            oData.AsignarDato("1", "nDirTrabTip4")
            oData.AsignarDato("1", "cDirTrab1")
            oData.AsignarDato("1", "cDirTrab2")
            oData.AsignarDato("1", "cDirTrab3")
            oData.AsignarDato("1", "cDirTrab4")
            oData.AsignarDato("1", "cDirTrabRef")
            oData.AsignarDato("1", "cZonaTrab")

            oData.AsignarDato("", "cNomEmpresa")
            oData.AsignarDato("", "cRucEmpresa")
            oData.AsignarDato("", "cTelefEmpresa")
            '-----


            'Informacion del Negocio 
            oData.AsignarDato("", "cNomNegocio")
            oData.AsignarDato("", "cRucNegocio")
            oData.AsignarDato("", "cTelefNegocio")

            oData.AsignarDato("0", "cTiempoNeg")
            oData.AsignarDato("", "cTiempoActualNeg")
            oData.AsignarDato("", "cNomPropNeg")
            oData.AsignarDato("", "cTelefNeg")

            oData.AsignarDato("", "nLocal")
            oData.AsignarDato("", "cCIIUNeg")
            oData.AsignarDato("", "cActUNeg")
            '-----

            oData.AsignarDato(.nClasific, "nTipoFteIngreso")
            oData.AsignarDato(.numsuministro, "cSuministroElectrico")
            oData.AsignarDato("0", "bApeCas")

            oDocumentos.AsignarDato("1", "nTipoDoc")
            oDocumentos.AsignarDato(.cDoc, "cValor")
            oDocumentos.AsignarDato("DNI", "cNomCod")

        End With



        nResp = 0
        sRespuesta = ""

        sCadConexion = oConect.gsCadConexDBSolida

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try

            oComm = New SqlClient.SqlCommand("App_Pers_RegistraPersona", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = oPresolicitudApp.oCoopSolicitud.oCoopPersona.cDoc
            oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = 1
            oComm.Parameters.Add("@pData", SqlDbType.Xml).Value = oData.DevuelveXML
            oComm.Parameters.Add("@pDocs", SqlDbType.Xml).Value = oDocumentos.DevuelveXML
            oComm.Parameters.Add("@pnCodPersNew", SqlDbType.Int).Direction = ParameterDirection.Output

            oComm.Parameters.Add("@pResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@psError", SqlDbType.VarChar, 250).Direction = ParameterDirection.Output
            oComm.ExecuteNonQuery()

            oPresolicitudApp.oCoopSolicitud.oCoopPersona.nID = oComm.Parameters("@pnCodPersNew").Value

            nResp = oComm.Parameters("@pResp").Value
            sRespuesta = oComm.Parameters("@psError").Value

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




    Public Function RegistraSolicitudApp() As outNegocio

        Dim R As New outNegocio
        Dim D(2) As String
        Dim Rela As New clsDataset
        Dim oCalData As New clsDataset
        Dim poDesemPar As New clsDataset
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nResp As Integer
        Dim sRespuesta As String
        Dim dr As SqlClient.SqlDataReader
        Dim nCodPers As Long
        Dim nCodPersTitular As Long

        Rela.AgregaColumna("Documento")
        Rela.AgregaColumna("Cliente")
        Rela.AgregaColumna("Relacion")
        Rela.AgregaColumna("ID")
        Rela.AgregaColumna("Codigo")




        nResp = 0
        sRespuesta = ""

        sCadConexion = oConect.gsCadConexDBSolida

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try
            oComm = New SqlClient.SqlCommand("PA_AppDevuelveCodPersonaTitular", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = oPresolicitudApp.oCoopSolicitud.oCoopPersona.cDoc
            oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()
            nCodPersTitular = oComm.Parameters("@nCodPers").Value

            Rela.AddRegistro()
            Rela.AsignarDato(oPresolicitudApp.oCoopSolicitud.oCoopPersona.cDoc, "Documento")
            Rela.AsignarDato("", "Cliente")
            Rela.AsignarDato("10", "Relacion")
            Rela.AsignarDato("10", "ID")
            Rela.AsignarDato(Str(nCodPersTitular), "Codigo")



            oComm = New SqlClient.SqlCommand("PA_AppDevuelveCodPersonaAnalista", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@psCodUsu", SqlDbType.VarChar, 50).Value = sCodUsu
            oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()
            nCodPers = oComm.Parameters("@nCodPers").Value


            Dim nCodGrupo As Integer
            oComm = New SqlClient.SqlCommand("App_RecuperaGrupo", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nID", SqlDbType.Int).Value = oPresolicitudApp.nCaptacion
            oComm.Parameters.Add("@nCodGrupo", SqlDbType.Int).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()
            nCodGrupo = oComm.Parameters("@nCodGrupo").Value


            Dim nMontoCuota As Double
            nMontoCuota = GeneraCuota(oPresolicitudApp.oCoopSolicitud.dFecDesem, oPresolicitudApp.oCoopSolicitud.dFecPrimCuota, oPresolicitudApp.oCoopSolicitud.nMonto, oPresolicitudApp.oCoopSolicitud.nNumCuo, oPresolicitudApp.oCoopSolicitud.nPeriodo, oPresolicitudApp.oCoopSolicitud.nTasa)

            oComm = New SqlClient.SqlCommand("CredW_ActualizaSolcitudCredito", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = 0
            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
            oComm.Parameters.Add("@nProd", SqlDbType.Int).Value = oPresolicitudApp.oCoopSolicitud.nOperacion
            oComm.Parameters.Add("@nSubProd", SqlDbType.Int).Value = oPresolicitudApp.oCoopSolicitud.nSubProd
            oComm.Parameters.Add("@nPrestamo", SqlDbType.Money).Value = oPresolicitudApp.oCoopSolicitud.nMonto
            oComm.Parameters.Add("@nCuotas", SqlDbType.Int).Value = oPresolicitudApp.oCoopSolicitud.nNumCuo
            oComm.Parameters.Add("@nPlazo", SqlDbType.Int).Value = IIf(oPresolicitudApp.oCoopSolicitud.nSubProd = 5, 1, 30)
            oComm.Parameters.Add("@nDiaFijo", SqlDbType.Int).Value = oPresolicitudApp.oCoopSolicitud.nDiaPago
            oComm.Parameters.Add("@nGrupo", SqlDbType.Int).Value = nCodGrupo
            oComm.Parameters.Add("@nCiclo", SqlDbType.Int).Value = 0
            oComm.Parameters.Add("@nCodLinea", SqlDbType.Int).Value = 0
            oComm.Parameters.Add("@nDestino", SqlDbType.Int).Value = oPresolicitudApp.oCoopSolicitud.nDestino
            oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = nCodPers
            oComm.Parameters.Add("@nTasaComp", SqlDbType.Money).Value = oPresolicitudApp.oCoopSolicitud.nTasa
            oComm.Parameters.Add("@nTasaMor", SqlDbType.Money).Value = 0
            oComm.Parameters.Add("@nMontoCuota", SqlDbType.Money).Value = nMontoCuota
            oComm.Parameters.Add("@nCondicion", SqlDbType.Int).Value = oPresolicitudApp.oCoopSolicitud.nModCredito
            oComm.Parameters.Add("@nRelaPers", SqlDbType.Xml).Value = Rela.DevuelveXML()
            oComm.Parameters.Add("@nTipoAct", SqlDbType.Int).Value = 1
            oComm.Parameters.Add("@nGastos", SqlDbType.Money).Value = 0
            oComm.Parameters.Add("@dFecDes", SqlDbType.DateTime).Value = oPresolicitudApp.oCoopSolicitud.dFecDesem
            oComm.Parameters.Add("@nDiasGracia", SqlDbType.Int).Value = 0
            oComm.Parameters.Add("@nTipoGracia", SqlDbType.Int).Value = 0
            oComm.Parameters.Add("@Cal_XML", SqlDbType.Xml).Value = oCalData.DevuelveXML
            oComm.Parameters.Add("@nRespCreditoNew", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nCodCamp", SqlDbType.Int).Value = 0
            oComm.Parameters.Add("@cCodAct", SqlDbType.VarChar).Value = Str(oPresolicitudApp.oCoopSolicitud.nActividadInterna)
            oComm.Parameters.Add("@nTipoPeriodo", SqlDbType.SmallInt).Value = 2
            oComm.Parameters.Add("@nMoneda", SqlDbType.SmallInt).Value = oPresolicitudApp.oCoopSolicitud.nMoneda
            oComm.Parameters.Add("@nOficina", SqlDbType.Int).Value = nCodAge
            oComm.Parameters.Add("@bTasaPref", SqlDbType.VarChar, 20).Value = 0

            oComm.Parameters.Add("@Cal_Des_Par", SqlDbType.Xml).Value = poDesemPar.DevuelveXML
            oComm.Parameters.Add("@nTipoDesemb", SqlDbType.Int).Value = 1
            oComm.Parameters.Add("@pnSoloIntereses", SqlDbType.Int).Value = 0
            oComm.Parameters.Add("@nTipoSegDesg", SqlDbType.Int).Value = 1

            oComm.Parameters.Add("@nPorCapTrab", SqlDbType.Money).Value = 0
            oComm.Parameters.Add("@nPorActFijo", SqlDbType.Money).Value = 0
            oComm.Parameters.Add("@nPorCompDeuda", SqlDbType.Money).Value = 0
            oComm.Parameters.Add("@dFecPrimPago", SqlDbType.DateTime).Value = oPresolicitudApp.oCoopSolicitud.dFecPrimCuota

            oComm.ExecuteNonQuery()

            Dim nCodCredNew As Integer
            nCodCredNew = oComm.Parameters("@nRespCreditoNew").Value



            '************************************************************************
            '***************  Si es Credito Grupal se registra 
            '************************************************************************

            If nCodGrupo > 0 Then

                oComm = New SqlClient.SqlCommand("CredW_RegistraCredGrupoExp", oConex)
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure

                oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = nCodPersTitular
                oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCredNew
                oComm.Parameters.Add("@bCoord", SqlDbType.Bit).Value = 0
                oComm.Parameters.Add("@nCodGrp", SqlDbType.Int).Value = nCodGrupo
                oComm.Parameters.Add("@nCiclo", SqlDbType.Int).Value = 0
                oComm.Parameters.Add("@Age", SqlDbType.Int).Value = nCodAge

                oComm.ExecuteNonQuery()

            End If

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


    Public Function RegistraPreSolicApp() As outNegocio

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

            oComm = New SqlClient.SqlCommand("PA_CoopRegistraPresolicitud", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@nTipoDoc", SqlDbType.Int).Value = oPresolicitudApp.nTipoDoc
            oComm.Parameters.Add("@cNroDoc", SqlDbType.VarChar, 50).Value = oPresolicitudApp.cNroDoc
            oComm.Parameters.Add("@cApePat", SqlDbType.VarChar, 50).Value = oPresolicitudApp.cApePat
            oComm.Parameters.Add("@cApeMat", SqlDbType.VarChar, 50).Value = oPresolicitudApp.cApeMat
            oComm.Parameters.Add("@cNombres", SqlDbType.VarChar, 150).Value = oPresolicitudApp.cNombres
            oComm.Parameters.Add("@cTelFij", SqlDbType.VarChar, 50).Value = oPresolicitudApp.cTelFij
            oComm.Parameters.Add("@cTelCel", SqlDbType.VarChar, 50).Value = oPresolicitudApp.cTelCel
            oComm.Parameters.Add("@cCodZona", SqlDbType.VarChar, 50).Value = oPresolicitudApp.cCodZona
            oComm.Parameters.Add("@cDireccion", SqlDbType.VarChar, 50).Value = oPresolicitudApp.cDireccion
            oComm.Parameters.Add("@cReferencia", SqlDbType.VarChar, 150).Value = oPresolicitudApp.cReferencia
            oComm.Parameters.Add("@nOperacion", SqlDbType.Int).Value = oPresolicitudApp.nOperacion
            oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = oPresolicitudApp.nMonto
            oComm.Parameters.Add("@nMoneda", SqlDbType.Int).Value = oPresolicitudApp.nMoneda
            oComm.Parameters.Add("@nPeriodo", SqlDbType.Int).Value = oPresolicitudApp.nPeriodo
            oComm.Parameters.Add("@nNumCuo", SqlDbType.Int).Value = oPresolicitudApp.nNumCuo
            oComm.Parameters.Add("@nDiaPago", SqlDbType.Int).Value = oPresolicitudApp.nDiaPago
            oComm.Parameters.Add("@cCodAsesor", SqlDbType.VarChar, 50).Value = oPresolicitudApp.cCodAsesor
            oComm.Parameters.Add("@nCaptacion", SqlDbType.Int).Value = oPresolicitudApp.nCaptacion
            oComm.Parameters.Add("@cCodPromotor", SqlDbType.VarChar, 50).Value = oPresolicitudApp.cCodPromotor
            oComm.Parameters.Add("@cComentarios", SqlDbType.VarChar, 150).Value = oPresolicitudApp.cComentarios
            oComm.Parameters.Add("@nEstado", SqlDbType.Int).Value = oPresolicitudApp.nEstado
            oComm.Parameters.Add("@nMotRechazo", SqlDbType.Int).Value = oPresolicitudApp.nMotRechazo
            oComm.Parameters.Add("@nCenNeg", SqlDbType.Int).Value = oPresolicitudApp.nCenNeg
            oComm.Parameters.Add("@cComRechazo", SqlDbType.VarChar, 150).Value = oPresolicitudApp.cComRechazo
            oComm.Parameters.Add("@nIDPresol", SqlDbType.Int).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            oPresolicitudApp.nIDPresol = oComm.Parameters("@nIDPresol").Value

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
                oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = oBandejaDesembolsoNuevo(k).ncodcred
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

                        i = i + 1

                    End While
                    dr.Close()
                    dr = Nothing

                Next


                For j = 0 To R.oAprobaciones.Count - 1

                    oComm = New SqlClient.SqlCommand("PA_RecuperaHistoriaSBS", oConex)
                    oComm.Parameters.Add("@psDoc", SqlDbType.Int).Value = R.oAprobaciones(j).cDoc
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
                    oComm.Parameters.Add("@psDoc", SqlDbType.Int).Value = R.oAprobaciones(j).cDoc
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
                    oComm.Parameters.Add("@psDoc", SqlDbType.Int).Value = R.oAprobaciones(j).cDoc
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
                    oComm.Parameters.Add("@psDoc", SqlDbType.Int).Value = R.oAprobaciones(j).cDoc
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

                oComm = New SqlClient.SqlCommand("PA_RegistraVerificacionNueva", oConex)
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

                Dim i As Integer

                For i = 0 To oVerifCliente(0).oImagenes.Count - 1


                    oComm = New SqlClient.SqlCommand("PA_RegistraVerificacionImagenNuevo", oConex)
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = oVerifCliente(0).nCodCred
                    oComm.Parameters.Add("@nImagen", SqlDbType.Image).Value = UnicodeStringToBytes(oVerifCliente(0).oImagenes(i).iImagen)
                    oComm.Parameters.Add("@cTipoFoto", SqlDbType.VarChar, 50).Value = oVerifCliente(0).oImagenes(i).cTipo


                    oComm.ExecuteNonQuery()

                Next



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

            oComm = New SqlClient.SqlCommand("PA_RecuperaBandejaVerificacion", oConex)
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

                        i = i + 1

                    End While
                    dr.Close()
                    dr = Nothing

                Next


                For j = 0 To R.oBandejaVerificacion.Count - 1

                    oComm = New SqlClient.SqlCommand("PA_RecuperaHistoriaSBS", oConex)
                    oComm.Parameters.Add("@psDoc", SqlDbType.Int).Value = R.oBandejaVerificacion(j).cDoc
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
                    oComm.Parameters.Add("@psDoc", SqlDbType.Int).Value = R.oBandejaVerificacion(j).cDoc
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

            oComm = New SqlClient.SqlCommand("PA_RecuperaBandejaVerificacionActualiza", oConex)
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


            oComm = New SqlClient.SqlCommand("PA_RecuperaBonoXColocacion", oConex)
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

            oComm = New SqlClient.SqlCommand("PA_RegistraMotivoNOPago", oConex)
            oComm.CommandType = CommandType.StoredProcedure
            '
            oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = oMotivoNoPago.nCodAge
            oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = oMotivoNoPago.nCodCred
            oComm.Parameters.Add("@psCodMotivo", SqlDbType.Int).Value = oMotivoNoPago.nMotivo
            oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = oMotivoNoPago.cUser
            oComm.Parameters.Add("@nLat", SqlDbType.Float).Value = oMotivoNoPago.nLatitud
            oComm.Parameters.Add("@nLong", SqlDbType.Float).Value = oMotivoNoPago.nLongitud
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

    Public Function RecuperaFeriados() As outNegocio
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

            oComm = New SqlClient.SqlCommand("PA_RecuperaFeriados", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            'oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

            Dim i As Integer = 0

            While dr.Read
                ReDim Preserve R.oFeriados(i)
                R.oFeriados(i) = New Date()
                R.oFeriados(i) = dr.Item("dFecha")

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

            oComm = New SqlClient.SqlCommand("PA_RecuperaAsesoresPromotores", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            'oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

            Dim i As Integer = 0

            While dr.Read
                ReDim Preserve R.oAsesor(i)
                R.oAsesor(i) = New clsItemAsesor()

                R.oAsesor(i).cCodUsu = dr.Item("cCodUsu")
                R.oAsesor(i).nCodAge = dr.Item("nCodAge")
                R.oAsesor(i).cNomApe = dr.Item("cNomApe")
                R.oAsesor(i).nAsesor = dr.Item("nAsesor")
                R.oAsesor(i).nPromotor = dr.Item("nPromotor")

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

            oComm = New SqlClient.SqlCommand("PA_Movil_AprobacionCredito_3", oConex)
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

    Public Sub RecuperaFeriadosDBSolida()

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

        sCadConexion = oConect.gsCadConexDBSolida

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try

            oComm = New SqlClient.SqlCommand("PA_RecuperaFeriados", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge

            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)

            Dim i As Integer = 0

            While dr.Read
                ReDim Preserve oListaFeriados(i)

                oListaFeriados(i) = CDate(dr.Item("dFecha").ToString)

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


    End Sub

    Public Function EsFeriado(ByVal dFecha As Date) As Boolean
        Dim i As Integer

        For i = 0 To oListaFeriados.Count - 1
            If DateDiff(DateInterval.Day, dFecha, oListaFeriados(i)) = 0 Then
                EsFeriado = True
                Exit Function
            Else
                EsFeriado = False
            End If
        Next

    End Function

    Public Function GeneraCuota(ByVal dFecDesem As Date, ByVal dFecPrimPago As Date,
                                ByVal nPrestamo As Double, ByVal nCuotas As Integer,
                                ByVal nPlazo As Integer, ByVal nTasa As Double) As Double


        Dim oListaPlazos() As Integer
        Dim nPosPlazos As Integer
        Dim oListaFactor() As Double
        Dim nPosFactor As Integer
        Dim i As Integer
        Dim dFecha As Date
        Dim dFechaAnt As Date
        Dim nDiasTrans As Integer

        RecuperaFeriadosDBSolida()
        dFecha = dFecDesem
        dFechaAnt = dFecDesem
        nPosPlazos = 0

        For i = 0 To nCuotas - 1

            If i = 0 Then
                nDiasTrans = DateDiff(DateInterval.Day, dFecha, dFecPrimPago)
                dFecha = dFecPrimPago
            Else
                dFecha = DateAdd(DateInterval.Day, nPlazo, dFecha)
            End If


            'Si es Domingo
            If DatePart(DateInterval.Weekday, dFecha) = 7 Then
                dFecha = DateAdd(DateInterval.Day, 1, dFecha)
            End If

            'Si considera Feriado
            If EsFeriado(dFecha) Then
                dFecha = DateAdd(DateInterval.Day, 1, dFecha)
            End If

            'Si es Domingo
            If DatePart(DateInterval.Weekday, dFecha) = 7 Then
                dFecha = DateAdd(DateInterval.Day, 1, dFecha)
            End If

            'Si considera Feriado
            If EsFeriado(dFecha) Then
                dFecha = DateAdd(DateInterval.Day, 1, dFecha)
            End If
            'Si es Domingo
            If DatePart(DateInterval.Weekday, dFecha) = 7 Then
                dFecha = DateAdd(DateInterval.Day, 1, dFecha)
            End If

            'Si considera Feriado
            If EsFeriado(dFecha) Then
                dFecha = DateAdd(DateInterval.Day, 1, dFecha)
            End If

            nDiasTrans = DateDiff(DateInterval.Day, dFechaAnt, dFecha)
            ReDim Preserve oListaPlazos(i)
            oListaPlazos(i) = nDiasTrans

        Next

        'Hallla Factor de Cada Cuota
        Dim nSumFactor As Double = 0
        Dim nSumaPlazo As Double = 0
        Dim nFactor As Double
        nPosFactor = 0
        For i = 0 To nCuotas - 1
            nSumaPlazo = nSumaPlazo + oListaPlazos(i)
            nFactor = 1.0 / Math.Pow((1 + nTasa / 100.0), (nSumaPlazo / 30.0))
            nSumFactor = nSumFactor + nFactor
            ReDim Preserve oListaFactor(i)
            oListaFactor(i) = nFactor
        Next

        GeneraCuota = nPrestamo / nSumFactor

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
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()

            R.nPorcComision = oComm.Parameters("@nPorcComision").Value
            R.nIncreMora = oComm.Parameters("@nIncreMora").Value
            R.nMoraAnt = oComm.Parameters("@nMoraAnt").Value
            R.nMorahoy = oComm.Parameters("@nMorahoy").Value


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

            oComm = New SqlClient.SqlCommand("PA_RegistraFotoYDeposito", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@pnMonto", SqlDbType.Money).Value = oVoucher.nMonto
            oComm.Parameters.Add("@iDoc", SqlDbType.Image).Value = oVoucher.iFoto
            oComm.Parameters.Add("@ncodage", SqlDbType.Int).Value = oVoucher.nCodAge
            oComm.Parameters.Add("@psUser", SqlDbType.VarChar, 50).Value = oVoucher.cCodUsu

            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output


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
        Dim oPersonas = From P In db.vPersonasBase
                        Where P.cDoc = cDoc.Trim
                        Select New dPersona With {.cApellidos = P.cApellidos, .cCodGenero = P.cGenero, .cDirecDom = P.cDireccionDom, .cDistritoDom = "", .cDoc = P.cDoc, .cDocConyuge = "", .cMovil1 = "", .cMovil2 = "", .cNomBanco = "", .cNombres = P.cNombres, .cNroCtaAho = "", .cRefDirDom = P.cReferenciaDom, .cTelefFijo = "", .cZonaDom = P.cCodZonaDom, .dFechaNac = P.dFechaNac, .nCodCred = P.nCodCred, .nCodEstadoCiv = P.nEstadoCiv, .nNumDep = 0, .nSaldoCap = P.nSaldoK}

        R.oPersonas = oPersonas

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

        If nProd = 6 Then

            If nTasaSemAnt <= 0 Then
                Dim oTasa = From T In db.TarifarioMicroSemanal
                            Where T.cCodCat = cCodCat And T.nCodAge = nCodAge And T.nPretamoIni <= nPrestamo And T.nPrestamoFin >= nPrestamo
                            Select T

                nTasa = oTasa.First.nTasaComp
            Else
                nTasa = nTasaSemAnt
            End If


        ElseIf nProd = 7 Or nProd = 8 Then

            Dim oTasa = From T In db.TarifarioMicroJuniorSemanal
                        Where T.cCodCat = cCodCat And T.nCodAge = nCodAge And T.nPretamoIni <= nPrestamo And T.nPrestamoFin >= nPrestamo
                        Select T

            nTasa = oTasa.First.nTasaComp

        Else

            Dim oTasa = From T In db.TarifarioMicro
                        Where T.cCodCat = cCodCat And T.nCodAge = nCodAge And T.nPretamoIni <= nPrestamo And T.nPrestamoFin >= nPrestamo
                        Select T

            nTasa = oTasa.First.nTasaComp

        End If

        sMontoCuota = D.MontoCuotaPorProd(Str(nPrestamo), Str(nCuotas), Str(nTasa), "1", "1", Str(nProd), "1")

        R.nResultado = 1
        R.sRespuesta = sMontoCuota(0)
        R.nTasaComp = nTasa

        Return R
    End Function

    Public Function RecuperaCuotaONLINE() As outNegocio
        Dim R As New outNegocio
        Dim nTasa As Double
        Dim D As New dDesembolso
        Dim sMontoCuota As String()

        Dim oTasa = From T In db.TarifarioAgencia
                    Where T.nProd = nProd And T.nCodAge = nCodAge
                    Select T

        nTasa = oTasa.First.nTasaComp

        sMontoCuota = D.MontoCuotaPorProd(Str(nPrestamo), Str(nCuotas), Str(nTasa), "1", "1", Str(nProd), "1")

        R.nResultado = 1
        R.sRespuesta = sMontoCuota(0)

        Return R
    End Function

    Public Function RecuperaDatosDeCredito() As outNegocio
        Dim R As New outNegocio

        Dim oData() As dPagoCredito

        Dim oListaPagos = From P In db.vListadoPagos
                          Where P.nCodCred = nCodcred
                          Select New dPagoCredito With {.bPagadoHoy = False, .cDireccionDom = P.cDireccionDom,
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
                              .cCategoria = P.cCategoria, .nProd = P.nProd, .nCiclos = P.nCiclos}

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

        Rtemp = RecuperaListaPagos()
        R.ListaPagosCredito = Rtemp.ListaPagosCredito

        Dim oCatalogocodigos = From C In db.CatalogoCodigos
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


        Dim oPersonas = From P In db.vPersonasBase
                        Where P.nCodAge = nCodAge
                        Select New dPersona With {.cApellidos = P.cApellidos, .cCodGenero = P.cGenero, .cDirecDom = P.cDireccionDom, .cDistritoDom = "", .cDoc = P.cDoc, .cDocConyuge = P.cDocConyuge, .cMovil1 = P.cMovil1, .cMovil2 = P.cMovil2, .cNomBanco = P.cNomBanco, .cNombres = P.cNombres, .cNroCtaAho = P.cctabanco, .cRefDirDom = P.cReferenciaDom, .cTelefFijo = P.cTelFijo, .cZonaDom = P.cCodZonaDom, .dFechaNac = P.dFechaNac, .nCodCred = P.nCodCred, .nCodEstadoCiv = P.nEstadoCiv, .nNumDep = P.nDependientes, .nSaldoCap = P.nSaldoK, .nMontoMaximo = P.nLineaCredito, .nPlazoMaximo = P.nCuotasMax, .cCategoria = P.cCategoria2}


        R.oPersonas = oPersonas

        Dim oNegocios = From E In db.vListaNegocios
                        Where E.nCodAge = nCodAge
                        Select New dNegocio With {.cCodActEcon = E.cCodActEcon, .cCodZona = E.cCodZonaNeg, .cDirecNeg = E.cDirecNeg, .cDoc = E.cDoc, .cRazonSocial = E.cRazonSocial, .cRefDirecNeg = E.cReferenciaNeg, .dFechaIniNeg = E.dFechaIniNeg, .dFechareg = E.dFechaEval, .nGastosMensFam = E.nGastosFam, .nMontoVentasDiarias = E.nMontoInventario, .nNumTrab = E.nNumTrab}


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

        'Carga Asesores
        Dim RAsesores As New outNegocio
        RAsesores = RecuperaAsesores()
        R.oAsesor = RAsesores.oAsesor

        'Carga Asesores
        Dim RFeriados As New outNegocio
        RFeriados = RecuperaFeriados()
        R.oFeriados = RFeriados.oFeriados

        'Carga Grupos
        Dim RGrupos As New outNegocio
        RGrupos = RecuperaListaGruposComunales()
        R.ListaGrupos = RGrupos.ListaGrupos

        Return R

    End Function

    Public Function RegistrarPago() As outNegocio
        Dim R As New outNegocio

        Try



            R = TranPago.RegistrarPago(sCodUsu)

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

            R.oUsuarioroles = oUsuarioRoles

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
    Public Function RecuperaListaPagos() As outNegocio
        Dim R As New outNegocio
        Dim oData() As dPagoCredito
        db.Database.CommandTimeout = 360
        Dim oListaPagos = From P In db.vListadoPagos
                          Where P.cCodUsuCob = sCodUsu
                          Select New dPagoCredito With {.bPagadoHoy = False, .cDireccionDom = P.cDireccionDom,
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
                              .sAmpliacion = P.Ampliacion, .UltFecCob = P.UltFecCob, .cCodUsuCob = P.cCodUsuCob, .nNumPagosDelDia = P.nNumPagosDelDia}

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
                                     Where GC.bActivo = True And CC.nCodigo = 1025
                                     Select New clsGrupoComunal With {.cNombre = GC.cNombre, .cLugarProxReunion = GC.cLugarProximaReunion, .cNomEstado = CC.cDescripcion,
                                         .dFechaReunion = GC.dFechaProxReunion, .nCodEstado = GC.nCodEstado, .nIdGrupo = GC.nID, .nNumCiclo = GC.nNumCiclo, .nNumSocios = GC.nNumSocios, .dFecVenc = Nothing, .nDiasAtraso = 0
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

End Class
