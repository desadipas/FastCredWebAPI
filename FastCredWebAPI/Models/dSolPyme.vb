Imports System.Data.SqlClient
Imports FastCredWebAPI.FastCredWebAPI

Public Class dSolPyme

    Private db As New DBNEGOCIOEntities

    '*************************************************
    'Datos de Persona
    '*************************************************
    Public Property cDoc As String
    Public Property nIDGrupo As Integer
    Public Property nCodCagrgoGrupo As Integer
    Public Property nCodDestinoCredito As Integer 'mdipas 2024

    Public Property cNombres As String
    Public Property cApellidos As String
    Public Property dFechaNac As String
    Public Property cGenero As String
    Public Property cDireccionDom As String
    Public Property cReferenciaDom As String
    Public Property cCodZonaDom As String
    Public Property cMovil1 As String
    Public Property cDocConyuge As String
    Public Property cTelFijo As String
    Public Property sEstadoCiv As String

    Public Property cNivelEducac As String
    Public Property cTipoVivienda As String
    Public Property cTiempoResidencia As String
    Public Property cExperienActiv As String

    Public Property sNumDep As String
    Public Property psNomBanco As String
    Public Property psNroCuenta As String
    Public Property psNroCuentaCCI As String
    Public Property nCodBanco As Integer

    Public Property psTipoTarifario As String
    Public Property psProd As String
    Public Property nActualizado As Integer


    Public Property nExcepcion As Integer
    Public Property cComentarioExcepcion As String



    '*************************************************
    'Datos de Evaluacion
    '*************************************************
    Public Property nCodEval As Integer
    Public Property psNuevo As String
    Public Property cRazonSocial As String
    Public Property cDirecNeg As String
    Public Property cCodZonaNeg As String
    Public Property cCodActEcon As String
    Public Property dFechaIniNeg As String
    Public Property nNumTrab As String
    Public Property nGastosFam As String
    Public Property nGastosMen As String
    Public Property cReferNeg As String
    Public Property cUser As String
    Public Property cLatitud As String
    Public Property cLong As String
    Public Property oImagenes As dFotosNegocio()
    Public Property nActualizadoNeg As Integer

    Public Property nVentadialv As Double

    Public Property nVentadiasd As Double
    Public Property nComprasSemanal As Double
    Public Property nCajaEfectivo As Double

    Public Property nGastoAlquiler As Double

    Public Property nGastoFinanciero As Double

    Public Property nGastoPrestamistas As Double

    Public Property nMontoInventario As Double
    ' mdipas 2024
    Public Property cComentarios As String
    Public Property nGastoFamAlquiler As Double
    Public Property nGastoFamTransporte As Double
    Public Property nGastoFamAlimentacion As Double
    Public Property nNumHijos As Integer
    Public Property nOtrosCostOpe As Double
    Public Property nMontoServiciosCostOpe As Double
    Public Property nMaquinYEquipos As Double
    Public Property nMontoDBueno As Double
    Public Property nNroDBuenoSeman As Integer
    Public Property nMontoDiaNormal As Double
    Public Property nGastoFamServicios As Double
    Public Property nOtrosIngresos As Double
    Public Property cComentariosOtrosIngresos As String
    Public Property nTrabajaConInformales As Integer
    Public Property nNroInformales As Integer
    Public Property nMontoInformales As Double

    'DATOS SOLICITUD DE CREDITO
    Public Property nIDSolicitud As Long
    Public Property nPrestamo As String
    Public Property nNumCuotas As String
    Public Property cCodUsuAna As String
    Public Property nMontoCuota As String
    Public Property pnTipoPlazo As String
    Public Property psFechaPago As String
    Public Property nAmpliado As Integer
    Public Property cUserSolicitud As String
    Public Property nCampCuotasSinInt As Integer

    '******************************************
    'Datos de Evaluacion de Consumo
    '***********************************************
    Public Property cNomBanco As String
    Public Property cCtaBanco As String
    Public Property cNomConyuge As String

    Public Property cIndepDep As String
    Public Property cNomEmpresa As String
    Public Property cTelefEmpresa As String
    Public Property cDirecEmpresa As String
    Public Property cCargo As String
    Public Property nSueldoBruto As String
    Public Property dFechaIngresoEmpresa As String
    Public Property bDesemnBanco As Integer


    Public Property oTasacion As dTasacion()

    Public Property nTipoDoc As Integer
    Public Property nTipoDocConyuge As Integer

    'mdipas diciembre
    Public Property nGasFamOtros As Double
    Public Property nCosOpeTransp As Double
    ' 
    Public Property nUtilidadBrutaTotal As Double
    Public Property nGastosNegocio As Double
    Public Property nGastosTotalFam As Double
    Public Property nDeudasFinancieras As Double
    Public Property nUtilidadNetaTotal As Double
    Public Property nEndeudamientoMax As Double
    Public Property nAsumeCuotaDiaria As Double
    Public Property nAsumeSemanal As Double
    Public Property nAsumeCatorcenal As Double
    Public Property nAsumeMensual As Double
    Public Property oNegociosRubros As Rubro()
    Public Property oDeudasEntidades As dEntidad()

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

    Public Function RegistrarTasacion(ByVal pnCodAge As Integer, ByVal pnCodcred As Integer) As String()

        Dim ATasac As New ArticuloTasacion
        Dim ATasacDet As New ArticuloTasacionDet
        Dim ATasacImg As New ArticuloTasacionImagen
        Dim ATasacCred As New ArticuloTasacionCred
        Dim AF As New ArticuloFamilia
        Dim AM As New ArticuloMarcas
        Dim AMo As New ArticuloModelo
        Dim nCodFamilia As Integer
        Dim nCodMarca As Integer
        Dim nCodModelo As Integer
        Dim nCodTasacion As Integer

        Dim D(2) As String

        Dim i As Integer


        i = 0
        For i = 0 To oTasacion.Length - 1
            If oTasacion(i).nCodFamilia = -1 Then
                AF.cNomFamilia = oTasacion(i).cNomFamilia
                AF.bEliminado = 0

                db.ArticuloFamilia.Add(AF)
                Try
                    db.SaveChanges()
                    nCodFamilia = AF.nID
                    D(0) = "1"
                    D(1) = "Registro Correcto"

                Catch ex As Exception
                    D(0) = "0"
                    D(1) = ex.Message
                    Return D
                End Try
            Else
                nCodFamilia = oTasacion(i).nCodFamilia
            End If

            If oTasacion(i).nCodMarca = -1 Then
                AM.nIdFamilia = nCodFamilia
                AM.cNomMarca = oTasacion(i).cNomMarca
                AM.bEliminado = 0

                db.ArticuloMarcas.Add(AM)
                Try
                    db.SaveChanges()
                    nCodMarca = AM.nCodMarca
                    D(0) = "1"
                    D(1) = "Registro Correcto"

                Catch ex As Exception
                    D(0) = "0"
                    D(1) = ex.Message
                    Return D
                End Try
            Else
                nCodMarca = oTasacion(i).nCodMarca
            End If

            If oTasacion(i).nCodModelo = -1 Then
                AMo.nCodMarca = nCodMarca
                AMo.cNomModelo = oTasacion(i).cNomModelo
                AMo.bEliminado = 0
                AMo.nValorXGrUni = oTasacion(i).nMontoTasado

                db.ArticuloModelo.Add(AMo)
                Try
                    db.SaveChanges()
                    nCodModelo = AMo.nCodModelo
                    D(0) = "1"
                    D(1) = "Registro Correcto"

                Catch ex As Exception
                    D(0) = "0"
                    D(1) = ex.Message
                    Return D
                End Try
            Else
                nCodModelo = oTasacion(i).nCodModelo
            End If

            'Registra Cabecera
            If i = 0 Then
                ATasac.dFecha = Now()
                ATasac.nPrestamoFinal = oTasacion(i).nMontoTasado
                db.ArticuloTasacion.Add(ATasac)
                Try
                    db.SaveChanges()
                    nCodTasacion = ATasac.nCodTasacion
                    D(0) = "1"
                    D(1) = "Registro Correcto"

                Catch ex As Exception
                    D(0) = "0"
                    D(1) = ex.Message
                    Return D
                End Try

            End If

            ATasacDet = New ArticuloTasacionDet
            ATasacDet.bEliminado = 0
            ATasacDet.cDescripcion = ""
            ATasacDet.nCantidad = 1
            ATasacDet.nCodMarca = nCodMarca
            ATasacDet.nCodModelo = nCodModelo
            ATasacDet.nCodTasacion = nCodTasacion
            ATasacDet.nItem = i + 1
            ATasacDet.nPorDesc = 0
            ATasacDet.nPrestamoBruto = oTasacion(i).nMontoTasado
            ATasacDet.nPrestamoFinal = oTasacion(i).nMontoTasado
            ATasacDet.nValorXGrUni = 1

            db.ArticuloTasacionDet.Add(ATasacDet)
            Try
                db.SaveChanges()

                D(0) = "1"
                D(1) = "Registro Correcto"

            Catch ex As Exception
                D(0) = "0"
                D(1) = ex.Message
                Return D
            End Try

            'Registra Imaganes de Tasacion
            For j = 0 To oTasacion(i).oImagenes.Length - 1

                ATasacImg = New ArticuloTasacionImagen
                ATasacImg.nCodTasacion = nCodTasacion
                ATasacImg.nImagen = oTasacion(i).oImagenes(j).nImagen
                ATasacImg.nItem = i + 1
                db.ArticuloTasacionImagen.Add(ATasacImg)
                Try
                    db.SaveChanges()

                    D(0) = "1"
                    D(1) = "Registro Correcto"

                Catch ex As Exception
                    D(0) = "0"
                    D(1) = ex.Message
                    Return D
                End Try
            Next

            'Registra Relacion con el Credito
            If i = 0 Then
                ATasacCred = New ArticuloTasacionCred
                ATasacCred.nCodAge = pnCodAge
                ATasacCred.nCodCred = pnCodcred
                ATasacCred.nCodTasacion = nCodTasacion

                db.ArticuloTasacionCred.Add(ATasacCred)
                Try
                    db.SaveChanges()

                    D(0) = "1"
                    D(1) = "Registro Correcto"

                Catch ex As Exception
                    D(0) = "0"
                    D(1) = ex.Message
                    Return D
                End Try

            End If

        Next

        Return D

    End Function

    Public Function RegistraEvaluacionConsumo() As String()

        Dim oConex As SqlConnection
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim D() As String
        Dim sCadConexion As String
        Dim oConect As New clsConecta
        Dim nCodEvalMax As Long

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        If Not IsNumeric(cLatitud) Then
            cLatitud = "0"
        End If
        If Not IsNumeric(cLong) Then
            cLong = "0"
        End If

        oComm = New SqlClient.SqlCommand("PA_Movil2_RegistraEvaluacionConsumo_2", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure


        oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = cDoc
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@psNuevo", SqlDbType.VarChar, 50).Value = psNuevo
        oComm.Parameters.Add("@cNomBanco", SqlDbType.VarChar, 150).Value = cNomBanco
        oComm.Parameters.Add("@cCtaBanco", SqlDbType.VarChar, 150).Value = cCtaBanco
        oComm.Parameters.Add("@cNomConyuge", SqlDbType.VarChar, 150).Value = cNomConyuge
        oComm.Parameters.Add("@cDNIConyuge", SqlDbType.VarChar, 50).Value = cDocConyuge
        oComm.Parameters.Add("@dnNumDep", SqlDbType.Int).Value = CInt(sNumDep)
        oComm.Parameters.Add("@cIndepDep", SqlDbType.VarChar, 100).Value = cIndepDep
        oComm.Parameters.Add("@cNomEmpresa", SqlDbType.VarChar, 100).Value = cNomEmpresa
        oComm.Parameters.Add("@cTelefEmpresa", SqlDbType.VarChar, 50).Value = cTelefEmpresa
        oComm.Parameters.Add("@cDirecEmpresa", SqlDbType.VarChar, 150).Value = cDirecEmpresa
        oComm.Parameters.Add("@cCargo", SqlDbType.VarChar, 150).Value = cCargo
        oComm.Parameters.Add("@nSueldoBruto", SqlDbType.Money).Value = CDbl(nSueldoBruto)
        oComm.Parameters.Add("@dFechaIngresoEmpresa", SqlDbType.Date).Value = CDate(dFechaIngresoEmpresa)
        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = cUser
        oComm.Parameters.Add("@nLatitud", SqlDbType.Float).Value = CDec(cLatitud)
        oComm.Parameters.Add("@nLong", SqlDbType.Float).Value = CDec(cLong)

        oComm.Parameters.Add("@cDescripGarantias", SqlDbType.VarChar, 3000).Value = ""
        oComm.Parameters.Add("@nMontoGarantias", SqlDbType.VarChar, 50).Value = 0


        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nCodEvalMax", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        ReDim Preserve D(2)

        D(0) = oComm.Parameters("@nResp").Value
        D(1) = oComm.Parameters("@Resultado").Value
        nCodEvalMax = oComm.Parameters("@nCodEvalMax").Value

        oConex.Close()

        'Registra Fotos
        Dim oImagen As dFotosNegocio

        If D(0) = "1" Then
            For Each oImagen In oImagenes
                D = GuardarFotoNegocio(nCodEvalMax, oImagen.nCodFoto, oImagen.nImg)
                If D(0) = "0" Then
                    Return D
                End If
            Next
        End If

        Return D



        Return D


    End Function


    Public Function RegistraNegocio() As String()
        Dim oConex As SqlConnection
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim D() As String
        Dim sCadConexion As String
        Dim oConect As New clsConecta
        Dim nCodEvalMax As Long

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        If Not IsNumeric(cLatitud) Then
            cLatitud = "0"
        End If
        If Not IsNumeric(cLong) Then
            cLong = "0"
        End If

        'Siempre es Nuevo
        nCodEval = 0

        oComm = New SqlClient.SqlCommand("PA_Movil2_RegistraNegocio_5", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@psDoc", SqlDbType.VarChar, 50).Value = cDoc
        oComm.Parameters.Add("@nCodEval", SqlDbType.Int).Value = nCodEval
        oComm.Parameters.Add("@psNuevo", SqlDbType.VarChar, 50).Value = "1"
        oComm.Parameters.Add("@cRazonSocial", SqlDbType.VarChar, 150).Value = cRazonSocial
        oComm.Parameters.Add("@cDirecNeg", SqlDbType.VarChar, 150).Value = cDirecNeg
        oComm.Parameters.Add("@cCodZonaNeg", SqlDbType.VarChar, 50).Value = cCodZonaNeg
        oComm.Parameters.Add("@cCodActEcon", SqlDbType.VarChar, 50).Value = cCodActEcon
        oComm.Parameters.Add("@dFechaIniNeg", SqlDbType.DateTime).Value = CDate(dFechaIniNeg)
        oComm.Parameters.Add("@nNumTrab", SqlDbType.Int).Value = CInt(nNumTrab)
        'oComm.Parameters.Add("@nGastosFam", SqlDbType.Money).Value = CDbl(nGastosFam)
        oComm.Parameters.Add("@nGastosMen", SqlDbType.Money).Value = CDbl(nGastosMen)
        oComm.Parameters.Add("@cReferNeg", SqlDbType.VarChar, 150).Value = cReferNeg
        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = cUser
        oComm.Parameters.Add("@nLatitud", SqlDbType.Float).Value = CDec(cLatitud)
        oComm.Parameters.Add("@nLong", SqlDbType.Float).Value = CDec(cLong)

        'oComm.Parameters.Add("@nVentadialv", SqlDbType.Money).Value = nVentadialv
        'oComm.Parameters.Add("@nVentadiasd", SqlDbType.Money).Value = nVentadiasd
        'oComm.Parameters.Add("@nComprasSemanal", SqlDbType.Money).Value = nComprasSemanal
        'oComm.Parameters.Add("@nGastoFinanciero", SqlDbType.Money).Value = nGastoFinanciero
        'oComm.Parameters.Add("@nGastoPrestamistas", SqlDbType.Money).Value = nGastoPrestamistas

        'mdipas 2024
        oComm.Parameters.Add("@cComentarioEvaluacion", SqlDbType.VarChar, 200).Value = cComentarios
        oComm.Parameters.Add("@nGastoFamAlquiler", SqlDbType.Money).Value = nGastoFamAlquiler
        oComm.Parameters.Add("@nGastoFamTransporte", SqlDbType.Money).Value = nGastoFamTransporte
        oComm.Parameters.Add("@nGastoFamAlimentacion", SqlDbType.Money).Value = nGastoFamAlimentacion
        oComm.Parameters.Add("@nNumHijos", SqlDbType.Int).Value = nNumHijos
        oComm.Parameters.Add("@nCosOpeOtros", SqlDbType.Money).Value = nOtrosCostOpe
        oComm.Parameters.Add("@nCosOpeServicios", SqlDbType.Money).Value = nMontoServiciosCostOpe
        oComm.Parameters.Add("@nMaquinYEquipos", SqlDbType.Money).Value = nMaquinYEquipos
        oComm.Parameters.Add("@nMontoDiaBueno", SqlDbType.Money).Value = nMontoDBueno
        oComm.Parameters.Add("@nDiasBuenos", SqlDbType.Int).Value = nNroDBuenoSeman
        oComm.Parameters.Add("@nMontoDiaNormal", SqlDbType.Money).Value = nMontoDiaNormal
        oComm.Parameters.Add("@nCajaEfectivo", SqlDbType.Money).Value = nCajaEfectivo 'mdipas 2024
        oComm.Parameters.Add("@nMontoInventario", SqlDbType.Money).Value = nMontoInventario 'mdipas 2024
        oComm.Parameters.Add("@nGastoAlquiler", SqlDbType.Money).Value = nGastoAlquiler
        oComm.Parameters.Add("@nGasFamServicios", SqlDbType.Money).Value = nGastoFamServicios

        oComm.Parameters.Add("@nOtrosIngresos", SqlDbType.Money).Value = nOtrosIngresos
        oComm.Parameters.Add("@cComentariosOtrosIngresos", SqlDbType.VarChar, 200).Value = cComentariosOtrosIngresos

        oComm.Parameters.Add("@nTrabajaConInformales", SqlDbType.Int).Value = nTrabajaConInformales
        oComm.Parameters.Add("@nNroInformales", SqlDbType.Int).Value = nNroInformales
        oComm.Parameters.Add("@nMontoInformales", SqlDbType.Money).Value = nMontoInformales
        '--
        oComm.Parameters.Add("@nCosOpeTransp", SqlDbType.Money).Value = nCosOpeTransp
        oComm.Parameters.Add("@nGasFamOtros", SqlDbType.Money).Value = nGasFamOtros
        '
        oComm.Parameters.Add("@nUtilidadBrutaTotal", SqlDbType.Money).Value = nUtilidadBrutaTotal
        oComm.Parameters.Add("@nGastosNegocio", SqlDbType.Money).Value = nGastosNegocio
        oComm.Parameters.Add("@nGastosTotalFam", SqlDbType.Money).Value = nGastosTotalFam
        oComm.Parameters.Add("@nDeudasFinancieras", SqlDbType.Money).Value = nDeudasFinancieras
        oComm.Parameters.Add("@nUtilidadNetaTotal", SqlDbType.Money).Value = nUtilidadNetaTotal
        oComm.Parameters.Add("@nEndeudamientoMax", SqlDbType.Money).Value = nEndeudamientoMax
        oComm.Parameters.Add("@nAsumeCuotaDiaria", SqlDbType.Money).Value = nAsumeCuotaDiaria
        oComm.Parameters.Add("@nAsumeSemanal", SqlDbType.Money).Value = nAsumeSemanal
        oComm.Parameters.Add("@nAsumeCatorcenal", SqlDbType.Money).Value = nAsumeCatorcenal
        oComm.Parameters.Add("@nAsumeMensual", SqlDbType.Money).Value = nAsumeMensual

        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nCodEvalMax", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()

        ReDim Preserve D(2)

        D(0) = oComm.Parameters("@nResp").Value
        D(1) = oComm.Parameters("@Resultado").Value
        nCodEvalMax = oComm.Parameters("@nCodEvalMax").Value

        oConex.Close()

        'Registra Fotos
        Dim oImagen As dFotosNegocio

        If D(0) = "1" Then
            For Each oImagen In oImagenes
                D = GuardarFotoNegocio(nCodEvalMax, oImagen.nCodFoto, oImagen.nImg)
                If D(0) = "0" Then
                    Return D
                End If
            Next
            'mdipas diciembre
            For Each oRubros In oNegociosRubros
                D = RegistrarRubroNegocio(nCodEvalMax, oRubros)
                If D(0) = "0" Then
                    Return D
                End If
            Next
            For Each oEntidad In oDeudasEntidades
                D = RegistrarEntidadDeudasNegocio(nCodEvalMax, oEntidad)
                If D(0) = "0" Then
                    Return D
                End If
            Next
            'end
        End If

        Return D



    End Function

    Public Function RegistrarEntidadDeudasNegocio(nCodEvalMax As Long, entidad As dEntidad) As String()
        'PA_RegistraImagen
        Dim oConex As SqlConnection
        Dim C As SqlCommand
        Dim oComm As SqlCommand
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


        oComm = New SqlCommand("[dbo].[PA_RegistraEntidadDeudaNegocio]", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodEval", SqlDbType.Int).Value = nCodEvalMax

        oComm.Parameters.Add("@cNomEntidad", SqlDbType.VarChar, 200).Value = entidad.nombre
        oComm.Parameters.Add("@nSaldoDeuda", SqlDbType.Money).Value = entidad.deuda

        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()

        ReDim Preserve D(2)

        D(0) = oComm.Parameters("@nResp").Value
        D(1) = oComm.Parameters("@Resultado").Value

        oConex.Close()

        Return D

    End Function


    Public Function RegistrarRubroNegocio(nCodEvalMax As Long, rubro As Rubro) As String()
        'PA_RegistraImagen
        Dim oConex As SqlConnection
        Dim C As SqlCommand
        Dim oComm As SqlCommand
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


        oComm = New SqlCommand("[dbo].[PA_RegistraRubroNegocio]", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@pnCodEval", SqlDbType.Int).Value = nCodEvalMax

        oComm.Parameters.Add("@cRubro", SqlDbType.VarChar, 200).Value = rubro.nombre
        oComm.Parameters.Add("@nVentasDiariasDN", SqlDbType.Money).Value = rubro.ventaDiariaDN
        oComm.Parameters.Add("@nVentasDiariasDB", SqlDbType.Money).Value = rubro.ventaDiariaDB
        oComm.Parameters.Add("@nVentasSemanalesDN", SqlDbType.Money).Value = rubro.ventaSemanalDN
        oComm.Parameters.Add("@nVentasSemanalesDB", SqlDbType.Money).Value = rubro.ventaSemanalDB
        oComm.Parameters.Add("@nVentasMensualesDN", SqlDbType.Money).Value = rubro.ventaMensualDN
        oComm.Parameters.Add("@nVentasMensualesDB", SqlDbType.Money).Value = rubro.ventaMensualDB
        oComm.Parameters.Add("@nPorcCostoDB", SqlDbType.Money).Value = rubro.CostoPorcDB
        oComm.Parameters.Add("@nPorcCostoDN", SqlDbType.Money).Value = rubro.CostoPorcDN
        oComm.Parameters.Add("@nCostoDB", SqlDbType.Money).Value = rubro.CostoDB
        oComm.Parameters.Add("@nCostoDN", SqlDbType.Money).Value = rubro.CostoDN
        oComm.Parameters.Add("@nPorcUtilidadDB", SqlDbType.Money).Value = rubro.UtilidadPorcDB
        oComm.Parameters.Add("@nPorcUtilidadDN", SqlDbType.Money).Value = rubro.UtilidadPorcDN
        oComm.Parameters.Add("@nUtilidadDB", SqlDbType.Money).Value = rubro.UtilidadDB
        oComm.Parameters.Add("@nUtilidadDN", SqlDbType.Money).Value = rubro.UtilidadDN
        oComm.Parameters.Add("@nTotalVentasMensual", SqlDbType.Money).Value = rubro.totalVentaMensual
        oComm.Parameters.Add("@nTotalUtilidadMensual", SqlDbType.Money).Value = rubro.totalUtilidadMensual

        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
        oComm.ExecuteNonQuery()

        ReDim Preserve D(2)

        D(0) = oComm.Parameters("@nResp").Value
        D(1) = oComm.Parameters("@Resultado").Value

        oConex.Close()

        Return D

    End Function

    Public Function RegistraPersonaMovilXProd() As String()

        Try

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


            oComm = New SqlClient.SqlCommand("PA_Movil2_RegistraPersonaXProd_4", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = cDoc
            oComm.Parameters.Add("@cRazonSocial", SqlDbType.VarChar, 50).Value = ""
            oComm.Parameters.Add("@cNombres", SqlDbType.VarChar, 50).Value = cNombres
            oComm.Parameters.Add("@cApellidos", SqlDbType.VarChar, 50).Value = cApellidos
            oComm.Parameters.Add("@dFechaNac", SqlDbType.DateTime).Value = CDate(dFechaNac)
            oComm.Parameters.Add("@cGenero", SqlDbType.VarChar, 50).Value = cGenero
            oComm.Parameters.Add("@cDireccionDom", SqlDbType.VarChar, 150).Value = cDireccionDom
            oComm.Parameters.Add("@cReferenciaDom", SqlDbType.VarChar, 150).Value = cReferenciaDom
            oComm.Parameters.Add("@cCodZonaDom", SqlDbType.VarChar, 50).Value = cCodZonaDom
            oComm.Parameters.Add("@cMovil1", SqlDbType.VarChar, 50).Value = cMovil1
            oComm.Parameters.Add("@cDocConyuge", SqlDbType.VarChar, 50).Value = cDocConyuge
            oComm.Parameters.Add("@cTelFijo", SqlDbType.VarChar, 50).Value = cTelFijo
            oComm.Parameters.Add("@nEstadoCiv", SqlDbType.Int).Value = CInt(sEstadoCiv)
            oComm.Parameters.Add("@nNumDep", SqlDbType.Int).Value = CInt(sNumDep)
            oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = cUser
            oComm.Parameters.Add("@psNomBanco", SqlDbType.VarChar, 150).Value = psNomBanco
            oComm.Parameters.Add("@psNumCuenta", SqlDbType.VarChar, 150).Value = psNroCuenta
            oComm.Parameters.Add("@sTipoTarifario", SqlDbType.VarChar, 50).Value = "3"
            oComm.Parameters.Add("@psProd", SqlDbType.VarChar, 50).Value = psProd
            oComm.Parameters.Add("@cNivelEducac", SqlDbType.VarChar, 50).Value = cNivelEducac
            oComm.Parameters.Add("@cTipoVivienda", SqlDbType.VarChar, 50).Value = cTipoVivienda
            oComm.Parameters.Add("@cTiempoResidencia", SqlDbType.VarChar, 50).Value = cTiempoResidencia
            oComm.Parameters.Add("@cExperienActiv", SqlDbType.VarChar, 50).Value = cExperienActiv
            oComm.Parameters.Add("@nTipoDoc", SqlDbType.Int).Value = CInt(nTipoDoc)
            oComm.Parameters.Add("@nTipoDocConyuge", SqlDbType.Int).Value = CInt(nTipoDocConyuge)
            oComm.Parameters.Add("@psNumCuentaCCI", SqlDbType.VarChar, 150).Value = psNroCuentaCCI
            oComm.Parameters.Add("@nCodBanco", SqlDbType.Int).Value = nCodBanco

            oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
            oComm.ExecuteNonQuery()

            ReDim Preserve D(2)

            D(0) = oComm.Parameters("@nResp").Value
            D(1) = oComm.Parameters("@Resultado").Value

            oConex.Close()

            Return D

        Catch ex As Exception
            Dim D(2) As String
            D(0) = "0"
            D(1) = ex.Message
            Return D

        End Try


    End Function

    Public Function Movil2_RegistraSolicitudAmpliacionPorProducto(ByVal psDoc As String, ByVal psPrestamo As String, ByVal psCuotas As String,
                                                ByVal psMontoCuota As String, ByVal psCodUsu As String, ByVal psTasa As String) As String()

        Dim dr As SqlClient.SqlDataReader

        Dim oConex As SqlConnection
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim D() As String
        Dim i As Integer
        Dim nCodAge As Integer
        Dim nCodCred As Integer
        Dim psResp As String
        Dim nResp As Integer
        Dim nProd As Integer
        Dim sCadConexion As String
        Dim oConect As New clsConecta

        sCadConexion = oConect.gsCadConex


        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        oComm = New SqlClient.SqlCommand("PA_RegistraCredito_Retanqueo_Movil_PorProducto_Solicitud_6", oConex)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@nIdSolicitud", SqlDbType.Int).Value = nIDSolicitud
        oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = psDoc
        oComm.Parameters.Add("@nPrestamo", SqlDbType.Money).Value = CDbl(psPrestamo)
        oComm.Parameters.Add("@nNumCuotas", SqlDbType.Int).Value = CInt(psCuotas)
        oComm.Parameters.Add("@cCodUsuAna", SqlDbType.VarChar, 50).Value = psCodUsu
        oComm.Parameters.Add("@nMontoCuota", SqlDbType.Money).Value = CDbl(psMontoCuota)
        oComm.Parameters.Add("@nTipoPlazo", SqlDbType.Int).Value = 1
        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = psCodUsu
        oComm.Parameters.Add("@TasaComp", SqlDbType.Money).Value = CDbl(psTasa)

        oComm.Parameters.Add("@nCodAgeResp", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nCodCredResp", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nProdResp", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@sRespuesta", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@bDesemnBanco", SqlDbType.Int).Value = bDesemnBanco

        If Not IsNumeric(psProd) Then
            psProd = "3"
        End If

        oComm.Parameters.Add("@pnProd", SqlDbType.Int).Value = CInt(psProd)

        oComm.ExecuteNonQuery()

        nCodAge = oComm.Parameters("@nCodAgeResp").Value
        nCodCred = oComm.Parameters("@nCodCredResp").Value
        nProd = oComm.Parameters("@nProdResp").Value

        oConex.Close()


        ReDim D(2)

        D(0) = Str(oComm.Parameters("@nResp").Value).Trim
        D(1) = oComm.Parameters("@sRespuesta").Value


        oConex.Close()

        Return D


    End Function

    Public Function Movil2_RegistraCreditoAmpliacion_PorExcepcion(ByVal psDoc As String, ByVal psPrestamo As String, ByVal psCuotas As String,
                                                ByVal psMontoCuota As String, ByVal psCodUsu As String, ByVal psTasa As String) As String()

        Dim dr As SqlClient.SqlDataReader

        Dim oConex As SqlConnection
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim D() As String
        Dim i As Integer
        Dim nCodAge As Integer
        Dim nCodCred As Integer
        Dim psResp As String
        Dim nResp As Integer
        Dim nProd As Integer
        Dim sCadConexion As String
        Dim oConect As New clsConecta

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        oComm = New SqlClient.SqlCommand("PA_RegistraCreditoAmpliacion_PorExcepcion_3", oConex)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@nIdSolicitud", SqlDbType.Int).Value = nIDSolicitud
        oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = psDoc
        oComm.Parameters.Add("@nPrestamo", SqlDbType.Money).Value = CDbl(psPrestamo)
        oComm.Parameters.Add("@nNumCuotas", SqlDbType.Int).Value = CInt(psCuotas)
        oComm.Parameters.Add("@cCodUsuAna", SqlDbType.VarChar, 50).Value = psCodUsu
        oComm.Parameters.Add("@nMontoCuota", SqlDbType.Money).Value = CDbl(psMontoCuota)
        oComm.Parameters.Add("@nTipoPlazo", SqlDbType.Int).Value = 1
        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = psCodUsu
        oComm.Parameters.Add("@TasaComp", SqlDbType.Money).Value = CDbl(psTasa)

        oComm.Parameters.Add("@nCodAgeResp", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nCodCredResp", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nProdResp", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@sRespuesta", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@bDesemnBanco", SqlDbType.Int).Value = bDesemnBanco

        If Not IsNumeric(psProd) Then
            psProd = "3"
        End If

        oComm.Parameters.Add("@pnProd", SqlDbType.Int).Value = CInt(psProd)
        oComm.Parameters.Add("@nRespNOAmpliacion", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cComentarioExcepcion", SqlDbType.VarChar, 1000).Value = cComentarioExcepcion

        oComm.ExecuteNonQuery()

        If oComm.Parameters("@nRespNOAmpliacion").Value = 0 Then

            nCodAge = oComm.Parameters("@nCodAgeResp").Value
            nCodCred = oComm.Parameters("@nCodCredResp").Value
            nProd = oComm.Parameters("@nProdResp").Value

            oConex.Close()


            ReDim D(2)

            D(0) = Str(oComm.Parameters("@nResp").Value).Trim
            D(1) = oComm.Parameters("@sRespuesta").Value

            oConex.Close()

            Return D


        End If



        oComm = New SqlClient.SqlCommand("PA_RegistraCreditoNuevo_PorAmpliacion", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = cDoc
        oComm.Parameters.Add("@nPrestamo", SqlDbType.Money).Value = CDbl(nPrestamo)
        oComm.Parameters.Add("@nNumCuotas", SqlDbType.Int).Value = CInt(nNumCuotas)
        oComm.Parameters.Add("@cCodUsuAna", SqlDbType.VarChar, 50).Value = cCodUsuAna
        oComm.Parameters.Add("@nMontoCuota", SqlDbType.Money).Value = CDbl(nMontoCuota)
        oComm.Parameters.Add("@nTipoPlazo", SqlDbType.VarChar, 50).Value = pnTipoPlazo
        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = "ADMIN"
        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@psProd", SqlDbType.VarChar, 50).Value = psProd
        oComm.Parameters.Add("@pdFecgaPago", SqlDbType.Date).Value = CDate(psFechaPago)
        oComm.Parameters.Add("@nIdSolicitud", SqlDbType.Int).Value = nIDSolicitud
        oComm.Parameters.Add("@nCampCuotasSinInt", SqlDbType.Int).Value = nCampCuotasSinInt
        oComm.Parameters.Add("@cComentarioExcepcion", SqlDbType.VarChar, 1000).Value = cComentarioExcepcion
        oComm.Parameters.Add("@bDesemnBanco", SqlDbType.Int).Value = bDesemnBanco

        oComm.ExecuteNonQuery()

        ReDim Preserve D(2)

        D(0) = oComm.Parameters("@nResp").Value
        D(1) = oComm.Parameters("@Resultado").Value

        oConex.Close()

        Return D


    End Function



    Public Function RegistraSolicitudCreditoXProducto() As String()

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

        oComm = New SqlClient.SqlCommand("PA_RegistraCreditoPorProducto_11", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = cDoc
        oComm.Parameters.Add("@nIDGrupo", SqlDbType.Int).Value = nIDGrupo
        oComm.Parameters.Add("@nCodCagrgoGrupo", SqlDbType.Int).Value = nCodCagrgoGrupo

        oComm.Parameters.Add("@nPrestamo", SqlDbType.Money).Value = CDbl(nPrestamo)
        oComm.Parameters.Add("@nNumCuotas", SqlDbType.Int).Value = CInt(nNumCuotas)
        oComm.Parameters.Add("@cCodUsuAna", SqlDbType.VarChar, 50).Value = cCodUsuAna
        oComm.Parameters.Add("@nMontoCuota", SqlDbType.Money).Value = CDbl(nMontoCuota)
        oComm.Parameters.Add("@nTipoPlazo", SqlDbType.VarChar, 50).Value = pnTipoPlazo
        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = "ADMIN"
        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@psProd", SqlDbType.VarChar, 50).Value = psProd
        oComm.Parameters.Add("@pdFecgaPago", SqlDbType.Date).Value = CDate(psFechaPago)
        oComm.Parameters.Add("@nIdSolicitud", SqlDbType.Int).Value = nIDSolicitud
        oComm.Parameters.Add("@nCampCuotasSinInt", SqlDbType.Int).Value = nCampCuotasSinInt
        oComm.Parameters.Add("@nCodDestinoCredito", SqlDbType.Int).Value = nCodDestinoCredito 'mdipas 2024
        oComm.Parameters.Add("@nCodTasacion", SqlDbType.Int).Value = oTasacion(0).nCodFamilia 'mdipas 2024
        oComm.Parameters.Add("@bDesemnBanco", SqlDbType.Int).Value = bDesemnBanco

        'bDesemnBanco

        oComm.ExecuteNonQuery()

        ReDim Preserve D(2)

        D(0) = oComm.Parameters("@nResp").Value
        D(1) = oComm.Parameters("@Resultado").Value

        oConex.Close()

        Return D


    End Function


    Public Function RegistraSolicitudCreditoXProductoEmpenio(ByRef pnCodCRed As Integer, ByRef pnCodAge As Integer) As String()

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


        oComm = New SqlClient.SqlCommand("PA_RegistraCreditoPorProductoEmpenio", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@cDoc", SqlDbType.VarChar, 50).Value = cDoc
        oComm.Parameters.Add("@nPrestamo", SqlDbType.Money).Value = CDbl(nPrestamo)
        oComm.Parameters.Add("@nNumCuotas", SqlDbType.Int).Value = CInt(nNumCuotas)
        oComm.Parameters.Add("@cCodUsuAna", SqlDbType.VarChar, 50).Value = cCodUsuAna
        oComm.Parameters.Add("@nMontoCuota", SqlDbType.Money).Value = CDbl(nMontoCuota)
        oComm.Parameters.Add("@nTipoPlazo", SqlDbType.VarChar, 50).Value = pnTipoPlazo
        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = "ADMIN"
        oComm.Parameters.Add("@Resultado", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nCodCredOut", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nCodAgeOut", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@psProd", SqlDbType.VarChar, 50).Value = psProd
        oComm.Parameters.Add("@pdFecgaPago", SqlDbType.Date).Value = CDate(psFechaPago)

        oComm.ExecuteNonQuery()

        ReDim Preserve D(2)

        D(0) = oComm.Parameters("@nResp").Value
        D(1) = oComm.Parameters("@Resultado").Value
        If D(0).Trim = "1" Then
            pnCodAge = oComm.Parameters("@nCodAgeOut").Value
            pnCodCRed = oComm.Parameters("@nCodCredOut").Value
        End If

        oConex.Close()

        Return D


    End Function


    Public Function RegistraSolcitudEmpenio() As String()

        Dim D(2) As String
        Dim nCodAge As Integer
        Dim nCodcred As Integer

        If ValidaIDTransaccion() Then
            D(0) = "1"
            D(1) = "OK"
        End If

        If nActualizado = 1 Then
            D = RegistraPersonaMovilXProd()
            If (D(0) = "0") Then
                Return D
            End If
        End If

        D = RegistraSolicitudCreditoXProductoEmpenio(nCodcred, nCodAge)

        If (D(0) = "0") Then
            Return D
        End If
        D = RegistrarTasacion(nCodAge, nCodcred)

        Return D

    End Function


    Public Function RegistraSolcitudConsumo() As String()


        Dim D() As String

        If ValidaIDTransaccion() Then
            D(0) = "1"
            D(1) = "OK"
        End If

        If nActualizado = 1 Then
            D = RegistraPersonaMovilXProd()
            If (D(0) = "0") Then
                Return D
            End If
        End If

        If nActualizadoNeg = 1 Then
            D = RegistraEvaluacionConsumo()
            If (D(0) = "0") Then
                Return D
            End If
        End If

        D = RegistraSolicitudCreditoXProducto()

        Return D

    End Function

    Public Function ValidaIDTransaccion() As Boolean

        Dim D = From S In db.SolicitudesProcesadas
                Where S.cDoc = cDoc And S.nId = nIDSolicitud
                Select S

        If D.Count() > 0 Then
            Return True
        End If

        Return False
    End Function

    Public Function RegistraSolcitudPyme() As String()

        Dim D(2) As String

        If nIDSolicitud <> -1 Then
            If ValidaIDTransaccion() Then
                D(0) = "1"
                D(1) = "OK"
            End If
        End If

        If nActualizado = 1 Then
            D = RegistraPersonaMovilXProd()
            If (D(0) = "0") Then
                Return D
            End If
        End If

        If nActualizadoNeg = 1 Then
            D = RegistraNegocio()
            If (D(0) = "0") Then
                Return D
            End If
        End If

        If nExcepcion = 1 Then
            D = Movil2_RegistraCreditoAmpliacion_PorExcepcion(cDoc, nPrestamo, nNumCuotas, nMontoCuota, cUserSolicitud, "0")
        Else
            If nAmpliado = 1 Then
                D = Movil2_RegistraSolicitudAmpliacionPorProducto(cDoc, nPrestamo, nNumCuotas, nMontoCuota, cUserSolicitud, "0")

            Else
                D = RegistraSolicitudCreditoXProducto()
            End If

        End If




        Return D

    End Function

    Public Function RegistraSolcitudPymeExcepcion() As String()

        Dim D(2) As String

        If nIDSolicitud <> -1 Then
            If ValidaIDTransaccion() Then
                D(0) = "1"
                D(1) = "OK"
            End If
        End If

        If nActualizado = 1 Then
            D = RegistraPersonaMovilXProd()
            If (D(0) = "0") Then
                Return D
            End If
        End If

        If nActualizadoNeg = 1 Then
            D = RegistraNegocio()
            If (D(0) = "0") Then
                Return D
            End If
        End If


        D = Movil2_RegistraCreditoAmpliacion_PorExcepcion(cDoc, nPrestamo, nNumCuotas, nMontoCuota, cUserSolicitud, "0")




        Return D

    End Function

End Class
